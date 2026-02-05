from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


@dataclass
class SupplierInfo:
    name: str
    customer_code: str = ""
    coord1: str = ""
    coord2: str = ""


def _supplier_lookup(suppliers: List[Dict[str, str]]) -> Dict[str, SupplierInfo]:
    out: Dict[str, SupplierInfo] = {}
    for s in suppliers or []:
        name = str(s.get("name", "") or "").strip()
        if not name:
            continue
        out[name] = SupplierInfo(
            name=name,
            customer_code=str(s.get("customer_code", "") or ""),
            coord1=str(s.get("coord1", "") or ""),
            coord2=str(s.get("coord2", "") or ""),
        )
    return out


def group_lines_for_order(df: pd.DataFrame) -> pd.DataFrame:
    """Bon fournisseur (lisible) :
    - Regroupe par Produit + Unité (évite de mélanger kg / pièces / L)
    - Somme des Quantités
    - Sortie : Produit | Quantité | Unité
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=["Produit", "Quantité", "Unité"])

    d = df.copy()

    # Produit : priorité au Libellé si dispo
    if "Libellé" in d.columns:
        d["Produit"] = d["Libellé"].fillna(d.get("Produit", ""))
    elif "Produit" not in d.columns:
        d["Produit"] = ""

    # Unité : tolère plusieurs noms de colonnes possibles
    unit_col = None
    for cand in ["Unité", "Unite", "Unité", "U", "Unites"]:
        if cand in d.columns:
            unit_col = cand
            break
    if unit_col is None:
        d["Unité"] = ""
    else:
        d["Unité"] = d[unit_col]

    d["Produit"] = d["Produit"].fillna("").astype(str).str.strip()
    d["Unité"] = d["Unité"].fillna("").astype(str).str.strip()

    d["Quantité"] = pd.to_numeric(d.get("Quantité", 0), errors="coerce").fillna(0)

    grouped = (
        d.groupby(["Produit", "Unité"], as_index=False)["Quantité"]
        .sum()
        .sort_values(["Produit", "Unité"])
        .reset_index(drop=True)
    )
    return grouped


# ----------------- Livraison (découpage par poids) -----------------

def _to_float(x) -> float:
    try:
        if x is None:
            return 0.0
        if isinstance(x, str):
            x = x.strip().replace(",", ".")
        return float(x)
    except Exception:
        return 0.0


def _infer_weight_kg(unit: str, qty: float) -> float:
    """Infère un poids (kg) à partir de l'unité quand aucun poids n'est fourni.

    Règles simples:
    - unité ~ kg => poids = quantité
    - unité ~ g  => poids = quantité/1000
    Sinon: 0
    """
    u = str(unit or "").strip().lower()
    if not u:
        return 0.0
    kg_alias = {"kg", "kilo", "kilos", "kilogramme", "kilogrammes"}
    g_alias = {"g", "gr", "gramme", "grammes"}
    if u in kg_alias:
        return max(qty, 0.0)
    if u in g_alias:
        return max(qty, 0.0) / 1000.0
    return 0.0


def split_bon_by_deliveries(
    bon_df: pd.DataFrame,
    delivery_labels: List[str],
    *,
    max_weight_kg: float = 50.0,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Assigne chaque ligne du bon à une livraison, sans dépasser max_weight_kg.

    - delivery_labels: liste fournie par l'utilisateur (ex: dates) dans l'ordre souhaité.
    - On fait un "first-fit decreasing" sur les lignes (poids décroissant).
    - Si pas assez de labels, on crée automatiquement "Livraison suppl. #".

    Retour:
      - df_assigné (avec colonnes Livraison, Poids_ligne_kg, Prix_cible_ligne)
      - résumé par livraison (Poids_total_kg, Prix_cible_total)
    """
    df = bon_df.copy() if bon_df is not None else pd.DataFrame()
    if df.empty:
        df["Livraison"] = ""
        return df, pd.DataFrame(columns=["Livraison", "Poids_total_kg", "Prix_cible_total"])

    # Colonnes tolérées
    unit_col = None
    for cand in ["Unité", "Unite", "Unité", "U", "Unites"]:
        if cand in df.columns:
            unit_col = cand
            break

    qty = pd.to_numeric(df.get("Quantité", 0), errors="coerce").fillna(0).astype(float)

    # Poids (priorité: Poids total (kg), sinon Quantité*Poids unitaire, sinon inférence unité)
    poids_total = pd.to_numeric(df.get("Poids total (kg)", None), errors="coerce")
    if poids_total is None:
        poids_total = pd.Series([float("nan")]*len(df), index=df.index)
    poids_unit = pd.to_numeric(df.get("Poids unitaire (kg)", None), errors="coerce")
    if poids_unit is None:
        poids_unit = pd.Series([float("nan")]*len(df), index=df.index)

    inferred = []
    if unit_col:
        for u, q in zip(df[unit_col], qty):
            inferred.append(_infer_weight_kg(u, float(q)))
    else:
        inferred = [0.0]*len(df)

    poids_line = poids_total.copy()
    # si poids_total vide -> qty*poids_unit
    mask_nan = poids_line.isna()
    poids_line.loc[mask_nan] = (qty * poids_unit).where(~poids_unit.isna(), float("nan"))
    # si encore nan -> inférence
    mask_nan2 = poids_line.isna()
    poids_line.loc[mask_nan2] = inferred
    poids_line = poids_line.fillna(0).astype(float)
    poids_line = poids_line.clip(lower=0)

    # Prix cible
    prix_u = pd.to_numeric(df.get("Prix cible unitaire", 0), errors="coerce").fillna(0).astype(float)
    prix_line = (qty * prix_u).fillna(0).astype(float)

    # Prépare packing
    max_w = float(max_weight_kg or 50.0)
    labels = [str(x).strip() for x in (delivery_labels or []) if str(x).strip()]

    bins = []  # list of dict: label, remaining
    def _ensure_bin(i: int):
        while len(bins) <= i:
            if len(labels) > len(bins):
                lab = labels[len(bins)]
            else:
                lab = f"Livraison suppl. {len(bins)+1}"
            bins.append({"label": lab, "remaining": max_w})

    # lignes triées par poids desc
    order = sorted(list(df.index), key=lambda ix: float(poids_line.loc[ix]), reverse=True)
    assign = {}
    for ix in order:
        w = float(poids_line.loc[ix])
        placed = False
        for b in bins:
            if w <= b["remaining"] + 1e-9:
                b["remaining"] -= w
                assign[ix] = b["label"]
                placed = True
                break
        if not placed:
            _ensure_bin(len(bins))
            bins[-1]["remaining"] -= w
            assign[ix] = bins[-1]["label"]

    df["Livraison"] = pd.Series(assign)
    df["Poids_ligne_kg"] = poids_line
    df["Prix_cible_ligne"] = prix_line

    summary = (
        df.groupby("Livraison", dropna=False)
        .agg(Poids_total_kg=("Poids_ligne_kg", "sum"), Prix_cible_total=("Prix_cible_ligne", "sum"))
        .reset_index()
    )

    # ordre des livraisons: labels puis suppl.
    order_labels = {lab: i for i, lab in enumerate(labels)}
    def _sort_key(lab: str):
        lab = str(lab)
        if lab in order_labels:
            return (0, order_labels[lab])
        return (1, lab)
    summary = summary.sort_values(by="Livraison", key=lambda s: s.map(_sort_key)).reset_index(drop=True)

    return df, summary

def export_orders_per_supplier_excel(
    bon_df: pd.DataFrame,
    out_path: str,
    *,
    suppliers: Optional[List[Dict[str, str]]] = None,
) -> None:
    """Crée un classeur Excel avec 1 feuille par fournisseur (version lisible & épurée).

    Colonnes : Produit | Quantité | Unité
    - Filtre les lignes Quantité <= 0
    - En-tête propre + coordonnées
    - Tableau zébré, quantités alignées à droite
    """
    suppliers = suppliers or []
    sup_map = _supplier_lookup(suppliers)

    df = bon_df.copy() if bon_df is not None else pd.DataFrame()
    if df.empty:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Aucun fournisseur"
        ws["A1"].value = "Aucune ligne dans le bon de commande"
        wb.save(out_path)
        return

    if "Fournisseur" not in df.columns:
        df["Fournisseur"] = ""
    df["Fournisseur"] = df["Fournisseur"].fillna("").astype(str).str.strip()

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    title_fill = PatternFill("solid", fgColor="2F5597")  # bleu soutenu
    title_font = Font(bold=True, size=14, color="FFFFFF")
    subtitle_font = Font(bold=False, size=10, color="333333")

    header_fill = PatternFill("solid", fgColor="F2F2F2")
    header_font = Font(bold=True, color="1F1F1F")
    header_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
    right_align = Alignment(horizontal="right", vertical="top")
    zebra_fill = PatternFill("solid", fgColor="FAFAFA")

    for sup_name, part in df.groupby("Fournisseur", dropna=False):
        sup_name = str(sup_name or "").strip() or "(sans fournisseur)"
        ws = wb.create_sheet(title=sup_name[:31])
        info = sup_map.get(sup_name, SupplierInfo(name=sup_name))

        # --- Bandeau titre
        ws.merge_cells("A1:C1")
        ws["A1"].value = "BON DE COMMANDE"
        ws["A1"].fill = title_fill
        ws["A1"].font = title_font
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 26

        # Fournisseur (gros)
        ws.merge_cells("A2:C2")
        ws["A2"].value = f"{info.name}"
        ws["A2"].font = Font(bold=True, size=12)
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 20

        # Infos
        r = 3
        if info.customer_code:
            ws.merge_cells(f"A{r}:C{r}")
            ws[f"A{r}"].value = f"Code client : {info.customer_code}"
            ws[f"A{r}"].font = subtitle_font
            ws[f"A{r}"].alignment = Alignment(horizontal="center")
            r += 1
        if info.coord1:
            ws.merge_cells(f"A{r}:C{r}")
            ws[f"A{r}"].value = info.coord1
            ws[f"A{r}"].font = subtitle_font
            ws[f"A{r}"].alignment = Alignment(horizontal="center")
            r += 1
        if info.coord2:
            ws.merge_cells(f"A{r}:C{r}")
            ws[f"A{r}"].value = info.coord2
            ws[f"A{r}"].font = subtitle_font
            ws[f"A{r}"].alignment = Alignment(horizontal="center")
            r += 1

        start_row = max(r + 1, 6)

        # données
        lines = group_lines_for_order(part)
        if "Quantité" in lines.columns:
            lines = lines[lines["Quantité"].astype(float) > 0].reset_index(drop=True)

        # En-tête tableau
        headers = ["Produit", "Quantité", "Unité"]
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=start_row, column=c)
            cell.value = h
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = border
        ws.row_dimensions[start_row].height = 18

        # Lignes
        for i, row in enumerate(lines.itertuples(index=False), start=1):
            rr = start_row + i
            prod = getattr(row, "Produit", "")
            qty = getattr(row, "Quantité", 0)
            unit = getattr(row, "Unité", "")

            c1 = ws.cell(row=rr, column=1, value=prod)
            c1.alignment = left_align
            c1.border = border

            c2 = ws.cell(row=rr, column=2, value=float(qty) if qty != "" else "")
            c2.number_format = "#,##0.##"
            c2.alignment = right_align
            c2.border = border

            c3 = ws.cell(row=rr, column=3, value=unit)
            c3.alignment = left_align
            c3.border = border

            if i % 2 == 0:
                c1.fill = zebra_fill
                c2.fill = zebra_fill
                c3.fill = zebra_fill

        # Mise en page
        ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
        ws.column_dimensions["A"].width = 52
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 12
        ws.print_title_rows = f"{start_row}:{start_row}"
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_view.showGridLines = False

    wb.save(out_path)



def export_orders_per_supplier_excel_by_delivery(
    bon_df: pd.DataFrame,
    out_path: str,
    *,
    suppliers: Optional[List[Dict[str, str]]] = None,
    delivery_label_col: str = "Livraison",
    weight_col: str = "Poids_ligne_kg",
    target_col: str = "Prix_cible_ligne",
) -> None:
    """Excel: 1 feuille par (Livraison, Fournisseur).

    - Conserve le même style que la version 'par fournisseur'
    - Ajoute 2 lignes d'info: Livraison + Total poids / prix cible (si dispo)
    """
    suppliers = suppliers or []
    sup_map = _supplier_lookup(suppliers)

    df = bon_df.copy() if bon_df is not None else pd.DataFrame()
    if df.empty:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Aucune livraison"
        ws["A1"].value = "Aucune ligne dans le bon de commande"
        wb.save(out_path)
        return

    if delivery_label_col not in df.columns:
        df[delivery_label_col] = ""
    if "Fournisseur" not in df.columns:
        df["Fournisseur"] = ""

    df[delivery_label_col] = df[delivery_label_col].fillna("").astype(str).str.strip()
    df["Fournisseur"] = df["Fournisseur"].fillna("").astype(str).str.strip()

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    title_fill = PatternFill("solid", fgColor="2F5597")
    title_font = Font(bold=True, size=14, color="FFFFFF")
    subtitle_font = Font(bold=False, size=10, color="333333")

    header_fill = PatternFill("solid", fgColor="F2F2F2")
    header_font = Font(bold=True, color="1F1F1F")
    header_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
    right_align = Alignment(horizontal="right", vertical="top")
    zebra_fill = PatternFill("solid", fgColor="FAFAFA")

    # ordre stable
    deliveries = list(dict.fromkeys(df[delivery_label_col].tolist()))

    for liv in deliveries:
        df_l = df[df[delivery_label_col] == liv]
        for sup_name, part in df_l.groupby("Fournisseur", dropna=False):
            sup_name = str(sup_name or "").strip() or "(sans fournisseur)"
            info = sup_map.get(sup_name, SupplierInfo(name=sup_name))

            # Nom feuille
            title = f"{liv} - {sup_name}".strip(" -")
            ws = wb.create_sheet(title=title[:31] or "Livraison")

            # Bandeau
            ws.merge_cells("A1:C1")
            ws["A1"].value = "BON DE COMMANDE"
            ws["A1"].fill = title_fill
            ws["A1"].font = title_font
            ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 26

            ws.merge_cells("A2:C2")
            ws["A2"].value = f"{info.name}"
            ws["A2"].font = Font(bold=True, size=12)
            ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[2].height = 20

            r = 3
            if liv:
                ws.merge_cells(f"A{r}:C{r}")
                ws[f"A{r}"].value = f"Livraison : {liv}"
                ws[f"A{r}"].font = subtitle_font
                ws[f"A{r}"].alignment = Alignment(horizontal="center")
                r += 1

            if info.customer_code:
                ws.merge_cells(f"A{r}:C{r}")
                ws[f"A{r}"].value = f"Code client : {info.customer_code}"
                ws[f"A{r}"].font = subtitle_font
                ws[f"A{r}"].alignment = Alignment(horizontal="center")
                r += 1
            if info.coord1:
                ws.merge_cells(f"A{r}:C{r}")
                ws[f"A{r}"].value = info.coord1
                ws[f"A{r}"].font = subtitle_font
                ws[f"A{r}"].alignment = Alignment(horizontal="center")
                r += 1
            if info.coord2:
                ws.merge_cells(f"A{r}:C{r}")
                ws[f"A{r}"].value = info.coord2
                ws[f"A{r}"].font = subtitle_font
                ws[f"A{r}"].alignment = Alignment(horizontal="center")
                r += 1

            # Totaux livraison (si colonnes présentes)
            w_tot = _to_float(part.get(weight_col, pd.Series([0])).sum()) if weight_col in part.columns else 0.0
            p_tot = _to_float(part.get(target_col, pd.Series([0])).sum()) if target_col in part.columns else 0.0
            if weight_col in part.columns or target_col in part.columns:
                ws.merge_cells(f"A{r}:C{r}")
                ws[f"A{r}"].value = f"Total livraison : {w_tot:.2f} kg — Prix cible : {p_tot:.2f} €"
                ws[f"A{r}"].font = subtitle_font
                ws[f"A{r}"].alignment = Alignment(horizontal="center")
                r += 1

            start_row = max(r + 1, 8)

            lines = group_lines_for_order(part)
            if "Quantité" in lines.columns:
                lines = lines[lines["Quantité"].astype(float) > 0].reset_index(drop=True)

            headers = ["Produit", "Quantité", "Unité"]
            for c, h in enumerate(headers, start=1):
                cell = ws.cell(row=start_row, column=c)
                cell.value = h
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align
                cell.border = border
            ws.row_dimensions[start_row].height = 18

            for i, row in enumerate(lines.itertuples(index=False), start=1):
                rr = start_row + i
                prod = getattr(row, "Produit", "")
                qtyv = getattr(row, "Quantité", 0)
                unit = getattr(row, "Unité", "")

                c1 = ws.cell(row=rr, column=1, value=prod)
                c1.alignment = left_align
                c1.border = border

                c2 = ws.cell(row=rr, column=2, value=float(qtyv) if qtyv != "" else "")
                c2.number_format = "#,##0.##"
                c2.alignment = right_align
                c2.border = border

                c3 = ws.cell(row=rr, column=3, value=unit)
                c3.alignment = left_align
                c3.border = border

                if i % 2 == 0:
                    c1.fill = zebra_fill
                    c2.fill = zebra_fill
                    c3.fill = zebra_fill

            ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
            ws.column_dimensions["A"].width = 52
            ws.column_dimensions["B"].width = 14
            ws.column_dimensions["C"].width = 12
            ws.print_title_rows = f"{start_row}:{start_row}"
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.sheet_view.showGridLines = False

    wb.save(out_path)
def export_orders_per_supplier_pdf(
    bon_df: pd.DataFrame,
    out_pdf_path: str,
    *,
    suppliers: Optional[List[Dict[str, str]]] = None,
) -> None:
    """Génère un PDF avec 1 page (ou plus) par fournisseur – version lisible & épurée.

    Colonnes : Produit | Quantité | Unité
    - Filtre les lignes Quantité <= 0
    - Bandeau titre + fournisseur
    - Tableau zébré, retour à la ligne propre, pagination
    - Watermark (icône) en fond si src/assets/watermark.png est présent
    """
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib.units import mm
    from reportlab.lib.utils import ImageReader
    from pathlib import Path

    suppliers = suppliers or []
    sup_map = _supplier_lookup(suppliers)

    df = bon_df.copy() if bon_df is not None else pd.DataFrame()
    if df.empty:
        c = canvas.Canvas(out_pdf_path, pagesize=A4)
        c.drawString(40, 800, "Aucune ligne dans le bon de commande")
        c.save()
        return

    if "Fournisseur" not in df.columns:
        df["Fournisseur"] = ""
    df["Fournisseur"] = df["Fournisseur"].fillna("").astype(str).str.strip()

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleBand",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=16,
        textColor=colors.white,
        spaceAfter=0,
        leading=18,
    )
    h2_style = ParagraphStyle(
        "SupplierName",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=12,
        textColor=colors.HexColor("#1A1A1A"),
        spaceAfter=4,
        leading=14,
    )
    meta_style = ParagraphStyle(
        "Meta",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        textColor=colors.HexColor("#444444"),
        leading=11,
    )
    cell_style = ParagraphStyle(
        "Cell",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        textColor=colors.black,
        leading=11,
    )

        # Watermark désactivé : aucun logo sur les bons de commande fournisseurs
    watermark_reader = None

    def _on_page(canvas_obj, doc_obj):
        canvas_obj.saveState()

        # Watermark (fond)
        if watermark_reader is not None:
            try:
                # transparence si dispo
                if hasattr(canvas_obj, "setFillAlpha"):
                    canvas_obj.setFillAlpha(0.12)
                w, h = A4
                img_w, img_h = 360, 360
                x = (w - img_w) / 2
                y = (h - img_h) / 2 - 10
                # drawImage "simple" (compatibilité maximale)
                canvas_obj.drawImage(watermark_reader, x, y, width=img_w, height=img_h, mask="auto")
                if hasattr(canvas_obj, "setFillAlpha"):
                    canvas_obj.setFillAlpha(1)
            except Exception:
                pass

        # pagination
        w, _h = A4
        canvas_obj.setFont("Helvetica", 8)
        canvas_obj.setFillColor(colors.HexColor("#666666"))
        canvas_obj.drawRightString(w - 18 * mm, 12 * mm, f"Page {doc_obj.page}")
        canvas_obj.restoreState()

    doc = SimpleDocTemplate(
        out_pdf_path,
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        title="Bons de commande fournisseurs",
    )

    story = []
    first = True

    for sup_name, part in df.groupby("Fournisseur", dropna=False):
        sup_name = str(sup_name or "").strip() or "(sans fournisseur)"
        info = sup_map.get(sup_name, SupplierInfo(name=sup_name))

        lines = group_lines_for_order(part)
        if "Quantité" in lines.columns:
            lines = lines[lines["Quantité"].astype(float) > 0].reset_index(drop=True)

        if not first:
            story.append(PageBreak())
        first = False

        band = Table([[Paragraph("BON DE COMMANDE", title_style)]], colWidths=[doc.width])
        band.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#2F5597")),
                    ("LEFTPADDING", (0, 0), (-1, -1), 10),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                    ("TOPPADDING", (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ]
            )
        )
        story.append(band)
        story.append(Spacer(1, 8))

        story.append(Paragraph(info.name, h2_style))

        meta_lines = []
        if info.customer_code:
            meta_lines.append(f"Code client : <b>{info.customer_code}</b>")
        if info.coord1:
            meta_lines.append(info.coord1)
        if info.coord2:
            meta_lines.append(info.coord2)
        if meta_lines:
            story.append(Paragraph("<br/>".join(meta_lines), meta_style))
            story.append(Spacer(1, 10))

        data = [["Produit", "Quantité", "Unité"]]
        for _, r in lines.iterrows():
            prod = str(r.get("Produit", "") or "").strip()
            qty = r.get("Quantité", 0)
            unit = str(r.get("Unité", "") or "")
            qty_txt = ""
            if pd.notna(qty):
                try:
                    qty_txt = f"{float(qty):g}"
                except Exception:
                    qty_txt = str(qty)
            data.append([Paragraph(prod, cell_style), qty_txt, unit])

        col_widths = [doc.width * 0.64, doc.width * 0.18, doc.width * 0.18]
        t = Table(data, colWidths=col_widths, repeatRows=1)

        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1F1F1F")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 9),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("ALIGN", (1, 1), (1, -1), "RIGHT"),
            ("ALIGN", (2, 1), (2, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D0D0D0")),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]
        for i in range(1, len(data)):
            if i % 2 == 0:
                style_cmds.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#FAFAFA")))

        t.setStyle(TableStyle(style_cmds))
        story.append(t)

        story.append(Spacer(1, 8))
        story.append(Paragraph(f"Total références : <b>{max(len(data) - 1, 0)}</b>", meta_style))

    doc.build(story, onFirstPage=_on_page, onLaterPages=_on_page)


def export_orders_per_supplier_pdf_by_delivery(
    bon_df: pd.DataFrame,
    out_pdf_path: str,
    *,
    suppliers: Optional[List[Dict[str, str]]] = None,
    delivery_label_col: str = "Livraison",
    weight_col: str = "Poids_ligne_kg",
    target_col: str = "Prix_cible_ligne",
) -> None:
    """PDF: 1 page (ou plus) par (Livraison, Fournisseur)."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib.units import mm

    suppliers = suppliers or []
    sup_map = _supplier_lookup(suppliers)

    df = bon_df.copy() if bon_df is not None else pd.DataFrame()
    if df.empty:
        c = canvas.Canvas(out_pdf_path, pagesize=A4)
        c.drawString(40, 800, "Aucune ligne dans le bon de commande")
        c.save()
        return

    if delivery_label_col not in df.columns:
        df[delivery_label_col] = ""
    if "Fournisseur" not in df.columns:
        df["Fournisseur"] = ""

    df[delivery_label_col] = df[delivery_label_col].fillna("").astype(str).str.strip()
    df["Fournisseur"] = df["Fournisseur"].fillna("").astype(str).str.strip()

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleBand",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=16,
        textColor=colors.white,
        spaceAfter=0,
        leading=18,
    )
    h2_style = ParagraphStyle(
        "SupplierName",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=12,
        textColor=colors.HexColor("#1A1A1A"),
        spaceAfter=4,
        leading=14,
    )
    meta_style = ParagraphStyle(
        "Meta",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        textColor=colors.HexColor("#333333"),
        leading=11,
    )

    doc = SimpleDocTemplate(
        out_pdf_path,
        pagesize=A4,
        rightMargin=14 * mm,
        leftMargin=14 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title="Bons fournisseurs (livraisons)",
    )

    story = []

    def _add_block(liv: str, sup_name: str, part: pd.DataFrame):
        info = sup_map.get(sup_name, SupplierInfo(name=sup_name))

        # Bandeau
        story.append(Table([[Paragraph("BON DE COMMANDE", title_style)]], colWidths=[doc.width],
                           style=TableStyle([
                               ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#2F5597")),
                               ("LEFTPADDING", (0, 0), (-1, -1), 8),
                               ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                               ("TOPPADDING", (0, 0), (-1, -1), 6),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                           ])))
        story.append(Spacer(1, 6))

        story.append(Paragraph(f"{info.name}", h2_style))
        if liv:
            story.append(Paragraph(f"Livraison : <b>{liv}</b>", meta_style))
        if info.customer_code:
            story.append(Paragraph(f"Code client : <b>{info.customer_code}</b>", meta_style))
        if info.coord1:
            story.append(Paragraph(info.coord1, meta_style))
        if info.coord2:
            story.append(Paragraph(info.coord2, meta_style))

        # Totaux
        w_tot = float(pd.to_numeric(part.get(weight_col, 0), errors="coerce").fillna(0).sum()) if weight_col in part.columns else 0.0
        p_tot = float(pd.to_numeric(part.get(target_col, 0), errors="coerce").fillna(0).sum()) if target_col in part.columns else 0.0
        if weight_col in part.columns or target_col in part.columns:
            story.append(Paragraph(f"Total livraison : <b>{w_tot:.2f} kg</b> — Prix cible : <b>{p_tot:.2f} €</b>", meta_style))

        story.append(Spacer(1, 8))

        lines = group_lines_for_order(part)
        if "Quantité" in lines.columns:
            lines = lines[lines["Quantité"].astype(float) > 0].reset_index(drop=True)

        data = [["Produit", "Quantité", "Unité"]]
        for row in lines.itertuples(index=False):
            data.append([getattr(row, "Produit", ""), getattr(row, "Quantité", 0), getattr(row, "Unité", "")])

        col_widths = [doc.width * 0.64, doc.width * 0.18, doc.width * 0.18]
        t = Table(data, colWidths=col_widths, repeatRows=1)
        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1F1F1F")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 9),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("ALIGN", (1, 1), (1, -1), "RIGHT"),
            ("ALIGN", (2, 1), (2, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D0D0D0")),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]
        for i in range(1, len(data)):
            if i % 2 == 0:
                style_cmds.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#FAFAFA")))
        t.setStyle(TableStyle(style_cmds))
        story.append(t)
        story.append(Spacer(1, 8))

    # ordre stable
    deliveries = list(dict.fromkeys(df[delivery_label_col].tolist()))
    first = True
    for liv in deliveries:
        df_l = df[df[delivery_label_col] == liv]
        for sup_name, part in df_l.groupby("Fournisseur", dropna=False):
            sup_name = str(sup_name or "").strip() or "(sans fournisseur)"
            if not first:
                story.append(PageBreak())
            first = False
            _add_block(liv, sup_name, part)

    doc.build(story)
