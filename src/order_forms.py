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
    - Regroupe par Produit + Unité + Livraison (+ Prix/Poids si présents)
    - Somme des Quantités (+ Prix cible total / Poids total si présents)
    - Sortie : Produit | Quantité | Unité | Livraison | Prix cible unitaire | Prix cible total | Poids total (kg)
    """
    base_cols = ["Produit", "Quantité", "Unité", "Livraison", "Prix cible unitaire", "Prix cible total", "Poids total (kg)"]

    if df is None or df.empty:
        return pd.DataFrame(columns=base_cols)

    d = df.copy()

    # Produit : priorité au Libellé si dispo
    if "Libellé" in d.columns:
        d["Produit"] = d["Libellé"].fillna(d.get("Produit", ""))
    elif "Produit" not in d.columns:
        d["Produit"] = ""

    # Unité
    unit_col = None
    for cand in ["Unité", "Unite", "Unité", "U", "Unites"]:
        if cand in d.columns:
            unit_col = cand
            break
    d["Unité"] = d[unit_col] if unit_col else ""

    # Livraison
    if "Livraison" not in d.columns:
        d["Livraison"] = ""

    # Prix / poids
    for c in ["Prix cible unitaire", "Prix cible total", "Poids unitaire (kg)", "Poids total (kg)"]:
        if c not in d.columns:
            d[c] = ""

    d["Produit"] = d["Produit"].fillna("").astype(str).str.strip()
    d["Unité"] = d["Unité"].fillna("").astype(str).str.strip()
    d["Livraison"] = d["Livraison"].fillna("").astype(str).str.strip()

    d["Quantité"] = pd.to_numeric(d.get("Quantité", 0), errors="coerce").fillna(0)

    # calc poids total si pas fourni
    wt = pd.to_numeric(d.get("Poids total (kg)"), errors="coerce")
    wu = pd.to_numeric(d.get("Poids unitaire (kg)"), errors="coerce").fillna(0)
    wt_calc = d["Quantité"] * wu
    d["Poids total (kg)"] = wt.where(wt.notna() & (wt != 0), wt_calc).fillna(0)

    # calc prix cible total si pas fourni
    pt = pd.to_numeric(d.get("Prix cible total"), errors="coerce")
    pu = pd.to_numeric(d.get("Prix cible unitaire"), errors="coerce").fillna(0)
    pt_calc = d["Quantité"] * pu
    d["Prix cible total"] = pt.where(pt.notna() & (pt != 0), pt_calc).fillna(0)

    # keep unit price/weight unit as key to avoid mixing different prices
    d["Prix cible unitaire"] = pd.to_numeric(d.get("Prix cible unitaire"), errors="coerce").fillna(0)
    d["Poids unitaire (kg)"] = pd.to_numeric(d.get("Poids unitaire (kg)"), errors="coerce").fillna(0)

    grouped = (
        d.groupby(["Produit", "Unité", "Livraison", "Prix cible unitaire", "Poids unitaire (kg)"], as_index=False)[
            ["Quantité", "Prix cible total", "Poids total (kg)"]
        ]
        .sum()
        .sort_values(["Livraison", "Produit", "Unité"])
        .reset_index(drop=True)
    )

    # reorder
    grouped = grouped.rename(columns={"Poids total (kg)": "Poids total (kg)"})
    return grouped[base_cols]


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
        ws.merge_cells("A1:G1")
        ws["A1"].value = "BON DE COMMANDE"
        ws["A1"].fill = title_fill
        ws["A1"].font = title_font
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 26

        # Fournisseur (gros)
        ws.merge_cells("A2:G2")
        ws["A2"].value = f"{info.name}"
        ws["A2"].font = Font(bold=True, size=12)
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 20

        # Infos
        r = 3
        if info.customer_code:
            ws.merge_cells(f"A{r}:G{r}")
            ws[f"A{r}"].value = f"Code client : {info.customer_code}"
            ws[f"A{r}"].font = subtitle_font
            ws[f"A{r}"].alignment = Alignment(horizontal="center")
            r += 1
        if info.coord1:
            ws.merge_cells(f"A{r}:G{r}")
            ws[f"A{r}"].value = info.coord1
            ws[f"A{r}"].font = subtitle_font
            ws[f"A{r}"].alignment = Alignment(horizontal="center")
            r += 1
        if info.coord2:
            ws.merge_cells(f"A{r}:G{r}")
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
        headers = ["Livraison","Produit", "Quantité", "Unité", "Prix cible unitaire", "Prix cible total", "Poids total (kg)"]
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
            liv = getattr(row, "Livraison", "")
            prod = getattr(row, "Produit", "")
            qty = getattr(row, "Quantité", 0)
            unit = getattr(row, "Unité", "")
            pcu = getattr(row, "Prix cible unitaire", 0)
            pct = getattr(row, "Prix cible total", 0)
            wt = getattr(row, "Poids total (kg)", 0)

            
c0 = ws.cell(row=rr, column=1, value=liv)
c0.alignment = left_align
c0.border = border

c1 = ws.cell(row=rr, column=2, value=prod)
c1.alignment = left_align
c1.border = border

c2 = ws.cell(row=rr, column=3, value=float(qty) if qty != "" else "")
c2.number_format = "#,##0.##"
c2.alignment = right_align
c2.border = border

c3 = ws.cell(row=rr, column=4, value=unit)
c3.alignment = left_align
c3.border = border

c4 = ws.cell(row=rr, column=5, value=float(pcu) if pcu != "" else "")
c4.number_format = "#,##0.00"
c4.alignment = right_align
c4.border = border

c5 = ws.cell(row=rr, column=6, value=float(pct) if pct != "" else "")
c5.number_format = "#,##0.00"
c5.alignment = right_align
c5.border = border

c6 = ws.cell(row=rr, column=7, value=float(wt) if wt != "" else "")
c6.number_format = "#,##0.00"
c6.alignment = right_align
c6.border = border

if i % 2 == 0:
    for cc in [c0,c1,c2,c3,c4,c5,c6]:
        cc.fill = zebra_fill

        # Mise en page
        ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 52
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 12
        ws.column_dimensions["E"].width = 16
        ws.column_dimensions["F"].width = 16
        ws.column_dimensions["G"].width = 16

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

        data = [["Livraison","Produit", "Quantité", "Unité", "Prix cible total", "Poids total (kg)"]]
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
            liv = str(r.get("Livraison","") or "").strip()
            pct = r.get("Prix cible total", 0)
            wt = r.get("Poids total (kg)", 0)
            pct_txt = ""
            wt_txt = ""
            try:
                pct_txt = f"{float(pct):.2f}" if pd.notna(pct) else ""
            except Exception:
                pct_txt = str(pct) if pct is not None else ""
            try:
                wt_txt = f"{float(wt):.2f}" if pd.notna(wt) else ""
            except Exception:
                wt_txt = str(wt) if wt is not None else ""
            data.append([liv, Paragraph(prod, cell_style), qty_txt, unit, pct_txt, wt_txt])

        col_widths = [doc.width * 0.16, doc.width * 0.40, doc.width * 0.12, doc.width * 0.12, doc.width * 0.10, doc.width * 0.10]
        t = Table(data, colWidths=col_widths, repeatRows=1)

        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1F1F1F")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 9),
            ("ALIGN", (0, 0), (1, -1), "LEFT"),
            ("ALIGN", (2, 1), (2, -1), "RIGHT"),
            ("ALIGN", (3, 1), (3, -1), "LEFT"),
            ("ALIGN", (4, 1), (5, -1), "RIGHT"),
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