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
    """Version allégée pour bon fournisseur :
<<<<<<< HEAD
    - Regroupe uniquement par Produit (en privilégiant 'Libellé' si présent)
    - Somme des Quantités
    - Sortie : Produit | Quantité
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=["Produit", "Quantité"])

    d = df.copy()

    # Produit : priorité au Libellé si dispo
    if "Libellé" in d.columns:
        d["Produit"] = d["Libellé"].fillna(d.get("Produit", ""))
    elif "Produit" not in d.columns:
        d["Produit"] = ""

    d["Produit"] = d["Produit"].fillna("").astype(str).str.strip()

    # Quantité numérique
    d["Quantité"] = pd.to_numeric(d.get("Quantité", 0), errors="coerce").fillna(0)

    grouped = (
        d.groupby("Produit", as_index=False)["Quantité"]
        .sum()
        .sort_values("Produit")
        .reset_index(drop=True)
    )
    return grouped

=======
    ➜ Regroupe uniquement par Produit
    ➜ Somme des Quantités
    ➜ Sortie = Produit | Quantité
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=["Produit", "Quantité"])
>>>>>>> 12dd7aae5f59b52da015657b2a0f487d3ea9d973

    d = df.copy()
    if "Libellé" in d.columns:
        d["Produit"] = d["Libellé"].fillna(d.get("Produit", ""))
    elif "Produit" not in d.columns:
        d["Produit"] = ""

    d["Quantité"] = pd.to_numeric(d.get("Quantité", 0), errors="coerce").fillna(0)

    grouped = (
        d.groupby("Produit", as_index=False)["Quantité"]
        .sum()
        .sort_values("Produit")
    )

    return grouped




def export_orders_per_supplier_excel(
    bon_df: pd.DataFrame,
    out_path: str,
    *,
    suppliers: Optional[List[Dict[str, str]]] = None,
) -> None:
    """Crée un classeur Excel avec 1 feuille par fournisseur (version lisible & épurée).

    Colonnes : Produit | Quantité
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
        ws.merge_cells("A1:B1")
        ws["A1"].value = "BON DE COMMANDE"
        ws["A1"].fill = title_fill
        ws["A1"].font = title_font
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 26

        # Fournisseur (gros)
        ws.merge_cells("A2:B2")
        ws["A2"].value = f"{info.name}"
        ws["A2"].font = Font(bold=True, size=12)
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 20

        # Infos
        r = 3
        if info.customer_code:
            ws.merge_cells(f"A{r}:B{r}")
            ws[f"A{r}"].value = f"Code client : {info.customer_code}"
            ws[f"A{r}"].font = subtitle_font
            ws[f"A{r}"].alignment = Alignment(horizontal="center")
            r += 1
        if info.coord1:
            ws.merge_cells(f"A{r}:B{r}")
            ws[f"A{r}"].value = info.coord1
            ws[f"A{r}"].font = subtitle_font
            ws[f"A{r}"].alignment = Alignment(horizontal="center")
            r += 1
        if info.coord2:
            ws.merge_cells(f"A{r}:B{r}")
            ws[f"A{r}"].value = info.coord2
            ws[f"A{r}"].font = subtitle_font
            ws[f"A{r}"].alignment = Alignment(horizontal="center")
            r += 1

        start_row = max(r + 1, 6)

        # données (épuration + filtre qté)
        lines = group_lines_for_order(part)
        if "Quantité" in lines.columns:
            lines = lines[lines["Quantité"].astype(float) > 0].reset_index(drop=True)

        # En-tête tableau
        headers = ["Produit", "Quantité"]
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

            c1 = ws.cell(row=rr, column=1, value=prod)
            c1.alignment = left_align
            c1.border = border

            c2 = ws.cell(row=rr, column=2, value=float(qty) if qty != "" else "")
            c2.number_format = "#,##0.##"
            c2.alignment = right_align
            c2.border = border

            if i % 2 == 0:
                c1.fill = zebra_fill
                c2.fill = zebra_fill

        # Mise en page
        ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
        ws.column_dimensions["A"].width = 52
        ws.column_dimensions["B"].width = 14
        ws.print_title_rows = f"{start_row}:{start_row}"
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_view.showGridLines = False

    wb.save(out_path)
        return

    if "Fournisseur" not in df.columns:
        df["Fournisseur"] = ""

    df["Fournisseur"] = df["Fournisseur"].fillna("").astype(str).str.strip()

    wb = openpyxl.Workbook()
    # supprime la feuille par défaut
    wb.remove(wb.active)

    thin = Side(style="thin", color="9E9E9E")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="EDEDED")
    header_font = Font(bold=True)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_align = Alignment(horizontal="left", vertical="top", wrap_text=True)

    for sup_name, part in df.groupby("Fournisseur", dropna=False):
        sup_name = str(sup_name or "").strip() or "(sans fournisseur)"
        ws = wb.create_sheet(title=sup_name[:31])

        info = sup_map.get(sup_name, SupplierInfo(name=sup_name))
        # en-tête
        ws.merge_cells("A1:H1")
        ws["A1"].value = f"BON DE COMMANDE – {info.name}"
        ws["A1"].font = Font(bold=True, size=14)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 24

        ws["A2"].value = f"Code client : {info.customer_code}" if info.customer_code else ""
        ws["A3"].value = info.coord1
        ws["A4"].value = info.coord2

        # données (fusion via Libellé)
        lines = group_lines_for_order(part)
        start_row = 6
        headers = list(lines.columns)
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=start_row, column=c)
            cell.value = h
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = border

        for r_idx, row in enumerate(lines.itertuples(index=False), start=start_row + 1):
            for c_idx, val in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.value = val
                cell.alignment = cell_align
                cell.border = border

        ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
        ws.auto_filter.ref = f"A{start_row}:{openpyxl.utils.get_column_letter(len(headers))}{start_row + len(lines)}"

        # largeurs
        for col in range(1, len(headers) + 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 18
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["D"].width = 34

    wb.save(out_path)



def export_orders_per_supplier_pdf(
    bon_df: pd.DataFrame,
    out_pdf_path: str,
    *,
    suppliers: Optional[List[Dict[str, str]]] = None,
) -> None:
    """Génère un PDF avec 1 page (ou plus) par fournisseur – version lisible & épurée.

    Colonnes : Produit | Quantité
    - Filtre les lignes Quantité <= 0
    - Bandeau titre + fournisseur
    - Tableau zébré, retour à la ligne propre, pagination
    """
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

    def _on_page(canvas_obj, doc_obj):
        canvas_obj.saveState()
        w, _h = A4
        canvas_obj.setFont("Helvetica", 8)
        canvas_obj.setFillColor(colors.HexColor("#666666"))
        canvas_obj.drawRightString(w - 18*mm, 12*mm, f"Page {doc_obj.page}")
        canvas_obj.restoreState()

    doc = SimpleDocTemplate(
        out_pdf_path,
        pagesize=A4,
        leftMargin=18*mm,
        rightMargin=18*mm,
        topMargin=18*mm,
        bottomMargin=18*mm,
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

        # Bandeau titre (table 1 cellule)
        band = Table([[Paragraph("BON DE COMMANDE", title_style)]], colWidths=[doc.width])
        band.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#2F5597")),
            ("LEFTPADDING", (0, 0), (-1, -1), 10),
            ("RIGHTPADDING", (0, 0), (-1, -1), 10),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]))
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

        data = [["Produit", "Quantité"]]
        for _, r in lines.iterrows():
            prod = str(r.get("Produit", "") or "").strip()
            qty = r.get("Quantité", 0)
            # format qty without trailing .0 when possible
            qty_txt = ""
            if pd.notna(qty):
                try:
                    q = float(qty)
                    qty_txt = f"{q:g}"
                except Exception:
                    qty_txt = str(qty)
            data.append([Paragraph(prod, cell_style), qty_txt])

        col_widths = [doc.width * 0.78, doc.width * 0.22]
        t = Table(data, colWidths=col_widths, repeatRows=1)

        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1F1F1F")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 9),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("ALIGN", (1, 1), (1, -1), "RIGHT"),
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
        story.append(Paragraph(f"Total références : <b>{max(len(data)-1, 0)}</b>", meta_style))

    doc.build(story, onFirstPage=_on_page, onLaterPages=_on_page)
        return

    if "Fournisseur" not in df.columns:
        df["Fournisseur"] = ""
    df["Fournisseur"] = df["Fournisseur"].fillna("").astype(str).str.strip()

    c = canvas.Canvas(out_pdf_path, pagesize=A4)
    width, height = A4

    for sup_name, part in df.groupby("Fournisseur", dropna=False):
        sup_name = str(sup_name or "").strip() or "(sans fournisseur)"
        info = sup_map.get(sup_name, SupplierInfo(name=sup_name))

        # Header
        y = height - 60
        c.setFont("Helvetica-Bold", 16)
        c.drawString(40, y, f"Bon de commande – {info.name}")
        y -= 22
        c.setFont("Helvetica", 10)
        if info.customer_code:
            c.drawString(40, y, f"Code client : {info.customer_code}")
            y -= 14
        if info.coord1:
            c.drawString(40, y, info.coord1)
            y -= 14
        if info.coord2:
            c.drawString(40, y, info.coord2)
            y -= 18

        # Table
        lines = group_lines_for_order(part)
        headers = list(lines.columns)
        cols = headers
        # simple layout
        col_widths = [70, 60, 80, 170, 55, 55, 55, 55]
        x0 = 40
        c.setFont("Helvetica-Bold", 9)
        for i, h in enumerate(cols):
            c.drawString(x0 + sum(col_widths[:i]) + 2, y, str(h))
        y -= 14
        c.setFont("Helvetica", 9)
        for _, row in lines.iterrows():
            if y < 60:
                c.showPage()
                y = height - 60
                c.setFont("Helvetica-Bold", 16)
                c.drawString(40, y, f"Bon de commande – {info.name}")
                y -= 26
                c.setFont("Helvetica-Bold", 9)
                for i, h in enumerate(cols):
                    c.drawString(x0 + sum(col_widths[:i]) + 2, y, str(h))
                y -= 14
                c.setFont("Helvetica", 9)

            values = [row.get(h, "") for h in cols]
            for i, v in enumerate(values):
                text = str(v)
                # tronque un peu
                if len(text) > 28 and i == 3:
                    text = text[:28] + "…"
                c.drawString(x0 + sum(col_widths[:i]) + 2, y, text)
            y -= 12

        c.showPage()

    c.save()
