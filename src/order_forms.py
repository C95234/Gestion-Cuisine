"""Génération de bons de commande par fournisseur (PDF).

But : après édition de l'Excel par l'utilisateur, il peut ré-uploader le fichier
et l'application produit un PDF par fournisseur, avec des zones à remplir au stylo
(dates de livraison, montant attendu).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional

import pandas as pd


def export_orders_per_supplier_excel(
    bon_commande_df: pd.DataFrame,
    out_xlsx_path: str,
    options: Optional[SupplierOrderPDFOptions] = None,
    supplier_infos: Optional[Dict[str, Dict[str, str]]] = None,
) -> List[str]:
    """Crée un classeur Excel : 1 feuille par fournisseur, avec un en-tête.

    Retourne la liste des fournisseurs présents.
    """
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    opt = options or SupplierOrderPDFOptions()
    df = bon_commande_df.copy()

    for col in [opt.supplier_col, opt.qty_col, opt.unit_col, opt.product_col]:
        if col not in df.columns:
            raise ValueError(f"Colonne manquante dans le bon de commande: '{col}'")
    df[opt.supplier_col] = df[opt.supplier_col].fillna("").astype(str).str.strip()
    df = df[df[opt.supplier_col] != ""]
    if df.empty:
        raise ValueError("Aucun fournisseur renseigné dans le bon de commande.")

    suppliers = sorted(df[opt.supplier_col].unique().tolist())
    supplier_infos = supplier_infos or {}

    wb = openpyxl.Workbook()
    # Supprime la feuille par défaut
    wb.remove(wb.active)

    thin = Side(style="thin", color="9E9E9E")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="EDEDED")
    header_font = Font(bold=True)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_align = Alignment(horizontal="left", vertical="top", wrap_text=True)

    def _safe_sheet_name(name: str) -> str:
        # Excel: max 31 chars, pas de []:*?/\
        s = "".join(ch for ch in name if ch not in "[]:*?/\\")
        s = s.strip() or "Fournisseur"
        s = s[:31]
        # unique
        base = s
        k = 2
        while s in wb.sheetnames:
            suf = f"_{k}"
            s = (base[: 31 - len(suf)] + suf)[:31]
            k += 1
        return s

    for supplier in suppliers:
        sh = wb.create_sheet(_safe_sheet_name(supplier))

        info = supplier_infos.get(supplier, {}) or {}
        code_client = (info.get("code_client") or "").strip()
        coord1 = (info.get("coord1") or "").strip()
        coord2 = (info.get("coord2") or "").strip()

        # En-tête
        sh["A1"].value = f"{opt.title} — {supplier}"
        sh["A1"].font = Font(bold=True, size=14)

        r = 2
        if code_client:
            sh[f"A{r}"].value = f"Code client : {code_client}"
            r += 1
        if coord1:
            sh[f"A{r}"].value = coord1
            r += 1
        if coord2:
            sh[f"A{r}"].value = coord2
            r += 1

        sh[f"A{r}"].value = "Date(s) de livraison : ________________________________"
        r += 1
        sh[f"A{r}"].value = "Montant attendu facture (€) : _________________________"
        r += 2

        start_row = r
        sub = df[df[opt.supplier_col] == supplier].copy()
        cols = [
            c for c in [opt.days_col, opt.meal_col, opt.typology_col, opt.product_col, opt.qty_col, opt.unit_col]
            if c in sub.columns
        ]

        # Table header
        for j, col in enumerate(cols, start=1):
            cell = sh.cell(row=start_row, column=j)
            cell.value = col
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = border

        # Rows
        for i, (_, row) in enumerate(sub[cols].iterrows(), start=1):
            for j, col in enumerate(cols, start=1):
                cell = sh.cell(row=start_row + i, column=j)
                val = row[col]
                if pd.isna(val):
                    val = ""
                cell.value = str(val)
                cell.alignment = cell_align
                cell.border = border

        # Ajuste largeur simple
        for j, col in enumerate(cols, start=1):
            max_len = max([len(str(col))] + [len(str(v)) for v in sub[col].fillna("").astype(str).tolist()])
            sh.column_dimensions[openpyxl.utils.get_column_letter(j)].width = min(max(10, max_len + 2), 45)

        sh.freeze_panes = sh["A" + str(start_row + 1)]

    wb.save(out_xlsx_path)
    return suppliers


@dataclass
class SupplierOrderPDFOptions:
    title: str = "Bon de commande"
    supplier_col: str = "Fournisseur"
    qty_col: str = "Quantité"
    unit_col: str = "Unité"
    product_col: str = "Produit"
    meal_col: str = "Repas"
    typology_col: str = "Typologie"
    days_col: str = "Jour(s)"


def export_orders_per_supplier_pdf(
    bon_commande_df: pd.DataFrame,
    out_pdf_path: str,
    options: Optional[SupplierOrderPDFOptions] = None,
    supplier_infos: Optional[Dict[str, Dict[str, str]]] = None,
) -> List[str]:
    """Crée un seul PDF multi-pages : 1 page par fournisseur.

    Retourne la liste des fournisseurs présents.
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    from reportlab.platypus import Table, TableStyle
    from reportlab.lib import colors

    opt = options or SupplierOrderPDFOptions()
    df = bon_commande_df.copy()

    # Normalisation basique
    for col in [opt.supplier_col, opt.qty_col, opt.unit_col, opt.product_col]:
        if col not in df.columns:
            raise ValueError(f"Colonne manquante dans le bon de commande: '{col}'")
    df[opt.supplier_col] = df[opt.supplier_col].fillna("").astype(str).str.strip()
    df = df[df[opt.supplier_col] != ""]
    if df.empty:
        raise ValueError("Aucun fournisseur renseigné dans le bon de commande.")

    suppliers = sorted(df[opt.supplier_col].unique().tolist())

    c = canvas.Canvas(out_pdf_path, pagesize=A4)
    w, h = A4

    supplier_infos = supplier_infos or {}

    def draw_header(supplier: str):
        top = h - 18 * mm
        c.setFont("Helvetica-Bold", 14)
        c.drawString(18 * mm, top, f"{opt.title} — {supplier}")
        c.setFont("Helvetica", 10)
        info = supplier_infos.get(supplier, {}) or {}
        code_client = (info.get("code_client") or "").strip()
        coord1 = (info.get("coord1") or "").strip()
        coord2 = (info.get("coord2") or "").strip()

        y = top - 8 * mm
        if code_client:
            c.drawString(18 * mm, y, f"Code client : {code_client}")
            y -= 6 * mm
        if coord1:
            c.drawString(18 * mm, y, coord1)
            y -= 5 * mm
        if coord2:
            c.drawString(18 * mm, y, coord2)
            y -= 5 * mm

        y -= 2 * mm
        c.drawString(18 * mm, y, "Date(s) de livraison : ________________________________")
        c.drawString(18 * mm, y - 6 * mm, "Montant attendu facture (€) : _________________________")
        c.line(18 * mm, y - 10 * mm, w - 18 * mm, y - 10 * mm)
        return y - 16 * mm

    for supplier in suppliers:
        y = draw_header(supplier)
        sub = df[df[opt.supplier_col] == supplier].copy()

        # Colonnes dans l'ordre (si présentes)
        cols = [
            c for c in [opt.days_col, opt.meal_col, opt.typology_col, opt.product_col, opt.qty_col, opt.unit_col]
            if c in sub.columns
        ]

        table_data: List[List[str]] = [cols]
        for _, row in sub[cols].iterrows():
            out_row = []
            for col in cols:
                val = row[col]
                if pd.isna(val):
                    val = ""
                out_row.append(str(val))
            table_data.append(out_row)

        # Table reportlab
        table = Table(table_data, repeatRows=1)
        table.setStyle(
            TableStyle(
                [
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 9),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ]
            )
        )

        # Positionnement
        avail_w = w - 36 * mm
        avail_h = y - 18 * mm
        tw, th = table.wrapOn(c, avail_w, avail_h)
        if th > avail_h:
            # Si trop long, on réduit un peu (fallback)
            table.setStyle(TableStyle([("FONTSIZE", (0, 0), (-1, -1), 8)]))
            tw, th = table.wrapOn(c, avail_w, avail_h)
        table.drawOn(c, 18 * mm, y - th)

        c.showPage()

    c.save()
    return suppliers
