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
    ➜ Regroupe uniquement par Produit
    ➜ Somme des Quantités
    ➜ Sortie = Produit | Quantité
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=["Produit", "Quantité"])

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
    """Crée un classeur Excel avec 1 feuille par fournisseur."""
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
    """Génère un PDF avec 1 page par fournisseur."""
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
