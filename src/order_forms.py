from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional

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
    if df is None or df.empty:
        return pd.DataFrame(columns=["Produit", "Quantité", "Unité"])

    d = df.copy()

    if "Libellé" in d.columns:
        d["Produit"] = d["Libellé"].fillna(d.get("Produit", ""))
    elif "Produit" not in d.columns:
        d["Produit"] = ""

    unit_col = None
    for cand in ["Unité", "Unite", "Unité", "U", "Unites"]:
        if cand in d.columns:
            unit_col = cand
            break

    d["Unité"] = d[unit_col] if unit_col else ""

    d["Produit"] = d["Produit"].fillna("").astype(str).str.strip()
    d["Unité"] = d["Unité"].fillna("").astype(str).str.strip()
    d["Quantité"] = pd.to_numeric(d.get("Quantité", 0), errors="coerce").fillna(0)

    return (
        d.groupby(["Produit", "Unité"], as_index=False)["Quantité"]
        .sum()
        .sort_values(["Produit", "Unité"])
        .reset_index(drop=True)
    )


# ======================= PDF AVEC WATERMARK =======================

def export_orders_per_supplier_pdf(
    bon_df: pd.DataFrame,
    out_pdf_path: str,
    *,
    suppliers: Optional[List[Dict[str, str]]] = None,
) -> None:

    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib.units import mm
    from reportlab.lib.utils import ImageReader
    from pathlib import Path

    suppliers = suppliers or []
    sup_map = _supplier_lookup(suppliers)

    df = bon_df.copy()
    df["Fournisseur"] = df.get("Fournisseur", "").fillna("").astype(str).str.strip()

    styles = getSampleStyleSheet()
    cell_style = styles["Normal"]

    # === CHARGEMENT LOGO ===
    base_dir = Path(__file__).resolve().parent
    watermark_path = base_dir / "assets" / "watermark.jpg"

    watermark = None
    if watermark_path.exists():
        watermark = ImageReader(str(watermark_path))

    def _on_page(c, doc):
        if watermark:
            w, h = A4
            img_w = 350
            img_h = 350
            x = (w - img_w) / 2
            y = (h - img_h) / 2
            c.drawImage(watermark, x, y, width=img_w, height=img_h)

        c.setFont("Helvetica", 8)
        c.drawRightString(A4[0] - 20, 15, f"Page {doc.page}")

    doc = SimpleDocTemplate(out_pdf_path, pagesize=A4)
    story = []

    for sup_name, part in df.groupby("Fournisseur"):
        info = sup_map.get(sup_name, SupplierInfo(name=sup_name))
        lines = group_lines_for_order(part)

        story.append(Paragraph("BON DE COMMANDE", styles["Title"]))
        story.append(Paragraph(info.name, styles["Heading2"]))
        story.append(Spacer(1, 12))

        data = [["Produit", "Quantité", "Unité"]]
        for _, r in lines.iterrows():
            data.append([r["Produit"], str(r["Quantité"]), r["Unité"]])

        t = Table(data, colWidths=[300, 80, 80])
        t.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ]))

        story.append(t)
        story.append(PageBreak())

    doc.build(story, onFirstPage=_on_page, onLaterPages=_on_page)
