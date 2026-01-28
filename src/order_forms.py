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
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib.utils import ImageReader
    from pathlib import Path

    suppliers = suppliers or []
    sup_map = _supplier_lookup(suppliers)

    df = bon_df.copy() if bon_df is not None else pd.DataFrame()
    df["Fournisseur"] = df.get("Fournisseur", "").fillna("").astype(str).str.strip()

    styles = getSampleStyleSheet()

    # ========= Watermark : recherche robuste =========
    def _find_watermark() -> Path | None:
        candidates = ["watermark.png", "watermark.jpg", "watermark.jpeg"]

        base_dir = Path(__file__).resolve().parent  # .../src

        # 1) src/assets
        for name in candidates:
            p = base_dir / "assets" / name
            if p.exists():
                return p

        # 2) assets à la racine du projet (au cas où)
        for name in candidates:
            p = base_dir.parent / "assets" / name
            if p.exists():
                return p

        # 3) cwd (streamlit peut changer le working dir)
        cwd = Path.cwd()
        for name in candidates:
            p = cwd / "src" / "assets" / name
            if p.exists():
                return p
            p = cwd / "assets" / name
            if p.exists():
                return p

        return None

    watermark_path = _find_watermark()
    watermark_reader = None
    if watermark_path is not None:
        # volontairement pas silencieux : si ImageReader casse, tu veux le savoir
        watermark_reader = ImageReader(str(watermark_path))

    def _on_page(c, doc):
        c.saveState()

        if watermark_reader is None:
            # DEBUG visible => si tu vois ça, c’est juste ton fichier image qui n’est pas au bon endroit / bon nom
            c.setFont("Helvetica", 10)
            c.setFillGray(0.85)
            c.drawString(40, A4[1] - 40, "WATERMARK INTROUVABLE (mets watermark.png/jpg dans src/assets)")
        else:
            w, h = A4
            img_w, img_h = 360, 360
            x = (w - img_w) / 2
            y = (h - img_h) / 2 - 20
            c.drawImage(watermark_reader, x, y, width=img_w, height=img_h)

        c.restoreState()

        # pagination
        c.setFont("Helvetica", 8)
        c.setFillColor(colors.black)
        c.drawRightString(A4[0] - 20, 15, f"Page {doc.page}")

    doc = SimpleDocTemplate(out_pdf_path, pagesize=A4)
    story = []

    # Evite une page blanche finale : on ne met PageBreak qu’entre fournisseurs
    suppliers_list = list(df.groupby("Fournisseur")) if not df.empty else []
    for idx, (sup_name, part) in enumerate(suppliers_list):
        sup_name = str(sup_name or "").strip() or "(sans fournisseur)"
        info = sup_map.get(sup_name, SupplierInfo(name=sup_name))
        lines = group_lines_for_order(part)

        story.append(Paragraph("BON DE COMMANDE", styles["Title"]))
        story.append(Paragraph(info.name, styles["Heading2"]))
        story.append(Spacer(1, 12))

        data = [["Produit", "Quantité", "Unité"]]
        for _, r in lines.iterrows():
            data.append([str(r.get("Produit", "")), str(r.get("Quantité", "")), str(r.get("Unité", ""))])

        t = Table(data, colWidths=[300, 80, 80])
        t.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("ALIGN", (1, 1), (1, -1), "RIGHT"),
                ]
            )
        )

        story.append(t)

        if idx < len(suppliers_list) - 1:
            story.append(PageBreak())

    doc.build(story, onFirstPage=_on_page, onLaterPages=_on_page)
