"""Génération de bons de commande par fournisseur (PDF).

But : après édition de l'Excel par l'utilisateur, il peut ré-uploader le fichier
et l'application produit un PDF par fournisseur, avec des zones à remplir au stylo
(dates de livraison, montant attendu).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional

import pandas as pd


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

    def draw_header(supplier: str):
        top = h - 18 * mm
        c.setFont("Helvetica-Bold", 14)
        c.drawString(18 * mm, top, f"{opt.title} — {supplier}")
        c.setFont("Helvetica", 10)
        c.drawString(18 * mm, top - 8 * mm, "Date(s) de livraison : ________________________________")
        c.drawString(18 * mm, top - 14 * mm, "Montant attendu facture (€) : _________________________")
        c.line(18 * mm, top - 18 * mm, w - 18 * mm, top - 18 * mm)
        return top - 24 * mm

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
