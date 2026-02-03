from __future__ import annotations

import datetime as dt
import json
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# PDF (factures simplifiées par site)
import io
import zipfile

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas


# -----------------------------------------------------------------------------
# Objectif
# -----------------------------------------------------------------------------
# Module "Facturation PDJ" (Petit-déjeuner / goûter)
# - Permet d'enregistrer plusieurs bons de commande (lignes produit / quantités)
# - Normalise le site de facturation (24 ter + 24 simple => Internat)
# - Gère des prix unitaires par produit (optionnellement spécifiques à un site)
# - Autorise des ajustements manuels (consommations supplémentaires) et des avoirs
# - Exporte un classeur Excel mensuel (détail + synthèse)


# -----------------------------------------------------------------------------
# Produits par défaut (issus du modèle de bon PDJ fourni)
# -----------------------------------------------------------------------------

DEFAULT_PRODUCTS: List[str] = [
    "Lait demi - écrémé",
    "Lait entier",
    "Céréales",
    "Biscotte",
    "Sucre en sachet",
    "Sucre en morceau",
    "Beurre, plaquettes de 250g",
    "Chocolat en poudre",
    "Brioche",
    "Bledine arome chocolat",
    "Bledine arome vanille",
    "Confiture en carton",
    "Thé en boite",
    "Café en carton",
    "Jus d'orange",
    "Jus de pomme",
    "Jus de raisin",
    "Mayonnaise",
    "Ketchup",
    "Sel",
    "Poivre",
    "Fromage blanc pot de 5kg",
    "Yaourt Nature",
]


# -----------------------------------------------------------------------------
# Storage
# -----------------------------------------------------------------------------


def _data_dir() -> Path:
    base = Path(__file__).resolve().parent.parent
    d = base / "data" / "facturation_pdj"
    try:
        d.mkdir(parents=True, exist_ok=True)
        return d
    except PermissionError:
        # Fallback (environnements en lecture seule)
        d2 = Path.home() / ".gestion_cuisine" / "facturation_pdj"
        d2.mkdir(parents=True, exist_ok=True)
        return d2


def _records_path() -> Path:
    return _data_dir() / "pdj_records.csv"


def _prices_path() -> Path:
    return _data_dir() / "pdj_unit_prices.csv"


def _money_adj_path() -> Path:
    return _data_dir() / "pdj_money_adjustments.csv"


def _meta_path() -> Path:
    return _data_dir() / "meta.json"


def _read_meta() -> dict:
    p = _meta_path()
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _write_meta(meta: dict) -> None:
    _meta_path().write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")


# -----------------------------------------------------------------------------
# Normalisation
# -----------------------------------------------------------------------------


def _norm(s: str) -> str:
    return (str(s or "")).strip().lower()


def norm_site_facturation(site: str) -> str:
    """Règle métier: '24 ter' + '24 simple' => colonne unique 'Internat'."""
    s0 = str(site or "").strip()
    sN = _norm(s0)

    if re.fullmatch(r"24\s*(ter|simple)", sN):
        return "Internat"
    if sN in {"24ter", "24simple"}:
        return "Internat"
    if "24" in sN and ("ter" in sN or "simple" in sN):
        return "Internat"

    return s0


def norm_product(p: str) -> str:
    # On conserve la casse d'origine la plupart du temps, mais on nettoie les espaces.
    return " ".join(str(p or "").strip().split())


# -----------------------------------------------------------------------------
# Chargement / sauvegarde
# -----------------------------------------------------------------------------


def load_pdj_records() -> pd.DataFrame:
    p = _records_path()
    if not p.exists():
        return pd.DataFrame(
            columns=[
                "date",
                "month",
                "site",
                "product",
                "qty",
                "kind",
                "comment",
                "source",
            ]
        )
    df = pd.read_csv(p, parse_dates=["date"])
    df["date"] = pd.to_datetime(df["date"]).dt.date
    df["month"] = df["month"].astype(str)
    df["site"] = df["site"].astype(str)
    df["product"] = df["product"].astype(str)
    df["qty"] = pd.to_numeric(df["qty"], errors="coerce").fillna(0)
    df["kind"] = df["kind"].astype(str)
    df["comment"] = df.get("comment", "").astype(str)
    df["source"] = df.get("source", "").astype(str)
    return df


def add_pdj_records(
    records: pd.DataFrame,
    *,
    source_filename: str = "",
) -> int:
    """Ajoute (append) des lignes PDJ.

    records attend au minimum: date, site, product, qty.
    Optionnel: kind (commande|manuel|avoir_qty)
    """
    if records is None or records.empty:
        return 0

    df = records.copy()
    if "date" not in df.columns:
        raise ValueError("La colonne 'date' est obligatoire")
    if "site" not in df.columns:
        raise ValueError("La colonne 'site' est obligatoire")
    if "product" not in df.columns:
        raise ValueError("La colonne 'product' est obligatoire")
    if "qty" not in df.columns:
        raise ValueError("La colonne 'qty' est obligatoire")

    df["date"] = pd.to_datetime(df["date"]).dt.date
    df["month"] = df["date"].map(lambda d: f"{d.year:04d}-{d.month:02d}")
    df["site"] = df["site"].astype(str).map(norm_site_facturation)
    df["product"] = df["product"].astype(str).map(norm_product)
    df["qty"] = pd.to_numeric(df["qty"], errors="coerce").fillna(0)
    if "kind" not in df.columns:
        df["kind"] = "commande"
    df["kind"] = df["kind"].astype(str)
    if "comment" not in df.columns:
        df["comment"] = ""
    df["comment"] = df["comment"].astype(str)
    df["source"] = source_filename

    df = df[["date", "month", "site", "product", "qty", "kind", "comment", "source"]]

    p = _records_path()
    if p.exists():
        old = pd.read_csv(p, parse_dates=["date"])
        old["date"] = pd.to_datetime(old["date"]).dt.date
    else:
        old = pd.DataFrame(columns=df.columns)

    merged = pd.concat([old, df], ignore_index=True)
    merged.to_csv(p, index=False)

    meta = _read_meta()
    meta.setdefault("last_update", dt.datetime.now().isoformat(timespec="seconds"))
    _write_meta(meta)
    return int(len(df))


def load_unit_prices() -> pd.DataFrame:
    """Tarifs unitaires.

    Colonnes:
    - product
    - site (optionnel; '__default__' si commun)
    - unit_price
    - unit (optionnel)
    """
    p = _prices_path()
    if not p.exists():
        # Initialise une grille avec les produits par défaut
        return pd.DataFrame(
            {
                "product": DEFAULT_PRODUCTS,
                "site": "__default__",
                "unit_price": 0.0,
                "unit": "",
            }
        )
    df = pd.read_csv(p)
    df["product"] = df["product"].astype(str).map(norm_product)
    df["site"] = df.get("site", "__default__").astype(str)
    df["unit_price"] = pd.to_numeric(df.get("unit_price", 0.0), errors="coerce").fillna(0.0)
    df["unit"] = df.get("unit", "").astype(str)
    return df[["product", "site", "unit_price", "unit"]]


def save_unit_prices(prices_df: pd.DataFrame) -> int:
    if prices_df is None or prices_df.empty:
        return 0
    df = prices_df.copy()
    if "product" not in df.columns:
        raise ValueError("La colonne 'product' est obligatoire")
    df["product"] = df["product"].astype(str).map(norm_product)
    if "site" not in df.columns:
        df["site"] = "__default__"
    df["site"] = df["site"].astype(str)
    df["unit_price"] = pd.to_numeric(df.get("unit_price", 0.0), errors="coerce").fillna(0.0)
    if "unit" not in df.columns:
        df["unit"] = ""
    df["unit"] = df["unit"].astype(str)
    df = df[["product", "site", "unit_price", "unit"]]

    _prices_path().write_text(df.to_csv(index=False), encoding="utf-8")
    return int(len(df))


def load_money_adjustments() -> pd.DataFrame:
    """Ajustements monétaires (avoirs / frais / corrections) en euros.

    Colonnes:
    - date
    - month
    - site
    - label
    - amount_eur (positif ou négatif)
    - comment
    """
    p = _money_adj_path()
    if not p.exists():
        return pd.DataFrame(columns=["date", "month", "site", "label", "amount_eur", "comment"])
    df = pd.read_csv(p, parse_dates=["date"])
    df["date"] = pd.to_datetime(df["date"]).dt.date
    df["month"] = df["month"].astype(str)
    df["site"] = df["site"].astype(str)
    df["label"] = df.get("label", "Ajustement").astype(str)
    df["amount_eur"] = pd.to_numeric(df.get("amount_eur", 0.0), errors="coerce").fillna(0.0)
    df["comment"] = df.get("comment", "").astype(str)
    return df[["date", "month", "site", "label", "amount_eur", "comment"]]


def add_money_adjustments(adj: pd.DataFrame) -> int:
    if adj is None or adj.empty:
        return 0
    df = adj.copy()
    if "date" not in df.columns:
        raise ValueError("La colonne 'date' est obligatoire")
    if "site" not in df.columns:
        raise ValueError("La colonne 'site' est obligatoire")
    if "amount_eur" not in df.columns:
        raise ValueError("La colonne 'amount_eur' est obligatoire")
    if "label" not in df.columns:
        df["label"] = "Ajustement"
    if "comment" not in df.columns:
        df["comment"] = ""

    df["date"] = pd.to_datetime(df["date"]).dt.date
    df["month"] = df["date"].map(lambda d: f"{d.year:04d}-{d.month:02d}")
    df["site"] = df["site"].astype(str).map(norm_site_facturation)
    df["label"] = df["label"].astype(str)
    df["amount_eur"] = pd.to_numeric(df["amount_eur"], errors="coerce").fillna(0.0)
    df["comment"] = df["comment"].astype(str)
    df = df[["date", "month", "site", "label", "amount_eur", "comment"]]

    p = _money_adj_path()
    if p.exists():
        old = pd.read_csv(p, parse_dates=["date"])
        old["date"] = pd.to_datetime(old["date"]).dt.date
    else:
        old = pd.DataFrame(columns=df.columns)
    merged = pd.concat([old, df], ignore_index=True)
    merged.to_csv(p, index=False)
    return int(len(df))


# -----------------------------------------------------------------------------
# Calcul facture
# -----------------------------------------------------------------------------


def _pick_unit_price(prices: pd.DataFrame, product: str, site: str) -> Tuple[float, str]:
    """Retourne (prix, unité).

    Priorité:
    1) (product, site)
    2) (product, '__default__')
    3) 0.0
    """
    if prices is None or prices.empty:
        return 0.0, ""
    p = norm_product(product)
    s = str(site)
    sub = prices[(prices["product"] == p) & (prices["site"] == s)]
    if not sub.empty:
        r = sub.iloc[0]
        return float(r.get("unit_price", 0.0)), str(r.get("unit", "") or "")
    sub = prices[(prices["product"] == p) & (prices["site"] == "__default__")]
    if not sub.empty:
        r = sub.iloc[0]
        return float(r.get("unit_price", 0.0)), str(r.get("unit", "") or "")
    return 0.0, ""


def compute_monthly_pdj(month: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Retourne (synthèse_site, détail_lignes, ajustements_monetaires) pour un mois YYYY-MM."""
    rec = load_pdj_records()
    rec = rec[rec["month"] == month].copy()
    rec["site"] = rec["site"].astype(str).map(norm_site_facturation)
    rec["product"] = rec["product"].astype(str).map(norm_product)
    prices = load_unit_prices()

    if rec.empty:
        detail = pd.DataFrame(
            columns=["date", "site", "product", "qty", "unit_price", "unit", "amount_eur", "kind", "comment", "source"]
        )
    else:
        ups = []
        units = []
        for _, r in rec.iterrows():
            up, u = _pick_unit_price(prices, r["product"], r["site"])
            ups.append(up)
            units.append(u)
        rec["unit_price"] = ups
        rec["unit"] = units
        rec["amount_eur"] = rec["qty"].astype(float) * rec["unit_price"].astype(float)
        detail = rec[["date", "site", "product", "qty", "unit_price", "unit", "amount_eur", "kind", "comment", "source"]].copy()

    # Pivot synthèse
    if detail.empty:
        synth = pd.DataFrame(columns=["site", "total_eur"])
    else:
        synth = (
            detail.groupby("site", as_index=False)["amount_eur"].sum().sort_values("amount_eur", ascending=False)
        )
        synth = synth.rename(columns={"amount_eur": "total_eur"})

    # Ajustements €
    adj = load_money_adjustments()
    adj = adj[adj["month"] == month].copy()
    if not adj.empty:
        # Intègre aux totaux
        add = adj.groupby("site", as_index=False)["amount_eur"].sum().rename(columns={"amount_eur": "adj_eur"})
        synth = synth.merge(add, on="site", how="outer")
        synth["total_eur"] = pd.to_numeric(synth.get("total_eur", 0.0), errors="coerce").fillna(0.0)
        synth["adj_eur"] = pd.to_numeric(synth.get("adj_eur", 0.0), errors="coerce").fillna(0.0)
        synth["total_facture_eur"] = synth["total_eur"] + synth["adj_eur"]
    else:
        synth["total_facture_eur"] = synth.get("total_eur", 0.0)

    return synth, detail, adj


# -----------------------------------------------------------------------------
# Export Excel
# -----------------------------------------------------------------------------


def export_monthly_pdj_workbook(month: str, out_path: str) -> str:
    """Exporte le classeur de facturation PDJ pour le mois YYYY-MM."""
    synth, detail, adj = compute_monthly_pdj(month)
    prices = load_unit_prices()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Facturation PDJ"

    # Styles
    header_fill = PatternFill("solid", fgColor="2F5597")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(
        left=Side(style="thin", color="999999"),
        right=Side(style="thin", color="999999"),
        top=Side(style="thin", color="999999"),
        bottom=Side(style="thin", color="999999"),
    )
    money_fmt = "#,##0.00 €"

    ws["A1"].value = f"Facturation PDJ — {month}"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:D1")

    # Synthèse
    start = 3
    ws.cell(row=start, column=1, value="Site")
    ws.cell(row=start, column=2, value="Total lignes (€/produits)")
    ws.cell(row=start, column=3, value="Ajustements (€)")
    ws.cell(row=start, column=4, value="Total facturé (€)")
    for c in range(1, 5):
        cell = ws.cell(row=start, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    if synth.empty:
        ws.cell(row=start + 1, column=1, value="(Aucune donnée)")
    else:
        for i, r in enumerate(synth.itertuples(index=False), start=1):
            row = start + i
            ws.cell(row=row, column=1, value=getattr(r, "site", ""))
            ws.cell(row=row, column=2, value=float(getattr(r, "total_eur", 0.0) or 0.0))
            ws.cell(row=row, column=3, value=float(getattr(r, "adj_eur", 0.0) or 0.0))
            ws.cell(row=row, column=4, value=float(getattr(r, "total_facture_eur", 0.0) or 0.0))
            for c in range(1, 5):
                cell = ws.cell(row=row, column=c)
                cell.border = border
                if c in (2, 3, 4):
                    cell.number_format = money_fmt

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 18

    # Détail
    ws2 = wb.create_sheet("Détail lignes")
    cols = ["date", "site", "product", "qty", "unit", "unit_price", "amount_eur", "kind", "comment", "source"]
    ws2.append(cols)
    for c in range(1, len(cols) + 1):
        cell = ws2.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    if not detail.empty:
        for r in detail[cols].itertuples(index=False):
            ws2.append(list(r))
        for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, min_col=1, max_col=len(cols)):
            for cell in row:
                cell.border = border
                if cell.column_letter in {"F", "G"}:
                    cell.number_format = money_fmt
                if cell.column_letter == "D":
                    cell.number_format = "0.00"
    ws2.freeze_panes = "A2"
    widths = {
        "A": 12,
        "B": 20,
        "C": 28,
        "D": 10,
        "E": 10,
        "F": 12,
        "G": 14,
        "H": 12,
        "I": 28,
        "J": 22,
    }
    for k, w in widths.items():
        ws2.column_dimensions[k].width = w

    # Ajustements monétaires
    ws3 = wb.create_sheet("Ajustements €")
    ws3.append(["date", "site", "label", "amount_eur", "comment"])
    for c in range(1, 6):
        cell = ws3.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    if not adj.empty:
        for r in adj[["date", "site", "label", "amount_eur", "comment"]].itertuples(index=False):
            ws3.append(list(r))
        for row in ws3.iter_rows(min_row=2, max_row=ws3.max_row, min_col=1, max_col=5):
            for cell in row:
                cell.border = border
                if cell.column_letter == "D":
                    cell.number_format = money_fmt
    ws3.column_dimensions["A"].width = 12
    ws3.column_dimensions["B"].width = 20
    ws3.column_dimensions["C"].width = 22
    ws3.column_dimensions["D"].width = 14
    ws3.column_dimensions["E"].width = 28

    # Tarifs
    ws4 = wb.create_sheet("Tarifs")
    ws4.append(["product", "site", "unit_price", "unit"])
    for c in range(1, 5):
        cell = ws4.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    if prices is not None and not prices.empty:
        for r in prices[["product", "site", "unit_price", "unit"]].itertuples(index=False):
            ws4.append(list(r))
        for row in ws4.iter_rows(min_row=2, max_row=ws4.max_row, min_col=1, max_col=4):
            for cell in row:
                cell.border = border
                if cell.column_letter == "C":
                    cell.number_format = money_fmt
    ws4.column_dimensions["A"].width = 32
    ws4.column_dimensions["B"].width = 18
    ws4.column_dimensions["C"].width = 14
    ws4.column_dimensions["D"].width = 10

    wb.save(out_path)
    return out_path


# -----------------------------------------------------------------------------
# Export PDF simplifié : 1 facture par site (mois)
# -----------------------------------------------------------------------------


def _safe_filename(s: str) -> str:
    s = re.sub(r"[^A-Za-z0-9._ -]+", "_", str(s or "")).strip()
    s = re.sub(r"\s+", " ", s)
    return s or "site"


def build_site_invoice_tables(month: str, site: str) -> Tuple[pd.DataFrame, pd.DataFrame, float]:
    """Retourne (lignes_produits, ajustements, total_facture).

    lignes_produits: product, qty, unit, unit_price, line_total
    ajustements: label, amount_eur
    """
    synth, detail, adj = compute_monthly_pdj(month)
    siteN = norm_site_facturation(site)

    d = detail[detail["site"] == siteN].copy()
    if d.empty:
        lines = pd.DataFrame(columns=["product", "qty", "unit", "unit_price", "line_total"])
    else:
        lines = (
            d.groupby(["product", "unit", "unit_price"], as_index=False)["qty"].sum()
            .sort_values("product")
        )
        lines["line_total"] = lines["qty"].astype(float) * lines["unit_price"].astype(float)

    a = adj[adj["site"] == siteN].copy()
    if a.empty:
        adj_tbl = pd.DataFrame(columns=["label", "amount_eur"])
        adj_sum = 0.0
    else:
        adj_tbl = a.groupby("label", as_index=False)["amount_eur"].sum().sort_values("label")
        adj_sum = float(adj_tbl["amount_eur"].sum())

    total = float(lines["line_total"].sum() if not lines.empty else 0.0) + adj_sum
    return lines, adj_tbl, total


def export_site_invoice_pdf(month: str, site: str, out_path) -> str:
    """
    Génère un PDF A4 (lisible et "compta") pour 1 site.

    - 1 page (ou multi-pages si besoin) avec un vrai tableau
    - Pas de colonnes superflues
    - Compatible avec out_path = chemin (str) OU buffer (BytesIO)
    """
    lines, adj_tbl, total = build_site_invoice_tables(month, site)
    siteN = norm_site_facturation(site)

    # --- reportlab (platypus) pour un rendu propre ---
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

    def money(v: float) -> str:
        try:
            v = float(v or 0.0)
        except Exception:
            v = 0.0
        return f"{v:,.2f} €".replace(",", " ").replace(".", ",")

    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "Title",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=16,
        leading=20,
        spaceAfter=6,
    )
    subtitle = ParagraphStyle(
        "SubTitle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=11,
        leading=14,
        spaceAfter=2,
    )
    small = ParagraphStyle(
        "Small",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
    )

    doc = SimpleDocTemplate(
        out_path,
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
        title=f"FACTURATION PDJ {month} - {siteN}",
        author="Gestion-Cuisine",
    )

    story = []
    story.append(Paragraph(f"FACTURATION PDJ — {month}", title))
    story.append(Paragraph(f"<b>Site :</b> {siteN}", subtitle))
    story.append(Paragraph(f"<b>Généré le :</b> {dt.date.today().strftime('%d/%m/%Y')}", ParagraphStyle("gen", parent=subtitle, spaceAfter=10)))

    # --- Tableau consommations ---
    header = ["Produit", "Qté", "PU", "Total"]
    data = [header]

    if lines.empty:
        data.append([Paragraph("(Aucune consommation enregistrée)", small), "", "", ""])
    else:
        for r in lines.itertuples(index=False):
            prod = str(getattr(r, "product", "") or "")
            qty = float(getattr(r, "qty", 0.0) or 0.0)
            pu = float(getattr(r, "unit_price", 0.0) or 0.0)
            lt = float(getattr(r, "line_total", 0.0) or 0.0)
            data.append([Paragraph(prod, small), f"{qty:g}", money(pu), money(lt)])

    # Largeurs colonnes (A4 moins marges)
    table = Table(
        data,
        colWidths=[110 * mm, 18 * mm, 22 * mm, 26 * mm],
        repeatRows=1,
        hAlign="LEFT",
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 10),
                ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                ("ALIGN", (0, 0), (0, -1), "LEFT"),
                ("FONTSIZE", (0, 1), (-1, -1), 10),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FBFBFB")]),
                ("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.HexColor("#C9C9C9")),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#DDDDDD")),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(table)

    # --- Ajustements (si présents) ---
    if not adj_tbl.empty:
        story.append(Spacer(1, 8))
        story.append(Paragraph("<b>Ajustements</b>", subtitle))

        adj_data = [["Libellé", "Montant"]]
        for r in adj_tbl.itertuples(index=False):
            lab = str(getattr(r, "label", "Ajustement") or "Ajustement")
            amt = float(getattr(r, "amount_eur", 0.0) or 0.0)
            adj_data.append([lab, money(amt)])

        adj_table = Table(adj_data, colWidths=[150 * mm, 26 * mm], repeatRows=1, hAlign="LEFT")
        adj_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 10),
                    ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#DDDDDD")),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        story.append(adj_table)

    # --- Total ---
    story.append(Spacer(1, 10))
    total_tbl = Table(
        [["TOTAL MENSUEL À RÉGLER", money(total)]],
        colWidths=[150 * mm, 26 * mm],
        hAlign="LEFT",
    )
    total_tbl.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (0, 0), "Helvetica-Bold"),
                ("FONTNAME", (1, 0), (1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 12),
                ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                ("LINEABOVE", (0, 0), (-1, 0), 1.0, colors.black),
                ("TOPPADDING", (0, 0), (-1, 0), 8),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
            ]
        )
    )
    story.append(total_tbl)

    # Pied de page discret (numéro de page)
    def _on_page(c, doc):
        c.saveState()
        c.setFont("Helvetica", 8)
        c.setFillColor(colors.HexColor("#666666"))
        c.drawRightString(doc.pagesize[0] - doc.rightMargin, 10 * mm, f"Page {doc.page}")
        c.restoreState()

    doc.build(story, onFirstPage=_on_page, onLaterPages=_on_page)

    # SimpleDocTemplate accepte un buffer ; la "valeur de retour" sert juste au code appelant
    return out_path if isinstance(out_path, str) else ""


def export_monthly_invoices_zip(month: str) -> bytes:
    """Retourne un ZIP en mémoire contenant 1 PDF par site présent sur le mois."""
    synth, detail, adj = compute_monthly_pdj(month)
    sites = set(detail["site"].dropna().astype(str).tolist()) | set(adj["site"].dropna().astype(str).tolist())
    sites = sorted([s for s in sites if str(s).strip()])

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        if not sites:
            # zip vide + README
            z.writestr("README.txt", f"Aucune donnée PDJ pour {month}.")
        for s in sites:
            pdf_bytes = io.BytesIO()
            # reportlab écrit sur un chemin ou file-like : on passe un buffer.
            export_site_invoice_pdf(month, s, pdf_bytes)  # type: ignore[arg-type]
            name = f"Facture_PDJ_{month}_{_safe_filename(s)}.pdf"
            z.writestr(name, pdf_bytes.getvalue())
    return buf.getvalue()


# -----------------------------------------------------------------------------
# Import helpers (Excel)
# -----------------------------------------------------------------------------


def parse_pdj_excel(path_or_buffer) -> pd.DataFrame:
    """Tente d'extraire une table (product, qty) depuis un fichier Excel.

    Cette fonction est volontairement "tolérante" : elle essaie quelques heuristiques
    sans casser si le format varie.
    """
    try:
        df = pd.read_excel(path_or_buffer, sheet_name=0)
    except Exception:
        # Tentative xls via xlrd si dispo
        df = pd.read_excel(path_or_buffer, sheet_name=0, engine="xlrd")

    if df is None or df.empty:
        return pd.DataFrame(columns=["product", "qty"])

    # Normalise en string
    cols = [str(c) for c in df.columns]
    df.columns = cols

    # Cherche colonnes candidates
    def _find_col(patterns: List[str]) -> Optional[str]:
        for c in cols:
            cn = _norm(c)
            if any(p in cn for p in patterns):
                return c
        return None

    col_prod = _find_col(["produit", "libell", "designation", "article"]) or cols[0]
    col_qty = _find_col(["qt", "quant", "qte", "commande", "consomm"]) 

    # Si pas trouvé, essaye la 2e colonne numérique
    if col_qty is None and len(cols) >= 2:
        col_qty = cols[1]

    out = pd.DataFrame({
        "product": df[col_prod].astype(str),
        "qty": pd.to_numeric(df[col_qty], errors="coerce") if col_qty in df.columns else 0,
    })
    out["product"] = out["product"].map(norm_product)
    out["qty"] = pd.to_numeric(out["qty"], errors="coerce").fillna(0.0)
    out = out[out["product"].astype(str).str.strip() != ""].copy()
    out = out[out["qty"] != 0].copy()
    return out[["product", "qty"]]


def parse_pdj_pdf(path_or_buffer) -> pd.DataFrame:
    """Extraction *best-effort* (product, qty) depuis un PDF.

    ⚠️ Beaucoup de bons PDJ sont scannés (image) : il n'existe pas de texte exploitable.
    On tente donc une OCR légère, puis on cherche les produits connus et un nombre
    présent sur la même ligne. Résultat à **vérifier/corriger** dans l'interface.
    """
    try:
        import pdfplumber
        import pytesseract
    except Exception:
        # Pas de dépendances : on ne bloque pas
        return pd.DataFrame(columns=["product", "qty"])

    # OCR page 1 (souvent suffisant)
    try:
        with pdfplumber.open(path_or_buffer) as pdf:
            if not pdf.pages:
                return pd.DataFrame(columns=["product", "qty"])
            img = pdf.pages[0].to_image(resolution=200).original
    except Exception:
        return pd.DataFrame(columns=["product", "qty"])

    try:
        raw = pytesseract.image_to_string(img, lang="eng", config="--psm 6")
    except Exception:
        return pd.DataFrame(columns=["product", "qty"])

    lines = [" ".join(l.strip().split()) for l in (raw or "").splitlines() if l.strip()]
    if not lines:
        return pd.DataFrame(columns=["product", "qty"])

    # Indexe par ligne normalisée
    ln_norm = [re.sub(r"[^a-z0-9' ]+", " ", l.lower()).strip() for l in lines]

    out_rows = []
    for prod in DEFAULT_PRODUCTS:
        p_norm = re.sub(r"[^a-z0-9' ]+", " ", prod.lower()).strip()
        # match approximatif: contient au moins 2 mots du produit
        tokens = [t for t in p_norm.split() if len(t) >= 3]
        if not tokens:
            continue
        best_i = None
        best_score = 0
        for i, l in enumerate(ln_norm):
            score = sum(1 for t in tokens if t in l)
            if score > best_score:
                best_score = score
                best_i = i
        if best_i is None or best_score < max(1, min(2, len(tokens))):
            continue

        original_line = lines[best_i]
        # Cherche un nombre (ex: 12, 12.5, 12,5) sur la ligne
        m = re.findall(r"-?\d+[\.,]?\d*", original_line)
        qty = 0.0
        if m:
            # prend le dernier nombre (souvent à droite)
            try:
                qty = float(m[-1].replace(",", "."))
            except Exception:
                qty = 0.0

        if qty != 0:
            out_rows.append({"product": norm_product(prod), "qty": qty})

    if not out_rows:
        return pd.DataFrame(columns=["product", "qty"])
    return pd.DataFrame(out_rows)[["product", "qty"]]
