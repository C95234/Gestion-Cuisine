
from __future__ import annotations

import json
import datetime as dt
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


# -----------------------------
# Storage
# -----------------------------

def _data_dir() -> Path:
    """Local persistent folder (next to the app) to store weekly saved plannings."""
    base = Path(__file__).resolve().parent.parent  # GenerateurBonCommande/
    d = base / "data" / "facturation"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _records_path() -> Path:
    return _data_dir() / "records.csv"


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


# -----------------------------
# Normalization helpers
# -----------------------------

def _norm(s: str) -> str:
    return (s or "").strip().lower()


def is_pdj_regime(regime: str) -> bool:
    r = _norm(regime)
    return ("pdj" in r) or ("petit" in r and "dej" in r) or ("gouter" in r) or ("goûter" in r)


def is_mixe_lisse_regime(regime: str) -> bool:
    r = _norm(regime)
    return ("mixe" in r) or ("mixé" in r) or ("lisse" in r) or ("m/l" in r) or ("ml" == r)


def planning_to_daily_totals(
    dej: pd.DataFrame,
    din: pd.DataFrame,
    week_monday: dt.date,
    *,
    exclude_pdj: bool = True,
    exclude_mixe_lisse: bool = True,
) -> pd.DataFrame:
    """
    Convert planning fabrication (déjeuner + dîner) into day-level totals per site.

    Output columns: date, site, qty_repas
    """
    def _one(df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame(columns=["Site", "Regime", "Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"])
        return df.copy()

    dej = _one(dej)
    din = _one(din)

    df = pd.concat([dej, din], ignore_index=True)

    if df.empty:
        return pd.DataFrame(columns=["date", "site", "qty_repas"])

    # Filters
    mask = pd.Series([True] * len(df))
    if exclude_pdj and "Regime" in df.columns:
        mask &= ~df["Regime"].astype(str).map(is_pdj_regime)
    if exclude_mixe_lisse and "Regime" in df.columns:
        mask &= ~df["Regime"].astype(str).map(is_mixe_lisse_regime)

    df = df.loc[mask].copy()

    day_cols = [c for c in ["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"] if c in df.columns]

    # Melt
    melted = df.melt(id_vars=["Site"], value_vars=day_cols, var_name="day_name", value_name="qty")
    melted["qty"] = pd.to_numeric(melted["qty"], errors="coerce").fillna(0).astype(int)

    day_index = {name: i for i, name in enumerate(["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"])}
    melted["date"] = melted["day_name"].map(lambda d: week_monday + dt.timedelta(days=day_index.get(d, 0)))
    melted = melted.drop(columns=["day_name"])

    out = melted.groupby(["date", "Site"], as_index=False)["qty"].sum()
    out = out.rename(columns={"Site": "site", "qty": "qty_repas"})
    return out


def mixe_lisse_to_daily_totals(
    dej: Optional[pd.DataFrame],
    din: Optional[pd.DataFrame],
    week_monday: dt.date,
) -> pd.DataFrame:
    """
    Convert planning mixé/lissé (déjeuner + dîner) into day-level totals per site.

    Output columns: date, site, qty_ml
    """
    frames = []
    for df in (dej, din):
        if df is not None and not df.empty:
            frames.append(df.copy())
    if not frames:
        return pd.DataFrame(columns=["date", "site", "qty_ml"])
    df = pd.concat(frames, ignore_index=True)

    day_cols = [c for c in ["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"] if c in df.columns]
    melted = df.melt(id_vars=["Site"], value_vars=day_cols, var_name="day_name", value_name="qty")
    melted["qty"] = pd.to_numeric(melted["qty"], errors="coerce").fillna(0).astype(int)

    day_index = {name: i for i, name in enumerate(["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"])}
    melted["date"] = melted["day_name"].map(lambda d: week_monday + dt.timedelta(days=day_index.get(d, 0)))
    out = melted.groupby(["date", "Site"], as_index=False)["qty"].sum()
    out = out.rename(columns={"Site": "site", "qty": "qty_ml"})
    return out


# -----------------------------
# Persist / load
# -----------------------------

def save_week(
    *,
    week_monday: dt.date,
    repas_daily: pd.DataFrame,
    ml_daily: pd.DataFrame,
    source_filename: str = "",
) -> Tuple[int, int]:
    """
    Append records for a week into storage (idempotent per date+site).
    Returns (n_repas, n_ml) saved rows.
    """
    # Normalize
    def _norm_df(df: pd.DataFrame, qty_col: str, category: str) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame(columns=["date","site","category","qty","week_monday","source"])
        out = df.copy()
        out["date"] = pd.to_datetime(out["date"]).dt.date
        out["site"] = out["site"].astype(str).str.strip()
        out["category"] = category
        out["qty"] = pd.to_numeric(out[qty_col], errors="coerce").fillna(0).astype(int)
        out["week_monday"] = week_monday.isoformat()
        out["source"] = source_filename
        return out[["date","site","category","qty","week_monday","source"]]

    a = _norm_df(repas_daily, "qty_repas", "repas")
    b = _norm_df(ml_daily, "qty_ml", "mixe_lisse")
    new = pd.concat([a,b], ignore_index=True)

    p = _records_path()
    if p.exists():
        old = pd.read_csv(p, parse_dates=["date"])
        old["date"] = pd.to_datetime(old["date"]).dt.date
    else:
        old = pd.DataFrame(columns=["date","site","category","qty","week_monday","source"])

    # Drop existing rows for same date+site+category within the saved week range (7 days)
    week_dates = {week_monday + dt.timedelta(days=i) for i in range(7)}
    mask_keep = ~(
        old["date"].isin(list(week_dates))
        & old["category"].isin(["repas","mixe_lisse"])
    )
    old = old.loc[mask_keep].copy()

    merged = pd.concat([old, new], ignore_index=True)
    merged.to_csv(p, index=False)

    meta = _read_meta()
    meta.setdefault("saved_weeks", [])
    if week_monday.isoformat() not in meta["saved_weeks"]:
        meta["saved_weeks"].append(week_monday.isoformat())
        meta["saved_weeks"] = sorted(meta["saved_weeks"])
    _write_meta(meta)

    return int(len(a)), int(len(b))


def load_records() -> pd.DataFrame:
    p = _records_path()
    if not p.exists():
        return pd.DataFrame(columns=["date","site","category","qty","week_monday","source"])
    df = pd.read_csv(p, parse_dates=["date"])
    df["date"] = pd.to_datetime(df["date"]).dt.date
    df["site"] = df["site"].astype(str)
    df["category"] = df["category"].astype(str)
    df["qty"] = pd.to_numeric(df["qty"], errors="coerce").fillna(0).astype(int)
    return df


# -----------------------------
# Export workbook (monthly)
# -----------------------------

DEFAULT_UNIT_PRICES = {
    "repas": {"__default__": 6.50, "MAS": 6.65},
    "mixe_lisse": {"__default__": 7.61, "MAS": 7.74},
}

def _unit_price(category: str, site: str, custom_prices: Optional[dict] = None) -> float:
    prices = (custom_prices or DEFAULT_UNIT_PRICES).get(category, {})
    s = (site or "").strip().upper()
    return float(prices.get(s, prices.get("__default__", 0.0)))


def export_monthly_workbook(
    records: pd.DataFrame,
    out_path: str,
    *,
    custom_prices: Optional[dict] = None,
) -> str:
    """
    Create an Excel workbook for billing.

    The workbook always contains **12 sheets (Janvier → Décembre)** for a single year.
    The exported year is the most recent year found in ``records``.

    Each month sheet contains 2 blocks: Repas and Mixé/Lissé.
    Layout is aligned with your monthly template (unit prices on the first row).
    """
    if records is None or records.empty:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Aucune donnée"
        ws["A1"] = "Aucune semaine mémorisée pour la facturation."
        wb.save(out_path)
        return out_path

    # Ensure dates
    records = records.copy()
    records["date"] = pd.to_datetime(records["date"]).dt.date

    # Normalize site names (facturation):
    # - "24 ter" + "24 simple" -> "Internat" (business rule)
    # - any site containing "internat" -> "Internat" (legacy / variants)
    def _norm_site_name(s: str) -> str:
        s0 = str(s).strip()
        sN = _norm(s0)
        if sN in {"24 ter", "24 simple", "24ter", "24simple"}:
            return "Internat"
        # tolerate variants like "24 ter - internat" etc.
        if ("24" in sN) and ("ter" in sN) and ("internat" in sN):
            return "Internat"
        if ("24" in sN) and ("simple" in sN) and ("internat" in sN):
            return "Internat"
        if "internat" in sN:
            return "Internat"
        return s0
    if "site" in records.columns:
        records["site"] = records["site"].map(_norm_site_name)

    # Months present
    records["year"] = records["date"].map(lambda d: d.year)
    records["month"] = records["date"].map(lambda d: d.month)

    # Export a full year (Janvier -> Décembre). We pick the most recent year found in records.
    export_year = int(records["year"].max())

    # Keep a stable set of sites across the year (so columns don't change month to month)
    year_records = records[records["year"] == export_year]
    sites = sorted(year_records["site"].dropna().unique().tolist(), key=lambda s: str(s).upper())

    wb = openpyxl.Workbook()
    # remove default sheet
    wb.remove(wb.active)

    for m in range(1, 13):
        y = export_year
        # Keep the same naming convention as the UI selector (YYYY-MM)
        sheet_name = f"{y}-{m:02d}"
        ws = wb.create_sheet(sheet_name)

        month_start = dt.date(y, m, 1)
        # month end
        if m == 12:
            month_end = dt.date(y+1, 1, 1) - dt.timedelta(days=1)
        else:
            month_end = dt.date(y, m+1, 1) - dt.timedelta(days=1)

        # Month slice (sites are kept stable across the year)
        sub = records[(records["year"] == y) & (records["month"] == m)].copy()

# Build two blocks
        r = 1
        r = _write_month_block(ws, r, "Repas", "repas", sub, sites, month_start, month_end, custom_prices=custom_prices)
        r += 2
        _write_month_block(ws, r, "Mixé/Lissé", "mixe_lisse", sub, sites, month_start, month_end, custom_prices=custom_prices)

        # Column widths
        ws.column_dimensions["A"].width = 10
        for i, _ in enumerate(sites, start=2):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 14
        ws.column_dimensions[openpyxl.utils.get_column_letter(len(sites)+2)].width = 12

        # Freeze panes at first day row (keeps headers + day column visible)
        ws.freeze_panes = "B4"

    wb.save(out_path)
    return out_path


def _style_header_cell(cell):
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.fill = PatternFill("solid", fgColor="EDEDED")
    thin = Side(style="thin", color="999999")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def _style_cell(cell):
    cell.alignment = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="CCCCCC")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def _write_month_block(
    ws,
    start_row: int,
    title: str,
    category: str,
    sub: pd.DataFrame,
    sites: List[str],
    month_start: dt.date,
    month_end: dt.date,
    *,
    custom_prices: Optional[dict] = None,
) -> int:
    """
    Writes a block and returns the last row used.
    Block layout:
    Row start: title + unit prices + month_start date in last col
    Row start+1: year + site names + TOTAL
    Row start+2: (optional label row) left blank except category label in last col
    Rows start+3..: day numbers and quantities
    Last row: TOTAL
    """
    n_sites = len(sites)
    total_col = n_sites + 2  # A + sites + TOTAL

    # Row: title and unit prices
    ws.cell(start_row, 1, title)
    _style_header_cell(ws.cell(start_row, 1))
    for i, site in enumerate(sites, start=2):
        # Unit price per site (keeps the same structure as your existing workbook)
        c = ws.cell(start_row, i, _unit_price(category, site, custom_prices))
        _style_header_cell(c)
        c.number_format = "0.00"
    ws.cell(start_row, total_col, month_start)
    _style_header_cell(ws.cell(start_row, total_col))
    ws.cell(start_row, total_col).number_format = "yyyy-mm-dd"

    # Row: year + sites + TOTAL
    ws.cell(start_row+1, 1, month_start.year)
    _style_header_cell(ws.cell(start_row+1, 1))
    for i, site in enumerate(sites, start=2):
        c = ws.cell(start_row+1, i, site)
        _style_header_cell(c)
    ctot = ws.cell(start_row+1, total_col, "TOTAL")
    _style_header_cell(ctot)

    # Label row
    ws.cell(start_row+2, 1, "")
    for i in range(2, total_col):
        _style_header_cell(ws.cell(start_row+2, i))
    ws.cell(start_row+2, total_col, title.upper())
    _style_header_cell(ws.cell(start_row+2, total_col))

    # Data map
    cat = sub[sub["category"] == category].copy()
    cat["date"] = pd.to_datetime(cat["date"]).dt.date

    # Build a pivot day x site
    days = pd.date_range(month_start, month_end, freq="D").date
    pivot = pd.DataFrame(index=days, columns=sites, data=0)
    if not cat.empty:
        for _, row in cat.iterrows():
            d = row["date"]
            s = row["site"]
            if d in pivot.index and s in pivot.columns:
                pivot.loc[d, s] += int(row["qty"])

    # Write rows for each day
    row = start_row + 3
    for d in days:
        ws.cell(row, 1, d.day)
        _style_cell(ws.cell(row, 1))
        day_total = 0
        for i, site in enumerate(sites, start=2):
            v = int(pivot.loc[d, site])
            day_total += v
            cell = ws.cell(row, i, v)
            _style_cell(cell)
        ws.cell(row, total_col, day_total)
        _style_cell(ws.cell(row, total_col))
        row += 1

    # TOTAL row
    ws.cell(row, 1, "TOTAL")
    _style_header_cell(ws.cell(row, 1))
    grand_total = 0
    for i, site in enumerate(sites, start=2):
        v = int(pivot[site].sum())
        grand_total += v
        cell = ws.cell(row, i, v)
        _style_header_cell(cell)
    ws.cell(row, total_col, grand_total)
    _style_header_cell(ws.cell(row, total_col))

    return row
