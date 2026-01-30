from __future__ import annotations

import datetime as dt
import json
import re
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# -----------------------------
# Storage
# -----------------------------

def _data_dir() -> Path:
    """Local persistent folder (next to the app) to store weekly saved plannings."""
    base = Path(__file__).resolve().parent.parent
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
    return (str(s or "")).strip().lower()


def norm_site_facturation(site: str) -> str:
    """Business rule: '24 ter' + '24 simple' must be billed as a single column 'Internat'."""
    s0 = str(site or "").strip()
    sN = _norm(s0)

    if re.fullmatch(r"24\s*(ter|simple)", sN):
        return "Internat"
    if sN in {"24ter", "24simple"}:
        return "Internat"
    if "24" in sN and ("ter" in sN or "simple" in sN):
        return "Internat"

    return s0


def is_pdj_regime(regime: str) -> bool:
    r = _norm(regime)
    return ("pdj" in r) or ("petit" in r and "dej" in r) or ("gouter" in r) or ("goûter" in r)


def is_mixe_lisse_regime(regime: str) -> bool:
    r = _norm(regime)
    return ("mixe" in r) or ("mixé" in r) or ("lisse" in r) or ("m/l" in r) or ("ml" == r)


# -----------------------------
# Planning conversion
# -----------------------------

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

    mask = pd.Series([True] * len(df))
    if exclude_pdj and "Regime" in df.columns:
        mask &= ~df["Regime"].astype(str).map(is_pdj_regime)
    if exclude_mixe_lisse and "Regime" in df.columns:
        mask &= ~df["Regime"].astype(str).map(is_mixe_lisse_regime)

    df = df.loc[mask].copy()

    day_cols = [c for c in ["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"] if c in df.columns]

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
    def _norm_df(df: pd.DataFrame, qty_col: str, category: str) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame(columns=["date","site","category","qty","week_monday","source"])
        out = df.copy()
        out["date"] = pd.to_datetime(out["date"]).dt.date
        out["site"] = out["site"].astype(str).map(norm_site_facturation)
        out["category"] = category
        out["qty"] = pd.to_numeric(out[qty_col], errors="coerce").fillna(0).astype(int)
        out["week_monday"] = week_monday.isoformat()
        out["source"] = source_filename
        return out[["date","site","category","qty","week_monday","source"]]

    a = _norm_df(repas_daily, "qty_repas", "repas")
    b = _norm_df(ml_daily, "qty_ml", "mixe_lisse")
    new = pd.concat([a, b], ignore_index=True)

    p = _records_path()
    if p.exists():
        old = pd.read_csv(p, parse_dates=["date"])
        old["date"] = pd.to_datetime(old["date"]).dt.date
    else:
        old = pd.DataFrame(columns=["date","site","category","qty","week_monday","source"])

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
# Pricing
# -----------------------------

DEFAULT_UNIT_PRICES = {
    # Tarifs 2026
    "repas": {"__default__": 6.60, "MAS": 6.65},
    "mixe_lisse": {"__default__": 7.85, "MAS": 7.74},
}


def _unit_price(category: str, site: str, custom_prices: Optional[dict] = None) -> float:
    """Return unit price for a category/site.

    Note: site names coming from Excel can be inconsistent ("MAS", "M.A.S", "Mas", etc.).
    We therefore normalise them before matching the pricing grid.
    """

    def _key(x: str) -> str:
        x = (x or "").strip().upper()
        # Keep letters/digits only (so "M.A.S" -> "MAS")
        return re.sub(r"[^A-Z0-9]", "", x)

    prices_raw = (custom_prices or DEFAULT_UNIT_PRICES).get(category, {})
    # Normalise pricing keys once (so user custom dicts still work with punctuation/spaces)
    prices = {(_key(k) if k != "__default__" else "__default__"): v for k, v in prices_raw.items()}

    s = _key(site)
    return float(prices.get(s, prices.get("__default__", 0.0)))


def _write_tarifs_sheet(wb: openpyxl.Workbook, *, custom_prices: Optional[dict] = None) -> str:
    """Create/replace a 'TARIFS' sheet.

    This sheet is used by Excel formulas (SUMIFS + VLOOKUP) so the workbook remains
    "croisé" and recalculable even if the user edits quantities or prices afterwards.
    """
    name = "TARIFS"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws["A1"] = "clé"
    ws["B1"] = "prix_unitaire"
    _style_header_cell(ws["A1"])
    _style_header_cell(ws["B1"])

    grid = custom_prices or DEFAULT_UNIT_PRICES

    r = 2
    for category, d in grid.items():
        # Ensure default first, then specifics
        items = []
        if "__default__" in d:
            items.append(("__default__", d["__default__"]))
        for k in sorted([x for x in d.keys() if x != "__default__"], key=lambda x: str(x)):
            items.append((k, d[k]))

        for site, price in items:
            # We normalise the site key the same way as _unit_price
            sk = re.sub(r"[^A-Z0-9]", "", str(site).strip().upper()) if site != "__default__" else "__default__"
            ws.cell(r, 1, f"{category}|{sk}")
            ws.cell(r, 2, float(price))
            _style_cell(ws.cell(r, 1))
            c = ws.cell(r, 2)
            _style_cell(c)
            c.number_format = "0.00"
            r += 1

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 16
    ws.freeze_panes = "A2"
    return name


def _write_data_sheet(wb: openpyxl.Workbook, records: pd.DataFrame) -> str:
    """Create/replace a 'DONNEES' sheet containing the raw facturation lines.

    Columns:
      A: date (true Excel date)
      B: site
      C: category
      D: qty

    All calculations in the other sheets can be done via Excel formulas.
    """
    name = "DONNEES"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)

    headers = ["date", "site", "category", "qty"]
    for j, h in enumerate(headers, start=1):
        ws.cell(1, j, h)
        _style_header_cell(ws.cell(1, j))

    if records is None or records.empty:
        ws.freeze_panes = "A2"
        return name

    rec = records.copy()
    rec["date"] = pd.to_datetime(rec["date"]).dt.date
    if "site" in rec.columns:
        rec["site"] = rec["site"].map(norm_site_facturation)
    rec = rec[["date", "site", "category", "qty"]].sort_values(["date", "category", "site"]).reset_index(drop=True)

    for i, row in enumerate(rec.itertuples(index=False), start=2):
        ws.cell(i, 1, row.date)
        ws.cell(i, 2, str(row.site))
        ws.cell(i, 3, str(row.category))
        ws.cell(i, 4, int(row.qty))
        for j in range(1, 5):
            _style_cell(ws.cell(i, j))
        ws.cell(i, 1).number_format = "yyyy-mm-dd"

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 10
    ws.freeze_panes = "A2"
    return name


# -----------------------------
# Excel styling helpers
# -----------------------------

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


# -----------------------------
# Excel export (monthly workbook)
# -----------------------------

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
    Row start+2: label row
    Rows start+3..: day numbers and quantities
    Last row: TOTAL
    """
    n_sites = len(sites)
    total_col = n_sites + 2  # A + sites + TOTAL

    ws.cell(start_row, 1, title)
    _style_header_cell(ws.cell(start_row, 1))
    for i, site in enumerate(sites, start=2):
        c = ws.cell(start_row, i, _unit_price(category, site, custom_prices))
        _style_header_cell(c)
        c.number_format = "0.00"
    ws.cell(start_row, total_col, month_start)
    _style_header_cell(ws.cell(start_row, total_col))
    ws.cell(start_row, total_col).number_format = "yyyy-mm-dd"

    ws.cell(start_row + 1, 1, month_start.year)
    _style_header_cell(ws.cell(start_row + 1, 1))
    for i, site in enumerate(sites, start=2):
        c = ws.cell(start_row + 1, i, site)
        _style_header_cell(c)
    ctot = ws.cell(start_row + 1, total_col, "TOTAL")
    _style_header_cell(ctot)

    ws.cell(start_row + 2, 1, "")
    for i in range(2, total_col):
        _style_header_cell(ws.cell(start_row + 2, i))
    ws.cell(start_row + 2, total_col, title.upper())
    _style_header_cell(ws.cell(start_row + 2, total_col))

    # --- Dynamic cross table ---
    # Quantities are computed by Excel (SUMIFS) from the raw 'DONNEES' sheet.
    # That way, users can edit quantities afterwards and totals update automatically.
    days = pd.date_range(month_start, month_end, freq="D").date

    month_anchor = f"{get_column_letter(total_col)}{start_row}"  # contains month_start date

    row = start_row + 3
    for d in days:
        ws.cell(row, 1, d.day)
        _style_cell(ws.cell(row, 1))

        for i, site in enumerate(sites, start=2):
            # date = DATE(YEAR($anchor), MONTH($anchor), $Arow)
            # qty  = SUMIFS(DONNEES!$D:$D, DONNEES!$A:$A, date, DONNEES!$B:$B, site, DONNEES!$C:$C, category)
            addr_day = f"$A{row}"
            date_expr = f"DATE(YEAR(${month_anchor}),MONTH(${month_anchor}),{addr_day})"
            formula = (
                f"=SUMIFS(DONNEES!$D:$D,"
                f"DONNEES!$A:$A,{date_expr},"
                f"DONNEES!$B:$B,\"{site}\","
                f"DONNEES!$C:$C,\"{category}\")"
            )
            cell = ws.cell(row, i, formula)
            _style_cell(cell)

        # Daily total = SUM of the site cells
        first_site = get_column_letter(2)
        last_site = get_column_letter(total_col - 1)
        ws.cell(row, total_col, f"=SUM({first_site}{row}:{last_site}{row})")
        _style_cell(ws.cell(row, total_col))
        row += 1

    # Totals row (per site + grand total)
    ws.cell(row, 1, "TOTAL")
    _style_header_cell(ws.cell(row, 1))
    for i, _ in enumerate(sites, start=2):
        col = get_column_letter(i)
        cell = ws.cell(row, i, f"=SUM({col}{start_row + 3}:{col}{row - 1})")
        _style_header_cell(cell)
    ws.cell(row, total_col, f"=SUM({get_column_letter(total_col)}{start_row + 3}:{get_column_letter(total_col)}{row - 1})")
    _style_header_cell(ws.cell(row, total_col))

    return row


def _write_recap_sheet(
    wb: openpyxl.Workbook,
    records: pd.DataFrame,
    sites: List[str],
    year: int,
    *,
    custom_prices: Optional[dict] = None,
) -> None:
    """Add a recap sheet with monthly billed totals per site (standard + mixé/lissé)."""
    ws = wb.create_sheet(f"RECAP {year}")

    n_sites = len(sites)
    total_col = n_sites + 2  # A + sites + TOTAL

    ws.cell(1, 1, f"Récapitulatif facturation {year}")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_col)
    _style_header_cell(ws.cell(1, 1))
    ws.row_dimensions[1].height = 22

    months = [
        "Janvier","Février","Mars","Avril","Mai","Juin",
        "Juillet","Août","Septembre","Octobre","Novembre","Décembre"
    ]

    def _block(start_row: int, label: str, category: str) -> int:
        ws.cell(start_row, 1, label)
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=total_col)
        _style_header_cell(ws.cell(start_row, 1))

        ws.cell(start_row + 1, 1, "Mois")
        _style_header_cell(ws.cell(start_row + 1, 1))
        for i, site in enumerate(sites, start=2):
            c = ws.cell(start_row + 1, i, site)
            _style_header_cell(c)
        ctot = ws.cell(start_row + 1, total_col, "TOTAL")
        _style_header_cell(ctot)

        r = start_row + 2
        for m in range(1, 13):
            ws.cell(r, 1, months[m - 1])
            _style_cell(ws.cell(r, 1))
            # Excel-driven recap (dynamic): amounts are computed with SUMIFS on DONNEES
            # and a VLOOKUP on TARIFS.
            row_total_cells = []

            start_date = f"DATE({year},{m},1)"
            # end_date = first day of next month (handles December)
            if m == 12:
                end_date = f"DATE({year + 1},1,1)"
            else:
                end_date = f"DATE({year},{m + 1},1)"

            for i, site in enumerate(sites, start=2):
                # qty = SUMIFS(DONNEES!$D:$D, DONNEES!$A:$A,">="&start, DONNEES!$A:$A,"<"&end, DONNEES!$B:$B,site, DONNEES!$C:$C,category)
                qty_expr = (
                    f"SUMIFS(DONNEES!$D:$D,"
                    f"DONNEES!$A:$A,\">=\"&{start_date},"
                    f"DONNEES!$A:$A,\"<\"&{end_date},"
                    f"DONNEES!$B:$B,\"{site}\","
                    f"DONNEES!$C:$C,\"{category}\")"
                )

                # price = IFERROR(VLOOKUP(category|site, TARIFS!$A$2:$B$200,2,FALSE), VLOOKUP(category|__default__, ...))
                sk = re.sub(r"[^A-Z0-9]", "", str(site).strip().upper())
                price_expr = (
                    f"IFERROR(VLOOKUP(\"{category}|{sk}\",TARIFS!$A$2:$B$200,2,FALSE),"
                    f"VLOOKUP(\"{category}|__default__\",TARIFS!$A$2:$B$200,2,FALSE))"
                )

                formula = f"=({qty_expr})*({price_expr})"
                cell = ws.cell(r, i, formula)
                _style_cell(cell)
                cell.number_format = '#,##0.00" €"'
                row_total_cells.append(f"{get_column_letter(i)}{r}")

            cell_total = ws.cell(r, total_col, f"=SUM({row_total_cells[0]}:{row_total_cells[-1]})" if row_total_cells else 0)
            _style_cell(cell_total)
            cell_total.number_format = '#,##0.00" €"'
            r += 1

        ws.cell(r, 1, "TOTAL ANNUEL")
        _style_header_cell(ws.cell(r, 1))

        # Annual totals = sum of the 12 monthly lines above
        for i, _ in enumerate(sites, start=2):
            col = get_column_letter(i)
            cell = ws.cell(r, i, f"=SUM({col}{start_row + 2}:{col}{r - 1})")
            _style_header_cell(cell)
            cell.number_format = '#,##0.00" €"'

        cell_total = ws.cell(r, total_col, f"=SUM({get_column_letter(2)}{r}:{get_column_letter(total_col - 1)}{r})")
        _style_header_cell(cell_total)
        cell_total.number_format = '#,##0.00" €"'

        return r + 2

    next_row = _block(3, "FACTURATION — REPAS STANDARD", "repas")
    _block(next_row, "FACTURATION — MIXÉ/LISSÉ", "mixe_lisse")

    ws.column_dimensions["A"].width = 16
    for i in range(2, total_col + 1):
        ws.column_dimensions[get_column_letter(i)].width = 18

    ws.freeze_panes = "B5"


def export_monthly_workbook(
    records: pd.DataFrame,
    out_path: str,
    *,
    custom_prices: Optional[dict] = None,
    year: Optional[int] = None,
) -> str:
    """
    Create an Excel workbook for billing.

    Workbook contains 12 sheets (YYYY-01 -> YYYY-12) for the selected year,
    plus a recap sheet at the end.
    """
    if records is None or records.empty:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Aucune donnée"
        ws["A1"] = "Aucune semaine mémorisée pour la facturation."
        wb.save(out_path)
        return out_path

    records = records.copy()
    records["date"] = pd.to_datetime(records["date"]).dt.date

    if "site" in records.columns:
        records["site"] = records["site"].map(norm_site_facturation)

    records["year"] = records["date"].map(lambda d: d.year)
    records["month"] = records["date"].map(lambda d: d.month)

    export_year = int(year) if year is not None else int(records["year"].max())
    year_records = records[records["year"] == export_year]
    sites = sorted(year_records["site"].dropna().unique().tolist(), key=lambda s: str(s).upper())

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Raw data + pricing grid used by Excel formulas (keeps the workbook dynamic)
    _write_data_sheet(wb, records)
    _write_tarifs_sheet(wb, custom_prices=custom_prices)

    for m in range(1, 13):
        y = export_year
        sheet_name = f"{y}-{m:02d}"
        ws = wb.create_sheet(sheet_name)

        month_start = dt.date(y, m, 1)
        if m == 12:
            month_end = dt.date(y + 1, 1, 1) - dt.timedelta(days=1)
        else:
            month_end = dt.date(y, m + 1, 1) - dt.timedelta(days=1)

        sub = records[(records["year"] == y) & (records["month"] == m)].copy()

        r = 1
        r = _write_month_block(ws, r, "Repas", "repas", sub, sites, month_start, month_end, custom_prices=custom_prices)
        r += 2
        _write_month_block(ws, r, "Mixé/Lissé", "mixe_lisse", sub, sites, month_start, month_end, custom_prices=custom_prices)

        ws.column_dimensions["A"].width = 10
        for i, _ in enumerate(sites, start=2):
            ws.column_dimensions[get_column_letter(i)].width = 14
        ws.column_dimensions[get_column_letter(len(sites) + 2)].width = 12

        ws.freeze_panes = "B4"

    _write_recap_sheet(wb, records, sites, export_year, custom_prices=custom_prices)

    wb.save(out_path)
    return out_path
