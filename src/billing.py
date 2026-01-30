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
    prices = (custom_prices or DEFAULT_UNIT_PRICES).get(category, {})
    s = (site or "").strip().upper()
    return float(prices.get(s, prices.get("__default__", 0.0)))


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

    # --- Remplissage des quantités (éditables) ---
    # On écrit des valeurs initiales (issues de la mémoire) mais on met les totaux en formules,
    # afin que le classeur devienne la source de vérité après correction manuelle.

    cat = sub[sub["category"] == category].copy()
    cat["date"] = pd.to_datetime(cat["date"]).dt.date

    days = pd.date_range(month_start, month_end, freq="D").date
    pivot = pd.DataFrame(index=days, columns=sites, data=0)
    if not cat.empty:
        for _, row0 in cat.iterrows():
            d = row0["date"]
            s = row0["site"]
            if d in pivot.index and s in pivot.columns:
                pivot.loc[d, s] += int(row0["qty"])

    first_day_row = start_row + 3
    row = first_day_row
    for d in days:
        ws.cell(row, 1, d.day)
        _style_cell(ws.cell(row, 1))
        for i, site in enumerate(sites, start=2):
            v = int(pivot.loc[d, site])
            cell = ws.cell(row, i, v)  # valeur éditable
            _style_cell(cell)

        # TOTAL jour = somme des sites (formule)
        left = get_column_letter(2)
        right = get_column_letter(total_col - 1)
        f = f"=SUM({left}{row}:{right}{row})"
        ws.cell(row, total_col, f)
        _style_cell(ws.cell(row, total_col))
        row += 1

    # TOTAL mois
    total_row = row
    ws.cell(total_row, 1, "TOTAL")
    _style_header_cell(ws.cell(total_row, 1))

    for i in range(2, total_col):
        col_letter = get_column_letter(i)
        f = f"=SUM({col_letter}{first_day_row}:{col_letter}{total_row-1})"
        cell = ws.cell(total_row, i, f)
        _style_header_cell(cell)

    # Total global = somme des totaux sites
    left = get_column_letter(2)
    right = get_column_letter(total_col - 1)
    ws.cell(total_row, total_col, f"=SUM({left}{total_row}:{right}{total_row})")
    _style_header_cell(ws.cell(total_row, total_col))

    return total_row


def import_billing_workbook(
    xlsx_path: str,
    *,
    replace_dates: bool = True,
) -> Tuple[int, int]:
    """Importe un classeur Facturation.xlsx corrigé et met à jour la mémoire (records.csv).

    Le classeur contient des feuilles mensuelles nommées 'YYYY-MM'.
    On relit les blocs 'Repas' et 'Mixé/Lissé' (quantités par jour et par site).

    Si replace_dates=True, on remplace toutes les lignes existantes pour les dates présentes
    dans le classeur (catégories repas + mixé/lissé).

    Retourne (n_repas, n_ml) : nombre de lignes jour/site importées.
    """
    # On charge 2 fois :
    # - data_only=False pour lire la structure (et les dates en en-tête)
    # - data_only=True pour récupérer les valeurs calculées si Excel a déjà évalué les formules
    wb = openpyxl.load_workbook(xlsx_path, data_only=False)
    wb_values = openpyxl.load_workbook(xlsx_path, data_only=True)

    # Collecte des enregistrements
    rows = []
    n_repas = 0
    n_ml = 0
    imported_dates = set()

    def _parse_month_sheet(ws) -> None:
        nonlocal n_repas, n_ml
        ws_val = wb_values[ws.title]

        # Sites sont sur la ligne start_row+1 (row 2 du bloc) colonnes B.. jusqu'à 'TOTAL'
        def _read_sites(header_row: int) -> Tuple[List[str], int]:
            sites_local = []
            col = 2
            while True:
                v = ws.cell(header_row, col).value
                if v is None:
                    break
                if str(v).strip().upper() == "TOTAL":
                    break
                sites_local.append(str(v).strip())
                col += 1
            total_col_local = col
            return sites_local, total_col_local

        def _parse_block(start_row: int, category: str) -> None:
            # start_row layout per export: start_row title, start_row+1 header with sites, start_row+3 first day row
            sites_local, total_col_local = _read_sites(start_row + 1)
            if not sites_local:
                return

            # month_start date stored in title row last col (TOTAL col)
            month_start_val = ws.cell(start_row, total_col_local).value
            if isinstance(month_start_val, dt.datetime):
                month_start = month_start_val.date()
            elif isinstance(month_start_val, dt.date):
                month_start = month_start_val
            else:
                # fallback: infer from sheet name
                m = re.match(r"^(\d{4})-(\d{2})$", ws.title.strip())
                if not m:
                    return
                month_start = dt.date(int(m.group(1)), int(m.group(2)), 1)

            # iterate day rows until TOTAL label in column A
            r = start_row + 3
            while True:
                a = ws.cell(r, 1).value
                if a is None:
                    break
                if str(a).strip().upper() == "TOTAL":
                    break
                try:
                    day = int(a)
                except Exception:
                    break
                try:
                    d = dt.date(month_start.year, month_start.month, day)
                except Exception:
                    r += 1
                    continue

                for i, site in enumerate(sites_local, start=2):
                    cell = ws.cell(r, i)
                    v = cell.value
                    if isinstance(v, str) and v.strip().startswith("="):
                        v = ws_val.cell(r, i).value
                    # si formule, on garde la formule (mais on ne peut pas recalculer ici) -> on lit la valeur si possible
                    # on privilégie un nombre direct
                    qty = 0
                    if isinstance(v, (int, float)):
                        qty = int(v)
                    else:
                        # si c'est une formule, on essaye data_only via second load? (trop coûteux) => 0
                        try:
                            qty = int(float(v))
                        except Exception:
                            qty = 0

                    rows.append({
                        "date": d,
                        "site": norm_site_facturation(site),
                        "category": category,
                        "qty": qty,
                        "week_monday": (d - dt.timedelta(days=d.weekday())).isoformat(),
                        "source": f"import:{Path(xlsx_path).name}",
                    })

                    if category == "repas":
                        n_repas += 1
                    else:
                        n_ml += 1

                imported_dates.add(d)
                r += 1

        # Repas block at row 1, mixé/lissé after it
        # We need to compute where the second block starts: depends on number of days.
        # Find it by searching for title 'Mixé/Lissé' in col A.
        repas_start = 1
        _parse_block(repas_start, "repas")

        mixe_start = None
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, 1).value
            if v is None:
                continue
            if _norm(v) == _norm("Mixé/Lissé"):
                mixe_start = r
                break
        if mixe_start:
            _parse_block(mixe_start, "mixe_lisse")

    # Parse monthly sheets
    month_re = re.compile(r"^\d{4}-\d{2}$")
    for name in wb.sheetnames:
        if month_re.match(name.strip()):
            _parse_month_sheet(wb[name])

    if not rows:
        return (0, 0)

    new = pd.DataFrame(rows)
    new["date"] = pd.to_datetime(new["date"]).dt.date
    new["qty"] = pd.to_numeric(new["qty"], errors="coerce").fillna(0).astype(int)

    p = _records_path()
    if p.exists():
        old = pd.read_csv(p, parse_dates=["date"])
        old["date"] = pd.to_datetime(old["date"]).dt.date
    else:
        old = pd.DataFrame(columns=["date","site","category","qty","week_monday","source"])

    if replace_dates and imported_dates:
        old = old.loc[~(old["date"].isin(list(imported_dates)) & old["category"].isin(["repas","mixe_lisse"]))].copy()

    merged = pd.concat([old, new], ignore_index=True)
    merged.to_csv(p, index=False)

    # meta: track weeks
    meta = _read_meta()
    meta.setdefault("saved_weeks", [])
    for d in sorted(imported_dates):
        w = (d - dt.timedelta(days=d.weekday())).isoformat()
        if w not in meta["saved_weeks"]:
            meta["saved_weeks"].append(w)
    meta["saved_weeks"] = sorted(set(meta["saved_weeks"]))
    _write_meta(meta)

    return (int(n_repas), int(n_ml))


def _write_recap_sheet(
    wb: openpyxl.Workbook,
    records: pd.DataFrame,
    sites: List[str],
    year: int,
    *,
    custom_prices: Optional[dict] = None,
) -> None:
    """Add a recap sheet with monthly billed totals per site.

    ⚠️ Important: le récap est généré en **formules Excel** qui pointent vers
    les feuilles mensuelles. Ainsi, si l'utilisateur corrige les effectifs dans
    le classeur, le récap se met à jour sans régénérer côté code.
    """
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
            sheet_name = f"{year}-{m:02d}"
            ws.cell(r, 1, months[m - 1])
            _style_cell(ws.cell(r, 1))

            # lignes TOTAL dans la feuille mensuelle
            month_start = dt.date(year, m, 1)
            if m == 12:
                month_end = dt.date(year + 1, 1, 1) - dt.timedelta(days=1)
            else:
                month_end = dt.date(year, m + 1, 1) - dt.timedelta(days=1)
            n_days = (month_end - month_start).days + 1

            if category == "repas":
                block_start = 1
            else:
                # bloc 2 commence 2 lignes après la fin du bloc 1
                block1_total_row = 1 + 3 + n_days
                block_start = block1_total_row + 2
            total_row_in_month = block_start + 3 + n_days  # ligne "TOTAL" du bloc

            # Pour chaque site: Montant = Qté totale du mois * prix unitaire (ligne titre)
            for i, site in enumerate(sites, start=2):
                col_letter = get_column_letter(i)
                qty_cell = f"'{sheet_name}'!{col_letter}{total_row_in_month}"
                price_cell = f"'{sheet_name}'!{col_letter}{block_start}"  # prix unitaire en ligne titre
                formula = f"={qty_cell}*{price_cell}"
                cell = ws.cell(r, i, formula)
                _style_cell(cell)
                cell.number_format = '#,##0.00" €"'

            # TOTAL ligne = somme des sites
            left = get_column_letter(2)
            right = get_column_letter(total_col - 1)
            cell_total = ws.cell(r, total_col, f"=SUM({left}{r}:{right}{r})")
            _style_cell(cell_total)
            cell_total.number_format = '#,##0.00" €"'
            r += 1

        # TOTAL annuel
        ws.cell(r, 1, "TOTAL ANNUEL")
        _style_header_cell(ws.cell(r, 1))
        for i in range(2, total_col):
            col_letter = get_column_letter(i)
            cell = ws.cell(r, i, f"=SUM({col_letter}{start_row+2}:{col_letter}{r-1})")
            _style_header_cell(cell)
            cell.number_format = '#,##0.00" €"'

        ws.cell(r, total_col, f"=SUM({get_column_letter(2)}{r}:{get_column_letter(total_col-1)}{r})")
        _style_header_cell(ws.cell(r, total_col))
        ws.cell(r, total_col).number_format = '#,##0.00" €"'

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
) -> str:
    """
    Create an Excel workbook for billing.

    Workbook contains 12 sheets (YYYY-01 -> YYYY-12) for the most recent year found in records,
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

    export_year = int(records["year"].max())
    year_records = records[records["year"] == export_year]
    sites = sorted(year_records["site"].dropna().unique().tolist(), key=lambda s: str(s).upper())

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

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


# -----------------------------
# Excel import (monthly workbook corrected by user)
# -----------------------------

def _iter_month_sheets(wb: openpyxl.Workbook) -> List[Tuple[int,int,openpyxl.worksheet.worksheet.Worksheet]]:
    """Yield (year, month, worksheet) for sheets named YYYY-MM."""
    out = []
    for ws in wb.worksheets:
        m = re.match(r"^(\d{4})-(\d{2})$", str(ws.title).strip())
        if not m:
            continue
        y = int(m.group(1))
        mo = int(m.group(2))
        out.append((y, mo, ws))
    out.sort(key=lambda t: (t[0], t[1]))
    return out


def _parse_block(ws, title: str) -> Tuple[dt.date, List[str], List[Tuple[dt.date, str, str, int]]]:
    """Parse a block ('Repas' or 'Mixé/Lissé') and return rows as (date, site, category, qty)."""
    # find title in column A
    start_row = None
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, 1).value).strip().lower() == title.strip().lower():
            start_row = r
            break
    if start_row is None:
        return None, [], []

    # Sites on row start_row+1, columns B.. until 'TOTAL'
    sites = []
    c = 2
    while c <= ws.max_column:
        v = ws.cell(start_row + 1, c).value
        if v is None:
            c += 1
            continue
        sv = str(v).strip()
        if sv.upper() == "TOTAL":
            break
        sites.append(sv)
        c += 1

    # Month start date is stored in last column (TOTAL col) on start_row
    total_col = 2 + len(sites)
    month_start_val = ws.cell(start_row, total_col).value
    if isinstance(month_start_val, dt.datetime):
        month_start = month_start_val.date()
    elif isinstance(month_start_val, dt.date):
        month_start = month_start_val
    else:
        # fallback from sheet name elsewhere handled by caller
        month_start = None

    # Daily rows start at start_row+3 until 'TOTAL' in col A
    rows = []
    r = start_row + 3
    while r <= ws.max_row:
        vday = ws.cell(r, 1).value
        if vday is None:
            r += 1
            continue
        if str(vday).strip().upper() == "TOTAL":
            break
        # day number expected
        try:
            day_num = int(vday)
        except Exception:
            r += 1
            continue
        if month_start is None:
            r += 1
            continue
        d = dt.date(month_start.year, month_start.month, day_num)
        category = "repas" if title.strip().lower().startswith("repas") else "mixe_lisse"
        for i, site in enumerate(sites, start=2):
            qty_val = ws.cell(r, i).value
            try:
                qty = int(qty_val) if qty_val is not None else 0
            except Exception:
                qty = 0
            rows.append((d, norm_site_facturation(site), category, qty))
        r += 1
    return month_start, sites, rows


def import_billing_workbook(
    in_path: str,
    *,
    replace_dates: bool = True,
) -> Tuple[int, int]:
    """Import a corrected Facturation.xlsx workbook and update records.csv.

    - Reads sheets named YYYY-MM.
    - Parses the two blocks: Repas + Mixé/Lissé.
    - If replace_dates=True: for any imported date, existing records for those dates/categories/sites are removed and replaced.
    Returns: (n_replaced, n_added)
    """
    wb = openpyxl.load_workbook(in_path, data_only=False)
    imported = []

    for y, m, ws in _iter_month_sheets(wb):
        # parse both blocks
        ms1, sites1, rows1 = _parse_block(ws, "Repas")
        ms2, sites2, rows2 = _parse_block(ws, "Mixé/Lissé")

        # If month_start missing, build from sheet name and re-parse day dates
        if ms1 is None:
            # We can still infer month_start
            month_start = dt.date(y, m, 1)
            # patch rows dates
            fixed = []
            for (d, site, cat, qty) in rows1:
                fixed.append((dt.date(y, m, d.day), site, cat, qty))
            rows1 = fixed
        if ms2 is None:
            fixed = []
            for (d, site, cat, qty) in rows2:
                fixed.append((dt.date(y, m, d.day), site, cat, qty))
            rows2 = fixed

        imported.extend(rows1)
        imported.extend(rows2)

    if not imported:
        return 0, 0

    # Build dataframe
    df_new = pd.DataFrame(imported, columns=["date","site","category","qty"])
    df_new["date"] = pd.to_datetime(df_new["date"]).dt.date
    df_new["qty"] = pd.to_numeric(df_new["qty"], errors="coerce").fillna(0).astype(int)
    df_new["week_monday"] = df_new["date"].map(lambda d: (d - dt.timedelta(days=d.weekday())))
    df_new["source"] = "import_facturation"

    # Load existing
    df_old = load_records()
    if df_old is None or df_old.empty:
        df_out = df_new.copy()
        df_out.to_csv(_records_path(), index=False)
        return 0, int(len(df_new))

    # Normalize
    df_old = df_old.copy()
    df_old["date"] = pd.to_datetime(df_old["date"]).dt.date
    if "site" in df_old.columns:
        df_old["site"] = df_old["site"].map(norm_site_facturation)
    if "category" in df_old.columns:
        df_old["category"] = df_old["category"].astype(str)
    df_old["qty"] = pd.to_numeric(df_old["qty"], errors="coerce").fillna(0).astype(int)

    # Replace dates present in import
    if replace_dates:
        key_dates = set(df_new["date"].unique().tolist())
        mask_keep = ~df_old["date"].isin(key_dates)
        n_removed = int((~mask_keep).sum())
        df_out = pd.concat([df_old.loc[mask_keep], df_new], ignore_index=True)
    else:
        n_removed = 0
        df_out = pd.concat([df_old, df_new], ignore_index=True)

    df_out.to_csv(_records_path(), index=False)
    return n_removed, int(len(df_new))
