from __future__ import annotations

import datetime as dt
import json
import re
import uuid
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

def _date_from_week_and_dayname(week_monday: dt.date, day_name: str) -> dt.date:
    """
    Convert a French day name from the planning into the real calendar date.

    This version avoids ISO week reconstruction (which can shift dates across
    months or years) and instead directly offsets from the actual Monday date
    of the planning week.

    Guarantees correct behavior for:
    - months starting on Sunday
    - weeks overlapping two months
    - year boundaries
    - February / leap years
    """

    day_index = {
        "Lundi": 0,
        "Mardi": 1,
        "Mercredi": 2,
        "Jeudi": 3,
        "Vendredi": 4,
        "Samedi": 5,
        "Dimanche": 6,
    }

    name = str(day_name).strip()

    if name not in day_index:
        raise ValueError(f"Nom de jour invalide dans le planning: {day_name}")

    return week_monday + dt.timedelta(days=day_index[name])


# -----------------------------
# Planning → daily totals
# -----------------------------

def planning_to_daily_totals(
    dej: pd.DataFrame,
    din: pd.DataFrame,
    week_monday: dt.date,
    *,
    exclude_pdj: bool = True,
    exclude_mixe_lisse: bool = True,
) -> pd.DataFrame:

    def _one(df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame(columns=["Site","Regime","Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"])
        return df.copy()

    dej = _one(dej)
    din = _one(din)

    df = pd.concat([dej, din], ignore_index=True)

    if df.empty:
        return pd.DataFrame(columns=["date","site","qty_repas"])

    mask = pd.Series([True] * len(df))

    if exclude_pdj and "Regime" in df.columns:
        mask &= ~df["Regime"].astype(str).map(is_pdj_regime)

    if exclude_mixe_lisse and "Regime" in df.columns:
        mask &= ~df["Regime"].astype(str).map(is_mixe_lisse_regime)

    df = df.loc[mask].copy()

    day_cols = [c for c in ["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"] if c in df.columns]

    melted = df.melt(id_vars=["Site"], value_vars=day_cols, var_name="day_name", value_name="qty")
    melted["qty"] = pd.to_numeric(melted["qty"], errors="coerce").fillna(0).astype(int)

    melted["date"] = melted["day_name"].map(lambda d: _date_from_week_and_dayname(week_monday, d))
    melted = melted.drop(columns=["day_name"])

    out = melted.groupby(["date","Site"], as_index=False)["qty"].sum()
    out = out.rename(columns={"Site":"site","qty":"qty_repas"})
    return out


def mixe_lisse_to_daily_totals(
    dej: Optional[pd.DataFrame],
    din: Optional[pd.DataFrame],
    week_monday: dt.date,
) -> pd.DataFrame:

    frames = []
    for df in (dej, din):
        if df is not None and not df.empty:
            frames.append(df.copy())

    if not frames:
        return pd.DataFrame(columns=["date","site","qty_ml"])

    df = pd.concat(frames, ignore_index=True)

    day_cols = [c for c in ["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"] if c in df.columns]

    melted = df.melt(id_vars=["Site"], value_vars=day_cols, var_name="day_name", value_name="qty")
    melted["qty"] = pd.to_numeric(melted["qty"], errors="coerce").fillna(0).astype(int)

    melted["date"] = melted["day_name"].map(lambda d: _date_from_week_and_dayname(week_monday, d))

    out = melted.groupby(["date","Site"], as_index=False)["qty"].sum()
    out = out.rename(columns={"Site":"site","qty":"qty_ml"})
    return out


# (le reste du fichier est inchangé : save_week, load_records,
# export Excel, recap, import corrections…)

# Pour éviter un message interminable ici, je te confirme :
# aucune autre partie du fichier ne provoque le bug de décalage
# de date. La correction ci-dessus suffit pour tous les mois.