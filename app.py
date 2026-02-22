# --- Persistence Layer (stable Streamlit Cloud) ---
import sys
import json
import os
import streamlit as st

# Toujours enregistrer le fichier au même endroit que app.py
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "data.json")

def save_data(data):
    try:
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception:
        pass  # évite tout affichage d'erreur Streamlit

def load_data():
    try:
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}

import pandas as pd
import datetime as dt
import importlib
import os
import shutil
import traceback
from pathlib import Path
def _purge_bytecode(root: Path) -> None:
    """Supprime __pycache__ et *.pyc sous root (best-effort)."""
    try:
        for p in root.rglob('__pycache__'):
            try:
                shutil.rmtree(p, ignore_errors=True)
            except Exception:
                pass
        for p in root.rglob('*.pyc'):
            try:
                p.unlink(missing_ok=True)
            except Exception:
                pass
    except Exception:
        pass




# -----------------------------
# Helpers import robustes
# -----------------------------
def _purge_src_modules() -> None:
    """Supprime du cache Python tous les modules commençant par 'src.' (et 'src')."""
    to_del = [k for k in list(sys.modules.keys()) if k == "src" or k.startswith("src.")]
    for k in to_del:
        sys.modules.pop(k, None)


def _import_or_stop():
    """Importe les modules src.* de façon robuste ; stop l'app si erreur."""
    try:
        # Purge bytecode + caches import avant tout
        sys.dont_write_bytecode = True
        importlib.invalidate_caches()
        _purge_bytecode(Path(__file__).resolve().parent)

        _purge_src_modules()

        # Import processor
        processor = importlib.import_module("src.processor")
        importlib.reload(processor)

        # Import config_store (puis reload)
        cs = importlib.import_module("src.config_store")
        importlib.reload(cs)

        # Vérification explicite (sinon erreur claire)
        if not hasattr(cs, "ConfigStore"):
            raise ImportError(
                f"src.config_store importé depuis {getattr(cs, '__file__', '<?>')} "
                f"mais ConfigStore est introuvable. Attributs: "
                f"{', '.join([k for k in dir(cs) if not k.startswith('__')])}"
            )

        # Import order_forms
        order_forms = importlib.import_module("src.order_forms")
        importlib.reload(order_forms)

        # Import bon_commande (BC)
        bon_commande = importlib.import_module("src.bon_commande")
        importlib.reload(bon_commande)

        # Import billing
        billing = importlib.import_module("src.billing")
        importlib.reload(billing)

        # Import facturation PDJ
        pdj_billing = importlib.import_module("src.pdj_billing")
        importlib.reload(pdj_billing)

        # Import parseur PDJ (lecture bons PDF/Excel + OCR best-effort)
        # Optionnel : certaines plateformes n'ont pas PyMuPDF (import fitz).
        # Si absent, l'app doit continuer à fonctionner (saisie manuelle toujours possible).
        pdj_facturation = None
        try:
            pdj_facturation = importlib.import_module("src.pdj_facturation")
            importlib.reload(pdj_facturation)
        except ModuleNotFoundError as e:
            # Ne bloque pas le démarrage si fitz/PyMuPDF manque.
            if "fitz" not in str(e):
                raise

        # Import allergènes
        learner = importlib.import_module("src.allergens.learner")
        importlib.reload(learner)

        generator = importlib.import_module("src.allergens.generator")
        importlib.reload(generator)

        return processor, cs, order_forms, billing, pdj_billing, pdj_facturation, learner, generator, bon_commande

    except Exception as e:
        st.error("💥 Erreur lors d’un import (module src.*)")
        st.code(repr(e))
        st.code(traceback.format_exc())
        st.stop()


processor, cs, order_forms, billing, pdj_billing, pdj_facturation, learner, generator, bon_commande = _import_or_stop()


# Exports processor
parse_planning_fabrication = processor.parse_planning_fabrication
parse_planning_mixe_lisse = processor.parse_planning_mixe_lisse
make_production_summary = processor.make_production_summary
make_production_pivot = processor.make_production_pivot
parse_menu = processor.parse_menu
build_bon_commande = bon_commande.build_bon_commande
export_excel = processor.export_excel
export_bons_livraison_pdf = processor.export_bons_livraison_pdf

# ConfigStore
ConfigStore = cs.ConfigStore

# order_forms
export_orders_per_supplier_excel = order_forms.export_orders_per_supplier_excel
export_orders_per_supplier_pdf = order_forms.export_orders_per_supplier_pdf

# billing
planning_to_daily_totals = billing.planning_to_daily_totals
mixe_lisse_to_daily_totals = billing.mixe_lisse_to_daily_totals
save_week = billing.save_week
load_records = billing.load_records
export_monthly_workbook = billing.export_monthly_workbook
apply_corrected_monthly_workbook = billing.apply_corrected_monthly_workbook
delete_billing_records = getattr(billing, "delete_billing_records", None)

# pdj_billing
pdj_default_products = pdj_billing.DEFAULT_PRODUCTS
pdj_load_records = pdj_billing.load_pdj_records
pdj_add_records = pdj_billing.add_pdj_records
pdj_load_prices = pdj_billing.load_unit_prices
pdj_save_prices = pdj_billing.save_unit_prices
pdj_add_money_adjustments = pdj_billing.add_money_adjustments
pdj_load_money_adjustments = pdj_billing.load_money_adjustments
pdj_export_monthly_workbook = pdj_billing.export_monthly_pdj_workbook
pdj_delete_records = getattr(pdj_billing, "delete_pdj_records", None)
pdj_delete_money_adjustments = getattr(pdj_billing, "delete_money_adjustments", None)

# allergènes
learn_from_filled_allergen_workbook = learner.learn_from_filled_allergen_workbook
generate_allergen_workbook = generator.generate_allergen_workbook


DAY_NAMES = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]


def format_pivot_for_display(piv: pd.DataFrame) -> pd.DataFrame:
    """Affichage: Régimes en lignes, jours en colonnes, + Totaux."""
    if piv is None or piv.empty:
        return piv
    df = piv.copy()

    cols = ["Regime"] + [d for d in DAY_NAMES if d in df.columns]
    if "Total" in df.columns:
        cols.append("Total")
    df = df[[c for c in cols if c in df.columns]]

    if "Total" in df.columns:
        df = df.rename(columns={"Total": "Total semaine"})

    if "Regime" in df.columns:
        df["Regime"] = df["Regime"].replace({"TOTAL": "TOTAL JOUR"})
    return df


def set_background():
    import base64

    img = Path(__file__).parent / "assets" / "background.jpg"
    if not img.exists():
        return
    b64 = base64.b64encode(img.read_bytes()).decode("utf-8")
    css = """
    <style>
    [data-testid="stAppViewContainer"], .stApp {
        background:
            linear-gradient(rgba(255,255,255,0.65), rgba(255,255,255,0.65)),
            url("data:image/jpeg;base64,IMGDATA");
        background-repeat: no-repeat;
        background-position: center 90px;
        background-size: 420px auto;
        background-attachment: fixed;
    }
    </style>
    """.replace("IMGDATA", b64)
    st.markdown(css, unsafe_allow_html=True)


def _save_uploaded_file(uploaded, suffix: str) -> str:
    """Save an UploadedFile to a temp file and return path."""
    import tempfile
    import os

    fd, path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    with open(path, "wb") as f:
        f.write(uploaded.getbuffer())
    return path


def _temp_out_path(suffix: str) -> str:
    """Create a unique temp output path (cloud-safe) and return it."""
    import tempfile
    import os

    fd, path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    return path


def _read_excel_any(file_obj, sheet_name=None):
    """Lecture Excel robuste: .xlsx/.xlsm/.xls.

    pandas choisit généralement le bon moteur, mais on force xlrd pour .xls.
    """
    name = str(getattr(file_obj, "name", "") or "").lower()
    if name.endswith(".xls"):
        return pd.read_excel(file_obj, sheet_name=sheet_name, engine="xlrd")
    return pd.read_excel(file_obj, sheet_name=sheet_name)


st.set_page_config(page_title="Gestion cuisine centrale", layout="wide")
set_background()

st.title("Gestion cuisine centrale")

with st.sidebar:
    st.header("Fichiers")
    planning_file = st.file_uploader("Planning fabrication (.xlsx)", type=["xlsx","xlsm"])
    menu_file = st.file_uploader("Menu (.xlsx)", type=["xlsx","xlsm"])
    st.markdown("---")
    st.caption(
        "Conseil : utilise les fichiers d’origine (avec formules) ; l’app récupère les valeurs correctement."
    )

    st.markdown("---")
    st.header("Paramètres — Bons de commande")

    store = ConfigStore()
    coeffs = store.load_coefficients()
    units = store.load_units()
    suppliers = store.load_suppliers()

    with st.expander("Configurer coefficients / unités / fournisseurs", expanded=False):
        st.caption(
            "Ces listes sont mémorisées (JSON) et utilisées pour les menus déroulants dans le bon de commande."
        )

        ctab1, ctab2, ctab3 = st.tabs(["Coefficients", "Unités", "Fournisseurs"])

        with ctab1:
            dfc = pd.DataFrame(
                [{"name": c.name, "value": c.value, "default_unit": c.default_unit} for c in coeffs]
            )
            if dfc.empty:
                dfc = pd.DataFrame([{"name": "1", "value": 1.0, "default_unit": "unité"}])
            dfc_edit = st.data_editor(
                dfc,
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "name": st.column_config.TextColumn("Nom coefficient"),
                    "value": st.column_config.NumberColumn("Valeur", step=0.01),
                    "default_unit": st.column_config.TextColumn("Unité par défaut"),
                },
                key="cfg_coeffs",
            )
            if st.button("Enregistrer les coefficients", key="save_coeffs"):
                try:
                    store.save_coefficients(dfc_edit.to_dict("records"))
                    st.success("Coefficients enregistrés.")
                except Exception as e:
                    st.error("❌ Impossible d'enregistrer les coefficients (écriture disque).")
                    st.caption(f"Dossier config: {store.info().get('base_dir','')}")
                    st.code(repr(e))

        with ctab2:
            dfu = pd.DataFrame({"unit": units})
            dfu_edit = st.data_editor(
                dfu,
                use_container_width=True,
                num_rows="dynamic",
                column_config={"unit": st.column_config.TextColumn("Unité")},
                key="cfg_units",
            )
            if st.button("Enregistrer les unités", key="save_units"):
                try:
                    store.save_units([u for u in dfu_edit["unit"].astype(str).tolist() if u.strip()])
                    st.success("Unités enregistrées.")
                except Exception as e:
                    st.error("❌ Impossible d'enregistrer les unités (écriture disque).")
                    st.caption(f"Dossier config: {store.info().get('base_dir','')}")
                    st.code(repr(e))

        with ctab3:
            # ✅ IMPORTANT : forcer les colonnes même si suppliers est vide,
            # sinon st.data_editor peut être "bloqué" sur l'ajout.
            dfs = pd.DataFrame(
                [
                    {
                        "name": s.name,
                        "customer_code": s.customer_code,
                        "coord1": s.coord1,
                        "coord2": s.coord2,
                    }
                    for s in suppliers
                ],
                columns=["name", "customer_code", "coord1", "coord2"],
            )

            # ✅ si vide, on met 1 ligne "starter" editable
            if dfs.empty:
                dfs = pd.DataFrame(
                    [{"name": "", "customer_code": "", "coord1": "", "coord2": ""}],
                    columns=["name", "customer_code", "coord1", "coord2"],
                )

            dfs_edit = st.data_editor(
                dfs,
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "name": st.column_config.TextColumn("Fournisseur"),
                    "customer_code": st.column_config.TextColumn("Code client"),
                    "coord1": st.column_config.TextColumn("Coordonnée 1"),
                    "coord2": st.column_config.TextColumn("Coordonnée 2"),
                },
                key="cfg_suppliers",
            )
            if st.button("Enregistrer les fournisseurs", key="save_suppliers"):
                # ✅ On vire les lignes "vides" (sans nom)
                recs = dfs_edit.fillna("").to_dict("records")
                recs = [r for r in recs if str(r.get("name", "")).strip()]
                try:
                    store.save_suppliers(recs)
                    st.success("Fournisseurs enregistrés.")
                except Exception as e:
                    st.error("❌ Impossible d'enregistrer les fournisseurs (écriture disque).")
                    st.caption(f"Dossier config: {store.info().get('base_dir','')}")
                    st.code(repr(e))

if not planning_file or not menu_file:
    st.info("Charge le planning et le menu pour afficher les tableaux et générer les documents.")
    st.stop()

try:
    # ---- Préparation fichiers temporaires (cloud-safe) ----
    planning_path = _save_uploaded_file(planning_file, suffix=Path(getattr(planning_file, "name", "")).suffix or ".xlsx")
    menu_path = _save_uploaded_file(menu_file, suffix=Path(getattr(menu_file, "name", "")).suffix or ".xlsx")

    # ✅ Parse planning depuis le PATH (plus robuste sur Streamlit Cloud)
    planning = parse_planning_fabrication(planning_path)

    # Optionnel : feuille mixé/lissé (si présente)
    mix_planning = {"dejeuner": pd.DataFrame(), "diner": pd.DataFrame()}
    try:
        mix_planning = parse_planning_mixe_lisse(planning_path)
    except Exception:
        pass

    # Parse menu items
    menu_items = parse_menu(menu_path)

    # Production (format long + pivot)
    prod_dej_long = make_production_summary(planning["dejeuner"])
    prod_din_long = make_production_summary(planning["diner"])
    prod_dej_piv = make_production_pivot(planning["dejeuner"])
    prod_din_piv = make_production_pivot(planning["diner"])

    # ---- UI ----
    tab_prod, tab_bc, tab_bl, tab_factu, tab_factu_pdj, tab_all = st.tabs(
        [
            "Production (Déj / Dîn)",
            "Bon de commande",
            "Bons de livraison",
            "Facturation mensuelle",
            "Facturation PDJ",
            "Allergènes",
        ]
    )

    with tab_prod:
        c1, c2 = st.columns(2)

        with c1:
            st.subheader("Déjeuner — tableau")
            st.dataframe(
                format_pivot_for_display(prod_dej_piv),
                use_container_width=True,
                hide_index=True,
            )

        with c2:
            st.subheader("Dîner — tableau")
            st.dataframe(
                format_pivot_for_display(prod_din_piv),
                use_container_width=True,
                hide_index=True,
            )

        st.markdown("### Graphe (totaux par jour)")

        def _totaux_jour(piv: pd.DataFrame) -> pd.Series:
            day_cols = [c for c in DAY_NAMES if c in piv.columns]
            if (
                piv is not None
                and not piv.empty
                and ("Regime" in piv.columns)
                and (piv["Regime"] == "TOTAL JOUR").any()
            ):
                row = piv[piv["Regime"] == "TOTAL JOUR"].iloc[0]
                return row[day_cols]
            if day_cols:
                return piv[day_cols].sum(numeric_only=True)
            return pd.Series(dtype=float)

        tot_dej = _totaux_jour(format_pivot_for_display(prod_dej_piv))
        tot_din = _totaux_jour(format_pivot_for_display(prod_din_piv))

        chart_df = pd.DataFrame({"Déjeuner": tot_dej, "Dîner": tot_din})
        st.bar_chart(chart_df)

        with st.expander("Comment est construit le graphe ?"):
            st.markdown(
                """Le graphe représente **les totaux par jour**.

- On prend le tableau Déjeuner (resp. Dîner).
- On récupère la ligne **TOTAL JOUR** (ou à défaut on additionne toutes les lignes régime).
- On trace une barre par jour, avec 2 séries : **Déjeuner** et **Dîner**."""
            )

    with tab_bc:
        st.subheader("Bon de commande")
        bon = build_bon_commande(planning, menu_items)
        st.caption(
            "Tu peux **fusionner/renommer des lignes** en modifiant la colonne *Libellé* "
            "(elles seront regroupées au moment des bons par fournisseur)."
        )

        bon_edit = st.data_editor(
            bon,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key="bc_editor",
        )

        store2 = ConfigStore()
        coeffs2 = [
            {"name": c.name, "value": float(c.value), "default_unit": getattr(c, "default_unit", "")}
            for c in store2.load_coefficients()
        ]
        units2 = store2.load_units()
        suppliers2 = [
            {
                "name": s.name,
                "customer_code": s.customer_code,
                "coord1": s.coord1,
                "coord2": s.coord2,
            }
            for s in store2.load_suppliers()
        ]

        cbc1, cbc2 = st.columns([1, 1])
        with cbc1:
            if st.button("Générer Bon de commande (Excel)", type="primary"):
                out_path = _temp_out_path(".xlsx")
                export_excel(
                    bon_edit,
                    prod_dej_long,
                    prod_din_long,
                    out_path,
                    coefficients=coeffs2,
                    units=units2,
                    suppliers=suppliers2,
                )
                with open(out_path, "rb") as f:
                    st.download_button(
                        "Télécharger Bon de commande.xlsx",
                        data=f,
                        file_name="Bon de commande.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

        with cbc2:
            st.markdown("**Bons par fournisseur (après édition du bon)**")
            st.caption(
                "1) Télécharge le bon Excel, complète les colonnes Coefficient/Unité/Fournisseur/Prix cible unitaire (et ajuste Quantité si besoin). "
                "2) Ré-uploade le fichier ici pour générer 1 bon par fournisseur."
            )
            bc_filled = st.file_uploader(
                "Bon de commande rempli (.xlsx/.xls)", type=["xlsx","xlsm","xls"], key="bc_filled"
            )
            if bc_filled is not None:
                try:
                    df_filled = _read_excel_any(bc_filled, sheet_name="Bon de commande")
                except Exception:
                    df_filled = _read_excel_any(bc_filled)

                out_xlsx = _temp_out_path(".xlsx")
                out_pdf = _temp_out_path(".pdf")

                # ✅ Nettoyage: on garde uniquement les colonnes utiles pour les bons fournisseurs
                # Colonnes utiles pour les bons fournisseurs.
                # On garde aussi les colonnes de calcul (prix/poids) si elles existent dans le fichier
                # afin qu'elles puissent être affichées telles quelles (sinon elles seront recalculées).
                # NB: on conserve aussi les infos de consommation si elles existent (Jour(s)/Date/Typologie)
                # afin de proposer des créneaux de livraison cohérents.
                cols_keep = [
                    "Jour(s)", "Jour", "Date", "Typologie",
                    "Produit", "Libellé", "Unité", "Quantité",
                    "Prix cible unitaire", "Prix cible total",
                    "Poids unitaire (kg)", "Poids total (kg)",
                    "Fournisseur",
                ]
                df_filled = df_filled[[c for c in cols_keep if c in df_filled.columns]].copy()

                # Créneaux de livraison par fournisseur (JJ/MM/YYYY, séparés par des virgules)
                suppliers_in_file = sorted(
                    [
                        s
                        for s in df_filled.get("Fournisseur", pd.Series(dtype=str))
                        .fillna("")
                        .astype(str)
                        .str.strip()
                        .unique()
                        .tolist()
                        if s
                    ]
                )
                delivery_dates_by_supplier = {}
                if suppliers_in_file:
                    st.markdown("**Créneaux de livraison par fournisseur (JJ/MM/YYYY)**")
                    st.caption(
                        "Tu peux préparer 2 créneaux (par défaut identiques pour tous les fournisseurs). "
                        "Le split se fait d’abord sur le créneau 1 (jusqu’au seuil de poids), puis sur le créneau 2."
                    )

                    # --- Proposition automatique des créneaux, basée sur les jours de consommation ---
                    # Règle demandée (PLATS UNIQUEMENT) :
                    # - livraison autorisée entre J-8 et J-3 avant consommation.
                    #   ("au maximum 8 jours avant" mais "au moins 3 jours avant".)
                    # - autres typologies : pas de restriction (n'influencent pas la proposition)
                    # IMPORTANT : dans le fichier Excel re-uploadé, il n'y a pas de colonne "Date".
                    # Les valeurs de "Jour(s)" (Lundi, Mardi, ...) correspondent toujours à la SEMAINE SUIVANTE.
                    def _typology_is_plat(typ: str) -> bool:
                        t = (typ or "").strip().lower()
                        # tolérance accents/variantes
                        t = (
                            t.replace("é", "e")
                            .replace("è", "e")
                            .replace("ê", "e")
                            .replace("à", "a")
                            .replace("ç", "c")
                        )
                        return ("plat" in t)

                    def _coerce_date(v):
                        if v is None or (isinstance(v, float) and pd.isna(v)):
                            return None
                        if isinstance(v, dt.datetime):
                            return v.date()
                        if isinstance(v, dt.date):
                            return v
                        s = str(v).strip()
                        if not s:
                            return None
                        # formats courants
                        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
                            try:
                                return dt.datetime.strptime(s, fmt).date()
                            except Exception:
                                pass
                        try:
                            return pd.to_datetime(s, dayfirst=True, errors="coerce").date()
                        except Exception:
                            return None

                    def _dates_from_jours_str(jours_raw: str, *, ref_date: dt.date):
                        if not jours_raw:
                            return []
                        # "Jour(s)" correspond toujours à la semaine suivante.
                        # On ancre donc sur le lundi de la semaine prochaine, quelle que soit la date du jour.
                        # 0 = lundi ... 6 = dimanche
                        days_until_next_monday = 7 - ref_date.weekday()
                        base_monday = ref_date + dt.timedelta(days=days_until_next_monday)
                        parts = [p.strip() for p in str(jours_raw).split(",") if p.strip()]
                        out: list[dt.date] = []
                        for p in parts:
                            # normalise (sans accents) et prend le début du mot
                            pl = (
                                p.lower()
                                .replace("é", "e")
                                .replace("è", "e")
                                .replace("ê", "e")
                                .replace("à", "a")
                                .replace("ç", "c")
                            )
                            idx = None
                            for i, dn in enumerate(DAY_NAMES):
                                dnl = (
                                    dn.lower()
                                    .replace("é", "e")
                                    .replace("è", "e")
                                    .replace("ê", "e")
                                    .replace("à", "a")
                                    .replace("ç", "c")
                                )
                                if pl.startswith(dnl[:3]):
                                    idx = i
                                    break
                            if idx is None:
                                continue
                            out.append(base_monday + dt.timedelta(days=int(idx)))
                        # unique + tri
                        out = sorted(set(out))
                        return out

                    def _suggest_slots_for_df(df_sub: pd.DataFrame):
                        today = dt.date.today()
                        if df_sub is None or df_sub.empty:
                            return today, today + dt.timedelta(days=1)

                        # On ne tient compte que des lignes "Plat".
                        # Pour un (ou plusieurs) jours de consommation C, une date de livraison D est valide si :
                        #   C-8 <= D <= C-3
                        # Pour une livraison unique qui couvre plusieurs consommations, on utilise l'intersection :
                        #   D >= max(C-8)  et  D <= min(C-3)
                        conso_dates_plats: list[dt.date] = []
                        jours_col = 'Jour(s)' if 'Jour(s)' in df_sub.columns else ('Jour' if 'Jour' in df_sub.columns else None)
                        for _, r in df_sub.iterrows():
                            if not _typology_is_plat(str(r.get("Typologie", "") or "")):
                                continue

                            # Dates de consommation de la ligne
                            conso_dates_row: list[dt.date] = []
                            if "Date" in df_sub.columns:
                                d = _coerce_date(r.get("Date"))
                                if d:
                                    conso_dates_row.append(d)
                            if not conso_dates_row and jours_col is not None:
                                conso_dates_row = _dates_from_jours_str(r.get(jours_col), ref_date=today)

                            conso_dates_plats.extend(conso_dates_row)

                        conso_dates_plats = sorted(set([d for d in conso_dates_plats if isinstance(d, dt.date)]))

                        # S'il n'y a aucun "Plat" (ou aucune info jours), on retombe sur le comportement simple.
                        if not conso_dates_plats:
                            return today, today + dt.timedelta(days=1)

                        c_min = min(conso_dates_plats)
                        c_max = max(conso_dates_plats)
                        lower = c_max - dt.timedelta(days=8)  # max(C-8)
                        upper = c_min - dt.timedelta(days=3)  # min(C-3)

                        # Si l'intersection est vide, on propose quand même une fourchette "raisonnable"
                        # (premier plat et dernier plat) afin que l'utilisateur ajuste manuellement si besoin.
                        if lower > upper:
                            d1 = max(today, c_min - dt.timedelta(days=8))
                            d2 = max(today, c_max - dt.timedelta(days=3))
                            if d2 < d1:
                                d2 = d1
                            return d1, d2

                        d1 = max(today, lower)
                        d2 = max(today, upper)
                        if d2 < d1:
                            d2 = d1
                        return d1, d2

                    # Proposition globale (sur tout le fichier)
                    suggested_g1, suggested_g2 = _suggest_slots_for_df(df_filled)

                    # Streamlit mémorise les widgets via leur key : si tu changes de fichier,
                    # les valeurs peuvent rester bloquées.
                    # On crée donc une clé dépendante du fichier uploadé pour forcer un recalcul.
                    import hashlib as _hashlib
                    _raw = None
                    try:
                        _raw = bc_filled.getvalue()
                    except Exception:
                        _raw = (getattr(bc_filled, 'name', '') + str(getattr(bc_filled, 'size', ''))).encode('utf-8', errors='ignore')
                    _sig = _hashlib.md5(_raw[:65536] if isinstance(_raw,(bytes,bytearray)) else bytes(str(_raw), 'utf-8')).hexdigest()[:8]
                    key_prefix = f"bc_{_sig}"


                    # Préremplissage global (modifiable)
                    dcol1, dcol2 = st.columns(2)
                    with dcol1:
                        d1 = st.date_input("Créneau 1 (date)", value=suggested_g1, key=f"{key_prefix}_global_slot_1")
                    with dcol2:
                        d2 = st.date_input("Créneau 2 (date)", value=suggested_g2, key=f"{key_prefix}_global_slot_2")

                    default_raw = f"{d1.strftime('%d/%m/%Y')}, {d2.strftime('%d/%m/%Y')}"

                    for sup_name in suppliers_in_file:
                        # Proposition par fournisseur (si on a des infos de consommation dans le fichier)
                        try:
                            df_sup = df_filled[df_filled.get("Fournisseur", "") == sup_name].copy()
                        except Exception:
                            df_sup = pd.DataFrame()
                        sd1, sd2 = _suggest_slots_for_df(df_sup) if not df_sup.empty else (d1, d2)
                        per_default = f"{sd1.strftime('%d/%m/%Y')}, {sd2.strftime('%d/%m/%Y')}"
                        raw = st.text_input(
                            f"{sup_name} — créneaux",
                            value=per_default or default_raw,
                            key=f"{key_prefix}_slots_{sup_name}",
                        )
                        slots = [x.strip() for x in str(raw).split(",") if x.strip()]
                        delivery_dates_by_supplier[sup_name] = slots
                else:
                    st.info("Aucun fournisseur renseigné dans le bon de commande rempli.")

                cbtn1, cbtn2 = st.columns(2)
                with cbtn1:
                    if st.button("Générer Excel (1 feuille / fournisseur)", key="gen_sup_xlsx"):
                        export_orders_per_supplier_excel(df_filled, out_xlsx, suppliers=suppliers2, delivery_dates_by_supplier=delivery_dates_by_supplier, max_weight_kg=600.0)
                        with open(out_xlsx, "rb") as f:
                            st.download_button(
                                "Télécharger Bons fournisseurs.xlsx",
                                data=f,
                                file_name="Bons fournisseurs.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )
                with cbtn2:
                    if st.button("Générer PDF (1 page / fournisseur)", key="gen_sup_pdf"):
                        export_orders_per_supplier_pdf(df_filled, out_pdf, suppliers=suppliers2, delivery_dates_by_supplier=delivery_dates_by_supplier, max_weight_kg=600.0)
                        with open(out_pdf, "rb") as f:
                            st.download_button(
                                "Télécharger Bons fournisseurs.pdf",
                                data=f,
                                file_name="Bons fournisseurs.pdf",
                                mime="application/pdf",
                            )

    with tab_bl:
        st.subheader("Bons de livraison (PDF)")
        st.caption("Les bons ne sont pas générés pour les jours où il n’y a aucune consommation.")

        sites_exclus_txt = st.text_input(
            "Sites exclus (séparés par des virgules)", value="24 ter, 24 simple, IME TL"
        )
        sites_exclus = [s.strip() for s in sites_exclus_txt.split(",") if s.strip()]

        if st.button("Générer Bons de livraison (PDF)", type="primary"):
            out_pdf = _temp_out_path(".pdf")
            export_bons_livraison_pdf(
                planning=planning,
                menu_path=menu_path,
                planning_path=planning_path,
                out_pdf_path=out_pdf,
                sheet_menu="Feuil2",
                sites_exclus=sites_exclus,
            )
            with open(out_pdf, "rb") as f:
                st.download_button(
                    "Télécharger Bons de livraison.pdf",
                    data=f,
                    file_name="Bons de livraison.pdf",
                    mime="application/pdf",
                )

    with tab_factu:
        st.subheader("Facturation mensuelle (mémoire des semaines)")
        st.caption(
            "Idée : à chaque semaine, tu peux mémoriser le planning. Ensuite tu exportes un classeur Excel par mois, "
            "avec 2 tableaux : Repas et Mixé/Lissé (sans PDJ)."
        )

        today = dt.date.today()
        default_monday = today - dt.timedelta(days=today.weekday())
        week_monday = st.date_input("Lundi de la semaine du planning", value=default_monday)

        repas_daily = planning_to_daily_totals(planning["dejeuner"], planning["diner"], week_monday)
        ml_daily = mixe_lisse_to_daily_totals(
            mix_planning.get("dejeuner"), mix_planning.get("diner"), week_monday
        )

        cA, cB = st.columns(2)
        with cA:
            st.markdown("**Aperçu — total Repas (semaine)**")
            if repas_daily.empty:
                st.info("Aucune donnée Repas détectée.")
            else:
                st.dataframe(
                    repas_daily.groupby("site", as_index=False)["qty_repas"]
                    .sum()
                    .sort_values("qty_repas", ascending=False),
                    use_container_width=True,
                    hide_index=True,
                )
        with cB:
            st.markdown("**Aperçu — total Mixé/Lissé (semaine)**")
            if ml_daily.empty:
                st.info("Aucune donnée Mixé/Lissé détectée (feuille absente ou vide).")
            else:
                st.dataframe(
                    ml_daily.groupby("site", as_index=False)["qty_ml"]
                    .sum()
                    .sort_values("qty_ml", ascending=False),
                    use_container_width=True,
                    hide_index=True,
                )

        st.divider()
        if st.button("📌 Mémoriser cette semaine", type="primary"):
            n_repas, n_ml = save_week(
                week_monday=week_monday,
                repas_daily=repas_daily,
                ml_daily=ml_daily,
                source_filename=getattr(planning_file, "name", ""),
            )
            st.success(f"Semaine mémorisée : {n_repas} lignes repas, {n_ml} lignes mixé/lissé.")

        st.markdown("### Mémoriser plusieurs semaines d'un coup")
        st.caption(
            "Optionnel : upload plusieurs plannings (1 fichier = 1 semaine), choisis le lundi de départ, "
            "et l'app mémorise tout d'un coup."
        )

        batch_files = st.file_uploader(
            "Plannings fabrication (plusieurs fichiers .xlsx)",
            type=["xlsx","xlsm"],
            accept_multiple_files=True,
            key="batch_plannings",
        )
        batch_monday = st.date_input(
            "Lundi de départ (pour le 1er fichier)",
            value=week_monday,
            key="batch_monday",
        )

        if st.button("📌 Mémoriser ces semaines", key="batch_save"):
            if not batch_files:
                st.error("Ajoute au moins 1 fichier planning (.xlsx).")
            else:
                total_repas = 0
                total_ml = 0
                for i, up in enumerate(batch_files):
                    w_mon = batch_monday + dt.timedelta(days=7 * i)

                    # ✅ parse depuis un fichier temp (robuste)
                    tmp_path_i = _save_uploaded_file(up, suffix=".xlsx")
                    plan_i = parse_planning_fabrication(tmp_path_i)

                    mix_i = {"dejeuner": pd.DataFrame(), "diner": pd.DataFrame()}
                    try:
                        mix_i = parse_planning_mixe_lisse(tmp_path_i)
                    except Exception:
                        pass

                    repas_i = planning_to_daily_totals(plan_i["dejeuner"], plan_i["diner"], w_mon)
                    ml_i = mixe_lisse_to_daily_totals(mix_i.get("dejeuner"), mix_i.get("diner"), w_mon)

                    n_r, n_m = save_week(
                        week_monday=w_mon,
                        repas_daily=repas_i,
                        ml_daily=ml_i,
                        source_filename=getattr(up, "name", ""),
                    )
                    total_repas += n_r
                    total_ml += n_m

                st.success(
                    f"Semaines mémorisées : {len(batch_files)} fichier(s) → "
                    f"{total_repas} lignes repas, {total_ml} lignes mixé/lissé."
                )

        st.markdown("### Export facturation")
        records = load_records()
        if records.empty:
            st.warning("Aucune semaine mémorisée pour le moment.")
        else:
            records = records.copy()
            records["date"] = pd.to_datetime(records["date"]).dt.date
            months = sorted({(d.year, d.month) for d in records["date"]})
            month_labels = [f"{y}-{m:02d}" for (y, m) in months]
            _ = st.multiselect("Mois présents (info)", options=month_labels, default=month_labels)

            if st.button("Générer le classeur Excel de facturation"):
                out_xlsx = _temp_out_path(".xlsx")
                export_monthly_workbook(records, out_xlsx)

                with open(out_xlsx, "rb") as f:
                    st.download_button(
                        "Télécharger Facturation.xlsx",
                        data=f,
                        file_name="Facturation.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

        with st.expander("🗑️ Suppression (facturation) — semaines / lignes", expanded=False):
            st.warning(
                "Zone sensible : ici tu peux supprimer des données mémorisées. "
                "Les exports Excel/PDF n'affectent pas la mémoire ; seule la suppression efface réellement."
            )
            if delete_billing_records is None:
                st.error("La fonction de suppression n'est pas disponible dans cette version.")
            else:
                rec = load_records()
                if rec.empty:
                    st.info("Aucune donnée à supprimer.")
                else:
                    weeks = sorted({str(w) for w in rec.get("week_monday", "").astype(str).tolist() if str(w).strip()})
                    sel_week = st.selectbox("Semaine à supprimer (lundi)", options=[""] + weeks, index=0)
                    if st.button("Supprimer cette semaine", key="del_billing_week"):
                        if not sel_week:
                            st.error("Choisis un lundi de semaine.")
                        else:
                            n = delete_billing_records(week_monday=dt.date.fromisoformat(sel_week))
                            st.success(f"✅ {n} ligne(s) supprimée(s) pour la semaine {sel_week}.")

        st.divider()
        st.markdown("### Importer une facturation corrigée (retour comptable)")
        st.caption(
            "Si le fichier Facturation.xlsx est corrigé (quantités modifiées), tu peux le réimporter ici. "
            "L'app mettra à jour sa mémoire (records.csv) pour les mois présents dans le fichier."
        )
        corrected_xlsx = st.file_uploader(
            "Facturation corrigée (.xlsx)",
            type=["xlsx","xlsm"],
            key="factu_corrected",
        )
        if st.button("✅ Appliquer les corrections", key="apply_factu_corrections"):
            if not corrected_xlsx:
                st.error("Upload d'abord un fichier Facturation.xlsx corrigé.")
            else:
                tmp_corr = _save_uploaded_file(corrected_xlsx, suffix=".xlsx")
                try:
                    n_removed, n_added = apply_corrected_monthly_workbook(tmp_corr)
                    if n_removed == 0 and n_added == 0:
                        st.warning("Aucune donnée importable détectée dans ce fichier (vérifie qu'il vient bien de l'app).")
                    else:
                        st.success(
                            f"Corrections appliquées : {n_removed} ligne(s) remplacée(s) / supprimée(s), "
                            f"{n_added} ligne(s) importée(s)."
                        )
                except Exception as e:
                    st.error("Impossible d'importer ce fichier. Il doit provenir de l'export de l'app.")
                    st.code(repr(e))

    with tab_factu_pdj:
        st.subheader("Facturation PDJ")
        st.caption(
            "Objectif : enregistrer plusieurs bons de commande PDJ (par site), définir les prix unitaires par produit, "
            "ajouter des consommations/avoirs manuels, puis exporter une facturation mensuelle détaillée."
        )

        # --- Sélection du mois (YYYY-MM)
        today = dt.date.today()
        default_month = f"{today.year:04d}-{today.month:02d}"
        month = st.text_input("Mois à facturer (YYYY-MM)", value=default_month, key="pdj_month")

        st.markdown("### 1) Saisie d'un bon PDJ")
        c1, c2, c3 = st.columns([1, 1, 2])
        with c1:
            pdj_date = st.date_input("Date du bon (commande/livraison)", value=today, key="pdj_date")
        with c2:
            pdj_site = st.text_input("Site (ex: 24 ter, 24 simple, MAS TL...)", value="", key="pdj_site")
        with c3:
            pdj_source = st.text_input("Référence (optionnel)", value="", key="pdj_source")

        pdj_file = st.file_uploader(
            "Importer un bon PDJ (PDF/Excel) — optionnel (archivage)",
            type=["xlsx", "xls", "xlsm", "pdf"],
            key="pdj_import",
        )

        st.caption(
            "Tu peux saisir **manuellement** (mode fiable) : une ligne = 1 produit, "
            "ou importer un bon (Excel/PDF) pour **pré-remplir** les quantités. "
            "Quel que soit le résultat, tout reste **modifiable** manuellement (utile pour la MAS)."
        )

        # Table de saisie pré-remplie (avec pré-lecture best-effort)
        base_rows = pd.DataFrame({"product": pdj_default_products, "qty": 0.0})


        if pdj_file is not None:
            try:
                parsed = pdj_facturation.parse_pdj_order_file(pdj_file)
                items = getattr(parsed, "items", []) or []
                detected_site = getattr(parsed, "site", None)

                # Pré-remplissage du champ site si vide
                if detected_site and not str(st.session_state.get("pdj_site", "")).strip():
                    st.session_state["pdj_site"] = str(detected_site)

                # Pré-remplissage des quantités (matching souple)
                def _k(s: str) -> str:
                    return " ".join(str(s or "").strip().lower().split())

                items_map = { _k(p): float(q) for (p, q) in items if p and q is not None }
                used = set()

                # Remplit d'abord les produits déjà dans la liste
                for i, prod in enumerate(list(base_rows["product"])):
                    k = _k(prod)
                    if k in items_map:
                        base_rows.loc[i, "qty"] = items_map[k]
                        used.add(k)
                    else:
                        # fallback: contient / ressemble (OCR)
                        for kk, vv in items_map.items():
                            if kk in k or k in kk:
                                base_rows.loc[i, "qty"] = vv
                                used.add(kk)
                                break

                # Ajoute les produits non reconnus dans la liste par défaut
                extra = [(p, q) for (p, q) in items if _k(p) not in used]
                if extra:
                    extra_df = pd.DataFrame({"product": [p for p, _ in extra], "qty": [float(q) for _, q in extra]})
                    base_rows = pd.concat([base_rows, extra_df], ignore_index=True)

                msg = "📎 Bon importé : pré-remplissage effectué (à vérifier / corriger si besoin)."
                if detected_site:
                    msg += f" Site détecté : **{detected_site}**."
                st.info(msg)
            except Exception as e:
                st.warning("📎 Bon importé, mais lecture automatique impossible. Saisie manuelle requise.")
                st.code(repr(e))

        pdj_table = st.data_editor(
            base_rows,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="pdj_order_editor",
        )

        kind = st.selectbox(
            "Type d'enregistrement",
            options=["commande", "manuel", "avoir_qty"],
            index=0,
            help="commande = bon de commande; manuel = consommation ajoutée; avoir_qty = avoir en quantités (valeurs négatives possibles)",
            key="pdj_kind",
        )
        comment = st.text_input("Commentaire (optionnel)", value="", key="pdj_comment")

        if st.button("➕ Enregistrer ce bon PDJ", type="primary", key="pdj_save_order"):
            if not str(pdj_site).strip():
                st.error("Renseigne un site pour enregistrer le bon.")
            else:
                df = pdj_table.copy()
                df["qty"] = pd.to_numeric(df["qty"], errors="coerce").fillna(0.0)
                df = df[df["qty"] != 0].copy()
                if df.empty:
                    st.warning("Aucune quantité non nulle : rien à enregistrer.")
                else:
                    df["date"] = pdj_date
                    df["site"] = pdj_site
                    df["kind"] = kind
                    df["comment"] = comment
                    n = pdj_add_records(df, source_filename=pdj_source)
                    st.success(f"✅ Bon enregistré : {n} ligne(s).")

        st.markdown("### 2) Tarifs unitaires")
        st.caption(
            "Renseigne les prix unitaires par produit. Si tous les sites ont le même tarif, laisse 'site' = __default__. "
            "Tu peux ajouter une ligne avec un site spécifique si besoin."
        )
        prices = pdj_load_prices()
        prices_edit = st.data_editor(
            prices,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="pdj_prices_editor",
        )
        if st.button("💾 Enregistrer les tarifs", key="pdj_save_prices"):
            try:
                n = pdj_save_prices(prices_edit)
                st.success(f"✅ Tarifs enregistrés ({n} lignes).")
            except Exception as e:
                st.error("Erreur lors de l'enregistrement des tarifs")
                st.code(repr(e))

        st.markdown("### 3) Ajustements monétaires (avoirs / corrections en €)")
        st.caption("Exemples : avoir global, correction sans quantité, frais exceptionnels. Montant négatif = avoir.")
        adj_base = pd.DataFrame(
            {
                "date": [today],
                "site": [""],
                "label": ["Avoir"],
                "amount_eur": [0.0],
                "comment": [""],
            }
        )
        adj_edit = st.data_editor(
            adj_base,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="pdj_adj_editor",
        )
        if st.button("➕ Ajouter ces ajustements", key="pdj_add_adj"):
            try:
                df = adj_edit.copy()
                df["amount_eur"] = pd.to_numeric(df["amount_eur"], errors="coerce").fillna(0.0)
                df = df[df["amount_eur"] != 0].copy()
                if df.empty:
                    st.warning("Aucun montant non nul : rien à ajouter.")
                else:
                    n = pdj_add_money_adjustments(df)
                    st.success(f"✅ Ajustements ajoutés : {n} ligne(s).")
            except Exception as e:
                st.error("Erreur lors de l'ajout des ajustements")
                st.code(repr(e))

        st.markdown("### 4) Aperçu & export mensuel")
        synth, detail, adj = pdj_billing.compute_monthly_pdj(month)
        cA, cB = st.columns(2)
        with cA:
            st.markdown("**Synthèse par site**")
            st.dataframe(synth, use_container_width=True, hide_index=True)
        with cB:
            st.markdown("**Détail lignes (mois)**")
            st.dataframe(detail, use_container_width=True, hide_index=True)

        c_exp1, c_exp2 = st.columns([1, 1])
        with c_exp1:
            do_xlsx = st.button("📤 Exporter Facturation PDJ (Excel)", type="primary", key="pdj_export")
        with c_exp2:
            do_pdf = st.button("📄 Exporter factures PDF par site (ZIP)", type="primary", key="pdj_export_pdf")

        if do_xlsx:
            out_xlsx = _temp_out_path(".xlsx")
            try:
                pdj_export_monthly_workbook(month, out_xlsx)
                with open(out_xlsx, "rb") as f:
                    st.download_button(
                        "Télécharger Facturation_PDJ.xlsx",
                        data=f,
                        file_name=f"Facturation_PDJ_{month}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            except Exception as e:
                st.error("Erreur lors de l'export PDJ")
                st.code(repr(e))

        if do_pdf:
            try:
                zip_bytes = pdj_billing.export_monthly_invoices_zip(month)
                st.download_button(
                    "Télécharger factures PDJ (ZIP)",
                    data=zip_bytes,
                    file_name=f"Factures_PDJ_{month}.zip",
                    mime="application/zip",
                )
            except Exception as e:
                st.error("Erreur lors de la génération des PDF")
                st.code(repr(e))

        # ------------------------------------------------------------------
        # Gestion des exports persistants (suppression)
        # ------------------------------------------------------------------
        with st.expander("📁 Exports PDJ enregistrés — suppression", expanded=False):
            st.caption(
                "Les exports (Excel / ZIP) générés ici sont maintenant enregistrés dans le dossier data/facturation_pdj/exports. "
                "Tu peux les supprimer à tout moment."
            )
            try:
                saved = pdj_billing.list_saved_exports(month)
            except Exception:
                saved = []

            if not saved:
                st.info("Aucun export enregistré pour ce mois.")
            else:
                options = {p.name: str(p) for p in saved}
                sel_name = st.selectbox("Choisir un export", options=list(options.keys()), key="pdj_saved_export_sel")
                c1, c2 = st.columns([1, 1])
                with c1:
                    st.code(sel_name)
                with c2:
                    if st.button("🗑️ Supprimer cet export", key="pdj_saved_export_del"):
                        ok = pdj_billing.delete_saved_export(options[sel_name])
                        if ok:
                            st.success("✅ Export supprimé.")
                        else:
                            st.error("Impossible de supprimer cet export (chemin non autorisé ou fichier absent).")

        with st.expander("🗑️ Suppression (PDJ) — lignes / ajustements", expanded=False):
            st.warning(
                "Zone sensible : tu peux supprimer des lignes enregistrées. "
                "Les exports Excel/PDF ne suppriment rien tout seuls."
            )

            if pdj_delete_records is None:
                st.error("La fonction de suppression PDJ n'est pas disponible dans cette version.")
            else:
                cDel1, cDel2 = st.columns(2)
                with cDel1:
                    del_site = st.text_input("Filtre site (optionnel)", value="", key="pdj_del_site")
                    del_source = st.text_input("Filtre référence/source (optionnel)", value="", key="pdj_del_source")
                    del_kind = st.selectbox(
                        "Filtre type (optionnel)",
                        options=["", "commande", "manuel", "avoir_qty"],
                        index=0,
                        key="pdj_del_kind",
                    )
                with cDel2:
                    st.caption("Supprime toutes les lignes correspondant aux filtres pour le mois sélectionné.")
                    if st.button("Supprimer lignes PDJ", key="pdj_delete_btn"):
                        n = pdj_delete_records(month=month, site=del_site, source=del_source, kind=del_kind)
                        st.success(f"✅ {n} ligne(s) PDJ supprimée(s) (mois {month}).")

            if pdj_delete_money_adjustments is None:
                st.info("Suppression des ajustements € indisponible dans cette version.")
            else:
                st.divider()
                st.caption("Ajustements en euros (avoirs/corrections) : suppression par mois et filtre site.")
                del_site2 = st.text_input("Filtre site ajustements (€) (optionnel)", value="", key="pdj_del_site2")
                if st.button("Supprimer ajustements €", key="pdj_delete_adj_btn"):
                    n2 = pdj_delete_money_adjustments(month=month, site=del_site2)
                    st.success(f"✅ {n2} ajustement(s) € supprimé(s) (mois {month}).")

    with tab_all:
        st.subheader("Tableaux allergènes (format EXACT)")
        st.caption(
            "Le logiciel génère **toujours** le tableau (plats + colonnes + bloc 'Origine des viandes') à partir du menu. "
            "L'apprentissage sert uniquement à **préremplir les croix (X)** à partir des classeurs de semaines précédentes."
        )

        base_dir = Path(__file__).parent
        template_dir = base_dir / "templates" / "allergen"

        c1, c2 = st.columns([2, 1])
        with c1:
            st.markdown("### 0) Référentiel maître (obligatoire)")
            master_upload = st.file_uploader(
                "Upload le référentiel maître (.xlsx) (celui que tu fais évoluer semaine après semaine)",
                type=["xlsx","xlsm"],
                key="master_upload",
            )
            st.caption(
                "Astuce : après avoir appris, télécharge le référentiel mis à jour et réutilise-le la semaine suivante."
            )

            st.markdown("### 1) Apprentissage (à partir d'un classeur allergènes rempli)")
            filled_allergen_wb = st.file_uploader(
                "Classeur allergènes rempli (ton format, avec des X) — optionnel (pour apprendre)",
                type=["xlsx","xlsm"],
                key="all_filled_upload",
            )

            if st.button("📚 Apprendre depuis ce classeur", type="primary"):
                if not master_upload:
                    st.error("Upload d'abord le référentiel maître (.xlsx).")
                elif not filled_allergen_wb:
                    st.error("Upload aussi un classeur allergènes rempli (.xlsx).")
                else:
                    tmp_master_in = _save_uploaded_file(master_upload, suffix=".xlsx")
                    tmp_filled = _save_uploaded_file(filled_allergen_wb, suffix=".xlsx")
                    tmp_master_out = _temp_out_path(".xlsx")

                    learn_from_filled_allergen_workbook(tmp_filled, tmp_master_in, tmp_master_out)

                    st.success("Référentiel maître mis à jour.")
                    with open(tmp_master_out, "rb") as f:
                        st.download_button(
                            "Télécharger le référentiel maître mis à jour",
                            data=f,
                            file_name="referentiel_allergenes_maitre.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

        with c2:
            st.markdown("### Templates allergènes")
            if (template_dir / "template_dejeuner.xlsx").exists():
                st.success("Templates présents")
            else:
                st.error("Templates allergènes manquants (templates/allergen).")
            st.caption("Ils doivent être présents dans ton repo GitHub.")

        st.divider()
        st.markdown("### 2) Générer les tableaux allergènes")
        if st.button("📄 Générer tableaux allergènes (Excel)", type="primary"):
            if not (template_dir / "template_dejeuner.xlsx").exists():
                st.error("Templates allergènes manquants (templates/allergen).")
            elif not master_upload:
                st.error("Upload d'abord le référentiel maître (colonne de gauche).")
            else:
                tmp_master = _save_uploaded_file(master_upload, suffix=".xlsx")
                out_all = _temp_out_path(".xlsx")

                out_xlsx, missing = generate_allergen_workbook(
                    menu_excel_path=menu_path,
                    allergen_ref_path=str(tmp_master),
                    out_xlsx_path=out_all,
                    template_dir=str(template_dir),
                )

                if missing:
                    st.warning(
                        "Certains plats n'ont pas été trouvés dans le référentiel. "
                        "Ils sont listés dans l'onglet _plats_non_trouves du classeur."
                    )
                    with st.expander("Voir la liste des plats non trouvés"):
                        st.write(sorted(set(missing)))

                with open(out_xlsx, "rb") as f:
                    st.download_button(
                        "Télécharger Tableaux_allergenes.xlsx",
                        data=f,
                        file_name="Tableaux_allergenes.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

except Exception:
    st.error("Une erreur est survenue pendant le calcul.")
    st.code(traceback.format_exc())


# --- Auto-save persistent data (safe) ---
if "persistent_data" in st.session_state:
    save_data(st.session_state["persistent_data"])
