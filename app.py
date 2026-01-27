import streamlit as st
import traceback
from pathlib import Path
import pandas as pd
import datetime as dt

from src.processor import (
    parse_planning_fabrication,
    parse_planning_mixe_lisse,
    make_production_summary,
    make_production_pivot,
    parse_menu,
    build_bon_commande,
    export_excel,
    export_bons_livraison_pdf,
)

# ✅ Remplacement : on "dé-masque" l'erreur réelle de Streamlit Cloud
try:
    from src.config_store import ConfigStore
except Exception as e:
    st.error("Erreur d'import de ConfigStore (src/config_store.py)")
    st.code(repr(e))
    st.code(traceback.format_exc())
    raise

from src.order_forms import export_orders_per_supplier_excel, export_orders_per_supplier_pdf
from src.billing import (
    planning_to_daily_totals,
    mixe_lisse_to_daily_totals,
    save_week,
    load_records,
    export_monthly_workbook,
)

# --- Nouveau : Allergènes (ajout sans modifier les fonctions existantes) ---
from src.allergens.learner import learn_from_filled_allergen_workbook
from src.allergens.generator import generate_allergen_workbook


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


st.set_page_config(page_title="Gestion cuisine centrale", layout="wide")
set_background()

st.title("Gestion cuisine centrale")

with st.sidebar:
    st.header("Fichiers")
    planning_file = st.file_uploader("Planning fabrication (.xlsx)", type=["xlsx"])
    menu_file = st.file_uploader("Menu (.xlsx)", type=["xlsx"])
    st.markdown("---")
    st.caption(
        "Conseil : utilise les fichiers d’origine (avec formules) ; l’app récupère les valeurs correctement."
    )

    st.markdown("---")
    st.header("Paramètres — Bons de commande")
    store = ConfigStore()
    # Charge la conf à chaque refresh (petits JSON)
    coeffs = store.load_coefficients()
    units = store.load_units()
    suppliers = store.load_suppliers()

    with st.expander("Configurer coefficients / unités / fournisseurs", expanded=False):
        st.caption(
            "Ces listes sont mémorisées (JSON) et utilisées pour les menus déroulants dans le bon de commande."
        )

        ctab1, ctab2, ctab3 = st.tabs(["Coefficients", "Unités", "Fournisseurs"])

        with ctab1:
            dfc = pd.DataFrame([{"name": c.name, "value": c.value, "default_unit": c.default_unit} for c in coeffs])
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
                store.save_coefficients(dfc_edit.to_dict("records"))
                st.success("Coefficients enregistrés.")

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
                store.save_units([u for u in dfu_edit["unit"].astype(str).tolist() if u.strip()])
                st.success("Unités enregistrées.")

        with ctab3:
            dfs = pd.DataFrame(
                [
                    {
                        "name": s.name,
                        "customer_code": s.customer_code,
                        "coord1": s.coord1,
                        "coord2": s.coord2,
                    }
                    for s in suppliers
                ]
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
                store.save_suppliers(dfs_edit.to_dict("records"))
                st.success("Fournisseurs enregistrés.")

if not planning_file or not menu_file:
    st.info("Charge le planning et le menu pour afficher les tableaux et générer les documents.")
    st.stop()

try:
    # ---- Préparation fichiers temporaires (cloud-safe) ----
    planning_path = _save_uploaded_file(planning_file, suffix=".xlsx")
    menu_path = _save_uploaded_file(menu_file, suffix=".xlsx")

    # Parse planning (openpyxl accepte aussi un file-like ; on garde ton comportement)
    planning = parse_planning_fabrication(planning_file)

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
    tab_prod, tab_bc, tab_bl, tab_factu, tab_all = st.tabs(
        [
            "Production (Déj / Dîn)",
            "Bon de commande",
            "Bons de livraison",
            "Facturation mensuelle",
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
                not piv.empty
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
- On trace une barre par jour, avec 2 séries : **Déjeuner** et **Dîner**.

Donc si Mardi = 120 au déjeuner et 95 au dîner, tu verras deux barres (ou deux segments) pour Mardi."""
            )

    with tab_bc:
        st.subheader("Bon de commande")
        bon = build_bon_commande(planning, menu_items)
        st.caption(
            "Tu peux **fusionner/renommer des lignes** en modifiant la colonne *Libellé* (elles seront regroupées au moment des bons par fournisseur)."
        )

        # Aperçu dans l'app (édition légère)
        bon_edit = st.data_editor(
            bon,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key="bc_editor",
        )

        # Charge les paramètres (au cas où la sidebar n'est pas ouverte)
        store2 = ConfigStore()
        coeffs2 = [
            {"name": c.name, "value": float(c.value), "default_unit": c.default_unit}
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
                "1) Télécharge le bon Excel, complète les colonnes Coefficient/Unité/Fournisseur/Libellé. "
                "2) Ré-uploade le fichier ici pour générer 1 bon par fournisseur."
            )
            bc_filled = st.file_uploader("Bon de commande rempli (.xlsx)", type=["xlsx"], key="bc_filled")
            if bc_filled is not None:
                try:
                    df_filled = pd.read_excel(bc_filled, sheet_name="Bon de commande")
                except Exception:
                    df_filled = pd.read_excel(bc_filled)

                out_xlsx = _temp_out_path(".xlsx")
                out_pdf = _temp_out_path(".pdf")

                cbtn1, cbtn2 = st.columns(2)
                with cbtn1:
                    if st.button("Générer Excel (1 feuille / fournisseur)", key="gen_sup_xlsx"):
                        export_orders_per_supplier_excel(df_filled, out_xlsx, suppliers=suppliers2)
                        with open(out_xlsx, "rb") as f:
                            st.download_button(
                                "Télécharger Bons fournisseurs.xlsx",
                                data=f,
                                file_name="Bons fournisseurs.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )
                with cbtn2:
                    if st.button("Générer PDF (1 page / fournisseur)", key="gen_sup_pdf"):
                        export_orders_per_supplier_pdf(df_filled, out_pdf, suppliers=suppliers2)
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

        repas_daily = planning_to_daily_totals(
            planning["dejeuner"], planning["diner"], week_monday
        )
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
            "Optionnel : upload plusieurs plannings (1 fichier = 1 semaine), choisis le lundi de départ, et l'app mémorise tout d'un coup."
        )

        batch_files = st.file_uploader(
            "Plannings fabrication (plusieurs fichiers .xlsx)",
            type=["xlsx"],
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

                    # Parse fabrication (openpyxl accepte aussi le file-like)
                    plan_i = parse_planning_fabrication(up)

                    # Mixé/Lissé : nécessite un path (on passe par un temp)
                    mix_i = {"dejeuner": pd.DataFrame(), "diner": pd.DataFrame()}
                    try:
                        tmp_path_i = _save_uploaded_file(up, suffix=".xlsx")
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
                    f"Semaines mémorisées : {len(batch_files)} fichier(s) → {total_repas} lignes repas, {total_ml} lignes mixé/lissé."
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
            # Export produces a full-year workbook (Jan → Dec) for the most recent year.
            # We keep the selector for information, but default to all months so users don't
            # accidentally export only the last one.
            choice = st.multiselect(
                "Mois présents (info)", options=month_labels, default=month_labels
            )

            if st.button("Générer le classeur Excel de facturation"):
                # Always export the full year so the workbook can be used from Jan to Dec.
                records_f = records

                out_xlsx = _temp_out_path(".xlsx")
                export_monthly_workbook(records_f, out_xlsx)

                with open(out_xlsx, "rb") as f:
                    st.download_button(
                        "Télécharger Facturation.xlsx",
                        data=f,
                        file_name="Facturation.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

    # ==============================
    # Allergènes
    # ==============================
    with tab_all:
        st.subheader("Tableaux allergènes (format EXACT)")
        st.caption(
            "Le logiciel génère **toujours** le tableau (plats + colonnes + bloc 'Origine des viandes') à partir du menu. "
            "L'apprentissage sert uniquement à **préremplir les croix (X)** à partir des classeurs de semaines précédentes que tu remplis."
        )

        base_dir = Path(__file__).parent
        template_dir = base_dir / "templates" / "allergen"

        # CLOUD-SAFE : on évite de dépendre d'un fichier local persistant.
        # On passe par upload / download du référentiel maître.
        c1, c2 = st.columns([2, 1])
        with c1:
            st.markdown("### 0) Référentiel maître (obligatoire)")
            master_upload = st.file_uploader(
                "Upload le référentiel maître (.xlsx) (celui que tu fais évoluer semaine après semaine)",
                type=["xlsx"],
                key="master_upload",
            )
            st.caption(
                "Astuce : après avoir appris, télécharge le référentiel mis à jour et réutilise-le la semaine suivante."
            )

            st.markdown("### 1) Apprentissage (à partir d'un classeur allergènes rempli)")
            filled_allergen_wb = st.file_uploader(
                "Classeur allergènes rempli (ton format, avec des X) — optionnel (pour apprendre)",
                type=["xlsx"],
                key="all_filled_upload",
            )
            st.markdown(
                "- Chaque semaine : tu exportes le classeur allergènes, tu complètes les X, puis tu l'upload ici.\n"
                "- Le logiciel met à jour le référentiel maître en faisant un **OR** (si un X existe, il reste).\n"
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

                    learn_from_filled_allergen_workbook(
                        tmp_filled, tmp_master_in, tmp_master_out
                    )

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
