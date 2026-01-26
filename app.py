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
        st.dataframe(bon, use_container_width=True, hide_index=True)

        if st.button("Générer Bon de commande (Excel)", type="primary"):
            out_path = _temp_out_path(".xlsx")
            export_excel(bon, prod_dej_long, prod_din_long, out_path)
            with open(out_path, "rb") as f:
                st.download_button(
                    "Télécharger Bon de commande.xlsx",
                    data=f,
                    file_name="Bon de commande.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
