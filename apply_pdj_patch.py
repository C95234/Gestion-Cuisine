#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Patch PDJ Import (anti-conflits)

But:
- Si tes fichiers contiennent des marqueurs de conflits Git (<<<<<<<), on les remet d'abord
  dans une version "clean" (fournie dans ce ZIP), puis on applique le patch.
- Ne touche pas à .git.
- Modifie seulement:
  - app.py (section Facturation PDJ)
  - (optionnel) src/pdj_billing.py si présence de conflits (restauration)

Usage:
  python apply_pdj_patch.py /chemin/vers/Gestion-Cuisine
"""
from __future__ import annotations
import os, re, sys, shutil, textwrap
from pathlib import Path

CLEAN_APP = "CLEAN_app.py"
CLEAN_PDJ_BILLING = "CLEAN_pdj_billing.py"

def backup(path: Path) -> None:
    bak = path.with_suffix(path.suffix + ".bak")
    if not bak.exists():
        shutil.copy2(path, bak)

def has_conflict_markers(text: str) -> bool:
    return "<<<<<<<" in text or ">>>>>>>" in text or "=======" in text

def restore_if_conflicted(target: Path, clean_text: str) -> bool:
    if not target.exists():
        return False
    txt = target.read_text(encoding="utf-8", errors="ignore")
    if has_conflict_markers(txt):
        backup(target)
        target.write_text(clean_text, encoding="utf-8")
        return True
    return False

def patch_app_py(app_path: Path) -> None:
    txt = app_path.read_text(encoding="utf-8")

    block_re = re.compile(
        r'st\.caption\(\n'
        r'\s*"Saisie \*\*manuelle\*\* \(mode fiable\) : une ligne = 1 produit\. "\n'
        r'\s*"Astuce : laisse à 0 les produits non consommés\. "\n'
        r'\s*"Le fichier importé sert uniquement de pièce jointe / référence \(pas de lecture OCR\)\."\n'
        r'\s*\)\n\n'
        r'\s*# Table de saisie pré-remplie.*?\n'
        r'\s*base_rows = pd\.DataFrame\(\{"product": pdj_default_products, "qty": 0\.0\}\)\n'
        r'\s*if pdj_file is not None:\n'
        r'\s*st\.info\("📎 Bon importé en référence\. Renseigne les quantités manuellement ci-dessous\."\)\n'
        r'\s*pdj_table = st\.data_editor\(\n'
        r'.*?key="pdj_order_editor",\n'
        r'\s*\)\n',
        re.S
    )

    replacement_raw = r"""
st.caption(
    "Import intelligent (Excel/PDF) : l'app tente de pré-remplir les quantités automatiquement. "
    "⚠️ Pour les PDF scannés (ex: MAS), le résultat peut être incomplet : tu pourras toujours corriger manuellement."
)

# Table de saisie pré-remplie (modifiable manuellement)
base_rows = pd.DataFrame({"product": pdj_default_products, "qty": 0.0})

suggested_site = ""
if pdj_file is not None:
    # Sauvegarde temporaire du fichier importé
    tmp_in = _save_uploaded_file(pdj_file, suffix=os.path.splitext(pdj_file.name)[1])
    ext = os.path.splitext(tmp_in.lower())[1]

    # Suggestion de site (best-effort) depuis le nom de fichier
    name_l = (pdj_file.name or "").lower()
    site_keywords = {
        "rosa": "24 ter",
        "24 ter": "24 ter",
        "24t": "24 ter",
        "internat": "Internat",
        "vinci": "Internat",
        "toulouse": "MAS",
        "lautrec": "MAS",
        "mas": "MAS",
    }
    for k, v in site_keywords.items():
        if k in name_l:
            suggested_site = v
            break

    # Extraction quantités (best-effort)
    try:
        if ext == ".pdf":
            parsed = pdj_billing.parse_pdj_pdf(tmp_in)
        else:
            parsed = pdj_billing.parse_pdj_excel(tmp_in)
    except Exception:
        parsed = pd.DataFrame(columns=["product", "qty"])

    if parsed is not None and not parsed.empty:
        # map product -> qty en normalisant
        try:
            parsed["product"] = parsed["product"].astype(str).map(pdj_billing.norm_product)
        except Exception:
            parsed["product"] = parsed["product"].astype(str)
        pmap = parsed.groupby("product")["qty"].sum().to_dict()
        base_rows["qty"] = base_rows["product"].map(lambda p: float(pmap.get(p, 0.0)) if p in pmap else 0.0)
        st.success("✅ Quantités pré-remplies depuis le fichier importé (à vérifier).")
    else:
        st.info("📎 Bon importé. Aucune quantité fiable détectée : saisie/correction manuelle ci-dessous.")

    if not str(pdj_site).strip() and suggested_site:
        st.info(f"💡 Site détecté (suggestion) : **{suggested_site}** — tu peux le modifier.")

pdj_table = st.data_editor(
    base_rows,
    use_container_width=True,
    hide_index=True,
    num_rows="dynamic",
    key="pdj_order_editor",
)
"""
    replacement = textwrap.dedent(replacement_raw).strip("\n") + "\n"

    m = block_re.search(txt)
    if not m:
        raise RuntimeError("Bloc PDJ (caption/table) introuvable dans app.py.")

    start = m.start()
    line_start = txt.rfind("\n", 0, start) + 1
    indent = re.match(r"[ \t]*", txt[line_start:start]).group(0)

    replace_start = line_start
    replace_end = m.end()

    rep_indented = "\n".join(indent + line if line.strip() else line for line in replacement.splitlines()) + "\n"
    txt2 = txt[:replace_start] + rep_indented + txt[replace_end:]

    save_re = re.compile(
        r'(if st\.button\("➕ Enregistrer ce bon PDJ".*?key="pdj_save_order"\):\n)'
        r'(?P<ind>[ \t]*)if not str\(pdj_site\)\.strip\(\):\n'
        r'[ \t]*st\.error\("Renseigne un site pour enregistrer le bon\."\)\n'
        r'(?P=ind)else:\n',
        re.S
    )
    m2 = save_re.search(txt2)
    if not m2:
        raise RuntimeError("Bloc PDJ (enregistrement) introuvable dans app.py.")

    ind = m2.group("ind")
    save_repl = (
        m2.group(1)
        + f"{ind}effective_site = str(pdj_site).strip() or suggested_site\n"
        + f"{ind}if not effective_site:\n"
        + f"{ind}    st.error(\"Renseigne un site pour enregistrer le bon.\")\n"
        + f"{ind}else:\n"
    )
    txt3 = txt2[:m2.start()] + save_repl + txt2[m2.end():]
    txt3 = txt3.replace('df["site"] = pdj_site', 'df["site"] = effective_site')

    backup(app_path)
    app_path.write_text(txt3, encoding="utf-8")

def main():
    if len(sys.argv) != 2:
        print("Usage: python apply_pdj_patch.py /chemin/vers/Gestion-Cuisine", file=sys.stderr)
        sys.exit(2)

    root = Path(sys.argv[1]).resolve()
    app_path = root / "app.py"
    pdj_billing_path = root / "src" / "pdj_billing.py"

    pkg_dir = Path(__file__).resolve().parent
    clean_app = (pkg_dir / CLEAN_APP).read_text(encoding="utf-8")
    clean_pdj = (pkg_dir / CLEAN_PDJ_BILLING).read_text(encoding="utf-8")

    restored_app = restore_if_conflicted(app_path, clean_app)
    restored_pdj = restore_if_conflicted(pdj_billing_path, clean_pdj)

    patch_app_py(app_path)

    print("OK ✅ Patch appliqué.")
    if restored_app:
        print(" - app.py contenait des conflits -> restauré depuis une version clean (backup .bak).")
    if restored_pdj:
        print(" - src/pdj_billing.py contenait des conflits -> restauré depuis une version clean (backup .bak).")
    if not restored_app and not restored_pdj:
        print(" - Aucun fichier n'avait de marqueurs de conflits.")

if __name__ == "__main__":
    main()
