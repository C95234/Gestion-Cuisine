import shutil
from pathlib import Path
import sys

PATCH_ROOT = Path(__file__).resolve().parent / "Gestion-Cuisine"

def main(project_dir: str):
    proj = Path(project_dir).resolve()
    if not (proj / "app.py").exists():
        raise SystemExit(f"Je ne trouve pas app.py dans {proj}\nLance: python apply_replace.py <dossier_qui_contient_app.py>")

    targets = [
        (PATCH_ROOT / "app.py", proj / "app.py"),
        (PATCH_ROOT / "src" / "pdj_billing.py", proj / "src" / "pdj_billing.py"),
    ]

    for src, dst in targets:
        dst.parent.mkdir(parents=True, exist_ok=True)
        if dst.exists():
            bak = dst.with_suffix(dst.suffix + ".bak")
            shutil.copy2(dst, bak)
            print(f"Backup: {bak}")
        shutil.copy2(src, dst)
        print(f"Replaced: {dst}")

    print("\nOK. Relance l'app Streamlit.")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: python apply_replace.py <dossier_qui_contient_app.py>")
    main(sys.argv[1])
