from __future__ import annotations
import shutil
from pathlib import Path


def apply(project_dir: str) -> None:
    root = Path(project_dir).resolve()
    src_root = Path(__file__).resolve().parent / "Gestion-Cuisine"

    targets = [
        (src_root / "app.py", root / "app.py"),
        (src_root / "src" / "pdj_billing.py", root / "src" / "pdj_billing.py"),
    ]

    for src, dst in targets:
        if not dst.exists():
            raise FileNotFoundError(f"Cible introuvable: {dst}")
        bak = dst.with_suffix(dst.suffix + ".bak")
        if not bak.exists():
            shutil.copy2(dst, bak)
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        print(f"OK: {dst} (backup: {bak.name})")


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("Usage: python apply_patch.py /chemin/vers/ton/Gestion-Cuisine")
        raise SystemExit(2)
    apply(sys.argv[1])
