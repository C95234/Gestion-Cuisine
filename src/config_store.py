from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import List, Dict, Any, Optional, Union


@dataclass
class Coefficient:
    name: str
    value: float
    default_unit: str = "unité"


@dataclass
class Supplier:
    name: str
    customer_code: str = ""
    coord1: str = ""
    coord2: str = ""


class ConfigStore:
    """Stockage persistant (JSON) des listes utilisées pour les bons de commande."""

    def __init__(self, base_dir: Optional[Union[Path, str]] = None) -> None:
        self.base_dir = self._init_base_dir(base_dir)

        self._coeff_path = self.base_dir / "coefficients.json"
        self._units_path = self.base_dir / "units.json"
        self._suppliers_path = self.base_dir / "suppliers.json"

        # IMPORTANT : aucune exception ne doit sortir d'ici
        try:
            self._ensure_defaults()
        except Exception:
            pass

    # ---------------- Dossier d'écriture sûr ----------------
    def _init_base_dir(self, base_dir: Optional[Union[Path, str]]) -> Path:
        candidates: List[Path] = []

        if base_dir is not None:
            candidates.append(Path(base_dir))

        # Cas "portable" (exécutable packagé) : dossier config à côté de l'exe
        try:
            if getattr(sys, "frozen", False):
                candidates.append(Path(sys.executable).resolve().parent / "config")
        except Exception:
            pass

        # Windows : privilégier APPDATA (stable, writable)
        try:
            appdata = os.environ.get("APPDATA")
            if appdata:
                candidates.append(Path(appdata) / "gestion-cuisine" / "config")
        except Exception:
            pass

        candidates.append(Path.home() / ".gestion-cuisine" / "config")
        candidates.append(Path("/tmp") / "gestion-cuisine" / "config")

        for p in candidates:
            try:
                p.mkdir(parents=True, exist_ok=True)
                test = p / ".write_test"
                test.write_text("ok", encoding="utf-8")
                try:
                    test.unlink()
                except Exception:
                    pass
                return p
            except Exception:
                continue

        # dernier recours (ne plante jamais l'import)
        fallback = Path("/tmp") / "gc_config_fallback"
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback

    # ---------------- Valeurs par défaut ----------------
    def _ensure_defaults(self) -> None:
        if not self._units_path.exists():
            self.save_units(["kg", "g", "L", "mL", "unité", "pièce", "barquette"])

        if not self._coeff_path.exists():
            self.save_coefficients([
                Coefficient(name="1", value=1.0, default_unit="unité"),
                Coefficient(name="1 kg", value=1.0, default_unit="kg"),
            ])

        if not self._suppliers_path.exists():
            self.save_suppliers([])

    # -------- Units --------
    def load_units(self) -> List[str]:
        return self._load_json_list(self._units_path, default=[])

    def save_units(self, units: List[str]) -> None:
        units = [str(u).strip() for u in units if str(u).strip()]
        self._safe_save(self._units_path, units)

    # -------- Coefficients --------
    def load_coefficients(self) -> List[Coefficient]:
        data = self._load_json_list(self._coeff_path, default=[])
        out: List[Coefficient] = []
        for row in data:
            if not isinstance(row, dict):
                continue
            name = str(row.get("name", "")).strip()
            if not name:
                continue
            try:
                value = float(row.get("value", 1.0))
            except Exception:
                value = 1.0
            default_unit = str(row.get("default_unit", "unité") or "unité").strip()
            out.append(Coefficient(name=name, value=value, default_unit=default_unit))
        return out

    def save_coefficients(self, coeffs: Union[List[Coefficient], List[Dict[str, Any]]]) -> None:
        payload: List[Dict[str, Any]] = []
        for c in coeffs:
            if isinstance(c, Coefficient):
                payload.append({"name": c.name, "value": float(c.value), "default_unit": c.default_unit})
            elif isinstance(c, dict):
                name = str(c.get("name", "")).strip()
                if not name:
                    continue
                try:
                    value = float(c.get("value", 1.0))
                except Exception:
                    value = 1.0
                default_unit = str(c.get("default_unit", "unité") or "unité").strip()
                payload.append({"name": name, "value": value, "default_unit": default_unit})
        self._safe_save(self._coeff_path, payload)

    # -------- Suppliers --------
    def load_suppliers(self) -> List[Supplier]:
        data = self._load_json_list(self._suppliers_path, default=[])
        out: List[Supplier] = []
        for row in data:
            if not isinstance(row, dict):
                continue
            name = str(row.get("name", "")).strip()
            if not name:
                continue
            out.append(
                Supplier(
                    name=name,
                    customer_code=str(row.get("customer_code", "") or ""),
                    coord1=str(row.get("coord1", "") or ""),
                    coord2=str(row.get("coord2", "") or ""),
                )
            )
        return out

    def save_suppliers(self, suppliers: Union[List[Supplier], List[Dict[str, Any]]]) -> None:
        payload: List[Dict[str, Any]] = []
        for s in suppliers:
            if isinstance(s, Supplier):
                payload.append({
                    "name": s.name,
                    "customer_code": s.customer_code,
                    "coord1": s.coord1,
                    "coord2": s.coord2,
                })
            elif isinstance(s, dict):
                name = str(s.get("name", "")).strip()
                if not name:
                    continue
                payload.append({
                    "name": name,
                    "customer_code": str(s.get("customer_code", "") or ""),
                    "coord1": str(s.get("coord1", "") or ""),
                    "coord2": str(s.get("coord2", "") or ""),
                })
        self._safe_save(self._suppliers_path, payload)

    # -------- Utils --------
    def _load_json_list(self, path: Path, default: List[Any]) -> List[Any]:
        try:
            if not path.exists():
                return default
            data = json.loads(path.read_text(encoding="utf-8"))
            return data if isinstance(data, list) else default
        except Exception:
            return default

    def _safe_save(self, path: Path, data: Any) -> None:
        """Sauvegarde atomique.

        IMPORTANT : on ne masque plus silencieusement les erreurs, sinon l'utilisateur
        pense que tout est mémorisé alors que rien n'est écrit.
        """
        # Sauvegarde atomique (évite un fichier JSON corrompu en cas de veille/coupure)
        tmp = path.with_suffix(path.suffix + ".tmp")
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp.replace(path)

    # Public helper (diagnostic)
    def info(self) -> Dict[str, str]:
        return {
            "base_dir": str(self.base_dir),
            "coefficients": str(self._coeff_path),
            "units": str(self._units_path),
            "suppliers": str(self._suppliers_path),
        }


# Expose explicit public API for `from src.config_store import ConfigStore`
__all__ = ["ConfigStore", "Coefficient", "Supplier"]
