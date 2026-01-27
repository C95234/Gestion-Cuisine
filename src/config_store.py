from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import List, Dict, Any, Optional, Union


@dataclass
class Coefficient:
    """Coefficient de conversion pour calculer les quantités."""
    name: str
    value: float
    default_unit: str = "unité"


@dataclass
class Supplier:
    """Fournisseur mémorisé."""
    name: str
    customer_code: str = ""
    coord1: str = ""
    coord2: str = ""


class ConfigStore:
    """Stockage persistant (JSON) des listes utilisées pour les bons de commande."""

    def __init__(self, base_dir: Optional[Union[Path, str]] = None) -> None:
        candidates: List[Path] = []

        if base_dir is not None:
            candidates.append(Path(base_dir))

        # Dossier projet (local)
        candidates.append(Path(__file__).resolve().parent.parent / "data" / "config")

        # HOME (souvent writable sur Streamlit Cloud)
        candidates.append(Path.home() / ".gestion-cuisine" / "config")

        # Secours (toujours writable mais non persistant)
        candidates.append(Path("/tmp") / "gestion-cuisine" / "config")

        self.base_dir = self._pick_writable_dir(candidates)

        self._coeff_path = self.base_dir / "coefficients.json"
        self._units_path = self.base_dir / "units.json"
        self._suppliers_path = self.base_dir / "suppliers.json"

        self._ensure_defaults()

    def _pick_writable_dir(self, candidates: List[Path]) -> Path:
        last_err: Optional[Exception] = None
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
            except Exception as e:
                last_err = e
                continue
        raise RuntimeError(
            "Impossible de trouver un dossier d'écriture pour la config: %r" % (last_err,)
        )

    def _ensure_defaults(self) -> None:
        if not self._units_path.exists():
            self.save_units(["kg", "g", "L", "mL", "unité", "pièce", "barquette"])
        if not self._coeff_path.exists():
            self.save_coefficients(
                [
                    Coefficient(name="1", value=1.0, default_unit="unité"),
                    Coefficient(name="1 kg", value=1.0, default_unit="kg"),
                ]
            )
        if not self._suppliers_path.exists():
            self.save_suppliers([])

    # -------- Units --------
    def load_units(self) -> List[str]:
        return self._load_json_list(self._units_path, default=[])

    def save_units(self, units: List[str]) -> None:
        units = [str(u).strip() for u in units if str(u).strip()]
        self._save_json(self._units_path, units)

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
                payload.append(
                    {"name": c.name, "value": float(c.value), "default_unit": c.default_unit}
                )
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
        self._save_json(self._coeff_path, payload)

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
                payload.append(
                    {
                        "name": s.name,
                        "customer_code": s.customer_code,
                        "coord1": s.coord1,
                        "coord2": s.coord2,
                    }
                )
            elif isinstance(s, dict):
                name = str(s.get("name", "")).strip()
                if not name:
                    continue
                payload.append(
                    {
                        "name": name,
                        "customer_code": str(s.get("customer_code", "") or ""),
                        "coord1": str(s.get("coord1", "") or ""),
                        "coord2": str(s.get("coord2", "") or ""),
                    }
                )
        self._save_json(self._suppliers_path, payload)

    # -------- Utils --------
    def _load_json_list(self, path: Path, default: List[Any]) -> List[Any]:
        try:
            if not path.exists():
                return default
            data = json.loads(path.read_text(encoding="utf-8"))
            return data if isinstance(data, list) else default
        except Exception:
            return default

    def _save_json(self, path: Path, data: Any) -> None:
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
