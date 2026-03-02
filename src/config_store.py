from __future__ import annotations

import json
import os
import sys
import base64
import urllib.request
import urllib.error
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

        # Sync optionnel depuis GitHub (persistance Streamlit Cloud / redémarrages)
        try:
            self._sync_from_github()
        except Exception:
            # ne jamais casser l'app au démarrage
            pass

    # ---------------- Dossier d'écriture sûr ----------------
    def _init_base_dir(self, base_dir: Optional[Union[Path, str]] = None) -> Path:
        candidates: List[Path] = []

        # 1) Dossier explicitement fourni
        if base_dir is not None:
            candidates.append(Path(base_dir))

        # 2) Dossier projet (idéal sur Streamlit) : <repo>/data/config
        try:
            project_root = Path(__file__).resolve().parents[1]  # .../src -> .../
            candidates.append(project_root / "data" / "config")
        except Exception:
            pass

        # 3) Cas "portable" (exécutable packagé) : dossier config à côté de l'exe
        try:
            if getattr(sys, "frozen", False):
                candidates.append(Path(sys.executable).resolve().parent / "config")
        except Exception:
            pass

        # 4) Windows : APPDATA (stable, writable)
        try:
            appdata = os.environ.get("APPDATA")
            if appdata:
                candidates.append(Path(appdata) / "gestion-cuisine" / "config")
        except Exception:
            pass

        # 5) Home (Linux/Mac/Streamlit)
        candidates.append(Path.home() / ".gestion-cuisine" / "config")

        # 6) Dernier recours
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

    # -------- GitHub persistence (optionnel) --------
    # Permet de conserver les JSON même si Streamlit met l'app en veille / redémarre le container.
    # Configuration via variables d'environnement (ou st.secrets):
    # - GC_GITHUB_TOKEN : token GitHub (scope "repo" ou "contents:write" selon type)
    # - GC_GITHUB_REPO  : "owner/repo"
    # - GC_GITHUB_BRANCH: (optionnel) ex "main"
    # - GC_GITHUB_PREFIX: (optionnel) dossier dans le repo, par défaut "data/config"

    def _github_enabled(self) -> bool:
        return bool(os.environ.get("GC_GITHUB_TOKEN") and os.environ.get("GC_GITHUB_REPO"))

    def _github_cfg(self) -> Dict[str, str]:
        return {
            "token": os.environ.get("GC_GITHUB_TOKEN", "").strip(),
            "repo": os.environ.get("GC_GITHUB_REPO", "").strip(),
            "branch": (os.environ.get("GC_GITHUB_BRANCH") or "main").strip(),
            "prefix": (os.environ.get("GC_GITHUB_PREFIX") or "data/config").strip().strip("/"),
        }

    def _http_json(self, url: str, method: str = "GET", payload: Optional[Dict[str, Any]] = None, headers: Optional[Dict[str, str]] = None) -> Any:
        data_bytes = None
        if payload is not None:
            data_bytes = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(url=url, data=data_bytes, method=method)
        hdrs = headers or {}
        for k, v in hdrs.items():
            if v is None:
                continue
            req.add_header(k, v)
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                raw = resp.read()
                if not raw:
                    return None
                return json.loads(raw.decode("utf-8"))
        except urllib.error.HTTPError as e:
            # essaye de lire le JSON d'erreur GitHub pour debug
            try:
                body = e.read().decode("utf-8")
                j = json.loads(body) if body else {}
                raise RuntimeError(f"HTTP {e.code} GitHub: {j}") from e
            except Exception:
                raise
        except Exception as e:
            raise

    def _github_get(self, rel_path: str) -> Optional[str]:
        cfg = self._github_cfg()
        api = f"https://api.github.com/repos/{cfg['repo']}/contents/{cfg['prefix']}/{rel_path}"
        url = f"{api}?ref={cfg['branch']}"
        headers = {
            "Authorization": f"token {cfg['token']}",
            "Accept": "application/vnd.github+json",
            "User-Agent": "gestion-cuisine-configstore",
        }
        try:
            data = self._http_json(url, method="GET", headers=headers)
            if not isinstance(data, dict):
                return None
            content_b64 = data.get("content", "")
            if not content_b64:
                return None
            # GitHub peut insérer des retours à la ligne dans le base64
            content_b64 = content_b64.replace("\n", "").replace("\r", "")
            return base64.b64decode(content_b64.encode("utf-8")).decode("utf-8")
        except Exception:
            return None

    def _github_put(self, rel_path: str, content_text: str, message: str) -> None:
        cfg = self._github_cfg()
        api = f"https://api.github.com/repos/{cfg['repo']}/contents/{cfg['prefix']}/{rel_path}"
        headers = {
            "Authorization": f"token {cfg['token']}",
            "Accept": "application/vnd.github+json",
            "User-Agent": "gestion-cuisine-configstore",
        }

        # récupérer sha si le fichier existe déjà
        sha = None
        try:
            existing = self._http_json(f"{api}?ref={cfg['branch']}", method="GET", headers=headers)
            if isinstance(existing, dict):
                sha = existing.get("sha")
        except Exception:
            sha = None

        payload: Dict[str, Any] = {
            "message": message,
            "content": base64.b64encode(content_text.encode("utf-8")).decode("utf-8"),
            "branch": cfg["branch"],
        }
        if sha:
            payload["sha"] = sha

        # PUT
        self._http_json(api, method="PUT", payload=payload, headers=headers)

    def _sync_from_github(self) -> None:
        if not self._github_enabled():
            return
        # On tente de récupérer les fichiers depuis le repo (sans écraser si le contenu est vide)
        mapping = {
            "coefficients.json": self._coeff_path,
            "units.json": self._units_path,
            "suppliers.json": self._suppliers_path,
        }
        for rel, local in mapping.items():
            txt = self._github_get(rel)
            if txt and txt.strip():
                try:
                    # Valide JSON avant d'écrire localement
                    json.loads(txt)
                    local.write_text(txt, encoding="utf-8")
                except Exception:
                    # ignore fichier distant corrompu
                    pass


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
        payload_text = json.dumps(data, ensure_ascii=False, indent=2)

        # Sauvegarde atomique (évite un fichier JSON corrompu en cas de veille/coupure)
        tmp = path.with_suffix(path.suffix + ".tmp")
        tmp.write_text(payload_text, encoding="utf-8")
        tmp.replace(path)

        # Sync GitHub si configuré (persistance au-delà des redémarrages)
        if self._github_enabled():
            rel = path.name
            try:
                self._github_put(rel, payload_text, message=f"Update {rel} (gestion-cuisine)")
            except Exception as e:
                # On remonte l'erreur : sinon l'utilisateur croit que c'est persistant alors que non
                raise RuntimeError(f"Sauvegarde locale OK mais échec sync GitHub pour {rel}: {e}") from e


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
