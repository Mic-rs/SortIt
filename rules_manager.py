import json
import winreg
import hashlib
from typing import Optional

REG_BASE = r"Software\SortIt"

DEFAULT_RULES = {
    "categories": [
        {"name": "Immagini", "folder": "Immagini",
         "extensions": [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tiff", ".svg"],
         "keywords": [], "rename_template": ""},
        {"name": "Fogli_Calcolo", "folder": "Fogli_Calcolo",
         "extensions": [".xlsx", ".xls", ".csv", ".ods"],
         "keywords": [], "rename_template": ""},
        {"name": "Presentazioni", "folder": "Presentazioni",
         "extensions": [".pptx", ".ppt", ".odp"],
         "keywords": [], "rename_template": ""},
        {"name": "Archivi", "folder": "Archivi",
         "extensions": [".zip", ".rar", ".7z", ".tar", ".gz"],
         "keywords": [], "rename_template": ""}
    ],
    "fallback_folder": "Da_Revisionare",
    "confidence_threshold": 0.5,
    "rename_enabled": True,
    "dry_run": False
}


def _reg_read(name: str, default=None):
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_BASE)
        value, _ = winreg.QueryValueEx(key, name)
        winreg.CloseKey(key)
        return value
    except Exception:
        return default


def _reg_write(name: str, value: str):
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_BASE)
    winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
    winreg.CloseKey(key)


def _hash_data(data: str) -> str:
    return hashlib.sha256(data.encode("utf-8")).hexdigest()


class RulesManager:
    def __init__(self):
        raw = _reg_read("rules")
        if raw is None:
            self.rules = DEFAULT_RULES.copy()
            self._save_to_registry()
        else:
            try:
                self.rules = json.loads(raw)
                self._check_integrity(raw)
            except Exception:
                self.rules = DEFAULT_RULES.copy()
                self._save_to_registry()

    def _check_integrity(self, raw: str):
        saved_hash = _reg_read("rules_hash")
        current_hash = _hash_data(raw)
        if saved_hash and saved_hash != current_hash:
            import tkinter.messagebox as mb
            mb.showwarning(
                "Attenzione — SortIt",
                "Le regole sono state modificate esternamente dal registro.\n"
                "Verifica che il contenuto sia corretto."
            )
        _reg_write("rules_hash", current_hash)

    def _save_to_registry(self):
        raw = json.dumps(self.rules, ensure_ascii=False)
        _reg_write("rules", raw)
        _reg_write("rules_hash", _hash_data(raw))

    def save(self):
        self._save_to_registry()

    def reload(self):
        raw = _reg_read("rules")
        if raw:
            self.rules = json.loads(raw)

    # --- Proprietà ---

    @property
    def categories(self) -> list:
        return self.rules.get("categories", [])

    @property
    def fallback_folder(self) -> str:
        return self.rules.get("fallback_folder", "Da_Revisionare")

    @property
    def confidence_threshold(self) -> float:
        return float(self.rules.get("confidence_threshold", 0.5))

    @property
    def rename_enabled(self) -> bool:
        return bool(self.rules.get("rename_enabled", True))

    @property
    def dry_run(self) -> bool:
        return bool(self.rules.get("dry_run", False))

    @dry_run.setter
    def dry_run(self, value: bool):
        self.rules["dry_run"] = value

    @rename_enabled.setter
    def rename_enabled(self, value: bool):
        self.rules["rename_enabled"] = value

    # --- Gestione categorie ---

    def add_category(self, name: str, folder: str, keywords: list,
                     extensions: list, rename_template: Optional[str] = None):
        self.rules["categories"].append({
            "name": name, "folder": folder,
            "keywords": keywords, "extensions": extensions,
            "rename_template": rename_template
        })
        self.save()

    def remove_category(self, name: str):
        self.rules["categories"] = [
            c for c in self.rules["categories"] if c["name"] != name
        ]
        self.save()

    def update_category(self, name: str, updated: dict):
        for i, cat in enumerate(self.rules["categories"]):
            if cat["name"] == name:
                self.rules["categories"][i] = updated
                self.save()
                return
        raise ValueError(f"Categoria '{name}' non trovata.")

    def get_category(self, name: str) -> Optional[dict]:
        for cat in self.categories:
            if cat["name"] == name:
                return cat
        return None
