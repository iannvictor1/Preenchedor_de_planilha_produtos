import json
import os
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
SETTINGS_FILE = Path(os.getenv("SETTINGS_FILE_PATH", BASE_DIR / "settings.local.json"))

DEFAULT_SETTINGS = {
    "folderPath": "Fotos Cod",
    "supplierPdfFolderPath": r"Fichas-20260609T161612Z-3-001\Fichas",
    "factoryRenameExcelPath": "produtos codigo fabrica.xlsx",
    "csvPath": "Listagem dos produtos.xlsx.csv",
    "zipPath": "",
    "pricePath": "preço.xlsx",
    "autoLoadProducts": True,
}


def normalize_settings(data=None):
    data = data if isinstance(data, dict) else {}
    normalized = {}
    for key, default in DEFAULT_SETTINGS.items():
        if isinstance(default, bool):
            normalized[key] = bool(data.get(key, default))
        else:
            value = data.get(key, default)
            normalized[key] = str(value if value is not None else default).strip()
    return normalized


def load_settings():
    if not SETTINGS_FILE.exists():
        return normalize_settings()
    try:
        data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return normalize_settings()
    return normalize_settings(data)


def save_settings(settings):
    normalized = normalize_settings(settings)
    SETTINGS_FILE.write_text(
        json.dumps(normalized, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return normalized
