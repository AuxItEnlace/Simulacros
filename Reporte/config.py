import configparser
from pathlib import Path
from typing import Dict

# Directorio donde vive este archivo (ICFES/Reporte/)
_HERE = Path(__file__).parent


def load_config(filepath: str | None = None) -> Dict[str, str]:
    # Si no se pasa ruta, usar config.ini junto a este módulo
    resolved = Path(filepath) if filepath else _HERE / "config.ini"
    config = configparser.ConfigParser()
    config.optionxform = str  # preserve original case of keys
    parsed = config.read(resolved, encoding="utf-8")
    if not parsed:
        raise FileNotFoundError(
            f"Archivo de configuración no encontrado: {resolved}"
        )
    if "Program" not in config:
        raise KeyError("El archivo config.ini debe tener una sección [Program]")
    section = config["Program"]
    required = ["SchoolName", "Date", "Prueba"]
    for key in required:
        if key not in section:
            raise KeyError(f"Falta clave requerida '{key}' en [Program]")

    result = dict(section)
    # Strip surrounding quotes from optional values (e.g. Sede = "" → "")
    for k, v in result.items():
        if len(v) >= 2 and v[0] == v[-1] and v[0] in ('"', "'"):
            result[k] = v[1:-1]
    return result
