import configparser
from typing import Dict


def load_config(filepath: str = "config.ini") -> Dict[str, str]:
    config = configparser.ConfigParser()
    parsed = config.read(filepath, encoding="utf-8")
    if not parsed:
        raise FileNotFoundError(
            f"Archivo de configuración no encontrado: {filepath}"
        )
    if "Program" not in config:
        raise KeyError("El archivo config.ini debe tener una sección [Program]")
    section = config["Program"]
    required = ["SchoolName", "Sede", "Date", "Prueba"]
    for key in required:
        if key not in section:
            raise KeyError(f"Falta clave requerida '{key}' en [Program]")
    return dict(section)
