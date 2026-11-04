"""Perfil tributario por empresa auditada (JSON local, clave = NIT sin DV)."""
import json

from core.config import ARCHIVO_PERFILES
from core.errors import logger
from shared.texto import normalizar_nit


def _cargar():
    try:
        with open(ARCHIVO_PERFILES, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}
    except (OSError, json.JSONDecodeError) as e:
        logger.warning(f"No se pudo leer {ARCHIVO_PERFILES}: {e}")
        return {}


def obtener(nit):
    return _cargar().get(normalizar_nit(nit))


def guardar(nit, perfil):
    clave = normalizar_nit(nit)
    if not clave:
        return
    datos = _cargar()
    datos[clave] = {k: v for k, v in perfil.items() if k != "nit"}
    try:
        ARCHIVO_PERFILES.parent.mkdir(parents=True, exist_ok=True)
        with open(ARCHIVO_PERFILES, "w", encoding="utf-8") as f:
            json.dump(datos, f, indent=2, ensure_ascii=False)
    except OSError as e:
        logger.warning(f"No se pudo guardar el perfil de la empresa {clave}: {e}")