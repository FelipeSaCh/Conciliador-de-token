import json
import re
from functools import lru_cache

from core.config import PUC_FILE
from core.errors import logger


def ordenar_por_prefijos(items, jerarquia):
    """Ordena por la posición del primer prefijo coincidente en `jerarquia`; el resto al final, alfabético."""
    def clave(item):
        s = str(item).strip()
        prioridad = next((i for i, p in enumerate(jerarquia) if s.startswith(p)), 999)
        return prioridad, s
    return sorted(items, key=clave)


def codigo_cuenta(texto):
    """Código numérico inicial de un serial ('240811019 Iva compras' -> '240811019')."""
    m = re.match(r'\s*(\d+)', str(texto))
    return m.group(1) if m else ''


@lru_cache(maxsize=1)
def _puc():
    try:
        with open(PUC_FILE, encoding='utf-8') as f:
            return {str(c['codigo']): str(c['nombre']).strip() for c in json.load(f)}
    except (OSError, ValueError, KeyError) as e:
        logger.warning(f"No se pudo cargar el PUC ({PUC_FILE}): {e}")
        return {}


def nombre_cuenta(texto):
    """Nombre PUC de la cuenta (o de su nivel superior más cercano)."""
    codigo, puc = codigo_cuenta(texto), _puc()
    while codigo:
        if codigo in puc:
            return puc[codigo]
        codigo = codigo[:-2] if len(codigo) > 4 else codigo[:-1]
    return ''


@lru_cache(maxsize=1)
def reglas_categorias():
    from data.parametros_repo import cargar_reglas
    return cargar_reglas()


def recargar_reglas():
    reglas_categorias.cache_clear()


def cuentas_puc():
    """{código: nombre} del PUC cargado."""
    return dict(_puc())


def clasificar(texto, categoria, reglas=None):
    """(pertenece, prefijo que decide o '') según el prefijo más largo entre 'incluir' y 'excluir'."""
    codigo = codigo_cuenta(texto)
    r = (reglas or reglas_categorias()).get(categoria)
    if not codigo or not r:
        return False, ''
    candidatos = [(len(p), True, p) for p in r['incluir'] if codigo.startswith(p)]
    candidatos += [(len(p), False, p) for p in r['excluir'] if codigo.startswith(p)]
    if not candidatos:
        return False, ''
    _, pertenece, prefijo = max(candidatos)
    return pertenece, prefijo


def pertenece_a_categoria(texto, categoria, reglas=None):
    return clasificar(texto, categoria, reglas)[0]