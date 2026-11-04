"""Parámetros tributarios editables: UVT por año, bases mínimas por vigencia y categorías de cuentas.

Se guardan en la carpeta compartida; si no está disponible se usa la copia local y, en último caso, config.py.
"""
import json
import re
from datetime import date, datetime

from core.config import (
    ARCHIVO_CATEGORIAS, ARCHIVO_CATEGORIAS_LOCAL, ARCHIVO_CONCEPTOS, ARCHIVO_CONCEPTOS_LOCAL, ARCHIVO_PARAMETROS,
    ARCHIVO_PARAMETROS_LOCAL, BASES_MINIMAS_RETENCION_UVT, CATEGORY_COLORS, CONCEPTO_POR_CUENTA_RETENCION,
    CONCEPTO_RETENCION_POR_CUENTA, CONCEPTOS_BASE_MINIMA, ETIQUETAS_CONCEPTOS, REGLAS_CATEGORIAS_PUC, UVT_POR_ANIO,
)
from core.errors import logger

FORMATOS_FECHA = ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y")


def por_defecto():
    return {
        "uvt": dict(UVT_POR_ANIO),
        "bases": [(desde, dict(bases)) for desde, bases in BASES_MINIMAS_RETENCION_UVT],
    }


def _desde_json(datos):
    uvt = {int(a): float(v) for a, v in datos["uvt"].items()}
    bases = [(b["desde"], {c: float(b.get(c, 0)) for c in CONCEPTOS_BASE_MINIMA}) for b in datos["bases"]]
    if not uvt or not bases:
        raise ValueError("Parámetros incompletos")
    return {"uvt": uvt, "bases": sorted(bases, key=lambda b: b[0])}


def cargar():
    for ruta in (ARCHIVO_PARAMETROS, ARCHIVO_PARAMETROS_LOCAL):
        try:
            with open(ruta, "r", encoding="utf-8") as f:
                return _desde_json(json.load(f))
        except FileNotFoundError:
            continue
        except (OSError, ValueError, KeyError, TypeError) as e:
            logger.warning(f"No se pudieron leer los parámetros tributarios de {ruta}: {e}")
    return por_defecto()


def normalizar_fecha(valor):
    texto = str(valor).strip()
    for fmt in FORMATOS_FECHA:
        try:
            return datetime.strptime(texto, fmt).date().isoformat()
        except ValueError:
            continue
    raise ValueError(f"Fecha inválida: '{texto}'. Usa AAAA-MM-DD o DD/MM/AAAA.")


def validar(uvt, bases):
    """uvt: {año: valor}; bases: [(desde, {concepto: uvts})]. Devuelve los datos normalizados o lanza ValueError."""
    if not uvt:
        raise ValueError("Debe existir al menos un año con valor de UVT.")
    limpio_uvt = {}
    for anio, valor in uvt.items():
        anio, valor = int(anio), float(valor)
        if not 2000 <= anio <= date.today().year + 1:
            raise ValueError(f"Año fuera de rango: {anio}.")
        if valor <= 0:
            raise ValueError(f"La UVT del año {anio} debe ser mayor a cero.")
        limpio_uvt[anio] = valor
    if not bases:
        raise ValueError("Debe existir al menos una vigencia de bases mínimas.")
    limpio_bases = {}
    for desde, valores in bases:
        desde = normalizar_fecha(desde)
        if desde in limpio_bases:
            raise ValueError(f"La vigencia {desde} está repetida.")
        fila = {c: float(valores.get(c, 0)) for c in CONCEPTOS_BASE_MINIMA}
        if any(v < 0 for v in fila.values()):
            raise ValueError(f"Las bases de la vigencia {desde} no pueden ser negativas.")
        limpio_bases[desde] = fila
    return {"uvt": limpio_uvt, "bases": sorted(limpio_bases.items())}


def guardar(datos):
    """Guarda en la carpeta compartida; si falla, localmente. Devuelve la ruta usada."""
    contenido = {
        "uvt": {str(a): v for a, v in sorted(datos["uvt"].items())},
        "bases": [{"desde": d, **b} for d, b in datos["bases"]],
    }
    for ruta in (ARCHIVO_PARAMETROS, ARCHIVO_PARAMETROS_LOCAL):
        try:
            ruta.parent.mkdir(parents=True, exist_ok=True)
            temporal = ruta.with_suffix(".tmp")
            with open(temporal, "w", encoding="utf-8") as f:
                json.dump(contenido, f, indent=2, ensure_ascii=False)
            temporal.replace(ruta)
            return ruta
        except OSError as e:
            logger.warning(f"No se pudieron guardar los parámetros en {ruta}: {e}")
    raise OSError("No se pudieron guardar los parámetros tributarios en ninguna ubicación.")


def uvt_vigente(datos, anio):
    return datos["uvt"].get(anio, datos["uvt"][max(datos["uvt"])])


# ---------------------------------------------------------------- reglas de categorías de cuentas
def reglas_por_defecto():
    return {cat: {"incluir": list(REGLAS_CATEGORIAS_PUC.get(cat, {}).get("incluir", ())),
                  "excluir": list(REGLAS_CATEGORIAS_PUC.get(cat, {}).get("excluir", ()))}
            for cat in CATEGORY_COLORS}


def cargar_reglas():
    reglas = reglas_por_defecto()
    for ruta in (ARCHIVO_CATEGORIAS, ARCHIVO_CATEGORIAS_LOCAL):
        try:
            with open(ruta, "r", encoding="utf-8") as f:
                datos = json.load(f)
            for cat in reglas:
                if cat in datos:
                    reglas[cat] = {k: [str(p) for p in datos[cat].get(k, [])] for k in ("incluir", "excluir")}
            return reglas
        except FileNotFoundError:
            continue
        except (OSError, ValueError, AttributeError, TypeError) as e:
            logger.warning(f"No se pudieron leer las categorías de cuentas de {ruta}: {e}")
    return reglas


def validar_reglas(reglas):
    limpio = {}
    for cat, r in reglas.items():
        incluir = sorted({str(p).strip() for p in r.get("incluir", []) if str(p).strip()})
        excluir = sorted({str(p).strip() for p in r.get("excluir", []) if str(p).strip()})
        malos = [p for p in incluir + excluir if not re.fullmatch(r"\d{1,12}", p)]
        if malos:
            raise ValueError(f"Prefijos inválidos en {cat}: {', '.join(malos)}. Solo se permiten dígitos.")
        repetidos = set(incluir) & set(excluir)
        if repetidos:
            raise ValueError(f"En {cat} hay prefijos en Incluir y Excluir a la vez: {', '.join(sorted(repetidos))}.")
        limpio[cat] = {"incluir": incluir, "excluir": excluir}
    return limpio


def guardar_reglas(reglas):
    for ruta in (ARCHIVO_CATEGORIAS, ARCHIVO_CATEGORIAS_LOCAL):
        try:
            ruta.parent.mkdir(parents=True, exist_ok=True)
            temporal = ruta.with_suffix(".tmp")
            with open(temporal, "w", encoding="utf-8") as f:
                json.dump(reglas, f, indent=2, ensure_ascii=False)
            temporal.replace(ruta)
            return ruta
        except OSError as e:
            logger.warning(f"No se pudieron guardar las categorías en {ruta}: {e}")
    raise OSError("No se pudieron guardar las categorías de cuentas en ninguna ubicación.")


# ---------------------------------------------------------------- cuentas por concepto de base mínima
def conceptos_por_defecto():
    return {"gasto": dict(CONCEPTO_RETENCION_POR_CUENTA), "retencion": dict(CONCEPTO_POR_CUENTA_RETENCION)}


def cargar_conceptos():
    for ruta in (ARCHIVO_CONCEPTOS, ARCHIVO_CONCEPTOS_LOCAL):
        try:
            with open(ruta, "r", encoding="utf-8") as f:
                datos = json.load(f)
            return {k: {str(p): str(c) for p, c in datos.get(k, {}).items()} for k in ("gasto", "retencion")}
        except FileNotFoundError:
            continue
        except (OSError, ValueError, AttributeError, TypeError) as e:
            logger.warning(f"No se pudieron leer los conceptos de retención de {ruta}: {e}")
    return conceptos_por_defecto()


def validar_conceptos(conceptos):
    limpio = {}
    for tipo in ("gasto", "retencion"):
        mapa = {}
        for prefijo, concepto in conceptos.get(tipo, {}).items():
            prefijo = str(prefijo).strip()
            if not re.fullmatch(r"\d{1,12}", prefijo):
                raise ValueError(f"Prefijo inválido: '{prefijo}'. Solo se permiten dígitos.")
            if concepto not in ETIQUETAS_CONCEPTOS:
                raise ValueError(f"Concepto desconocido para {prefijo}: {concepto}.")
            if tipo == "retencion" and concepto == "mixto":
                raise ValueError(f"Una cuenta de retención no puede ser 'Mixto' ({prefijo}).")
            mapa[prefijo] = concepto
        limpio[tipo] = dict(sorted(mapa.items()))
    return limpio


def guardar_conceptos(conceptos):
    for ruta in (ARCHIVO_CONCEPTOS, ARCHIVO_CONCEPTOS_LOCAL):
        try:
            ruta.parent.mkdir(parents=True, exist_ok=True)
            temporal = ruta.with_suffix(".tmp")
            with open(temporal, "w", encoding="utf-8") as f:
                json.dump(conceptos, f, indent=2, ensure_ascii=False)
            temporal.replace(ruta)
            return ruta
        except OSError as e:
            logger.warning(f"No se pudieron guardar los conceptos en {ruta}: {e}")
    raise OSError("No se pudieron guardar los conceptos de retención en ninguna ubicación.")


def concepto_de(codigo, mapa, defecto):
    """Concepto por el prefijo más largo: (concepto, prefijo que decide o '')."""
    codigo = str(codigo).strip()
    coincidencias = [p for p in mapa if codigo.startswith(p)]
    if not coincidencias:
        return defecto, ''
    prefijo = max(coincidencias, key=len)
    return mapa[prefijo], prefijo