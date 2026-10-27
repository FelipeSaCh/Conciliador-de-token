import json
import os
from datetime import date, datetime

import pandas as pd

from core.config import ARCHIVO_AUTORRETENEDORES, CARPETA_AUTORRETENEDORES
from core.errors import logger

COL_NIT = "NIT"
PREFIJO_NOMBRE = "NOMBRE O RAZON SOCIAL"
COL_NOMBRE = "NOMBRE O RAZON SOCIAL - AUTORRETENEDORES"
COL_RESOLUCION = "NUMERO DE RESOLUCION"
COL_FECHA = "FECHA DE RESOLUCION"
COL_COMENTARIO = "COMENT/COMENTARIO"
CLAVES_LEGADO = ("nit", "Nit", "nombre", "NOMBRE", "comentario", "COMENTARIO")
FORMATOS_FECHA = ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y")


def _texto(valor):
    if valor is None:
        return ""
    if isinstance(valor, float):
        if valor != valor:
            return ""
        if valor.is_integer():
            return str(int(valor))
    s = str(valor).strip()
    return "" if s.lower() in ("nan", "none") else s


def _digitos(texto):
    return "".join(ch for ch in texto if ch.isdigit())


def _clave_nit(valor):
    return _digitos(_texto(valor).split("-")[0])


def _entero_o_texto(texto):
    return int(texto) if texto.isdigit() else texto


def _normalizar_fecha(valor):
    s = _texto(valor).split(" ")[0]
    if not s:
        return ""
    for fmt in FORMATOS_FECHA:
        try:
            fecha = datetime.strptime(s, fmt).date()
        except ValueError:
            continue
        if fecha > date.today():
            raise ValueError("La fecha de resolución no puede ser posterior a hoy.")
        return fecha.isoformat()
    raise ValueError("La fecha de resolución debe tener el formato AAAA-MM-DD o DD/MM/AAAA.")


class GestorAutorretenedores:
    """Repositorio del listado compartido de autorretenedores (JSON en red)."""

    CARPETA_DATOS = CARPETA_AUTORRETENEDORES
    ARCHIVO_DATOS = CARPETA_DATOS / ARCHIVO_AUTORRETENEDORES
    DATOS_PREDETERMINADOS = []
    COMENTARIO_DEFECTO = "AUTORRETENEDORES"

    # ------------------------------------------------------------ archivo
    @classmethod
    def _asegurar_carpeta(cls):
        try:
            if not cls.CARPETA_DATOS.exists():
                raise FileNotFoundError(
                    f"No se puede acceder a la carpeta de red:\n{cls.CARPETA_DATOS}\n\n"
                    "Verifica que estés conectado a la red corporativa y que tengas permisos."
                )
        except OSError as e:
            if isinstance(e, FileNotFoundError):
                raise
            raise ConnectionError(f"Error al acceder a la carpeta de red:\n{cls.CARPETA_DATOS}\n\n{e}")

    @classmethod
    def _escribir(cls, datos):
        temporal = cls.ARCHIVO_DATOS.with_suffix(".tmp")
        try:
            with open(temporal, "w", encoding="utf-8") as f:
                json.dump(datos, f, indent=4, ensure_ascii=False)
            os.replace(temporal, cls.ARCHIVO_DATOS)
        except PermissionError:
            raise PermissionError(f"No tienes permisos para escribir en:\n{cls.ARCHIVO_DATOS}")

    @classmethod
    def _cargar_datos(cls):
        cls._asegurar_carpeta()
        if not cls.ARCHIVO_DATOS.exists():
            cls._escribir(cls.DATOS_PREDETERMINADOS)
        try:
            with open(cls.ARCHIVO_DATOS, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError as e:
            raise ValueError(f"El archivo '{cls.ARCHIVO_DATOS.name}' está corrupto o mal formado:\n{e}")
        except PermissionError:
            raise PermissionError(f"No tienes permisos para leer:\n{cls.ARCHIVO_DATOS}")

    @classmethod
    def _guardar_datos(cls, datos):
        cls._asegurar_carpeta()
        col_nombre = cls.columna_nombre(datos)
        cls._escribir([cls._canonico(d, col_nombre) for d in datos])

    # ------------------------------------------------------------ formato
    @staticmethod
    def columna_nombre(datos):
        for d in datos:
            for k in d:
                if str(k).upper().startswith(PREFIJO_NOMBRE):
                    return k
        return COL_NOMBRE

    @staticmethod
    def _desde_json(d, col_nombre):
        def valor(*claves):
            for k in claves:
                v = _texto(d.get(k))
                if v:
                    return v
            return ""

        return {
            "nit": valor(COL_NIT, "nit", "Nit"),
            "nombre": valor(col_nombre, "nombre", "NOMBRE"),
            "resolucion": valor(COL_RESOLUCION),
            "fecha": valor(COL_FECHA).split(" ")[0],
            "comentario": valor(COL_COMENTARIO, "comentario", "COMENTARIO"),
        }

    @classmethod
    def _canonico(cls, d, col_nombre):
        r = cls._desde_json(d, col_nombre)
        base = {
            COL_NIT: _entero_o_texto(_clave_nit(r["nit"]) or r["nit"]),
            col_nombre: r["nombre"],
            COL_RESOLUCION: _entero_o_texto(r["resolucion"]),
            COL_FECHA: r["fecha"],
            COL_COMENTARIO: r["comentario"],
        }
        extras = {k: v for k, v in d.items()
                  if k not in base and k not in CLAVES_LEGADO and not str(k).upper().startswith(PREFIJO_NOMBRE)}
        return {**base, **extras}

    @classmethod
    def _validar(cls, registro):
        nit = _clave_nit(registro.get("nit"))
        if not 5 <= len(nit) <= 10:
            raise ValueError("El NIT debe tener entre 5 y 10 dígitos, sin dígito de verificación.")
        nombre = " ".join(_texto(registro.get("nombre")).split()).upper()
        if not nombre:
            raise ValueError("Debes ingresar el nombre o razón social.")
        resolucion = _texto(registro.get("resolucion"))
        if resolucion and not resolucion.isdigit():
            raise ValueError("El número de resolución debe ser numérico.")
        return {
            "nit": nit,
            "nombre": nombre,
            "resolucion": resolucion,
            "fecha": _normalizar_fecha(registro.get("fecha")),
            "comentario": " ".join(_texto(registro.get("comentario")).split()).upper() or cls.COMENTARIO_DEFECTO,
        }

    @staticmethod
    def _a_json(r, col_nombre):
        return {
            COL_NIT: int(r["nit"]),
            col_nombre: r["nombre"],
            COL_RESOLUCION: _entero_o_texto(r["resolucion"]),
            COL_FECHA: r["fecha"],
            COL_COMENTARIO: r["comentario"],
        }

    @classmethod
    def _indice(cls, datos, nit):
        clave = _clave_nit(nit)
        col_nombre = cls.columna_nombre(datos)
        for i, d in enumerate(datos):
            if _clave_nit(cls._desde_json(d, col_nombre)["nit"]) == clave:
                return i
        return None

    # ------------------------------------------------------------ consultas
    @classmethod
    def obtener_todos(cls):
        return cls._cargar_datos()

    @classmethod
    def obtener_registros(cls):
        datos = cls._cargar_datos()
        col_nombre = cls.columna_nombre(datos)
        return [cls._desde_json(d, col_nombre) for d in datos]

    @classmethod
    def obtener_columnas(cls, datos=None):
        datos = cls._cargar_datos() if datos is None else datos
        return list(datos[0].keys()) if datos else [COL_NIT, COL_NOMBRE, COL_RESOLUCION, COL_FECHA, COL_COMENTARIO]

    # ------------------------------------------------------------ escritura
    @classmethod
    def agregar(cls, registro):
        r = cls._validar(registro)
        datos = cls._cargar_datos()
        if cls._indice(datos, r["nit"]) is not None:
            raise ValueError(f"El NIT '{r['nit']}' ya se encuentra registrado.")
        datos.append(cls._a_json(r, cls.columna_nombre(datos)))
        cls._guardar_datos(datos)
        return r

    @classmethod
    def actualizar(cls, nit_original, registro):
        r = cls._validar(registro)
        datos = cls._cargar_datos()
        i = cls._indice(datos, nit_original)
        if i is None:
            raise ValueError(f"El NIT '{nit_original}' ya no existe en el listado. Recarga la vista.")
        j = cls._indice(datos, r["nit"])
        if j is not None and j != i:
            raise ValueError(f"El NIT '{r['nit']}' ya se encuentra registrado.")
        col_nombre = cls.columna_nombre(datos)
        extras = {k: v for k, v in cls._canonico(datos[i], col_nombre).items()
                  if k not in (COL_NIT, col_nombre, COL_RESOLUCION, COL_FECHA, COL_COMENTARIO)}
        datos[i] = {**cls._a_json(r, col_nombre), **extras}
        cls._guardar_datos(datos)
        return r

    @classmethod
    def eliminar(cls, nits):
        claves = {_clave_nit(n) for n in ([nits] if isinstance(nits, (str, int)) else nits)}
        datos = cls._cargar_datos()
        col_nombre = cls.columna_nombre(datos)
        nuevos = [d for d in datos if _clave_nit(cls._desde_json(d, col_nombre)["nit"]) not in claves]
        eliminados = len(datos) - len(nuevos)
        if eliminados:
            cls._guardar_datos(nuevos)
        return eliminados

    @classmethod
    def exportar_excel(cls, ruta_salida):
        datos = cls._cargar_datos()
        if not datos:
            raise ValueError("No hay autorretenedores registrados para exportar.")
        pd.DataFrame(datos).to_excel(ruta_salida, index=False, engine='openpyxl')

    @classmethod
    def como_dataframe(cls):
        """DataFrame normalizado NIT/COMENTARIO para el motor de conciliación. Nunca lanza."""
        def _valor(d, *claves):
            for k in claves:
                if k in d and str(d[k]).strip().lower() not in ('', 'nan', 'none'):
                    return str(d[k]).strip()
            return ''

        try:
            registros = [
                {
                    'NIT': _valor(d, 'nit', 'NIT', 'Nit'),
                    'COMENTARIO': _valor(d, 'COMENT/COMENTARIO', 'comentario', 'COMENTARIO', 'nombre', 'NOMBRE'),
                }
                for d in cls._cargar_datos()
            ]
        except Exception as e:
            logger.warning(f"No se pudo cargar el listado de autorretenedores desde JSON: {e}")
            registros = []
        return pd.DataFrame(registros, columns=['NIT', 'COMENTARIO'])

    @classmethod
    def ruta_archivo(cls):
        return cls.ARCHIVO_DATOS