import json
import os
import sys
import pandas as pd
from pathlib import Path


# 📁 Ruta de red corporativa donde vive el JSON compartido
CARPETA_RED = Path(r"\\SRV-CONTAB\Servidor de Archivos\2025 SERVIDOR\APLICATIVO FELIPE\Documentos_auditor")
NOMBRE_ARCHIVO_JSON = "autorretenedores.json"


class GestorAutorretenedores:
    """Clase backend para manejar la lógica de los autorretenedores."""

    CARPETA_DATOS = CARPETA_RED
    ARCHIVO_DATOS = CARPETA_DATOS / NOMBRE_ARCHIVO_JSON

    # Datos iniciales SOLO si el JSON no existe todavía
    DATOS_PREDETERMINADOS = [
        # {"nit": "900123456", "nombre": "EMPRESA EJEMPLO S.A.S."},
        # {"nit": "800987654", "nombre": "OTRA EMPRESA LTDA."},
    ]

    # ------------------------------------------------------------------
    # Utilidades internas
    # ------------------------------------------------------------------
    @classmethod
    def _asegurar_carpeta(cls):
        """Verifica que la carpeta de red esté accesible. Lanza error si no."""
        try:
            if not cls.CARPETA_DATOS.exists():
                raise FileNotFoundError(
                    f"No se puede acceder a la carpeta de red:\n{cls.CARPETA_DATOS}\n\n"
                    "Verifica que estés conectado a la red corporativa y que tengas permisos."
                )
        except OSError as e:
            raise ConnectionError(
                f"Error al acceder a la carpeta de red:\n{cls.CARPETA_DATOS}\n\n{e}"
            )

    @classmethod
    def _asegurar_archivo(cls):
        """Si el JSON no existe, lo crea con los predeterminados en la carpeta de red."""
        cls._asegurar_carpeta()
        if not cls.ARCHIVO_DATOS.exists():
            try:
                with open(cls.ARCHIVO_DATOS, "w", encoding="utf-8") as f:
                    json.dump(cls.DATOS_PREDETERMINADOS, f, indent=4, ensure_ascii=False)
            except PermissionError:
                raise PermissionError(
                    f"No tienes permisos para escribir en:\n{cls.ARCHIVO_DATOS}"
                )

    @classmethod
    def _cargar_datos(cls):
        cls._asegurar_archivo()
        try:
            with open(cls.ARCHIVO_DATOS, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError as e:
            raise ValueError(
                f"El archivo '{cls.ARCHIVO_DATOS.name}' está corrupto o mal formado:\n{e}"
            )
        except PermissionError:
            raise PermissionError(
                f"No tienes permisos para leer:\n{cls.ARCHIVO_DATOS}"
            )

    @classmethod
    def _guardar_datos(cls, datos):
        cls._asegurar_carpeta()
        try:
            with open(cls.ARCHIVO_DATOS, "w", encoding="utf-8") as f:
                json.dump(datos, f, indent=4, ensure_ascii=False)
        except PermissionError:
            raise PermissionError(
                f"No tienes permisos para escribir en:\n{cls.ARCHIVO_DATOS}"
            )

    # ------------------------------------------------------------------
    # API pública
    # ------------------------------------------------------------------
    @classmethod
    def obtener_todos(cls):
        return cls._cargar_datos()

    @classmethod
    def obtener_columnas(cls):
        datos = cls._cargar_datos()
        if not datos:
            return ["nit", "nombre"]
        return list(datos[0].keys())

    @classmethod
    def agregar(cls, nit, nombre):
        datos = cls._cargar_datos()
        if any(d.get("nit") == nit for d in datos):
            raise ValueError(f"El NIT '{nit}' ya se encuentra registrado.")
        datos.append({"nit": nit, "nombre": nombre})
        cls._guardar_datos(datos)

    @classmethod
    def eliminar(cls, nit):
        datos = cls._cargar_datos()
        nuevos = [d for d in datos if d.get("nit") != nit]
        if len(nuevos) == len(datos):
            return False
        cls._guardar_datos(nuevos)
        return True

    @classmethod
    def actualizar(cls, nit, nuevo_nombre):
        datos = cls._cargar_datos()
        for d in datos:
            if d.get("nit") == nit:
                d["nombre"] = nuevo_nombre
                cls._guardar_datos(datos)
                return True
        return False

    @classmethod
    def exportar_excel(cls, ruta_salida):
        datos = cls._cargar_datos()
        if not datos:
            raise ValueError("No hay autorretenedores registrados para exportar.")
        df = pd.DataFrame(datos)
        df.to_excel(ruta_salida, index=False, engine='openpyxl')

    @classmethod
    def ruta_archivo(cls) -> Path:
        return cls.ARCHIVO_DATOS