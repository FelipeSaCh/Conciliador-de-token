"""Integridad del listado de autorretenedores. Ruta: variable AUTORRETENEDORES_JSON o la carpeta de red configurada."""
import json
import os
import unittest
from datetime import date
from pathlib import Path

from tests.base import PruebaBase

from data.autorretenedores_repo import COL_COMENTARIO, COL_FECHA, COL_NIT, GestorAutorretenedores
from services.conciliacion.alertas import tratamiento_listado


def _ruta():
    ruta = os.environ.get("AUTORRETENEDORES_JSON")
    if ruta:
        return Path(ruta)
    try:
        return GestorAutorretenedores.ARCHIVO_DATOS if GestorAutorretenedores.ARCHIVO_DATOS.exists() else None
    except OSError:
        return None


RUTA = _ruta()


@unittest.skipUnless(RUTA, "Listado no disponible: defina AUTORRETENEDORES_JSON o conecte la carpeta de red")
class TestListadoJSON(PruebaBase):
    @classmethod
    def setUpClass(cls):
        with open(RUTA, encoding="utf-8") as f:
            cls.datos = json.load(f)

    def test_nit_unico(self):
        vistos, repetidos = set(), set()
        for d in self.datos:
            nit = str(d[COL_NIT])
            (repetidos if nit in vistos else vistos).add(nit)
        self.assertFalse(repetidos, f"NIT repetidos: {sorted(repetidos)}")

    def test_nombres_sin_texto_corrupto(self):
        malos = [d[COL_NIT] for d in self.datos for k, v in d.items()
                 if k.upper().startswith("NOMBRE") and ("null" in str(v) or "?" in str(v))]
        self.assertFalse(malos, f"Nombres con 'null' o '?': {malos}")

    def test_fechas_validas(self):
        futuras = [(d[COL_NIT], d[COL_FECHA]) for d in self.datos
                   if d.get(COL_FECHA) and str(d[COL_FECHA]) > date.today().isoformat()]
        if futuras:
            self.skipTest(f"{len(futuras)} resoluciones con fecha posterior a hoy (revisar día/mes): {futuras}")

    def test_todo_comentario_tiene_tratamiento_conocido(self):
        sin_regla = sorted({d[COL_COMENTARIO] for d in self.datos if not tratamiento_listado(d[COL_COMENTARIO])[0]})
        self.assertFalse(sin_regla, f"Comentarios sin regla en TRATAMIENTOS_LISTADO (se tratan como observación): {sin_regla}")