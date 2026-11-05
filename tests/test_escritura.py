import pandas as pd
from openpyxl import Workbook

from tests.base import PruebaBase

from services.conciliacion import escritura


def libro(filas):
    wb = Workbook()
    ws = wb.active
    ws.title = "MOVS_AUD"
    ws.append(["Fecha", "Tipo", "Numero", "Nit/C.C.", "DETALLE", "Cuenta", "143505"])
    for nit, tipo in filas:
        ws.append(["2026-09-10", tipo, "1", nit, f"TERCERO {nit}", "x", 1000])
    return wb


def morado(ws, fila):
    return str(ws.cell(row=fila, column=1).fill.fgColor.rgb or "").upper().endswith(escritura.COLOR_AUTORRETENEDOR.upper())


class TestEscritura(PruebaBase):
    LISTADO = {"860001942": "AUTORRETENEDORES", "890911625": "AGRICOLA HUEVOS", "800237815": "RGB"}

    def _procesar(self, filas):
        wb = libro(filas)
        df = pd.DataFrame([{"BASE": 1000, "ALERTA": ""} for _ in filas])
        escritura.inyectar_columnas(wb, "MOVS_AUD", df, self.LISTADO)
        ws = wb["MOVS_AUD"]
        col_nit = next(c for c in range(1, ws.max_column + 1) if ws.cell(row=1, column=c).value == "Nit/C.C.")
        datos = [(str(ws.cell(row=r, column=col_nit).value), morado(ws, r)) for r in range(2, len(filas) + 2)]
        return ws, datos

    def test_solo_autorretenedores_en_morado_y_arriba(self):
        _, datos = self._procesar([("999999999", "FC1"), ("890911625", "FC1"), ("860001942", "FC1"), ("800237815", "FC1")])
        self.assertEqual(datos[0], ("860001942", True))
        self.assertTrue(all(not m for _, m in datos[1:]))

    def test_nit_guardado_como_decimal(self):
        _, datos = self._procesar([(860001942.0, "FC1"), ("999999999", "FC1")])
        self.assertTrue(datos[0][1], "el NIT 860001942.0 debe reconocerse como autorretenedor")

    def test_recuadro_del_bloque(self):
        ws, _ = self._procesar([("860001942", "FC1"), ("860001942", "FC2"), ("999999999", "FC1")])
        self.assertEqual(ws.cell(row=2, column=1).border.top.style, "medium")
        self.assertEqual(ws.cell(row=3, column=1).border.bottom.style, "medium")
        self.assertIsNone(ws.cell(row=2, column=1).border.bottom.style)

    def test_reproceso_no_duplica_ni_pierde_marcas(self):
        wb = libro([("860001942", "FC1"), ("999999999", "FC1")])
        df = pd.DataFrame([{"BASE": 1000, "ALERTA": ""}] * 2)
        for _ in range(2):
            escritura.inyectar_columnas(wb, "MOVS_AUD", df, self.LISTADO)
        ws = wb["MOVS_AUD"]
        encabezados = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        self.assertEqual(encabezados.count("AUTORRETENCION"), 1)
        self.assertTrue(morado(ws, 2))
        self.assertFalse(morado(ws, 3))