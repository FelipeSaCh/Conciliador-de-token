from tests.base import PruebaBase

from core.config import ETIQUETAS_TRATAMIENTO, GUIA_ALERTAS, TRATAMIENTOS_LISTADO
from services.conciliacion.alertas import resumen_tratamientos, tratamiento_listado


class TestTratamientoListado(PruebaBase):
    CASOS = {
        "AUTORRETENEDORES": ("AUTORRET", "no_sujeto"),
        "REGIMEN SIMPLE ": ("SIMPLE", "no_sujeto"),
        "PATRIMONIOS AUTONOMOS": ("PATRIMONIO", "no_sujeto"),
        "FIDEICOMISOS": ("FIDEICOMISO", "no_sujeto"),
        "ABSTENERSE DE PRACTICAR RETENCIÓN": ("ABSTEN", "no_sujeto"),
        "NO APLICA POR DECRETO 849 DE 2020": ("NO APLICA POR", "no_sujeto"),
        "AGRICOLA HUEVOS": ("AGRICOLA", "agropecuario"),
        "agricola": ("AGRICOLA", "agropecuario"),
        "REGIMEN ESPECIAL ZESE": ("ZESE", "tarifa_reducida"),
        "ZOMAC": ("ZOMAC", "tarifa_reducida"),
        "solicitamos ... se aplique el cincuenta por ciento (50%) de la tarifa": ("50%", "tarifa_reducida"),
        "NO APLICAR RETENCION EN LA FUENTE SE ACOGE A LA LEY 1819 DE 2016": ("LEY 1819", "observacion"),
        "REVISAR FACTURA- REGISTRO EN BOLSA": ("BOLSA", "observacion"),
        "REGISTRA EN BOLSA": ("BOLSA", "observacion"),
        "RGB": ("RGB", "observacion"),
        "SOCIEDAD SIN ANIMO DE LUCRO": ("SIN ANIMO", "observacion"),
        "REVISAR FACTURA": ("REVISAR FACTURA", "observacion"),
        "COMENTARIO DESCONOCIDO": ("", "observacion"),
    }

    def test_clasificacion_de_cada_comentario(self):
        for comentario, esperado in self.CASOS.items():
            with self.subTest(comentario=comentario):
                clave, tratamiento, _ = tratamiento_listado(comentario)
                self.assertEqual((clave, tratamiento), esperado)

    def test_ley_1819_no_se_confunde_con_no_aplica(self):
        self.assertEqual(tratamiento_listado("NO APLICAR RETENCION ... LEY 1819")[1], "observacion")

    def test_comentario_vacio(self):
        self.assertEqual(tratamiento_listado("  "), ("", "", ""))

    def test_todo_tratamiento_tiene_etiqueta_y_explicacion(self):
        for clave, tratamiento, explicacion in TRATAMIENTOS_LISTADO:
            with self.subTest(clave=clave):
                self.assertIn(tratamiento, ETIQUETAS_TRATAMIENTO)
                self.assertTrue(explicacion.strip())
                self.assertEqual(clave, clave.upper())

    def test_guia_documenta_cada_comentario(self):
        seccion = next(a for s, _, a in GUIA_ALERTAS if s.startswith("Comentarios del listado"))
        documentados = {nombre for nombre, _, _ in seccion}
        for clave, _, _ in TRATAMIENTOS_LISTADO:
            self.assertIn(f'Comentario con "{clave}"', documentados)

    def test_resumen_de_consola(self):
        lineas = resumen_tratamientos({"1": "AUTORRETENEDORES", "2": "AUTORRETENEDORES", "3": "AGRICOLA PAPA",
                                       "4": "XYZ"})
        self.assertIn("4 terceros", lineas[0])
        texto = "\n".join(lineas)
        self.assertIn('"AUTORRET" (2 terceros): No se exige retención', texto)
        self.assertIn('"AGRICOLA" (1 tercero)', texto)
        self.assertIn('"XYZ" (1 tercero): Solo informa', texto)
        self.assertNotIn("1.5 %", texto)