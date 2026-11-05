import pandas as pd

from tests.base import PruebaBase

from core.config import GUIA_ALERTAS
from services.conciliacion.alertas import alertas_retencion, explicacion_alerta, tipo_alerta

GASTO = "143505 - MERCANCIAS"
RETE = "236540 - Compras (2.5%)"
LISTADO = {
    "860001942": "AUTORRETENEDORES",
    "901000001": "REGIMEN SIMPLE",
    "890911625": "AGRICOLA HUEVOS",
    "830053812": "FIDEICOMISOS",
    "901292639": "NO APLICA POR DECRETO 849 DE 2020",
    "900927871": "solicitamos ... cincuenta por ciento (50%) ...",
    "901241091": "ZOMAC",
    "800237815": "RGB",
    "1017126086": "NO APLICAR RETENCION EN LA FUENTE SE ACOGE A LA LEY 1819 DE 2016",
}


def evaluar(filas, fecha="2026-09-10", cuenta_ret=RETE):
    """filas: [(nit, base, retención)] -> lista de alertas (lista de textos por comprobante)."""
    df = pd.DataFrame([{"Nit/C.C.": n, "Tipo": "FC1", "Fecha": fecha, GASTO: b, cuenta_ret: r} for n, b, r in filas])
    res = alertas_retencion(df, [GASTO], [], [cuenta_ret], [], LISTADO, {})
    return [[t for t in a.split(" / ") if t] for a in res["ALERTA"]]


def tipos(alertas):
    return [tipo_alerta(t) for t in alertas]


class TestAlertasRetencion(PruebaBase):
    # ---------------------------------------------------------------- terceros comunes
    def test_tercero_comun_sin_retencion_sobre_base(self):
        [a] = evaluar([("999999999", 5_000_000, 0)])
        self.assertEqual(tipos(a), ["Sin retención"])
        self.assertIn("compras 10 UVT = $523.740", a[0])

    def test_tercero_comun_bajo_base_sin_alerta(self):
        self.assertEqual(evaluar([("999999999", 300_000, 0)]), [[]])

    def test_tercero_comun_retencion_bajo_base(self):
        [a] = evaluar([("999999999", 300_000, 7_500)])
        self.assertIn("Retención con base bajo el mínimo", tipos(a))

    def test_tarifa_distinta_a_la_cuenta(self):
        [a] = evaluar([("999999999", 5_000_000, 50_000)])
        self.assertIn("Tarifa aplicada distinta a la de la cuenta", tipos(a))

    # ---------------------------------------------------------------- autorretenedor y SIMPLE: sin novedad visible
    def test_autorretenedor_sin_retencion_no_alerta(self):
        self.assertEqual(evaluar([("860001942", 5_000_000, 0)]), [[]])

    def test_autorretenedor_con_retencion(self):
        [a] = evaluar([("860001942", 5_000_000, 125_000)])
        self.assertEqual(tipos(a), ["Retención practicada a no sujeto"])
        self.assertIn("es autorretenedor", a[0])

    def test_simple_por_tipo_de_documento(self):
        df = pd.DataFrame([{"Nit/C.C.": "888888888", "Tipo": "FC REGIMEN SIMPLE", "Fecha": "2026-09-10",
                            GASTO: 5_000_000, RETE: 125_000}])
        a = alertas_retencion(df, [GASTO], [], [RETE], [], LISTADO, {})["ALERTA"].iat[0]
        self.assertIn("practicada a no sujeto (RÉGIMEN SIMPLE)", a)
        self.assertIn("según el tipo de documento", a)

    # ---------------------------------------------------------------- no sujetos nuevos: siempre visibles
    def test_fideicomiso_sin_retencion_deja_constancia(self):
        [a] = evaluar([("830053812", 5_000_000, 0)])
        self.assertEqual(tipos(a), ["Tratamiento especial por listado de autorretenedores"])
        self.assertIn("no se le exige retención", a[0])

    def test_fideicomiso_con_retencion(self):
        [a] = evaluar([("830053812", 5_000_000, 125_000)])
        self.assertIn("Retención practicada a no sujeto", tipos(a))
        self.assertIn("Tratamiento especial por listado de autorretenedores", tipos(a))
        self.assertNotIn("Tarifa aplicada distinta a la de la cuenta", tipos(a))

    def test_no_aplica_por_decreto(self):
        [a] = evaluar([("901292639", 5_000_000, 0)])
        self.assertEqual(tipos(a), ["Tratamiento especial por listado de autorretenedores"])
        self.assertIn('"NO APLICA POR"', a[0])

    # ---------------------------------------------------------------- agropecuarios
    def test_agro_sin_retencion_sobre_base(self):
        [a] = evaluar([("890911625", 8_000_000, 0)])
        self.assertEqual(sorted(tipos(a)), ["Sin retención", "Tratamiento especial por listado de autorretenedores"])
        sin = next(t for t in a if t.startswith("Sin retención"))
        self.assertIn("productos agropecuarios 70 UVT = $3.666.180", sin)
        self.assertIn("retención esperada al 1,5 %: $120.000", sin)

    def test_agro_bajo_base_no_requiere_retencion(self):
        [a] = evaluar([("890911625", 2_000_000, 0)])
        self.assertEqual(tipos(a), ["Tratamiento especial por listado de autorretenedores"])
        self.assertIn("no la alcanza: no requiere retención", a[0])

    def test_agro_retencion_correcta(self):
        [a] = evaluar([("890911625", 8_000_000, 120_000)])
        self.assertEqual(tipos(a), ["Tratamiento especial por listado de autorretenedores"])
        self.assertIn("esta compra la supera", a[0])

    def test_agro_tarifa_de_compras_generales(self):
        [a] = evaluar([("890911625", 8_000_000, 200_000)])
        tarifa = next(t for t in a if t.startswith("Tarifa aplicada distinta"))
        self.assertIn("vs 1,50 % de productos agropecuarios", tarifa)

    def test_agro_base_92_uvt_durante_suspension(self):
        [a] = evaluar([("890911625", 4_000_000, 0)], fecha="2026-06-15")
        self.assertIn("92 UVT", a[0])
        self.assertNotIn("Sin retención", tipos(a))

    def test_agro_base_70_uvt_antes_de_suspension(self):
        [a] = evaluar([("890911625", 4_000_000, 0)], fecha="2025-08-01")
        self.assertIn("Sin retención", tipos(a))

    # ---------------------------------------------------------------- tarifa reducida
    def test_tarifa_reducida_no_compara_tarifa(self):
        [a] = evaluar([("900927871", 5_000_000, 62_500)])
        self.assertEqual(tipos(a), ["Tratamiento especial por listado de autorretenedores"])
        self.assertIn("no se compara la tarifa", a[0])

    def test_tarifa_reducida_exige_retencion(self):
        [a] = evaluar([("901241091", 5_000_000, 0)])
        self.assertIn("Sin retención", tipos(a))
        self.assertIn("Tratamiento especial por listado de autorretenedores", tipos(a))

    # ---------------------------------------------------------------- observación
    def test_observacion_explica_y_no_alerta_sin_retencion(self):
        for nit, palabra in (("800237815", "RGB"), ("1017126086", "art. 383 E.T.")):
            with self.subTest(nit=nit):
                [a] = evaluar([(nit, 5_000_000, 0)])
                self.assertEqual(tipos(a), ["Tercero con observación en listado de autorretenedores"])
                self.assertIn(palabra, a[0])
                self.assertIn('No se evalúa "Sin retención"', a[0])

    # ---------------------------------------------------------------- NIT en distintos formatos
    def test_nit_con_decimal_y_digito_de_verificacion(self):
        for nit in (860001942.0, "860001942-1", " 860001942 "):
            with self.subTest(nit=nit):
                self.assertEqual(evaluar([(nit, 5_000_000, 0)]), [[]])

    # ---------------------------------------------------------------- guía y acta
    def test_toda_alerta_esta_documentada_en_la_guia(self):
        filas = [("999999999", 5_000_000, 0), ("999999999", 300_000, 7_500), ("999999999", 5_000_000, 50_000),
                 ("860001942", 5_000_000, 125_000), ("830053812", 5_000_000, 125_000), ("890911625", 8_000_000, 0),
                 ("900927871", 5_000_000, 62_500), ("800237815", 5_000_000, 0)]
        nombres_guia = {n for _, _, alertas in GUIA_ALERTAS for n, _, _ in alertas}
        for alertas in evaluar(filas):
            for texto in alertas:
                with self.subTest(texto=texto[:60]):
                    nombre = tipo_alerta(texto)
                    self.assertIn(nombre, nombres_guia)
                    self.assertTrue(explicacion_alerta(nombre)[2])