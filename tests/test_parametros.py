from tests.base import PruebaBase

from core.config import BASES_MINIMAS_RETENCION_UVT, CONCEPTOS_BASE_MINIMA
from data import parametros_repo


class TestParametros(PruebaBase):
    def test_fabrica_trae_todos_los_conceptos(self):
        for desde, bases in BASES_MINIMAS_RETENCION_UVT:
            with self.subTest(desde=desde):
                self.assertEqual(set(bases), set(CONCEPTOS_BASE_MINIMA))

    def test_json_antiguo_sin_agropecuario_usa_fabrica(self):
        datos = {"uvt": {"2026": 52374}, "bases": [
            {"desde": d, **{c: v for c, v in b.items() if c != "agropecuario"}} for d, b in BASES_MINIMAS_RETENCION_UVT]}
        leido = parametros_repo._desde_json(datos)
        self.assertEqual([b["agropecuario"] for _, b in leido["bases"]], [92, 70, 92, 70])

    def test_valor_guardado_prima_sobre_fabrica(self):
        datos = {"uvt": {"2026": 52374}, "bases": [{"desde": "2025-06-01", "compras": 10, "agropecuario": 50}]}
        self.assertEqual(parametros_repo._desde_json(datos)["bases"][0][1]["agropecuario"], 50)

    def test_vigencia_intermedia(self):
        self.assertEqual(parametros_repo._base_fabrica("2026-06-01", "agropecuario"), 92)
        self.assertEqual(parametros_repo._base_fabrica("1999-01-01", "agropecuario"), 92)