"""Pruebas del cuadre de retenciones con casos reales de la ronda 3501-4000.

Ejecutar:  python -m unittest discover -s tests      (o: python -m pytest tests)
Para ajustar tarifas con casos nuevos, agregue aquí el caso y su conclusión manual.
"""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from auditoria import reconciliacion as rec  # noqa: E402


class CasosReales(unittest.TestCase):
    """base, IVA y valor registrado tomados de UNIVERSALIDAD; conclusión de la revisión manual."""

    def test_3505_bs_ingenieria_retefuente_y_reteiva(self):
        r = rec.reconciliar(11_520_900, 2_188_971, -12_978_293)
        self.assertEqual(r.estado, rec.COMBINACION)
        c = r.combinaciones[0]
        self.assertEqual((c.retefuente_pct, c.reteiva_pct), (3.5, 15.0))
        self.assertAlmostEqual(c.total_retenido, 731_577.15, places=2)

    def test_3514_serviparamo_reteiva15_reteica8(self):
        r = rec.reconciliar(11_128_000, 2_114_320, -12_836_148)
        self.assertEqual(r.estado, rec.COMBINACION)
        c = r.combinaciones[0]
        self.assertEqual((c.retefuente_pct, c.reteiva_pct, c.reteica_por_mil), (0.0, 15.0, 8.0))

    def test_3515_honorarios_11_15_8(self):
        r = rec.reconciliar(12_600_000, 2_016_000, -12_826_800)
        c = r.combinaciones[0]
        self.assertEqual((c.retefuente_pct, c.reteiva_pct, c.reteica_por_mil), (11.0, 15.0, 8.0))
        self.assertEqual(c.total_retenido, 1_789_200.0)

    def test_1677_banco_occidente_exacto(self):
        r = rec.reconciliar(51_725_560, 9_827_856, -61_553_416)
        self.assertEqual(r.estado, rec.EXACTO)

    def test_3504_mc_universal_no_cierra_con_tarifas_estandar(self):
        # Revisión manual: "diferencia $708.700 sin combinación estándar"
        r = rec.reconciliar(11_803_800, 1_888_608, -12_983_708)
        self.assertNotEqual(r.estado, rec.COMBINACION)
        self.assertIn(r.estado, (rec.COMBINACION_ATIPICA, rec.SIN_COMBINACION))

    def test_3509_sin_iva_diferencia_sin_explicar(self):
        r = rec.reconciliar(13_673_474.79, None, -12_900_719.79)
        self.assertEqual(r.estado, rec.SIN_COMBINACION)
        self.assertTrue(r.sin_desglose_iva)
        self.assertIn("diferencia sin explicar", r.describir())


class Reglas(unittest.TestCase):
    def test_tolerancia_1_50(self):
        self.assertEqual(rec.reconciliar(1000.0, 0.0, 998.6).estado, rec.EXACTO)
        self.assertNotEqual(rec.reconciliar(1000.0, 0.0, 998.4).estado, rec.EXACTO)

    def test_registrado_mayor_que_bruto_nunca_se_fuerza(self):
        r = rec.reconciliar(10_000_000, 1_900_000, -12_000_000)
        self.assertEqual(r.estado, rec.SIN_COMBINACION)
        self.assertIn("MAYOR", r.nota)

    def test_sin_datos(self):
        self.assertEqual(rec.reconciliar(None, None, 100).estado, rec.SIN_DATOS)

    def test_combinaciones_ordenadas_por_simplicidad(self):
        combos = rec.buscar_combinaciones(10_000_000, 1_900_000, 250_000)  # 2,5% exacto
        self.assertEqual(combos[0].componentes, 1)
        self.assertEqual(combos[0].retefuente_pct, 2.5)

    def test_retenciones_impresas(self):
        self.assertTrue(rec.cuadrar_retenciones_impresas(1_190_000, {"retefuente": 25_000, "reteica": 8_000}, -1_157_000))
        self.assertFalse(rec.cuadrar_retenciones_impresas(1_190_000, {"retefuente": 25_000}, -1_157_000))
        self.assertIsNone(rec.cuadrar_retenciones_impresas(None, {"retefuente": 1}, 1))


if __name__ == "__main__":
    unittest.main()
