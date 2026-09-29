"""Reglas duras de la fase 4 y utilidades de taxonomía / RETEG / partes relacionadas."""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from auditoria.config import (  # noqa: E402
    COHERENTE, COHERENTE_VER_NOTA, INCONCLUSO, NO_CORRESPONDE, detectar_parte_relacionada,
    es_traslado_reteg, importar_script, nits_coinciden, normalizar_etiqueta, similitud_nombres,
)
from auditoria.modelos_llm import ErrorLLM  # noqa: E402

m04 = importar_script("04_clasificacion_auditoria")


def senales(**kw):
    base = {"veredicto_valor": "EXACTO_DOCUMENTO", "calidad": {"documento_legible": True},
            "moneda": {"usd": False}, "tercero": {"coincide_algun_documento": True},
            "partes_relacionadas_documento": [], "anomalias_datos": []}
    base.update(kw)
    return base


def resp(etiqueta, **kw):
    r = {"etiqueta": etiqueta, "categoria_ar": "h_operacion_administracion_impuestos",
         "observacion": "Orden de Operación a X por $1.000.000.", "requisito_cierre": "", "tags": [],
         "anomalia_datos": ""}
    r.update(kw)
    return r


class Taxonomia(unittest.TestCase):
    def test_normalizar_etiqueta(self):
        self.assertEqual(normalizar_etiqueta("coherente - ver nota"), COHERENTE_VER_NOTA)
        self.assertEqual(normalizar_etiqueta("COHERENTE–VER NOTA"), COHERENTE_VER_NOTA)
        self.assertEqual(normalizar_etiqueta(" NO CORRESPONDE "), NO_CORRESPONDE)
        self.assertIsNone(normalizar_etiqueta("RECONOCIBLE"))
        self.assertIsNone(normalizar_etiqueta("AR7"))

    def test_etiqueta_fuera_de_taxonomia_deja_fila_pendiente(self):
        with self.assertRaises(ErrorLLM):
            m04.ajustar_clasificacion(resp("RECONOCIBLE"), senales())


class ReglasDuras(unittest.TestCase):
    def test_no_inventar_cifras(self):
        s = senales(veredicto_valor="SIN_CIFRAS_DOCUMENTO", calidad={"documento_legible": False})
        r = m04.ajustar_clasificacion(resp(COHERENTE), s)
        self.assertEqual(r["etiqueta"], INCONCLUSO)
        self.assertIn("Para cerrarlo", r["observacion"])

    def test_coherente_con_diferencia_sin_explicar_baja_a_ver_nota(self):
        r = m04.ajustar_clasificacion(resp(COHERENTE), senales(veredicto_valor="SIN_COMBINACION_ESTANDAR"))
        self.assertEqual(r["etiqueta"], COHERENTE_VER_NOTA)
        self.assertIn("diferencia-no-explicada", r["tags"])

    def test_usd_sin_trm(self):
        r = m04.ajustar_clasificacion(resp(COHERENTE), senales(moneda={"usd": True, "trm_confirmable": False}))
        self.assertEqual(r["etiqueta"], COHERENTE_VER_NOTA)
        self.assertIn("verificar-usd-trm", r["tags"])

    def test_inconcluso_dice_que_falta(self):
        r = m04.ajustar_clasificacion(resp(INCONCLUSO, requisito_cierre="la factura completa legible"), senales())
        self.assertIn("la factura completa legible", r["observacion"])

    def test_reteg_es_no_corresponde_sin_modelo(self):
        ctx = {"tercero": "BANCOLOMBIA SA", "cuenta_contable": "AHO 47760972700 PA ERNESTO CO",
               "etiqueta_soporte": "RETEG OP 3807", "valor_registrado": -20_000_000, "es_reteg": True}
        r = m04.clasificar_fila(ctx, {}, {}, cliente=None)
        self.assertEqual(r["etiqueta"], NO_CORRESPONDE)
        self.assertIn("RETEG OP 3807", r["observacion"])

    def test_patron_reteg(self):
        self.assertTrue(es_traslado_reteg("BANCOLOMBIA SA", "AHO 47760972700 PA ERNESTO CO", None))
        self.assertTrue(es_traslado_reteg("BANCOLOMBIA SA", "AHO 47703093651 P A AEROPUERTO", None))
        self.assertTrue(es_traslado_reteg("X", "Y", "RETEG OP 12"))
        self.assertFalse(es_traslado_reteg("BANCOLOMBIA SA", "COMISIONES", "OP 12"))

    def test_partes_relacionadas_y_terceros(self):
        self.assertEqual(detectar_parte_relacionada("GRUPO AEROPORTUARIO DEL CARIBE S.A.S."), "GAC")
        self.assertEqual(detectar_parte_relacionada(None, "900.913.341-1"), "NAB")
        self.assertIsNone(detectar_parte_relacionada("CHUBB SEGUROS COLOMBIA SA"))
        self.assertTrue(nits_coinciden("900.525.886-1", "900525886"))
        self.assertFalse(nits_coinciden("830004748", "860002062"))
        self.assertGreaterEqual(similitud_nombres("CENTRO INDUSTRIAL MECANICO CIMEC LTDA",
                                                  "CENTRO INDUSTRIAL MECANICO CIMEC SAS"), 0.8)


class Hallazgos(unittest.TestCase):
    def test_agrupa_por_tag_no_por_rank(self):
        regs = [{"rank": r, "op": str(r), "tercero": "BANCOLOMBIA SA", "valor": -1e6,
                 "clasificacion": {"etiqueta": NO_CORRESPONDE, "observacion": "traslado", "tags": ["reteg-interno"]}}
                for r in (3531, 3590, 3676)]
        entradas = m04.construir_hallazgos(regs, pendientes=[{"rank": 9, "op": "1", "motivo": "sin PDF"}])
        titulos = [t for t, _ in entradas]
        self.assertEqual(sum("RETEG" in t for t in titulos), 1)
        self.assertTrue(any("#3531 / OP 3531" in d and "#3676" in d for t, d in entradas if "RETEG" in t))
        self.assertTrue(any("pendientes" in t.lower() for t in titulos))


if __name__ == "__main__":
    unittest.main()
