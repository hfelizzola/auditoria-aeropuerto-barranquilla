"""Prueba de punta a punta del orquestador con un modelo falso (sin red ni API reales).

Cubre: selección Pareto, creación del Excel, sub-lotes, regla RETEG sin modelo, fila sin PDF
que queda EN BLANCO, guardado por fila, hallazgos agrupados, control de calidad, reanudación
sin repetir llamadas al modelo y detención limpia por cuota agotada.
Requiere PyMuPDF para generar/rasterizar PDFs de prueba; si no está, se omite.
"""

import os
import re
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd
import requests

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from auditoria.config import importar_script  # noqa: E402
from auditoria.lote_excel import HOJA_CONTROL, HOJA_HALLAZGOS, LoteExcel  # noqa: E402
from auditoria.modelos_llm import ClienteLLM, ErrorCuotaLLM  # noqa: E402

try:
    import fitz  # noqa: F401
except ImportError:  # pragma: no cover
    fitz = None

orq = importar_script("00_orquestar_auditoria")
m01 = importar_script("01_pareto_universalidad")
m02 = importar_script("02_descargar_soportes_excel")

COLS = ["IDENTIFICACION TERCERO", "TERCERO", "NUMERO DE CUENTA CONTABLE", "NOMBRE DE CUENTA CONTABLE", "VALOR BASE",
        "VALOR IVA", "FECHA DE PAGO O DESEMBOLSO", "VALOR DEBITADO O ACREDITADO", "NOMBRE DE LA CUENTA BANCARIA",
        "TIPO DE SOPORTE", "SOPORTE VALIDADO", "NUMERO OP", "Validación", "URL"]
URL = "https://example.invalid/soporte{}.pdf"


def universalidad():
    filas = [
        ["900111222", "PROVEEDOR UNO SAS", 5130, "MANTENIMIENTO", 20_000_000, 3_800_000, "2024-03-04", -23_800_000,
         "AHO 1 PA E.CORT OPEX", "CON ORDEN DE OPERACIÓN", "OP 100", 100, "OK", URL.format(1)],
        ["890903938", "BANCOLOMBIA SA", 1, "AHO 47760972700 PA ERNESTO CO", None, None, "2024-03-05", -15_000_000,
         "AHO 1 PA E.CORT OPEX", "CON ORDEN DE OPERACIÓN", "RETEG OP 100", 100, None, URL.format(2)],
        ["890903938", "BANCOLOMBIA SA", 2, "GRAVAMEN", 57_731, None, "2024-03-05", -57_731,
         "AHO 1", "GMF - SERVICIOS BANCARIOS", "GMF", "05/2015", None, None],
        ["900333444", "PROVEEDOR TRES SAS", 5130, "MANTENIMIENTO", 11_520_900, 2_188_971, "2024-03-06", -12_978_293,
         "AHO 1 PA E.CORT OPEX", "CON ORDEN DE OPERACIÓN", "OP 300", 300, None, URL.format(3)],
        ["900444555", "OTRO SAS", 5130, "TERMINAL", 10_000_000, 1_900_000, "2024-03-07", -11_900_000,
         "AHO 1 PA E.CORT OPEX", "CON ORDEN DE OPERACIÓN", "OP 400", 400, None, URL.format(4)],
    ]
    df = pd.DataFrame(filas, columns=COLS)
    df["FECHA DE PAGO O DESEMBOLSO"] = pd.to_datetime(df["FECHA DE PAGO O DESEMBOLSO"])
    return df


HECHOS = {
    "100": ("PROVEEDOR UNO SAS", "900.111.222-3", 20_000_000, 3_800_000, 23_800_000),
    "400": ("ADELTE SERVICIOS COLOMBIA SAS", "900.999.888-1", 10_000_000, 1_900_000, 11_900_000),
}


class ModeloFalso(ClienteLLM):
    proveedor, modelo = "falso", "v1"

    def __init__(self, cuota_hasta=None):
        self.llamadas = []
        self.cuota_hasta = cuota_hasta

    def generar_json(self, instrucciones, texto, imagenes_jpeg, esquema, max_tokens=16000):
        if self.cuota_hasta is not None and len(self.llamadas) >= self.cuota_hasta:
            raise ErrorCuotaLLM("429 simulado")
        assert "Resultado de validación" not in texto and '"Validación"' not in texto  # independencia
        es_extraccion = "documentos" in esquema["properties"]
        self.llamadas.append("extraccion" if es_extraccion else "clasificacion")
        if es_extraccion:
            op = re.search(r"orden de pago (\d+)", texto).group(1)
            nombre, nit, base, iva, neto = HECHOS[op]
            return {"numero_op_visible": op, "paginas_legibles": 1, "documento_legible": True, "nota_calidad": "",
                    "documentos": [{"tipo": "orden_operacion_fiduciaria", "pagina_inicio": 1, "pagina_fin": 1,
                                    "legible": True, "numero_documento": op, "fecha": "2024-03-01",
                                    "emisor_nombre": "FIDUCIARIA", "emisor_nit": "", "beneficiario_nombre": nombre,
                                    "beneficiario_nit": nit, "concepto": "Mantenimiento", "moneda": "COP",
                                    "trm_impresa": None, "valores": [
                                        {"tipo": "subtotal_base", "monto": base, "moneda": "COP", "pagina": 1, "texto_literal": ""},
                                        {"tipo": "iva", "monto": iva, "moneda": "COP", "pagina": 1, "texto_literal": ""},
                                        {"tipo": "neto_a_pagar", "monto": neto, "moneda": "COP", "pagina": 1, "texto_literal": ""}]}],
                    "partes_relacionadas": [], "endoso_o_cesion": {"presente": False, "detalle": "", "pagina": 0},
                    "menciona_servicio_deuda": {"presente": False, "detalle": "", "pagina": 0},
                    "observaciones_objetivas": ""}
        if "ADELTE" in texto:
            return {"etiqueta": "TERCERO NO COINCIDE", "categoria_ar": "h_operacion_administracion_impuestos",
                    "observacion": "La orden gira a ADELTE SERVICIOS COLOMBIA SAS, no a OTRO SAS.",
                    "requisito_cierre": "", "tags": ["tercero-no-coincide"], "anomalia_datos": ""}
        return {"etiqueta": "COHERENTE", "categoria_ar": "h_operacion_administracion_impuestos",
                "observacion": "Orden de Operación a PROVEEDOR UNO SAS, $23.800.000, concilia exacto.",
                "requisito_cierre": "", "tags": [], "anomalia_datos": ""}


@unittest.skipIf(fitz is None, "PyMuPDF no instalado")
class Punta_a_punta(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        base = Path(self.tmp.name)
        os.environ["AUDITORIA_DIR_ESTADO"] = str(base / "estado")
        self.excel = base / "Validacion_Soportes_Rango1-4_UNIVERSALIDAD_ABAS1.xlsx"
        self.soportes = base / "soportes"
        self.soportes.mkdir()
        self.df = universalidad()
        self.ranking = m01.rankear_op_url(self.df)
        # Sin red real: sesión anónima (nunca SAML con credenciales del .env)
        self._sesion_original = m02.crear_sesion_autenticada
        m02.crear_sesion_autenticada = lambda **kw: requests.Session()
        self._crear_cliente_original = orq.crear_cliente
        rk = self.ranking.dropna(subset=["RANK"]).set_index("RANK")
        for rank in (1, 2, 4):  # el rank 3 (OP 300) no tiene PDF
            r = rk.loc[rank]
            self._pdf(self.soportes / m02.nombre_archivo_soporte(rank, r.OP_NORM, r.TERCERO))

    def tearDown(self):
        m02.crear_sesion_autenticada = self._sesion_original
        orq.crear_cliente = self._crear_cliente_original
        os.environ.pop("AUDITORIA_DIR_ESTADO", None)
        self.tmp.cleanup()

    @staticmethod
    def _pdf(ruta):
        doc = fitz.open()
        pag = doc.new_page()
        for i in range(60):
            pag.insert_text((40, 20 + 12 * i), f"ORDEN DE OPERACION GENERAL linea {i} valor $23.800.000")
        doc.save(str(ruta))

    def _correr(self, modelo, *extra):
        orq.crear_cliente = lambda proveedor, modelo_id=None: modelo
        args = orq.parsear_args(["--modo", "pareto", "--desde", "1", "--hasta", "4", "--excel", str(self.excel),
                                 "--dir-soportes", str(self.soportes), "--tamano-lote", "2", "--reintentos", "1",
                                 "--pausa", "0", *extra])
        lote = orq.fase1(args, self.ranking, self.df, Path("UNIV.xlsx"))
        o = orq.Orquestador(args, lote, self.ranking, {"02": m02, "03": importar_script("03_extraccion_documental"),
                                                       "04": importar_script("04_clasificacion_auditoria")})
        o._descargador = m02.Descargador(self.soportes, reintentos=1, buscar_en_otros_lotes=False)
        o._descargador.respaldo = m02.IndiceRespaldo([])
        return o.ejecutar()

    def _resultados(self):
        return {f.rank: f.resultado for f in LoteExcel.abrir(self.excel).filas()}

    def test_lote_completo_y_reanudacion(self):
        modelo = ModeloFalso()
        self.assertEqual(self._correr(modelo), 0)
        self.assertEqual(self._resultados(), {1: "COHERENTE", 2: "NO CORRESPONDE", 3: None, 4: "TERCERO NO COINCIDE"})
        # RETEG sin modelo: 2 extracciones + 2 clasificaciones (ranks 1 y 4)
        self.assertEqual(sorted(modelo.llamadas), ["clasificacion"] * 2 + ["extraccion"] * 2)

        wb = LoteExcel.abrir(self.excel).wb
        hallazgos = " ".join(str(c.value) for fila in wb[HOJA_HALLAZGOS].iter_rows() for c in fila if c.value)
        self.assertIn("RETEG", hallazgos)
        self.assertIn("TERCERO NO COINCIDE", hallazgos)
        self.assertIn("#3 / OP 300", hallazgos)  # pendiente sin PDF
        control = " ".join(str(c.value) for fila in wb[HOJA_CONTROL].iter_rows() for c in fila if c.value)
        self.assertIn("OP validada en más de un rank", control)  # OP 100 en ranks 1 y 2

        # Reanudar: nada que re-extraer ni re-clasificar; la fila sin PDF sigue en blanco
        modelo2 = ModeloFalso()
        self.assertEqual(self._correr(modelo2), 0)
        self.assertEqual(modelo2.llamadas, [])
        self.assertIsNone(self._resultados()[3])

    def test_cuota_agotada_detiene_y_conserva_lo_hecho(self):
        modelo = ModeloFalso(cuota_hasta=2)  # alcanza para el rank 1 (extracción + clasificación)
        self.assertEqual(self._correr(modelo), 2)
        res = self._resultados()
        self.assertEqual(res[1], "COHERENTE")
        self.assertEqual(res[2], "NO CORRESPONDE")
        self.assertIsNone(res[4])
        # Al reanudar con cuota se completa sin repetir el rank 1
        modelo2 = ModeloFalso()
        self.assertEqual(self._correr(modelo2), 0)
        self.assertEqual(self._resultados()[4], "TERCERO NO COINCIDE")
        self.assertEqual(len(modelo2.llamadas), 2)


if __name__ == "__main__":
    unittest.main()
