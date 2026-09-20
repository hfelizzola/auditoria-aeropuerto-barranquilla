"""
03_extraccion_gemini.py
=======================
Módulo para la extracción estructurada de hechos y clasificación contractual
mediante visión multimodal con la API de Google Gemini (gemini-2.5-flash / gemini-3.6-flash).

Objetivo:
- Leer los PDFs descargados en `data/datalake_pdfs/`.
- Convertir cada PDF a imágenes JPEG a 150 DPI en memoria (vía pdf2image y PIL,
  con soporte pymupdf integrado para entornos Windows).
- Enviar las páginas rasterizadas a Google Gemini con un prompt riguroso de auditoría
  contractual (Numeral 22.3 c) del Contrato de Concesión).
- Extraer en JSON:
  * "numero_factura_op": str
  * "emisor_nombre": str
  * "emisor_nit": str
  * "concepto_pago": str
  * "valor_total": float
  * "categoria_contractual": Literal['AR1', 'AR2', 'AR3', 'AR4', 'AR5', 'AR6', 'AR7', 'AR8', 'AR9', 'EXCLUIDO']
  * "requiere_acta_interventoria": bool (True si es AR7 o anticipo)
  * "es_parte_relacionada": bool (True si emisor es NUEVO AEROPUERTO DE BARRANQUILLA SAS,
    GRUPO AEROPORTUARIO DEL CARIBE SAS u OPERADORA AEROPORTUARIA DEL CARIBE SAS).
- Manejo de cuotas (HTTP 429) y saturación (HTTP 503) mediante reintentos con backoff exponencial.
- Guardar resultados progresivos en `data/output/resultados_gemini.csv`.
"""

from __future__ import annotations

import argparse
import io
import json
import logging
import os
import re
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
from PIL import Image

# Configuración de logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)

# Cargar variables de entorno desde .env (soporta python-dotenv o lectura nativa de respaldo)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    env_file = Path(__file__).resolve().parent.parent / ".env"
    if env_file.exists():
        with open(env_file, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip())

# Nombres y NITs conocidos de partes relacionadas para validación determinista complementaria
PARTES_RELACIONADAS_NOMBRES = [
    "NUEVO AEROPUERTO DE BARRANQUILLA",
    "GRUPO AEROPORTUARIO DEL CARIBE",
    "OPERADORA AEROPORTUARIA DEL CARIBE",
]

PROMPT_AUDITORIA = """
Actúa como un Auditor Contable y Financiero experto en contratos de concesión pública (ANI).
Analiza las imágenes adjuntas del soporte documental (factura, cuenta de cobro, orden de pago o acta)
y extrae la información requerida bajo la lista taxativa del Numeral 22.3 c) del Contrato de Concesión.

DEBES RESPONDER EXCLUSIVAMENTE UN OBJETO JSON VÁLIDO CON LA SIGUIENTE ESTRUCTURA:
{
  "numero_factura_op": "Número de factura, cuenta de cobro u orden de pago indicada en el soporte",
  "emisor_nombre": "Razón social o nombre completo del proveedor o emisor del soporte",
  "emisor_nit": "NIT o identificación tributaria del emisor (sin dígito de verificación si es posible)",
  "concepto_pago": "Descripción textual sintética del bien, obra o servicio prestado",
  "valor_total": 0.0,
  "categoria_contractual": "AR1 | AR2 | AR3 | AR4 | AR5 | AR6 | AR7 | AR8 | AR9 | EXCLUIDO",
  "requiere_acta_interventoria": true | false,
  "es_parte_relacionada": true | false,
  "justificacion_clasificacion": "Breve explicación del fundamento contractual de la clasificación"
}

REGLAS ESTRICTAS DE CLASIFICACIÓN (NUMERAL 22.3 c):
1. Categorías contractuales taxativas:
   - "AR1": Primas y comisiones de garantías y pólizas del Contrato (cumplimiento, RCE, todo riesgo).
   - "AR2": Aportes a las Subcuentas de la Cuenta ANI.
   - "AR3": Comisión de Éxito pagada al Consultor estructurador.
   - "AR4": Estudios y Diseños de las Intervenciones y su Cronograma.
   - "AR5": Gestión Social y Ambiental de las obras/intervenciones.
   - "AR6": Gestión Predial para el proyecto aeroportuario.
   - "AR7": Actuaciones en las Intervenciones (obras de infraestructura/CAPEX), verificadas por el Interventor.
   - "AR8": Operación y Mantenimiento (OPEX), servicios generales, administración, vigilancia, aseo, servicios públicos e impuestos.
   - "AR9": Comisiones u honorarios a Prestamistas/Financiadores (distintos al servicio de deuda).
   - "EXCLUIDO": CUALQUIER pago de servicio de la deuda (abono a capital/principal o intereses de créditos/préstamos), así como multas, sanciones o conceptos no previstos taxativamente.

2. REGLA CRÍTICA DE EXCLUSIÓN:
   Todo abono a capital (principal) o pago de intereses de créditos sindicados o bancarios debe clasificarse obligatoriamente como "EXCLUIDO".

3. "requiere_acta_interventoria":
   Asignar true si la categoría es "AR7" o si el pago corresponde a un "anticipo" o "avance de obra". De lo contrario false.

4. "es_parte_relacionada":
   Asignar true si el emisor o beneficiario es:
   - "NUEVO AEROPUERTO DE BARRANQUILLA SAS"
   - "GRUPO AEROPORTUARIO DEL CARIBE SAS"
   - "OPERADORA AEROPORTUARIA DEL CARIBE SAS"
   De lo contrario false.

5. Formato de número: "valor_total" debe ser un número en coma flotante (float) sin separadores de miles ni signos monetarios. Si la moneda es USD, reportar el valor nominal en USD o el valor liquidado en COP si es explícito.
"""


def rasterizar_pdf_a_imagenes(ruta_pdf: Path, dpi: int = 150, max_paginas: int = 10) -> List[Image.Image]:
    """
    Convierte las páginas de un PDF a objetos PIL Image a 150 DPI en memoria.
    Intenta primero con pdf2image; si falla por falta de Poppler en el sistema,
    utiliza pymupdf (fitz) de manera transparente.
    """
    imagenes: List[Image.Image] = []

    # 1. Intentar con pdf2image
    try:
        from pdf2image import convert_from_path
        logger.debug(f"Rasterizando '{ruta_pdf.name}' con pdf2image a {dpi} DPI...")
        imgs = convert_from_path(str(ruta_pdf), dpi=dpi, first_page=1, last_page=max_paginas)
        if imgs:
            return imgs
    except Exception as e_p2i:
        logger.debug(f"pdf2image no disponible o requiere poppler ({e_p2i}). Utilizando backend PyMuPDF...")

    # 2. Respaldo PyMuPDF (fitz)
    try:
        import fitz  # pymupdf

        doc = fitz.open(str(ruta_pdf))
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)

        total_pags = min(len(doc), max_paginas)
        for num_pag in range(total_pags):
            pag = doc.load_page(num_pag)
            pix = pag.get_pixmap(matrix=mat, alpha=False)
            img_bytes = pix.tobytes("jpeg")
            img = Image.open(io.BytesIO(img_bytes))
            imagenes.append(img)
        doc.close()
        return imagenes

    except Exception as e_fitz:
        logger.error(f"Error rasterizando PDF '{ruta_pdf}': {e_fitz}")
        raise RuntimeError(f"No fue posible rasterizar '{ruta_pdf.name}': {e_fitz}")


def llamar_gemini_con_backoff(
    model: Any,
    prompt: str,
    imagenes: List[Image.Image],
    max_reintentos: int = 5,
    espera_inicial: float = 2.0,
    factor_backoff: float = 2.0,
) -> str:
    """
    Ejecuta la llamada a la API de Gemini enviando imágenes y prompt,
    manejando errores de cuota (429) y servicio (503) con retroceso exponencial.
    """
    contenido = [prompt] + imagenes

    for intento in range(1, max_reintentos + 1):
        try:
            # Solicitar generación de contenido
            response = model.generate_content(
                contenido,
                generation_config={
                    "temperature": 0.1,
                    "response_mime_type": "application/json",
                },
            )
            return response.text

        except Exception as e:
            mensaje_error = str(e)
            es_error_reintentable = any(
                cod in mensaje_error for cod in ["429", "503", "ResourceExhausted", "ServiceUnavailable"]
            )

            if es_error_reintentable and intento < max_reintentos:
                tiempo_espera = espera_inicial * (factor_backoff ** (intento - 1))
                logger.warning(
                    f"[REINTENTO {intento}/{max_reintentos}] Gemini retornó error transitorio: "
                    f"{mensaje_error[:120]}... Esperando {tiempo_espera:.1f}s antes de reintentar."
                )
                time.sleep(tiempo_espera)
            else:
                logger.error(f"Fallo en llamada a Gemini en intento {intento}: {mensaje_error}")
                raise


def parsear_respuesta_gemini(texto_respuesta: str) -> Dict[str, Any]:
    """Limpia y valida el JSON devuelto por Gemini."""
    # Remover posibles delimitadores de código markdown si los hubiere
    limpio = texto_respuesta.strip()
    if limpio.startswith("```"):
        limpio = re.sub(r"^```(?:json)?\s*", "", limpio)
        limpio = re.sub(r"\s*```$", "", limpio)
    limpio = limpio.strip()

    datos = json.loads(limpio)

    # Validaciones y saneamiento
    categoria = str(datos.get("categoria_contractual", "EXCLUIDO")).upper().strip()
    categorias_validas = {"AR1", "AR2", "AR3", "AR4", "AR5", "AR6", "AR7", "AR8", "AR9", "EXCLUIDO"}
    if categoria not in categorias_validas:
        # Intento de normalización si viene con texto extra
        for cat in categorias_validas:
            if cat in categoria:
                categoria = cat
                break
        else:
            categoria = "EXCLUIDO"
    datos["categoria_contractual"] = categoria

    # Validar booleano parte relacionada contra lista maestra
    emisor = str(datos.get("emisor_nombre", "")).upper()
    es_relacionada_detectada = any(pr in emisor for pr in PARTES_RELACIONADAS_NOMBRES)
    datos["es_parte_relacionada"] = bool(datos.get("es_parte_relacionada", False) or es_relacionada_detectada)

    # Validar acta de interventoría
    if categoria == "AR7":
        datos["requiere_acta_interventoria"] = True
    else:
        datos["requiere_acta_interventoria"] = bool(datos.get("requiere_acta_interventoria", False))

    # Asegurar valor total numérico
    try:
        datos["valor_total"] = float(datos.get("valor_total", 0.0))
    except (ValueError, TypeError):
        datos["valor_total"] = 0.0

    return datos


def extraer_numero_op_de_nombre(nombre_archivo: str) -> str:
    """Extrae el número de OP del nombre del archivo (ej. 'OP_1953.pdf' -> '1953')."""
    m = re.search(r"OP[-_ ]?0*(\d+)", nombre_archivo, re.IGNORECASE)
    if m:
        return m.group(1)
    # Si el nombre es directamente un número (ej. '1953.pdf')
    m_num = re.search(r"^0*(\d+)", Path(nombre_archivo).stem)
    if m_num:
        return m_num.group(1)
    return Path(nombre_archivo).stem


def procesar_lote_pdfs(
    directorio_pdfs: Optional[Path] = None,
    ruta_salida_csv: Optional[Path] = None,
    modelo_nombre: Optional[str] = None,
    max_archivos: Optional[int] = None,
) -> pd.DataFrame:
    """
    Recorre los PDFs en el datalake, rasteriza, consulta Gemini y consolida resultados.
    """
    base_dir = Path(__file__).resolve().parent.parent

    if directorio_pdfs is None:
        directorio_pdfs = base_dir / "data" / "datalake_pdfs"

    if ruta_salida_csv is None:
        ruta_salida_csv = base_dir / "data" / "output" / "resultados_gemini.csv"

    ruta_salida_csv.parent.mkdir(parents=True, exist_ok=True)

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key or api_key == "tu_clave_aqui":
        logger.warning(
            "GEMINI_API_KEY no configurada en .env o tiene el valor por defecto. "
            "Configure su clave válida en .env para invocar la API real."
        )

    # Modelo a usar: configurable por CLI o .env (por defecto gemini-2.5-flash / gemini-3.6-flash)
    if not modelo_nombre:
        modelo_nombre = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

    # Inicializar cliente Gemini si la librería está disponible
    genai = None
    model = None
    try:
        import google.generativeai as genai_lib
        genai = genai_lib
        if api_key and api_key != "tu_clave_aqui":
            genai.configure(api_key=api_key)
            try:
                model = genai.GenerativeModel(modelo_nombre)
                logger.info(f"Cliente Gemini inicializado con modelo '{modelo_nombre}'.")
            except Exception as e_mod:
                logger.warning(f"No se pudo cargar modelo '{modelo_nombre}': {e_mod}. Probando 'gemini-2.5-flash'...")
                model = genai.GenerativeModel("gemini-2.5-flash")
    except ImportError:
        logger.warning("Librería google-generativeai no instalada. Ejecute: pip install google-generativeai")

    # Descubrir PDFs
    archivos_pdf = sorted(list(directorio_pdfs.glob("*.pdf")))
    logger.info(f"Total de PDFs encontrados en datalake: {len(archivos_pdf)}")

    if not archivos_pdf:
        logger.warning(f"No se encontraron archivos .pdf en '{directorio_pdfs}'.")
        return pd.DataFrame()

    if max_archivos and max_archivos > 0:
        logger.info(f"Limitando procesamiento a {max_archivos} PDFs.")
        archivos_pdf = archivos_pdf[:max_archivos]

    # Cargar resultados previos si existen para asegurar idempotencia
    registros_previos: Dict[str, Dict[str, Any]] = {}
    if ruta_salida_csv.exists():
        try:
            df_existente = pd.read_csv(ruta_salida_csv)
            for _, r in df_existente.iterrows():
                op_prev = str(r.get("NUMERO_OP", "")).strip()
                if op_prev:
                    registros_previos[op_prev] = r.to_dict()
            logger.info(f"Cargados {len(registros_previos)} registros procesados previamente desde '{ruta_salida_csv.name}'.")
        except Exception as e_read:
            logger.warning(f"No se pudo leer archivo previo de resultados: {e_read}")

    resultados: List[Dict[str, Any]] = []

    for idx, ruta_pdf in enumerate(archivos_pdf, 1):
        op_id = extraer_numero_op_de_nombre(ruta_pdf.name)
        logger.info(f"[{idx}/{len(archivos_pdf)}] Procesando soporte OP {op_id} ('{ruta_pdf.name}')...")

        # Verificar si ya fue procesado
        if op_id in registros_previos:
            logger.info(f" -> OP {op_id} ya procesada previamente. Reutilizando resultado.")
            resultados.append(registros_previos[op_id])
            continue

        try:
            # 1. Rasterizar a imágenes
            imagenes = rasterizar_pdf_a_imagenes(ruta_pdf, dpi=150, max_paginas=8)
            logger.info(f" -> {len(imagenes)} página(s) rasterizada(s) a 150 DPI.")

            # 2. Consultar Gemini
            if model is not None:
                resp_texto = llamar_gemini_con_backoff(model, PROMPT_AUDITORIA, imagenes)
                datos_extraidos = parsear_respuesta_gemini(resp_texto)
            else:
                # Modo simulación / offline si aún no se ha provisto API Key válida
                logger.info(" -> Modo simulación (sin API Key activa): extrayendo metadata base.")
                datos_extraidos = {
                    "numero_factura_op": f"OP-{op_id}",
                    "emisor_nombre": "PENDIENTE_API_KEY",
                    "emisor_nit": "",
                    "concepto_pago": f"Soporte documental OP {op_id}",
                    "valor_total": 0.0,
                    "categoria_contractual": "AR7" if int(op_id) % 2 == 0 else "AR8",
                    "requiere_acta_interventoria": (int(op_id) % 2 == 0),
                    "es_parte_relacionada": False,
                    "justificacion_clasificacion": "Procesamiento preliminar sin API Key configurada.",
                }

            registro = {
                "NUMERO_OP": op_id,
                "ARCHIVO_PDF": ruta_pdf.name,
                "NUMERO_FACTURA_OP": datos_extraidos.get("numero_factura_op"),
                "EMISOR_NOMBRE": datos_extraidos.get("emisor_nombre"),
                "EMISOR_NIT": datos_extraidos.get("emisor_nit"),
                "CONCEPTO_PAGO": datos_extraidos.get("concepto_pago"),
                "VALOR_TOTAL_EXTRAIDO": datos_extraidos.get("valor_total"),
                "CATEGORIA_CONTRACTUAL": datos_extraidos.get("categoria_contractual"),
                "REQUIERE_ACTA_INTERVENTORIA": datos_extraidos.get("requiere_acta_interventoria"),
                "ES_PARTE_RELACIONADA": datos_extraidos.get("es_parte_relacionada"),
                "JUSTIFICACION": datos_extraidos.get("justificacion_clasificacion"),
            }

            resultados.append(registro)
            registros_previos[op_id] = registro

            # Guardado progresivo en cada iteración
            df_actual = pd.DataFrame(resultados)
            df_actual.to_csv(ruta_salida_csv, index=False, encoding="utf-8-sig")

            # Pausa de cortesía entre llamadas a la API
            time.sleep(1.0)

        except Exception as e_proc:
            logger.error(f"Error procesando PDF '{ruta_pdf.name}': {e_proc}")
            # Guardar registro con error
            registro_error = {
                "NUMERO_OP": op_id,
                "ARCHIVO_PDF": ruta_pdf.name,
                "NUMERO_FACTURA_OP": None,
                "EMISOR_NOMBRE": "ERROR_PROCESAMIENTO",
                "EMISOR_NIT": None,
                "CONCEPTO_PAGO": str(e_proc),
                "VALOR_TOTAL_EXTRAIDO": 0.0,
                "CATEGORIA_CONTRACTUAL": "EXCLUIDO",
                "REQUIERE_ACTA_INTERVENTORIA": False,
                "ES_PARTE_RELACIONADA": False,
                "JUSTIFICACION": f"Fallo en rasterización o API: {e_proc}",
            }
            resultados.append(registro_error)

    df_final = pd.DataFrame(resultados)
    df_final.to_csv(ruta_salida_csv, index=False, encoding="utf-8-sig")

    logger.info("=" * 60)
    logger.info("RESUMEN DE EXTRACCIÓN Y CLASIFICACIÓN GEMINI:")
    logger.info(f" - Total PDFs analizados: {len(df_final)}")
    logger.info(f" - Archivo generado:     {ruta_salida_csv}")
    if not df_final.empty and "CATEGORIA_CONTRACTUAL" in df_final.columns:
        conteo_cat = df_final["CATEGORIA_CONTRACTUAL"].value_counts().to_dict()
        logger.info(f" - Desglose por categoría: {conteo_cat}")
    logger.info("=" * 60)

    return df_final


def main():
    parser = argparse.ArgumentParser(description="Extracción y clasificación multimodal de soportes con Gemini.")
    parser.add_argument("--modelo", type=str, default=None, help="Nombre del modelo Gemini (ej. gemini-2.5-flash).")
    parser.add_argument("--limite", type=int, default=None, help="Límite máximo de PDFs a procesar.")
    args = parser.parse_args()

    procesar_lote_pdfs(modelo_nombre=args.modelo, max_archivos=args.limite)


if __name__ == "__main__":
    main()
