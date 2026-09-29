"""
03_extraccion_documental.py
===========================
Fase 3 — Extracción documental con visión + cruce determinístico con UNIVERSALIDAD.

Principio (ya adoptado en el proyecto): el modelo de visión SOLO transcribe hechos
verificables del soporte con un esquema JSON estricto (sin clasificar ni calcular);
una capa determinística en Python cruza esos hechos contra la fila de UNIVERSALIDAD
(valor, tercero, fecha, moneda, partes relacionadas, RETEG) y hace el cuadre de
retenciones (auditoria/reconciliacion.py). Así cada señal es reproducible y
discutible ante la ANI.

1. Rasteriza el/los PDF del rank a imágenes en memoria (150 DPI; pdf2image o PyMuPDF
   como respaldo). Los soportes son escaneos sin capa de texto: no se usa OCR local,
   el modelo de visión lee la imagen directamente.
2. Llama al modelo (--modelo-vision gemini|claude) con ESQUEMA_HECHOS.
3. Calcula las señales determinísticas (`cruzar_con_universalidad`).
4. Guarda hechos + señales en el checkpoint del rank (data/output/estado/<lote>/ranks/).

Uso independiente (normalmente lo invoca 00_orquestar_auditoria.py):
    python src/03_extraccion_documental.py --excel "data/output/Validacion_...xlsx" --limite 5
"""

from __future__ import annotations

import argparse
import io
import logging
import re
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import pandas as pd
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parent))
from auditoria import reconciliacion as rec  # noqa: E402
from auditoria.config import (  # noqa: E402
    a_float, detectar_parte_relacionada, es_traslado_reteg, fmt_cop, nits_coinciden,
    normalizar_nit, normalizar_op, quitar_tildes, similitud_nombres,
)
from auditoria.estado import ahora  # noqa: E402
from auditoria.modelos_llm import ClienteLLM  # noqa: E402

logger = logging.getLogger(__name__)

UMBRAL_SIMILITUD_TERCERO = 0.80
LADO_MAXIMO_IMAGEN = 1600  # px; a 150 DPI una página carta mide ~1275x1650

# ---------------------------------------------------------------------------
# Rasterizado
# ---------------------------------------------------------------------------


def contar_paginas_pdf(ruta_pdf: Path) -> Optional[int]:
    try:
        import fitz  # pymupdf
        with fitz.open(str(ruta_pdf)) as doc:
            return len(doc)
    except Exception:  # noqa: BLE001
        pass
    try:
        from pdf2image import pdfinfo_from_path
        return int(pdfinfo_from_path(str(ruta_pdf))["Pages"])
    except Exception:  # noqa: BLE001
        return None


def rasterizar_pdf_a_imagenes(ruta_pdf: Path, dpi: int = 150, max_paginas: int = 10) -> List[Image.Image]:
    """Convierte las páginas de un PDF a objetos PIL Image en memoria.

    Intenta primero con pdf2image; si falla (p. ej. falta Poppler) usa PyMuPDF.
    PyMuPDF aplica la rotación declarada de la página (escaneos a 270°).
    """
    try:
        from pdf2image import convert_from_path
        imgs = convert_from_path(str(ruta_pdf), dpi=dpi, first_page=1, last_page=max_paginas)
        if imgs:
            return imgs
    except Exception as e_p2i:  # noqa: BLE001
        logger.debug(f"pdf2image no disponible o requiere Poppler ({e_p2i}); uso PyMuPDF.")

    try:
        import fitz  # pymupdf
        imagenes: List[Image.Image] = []
        with fitz.open(str(ruta_pdf)) as doc:
            mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
            for num_pag in range(min(len(doc), max_paginas)):
                pix = doc.load_page(num_pag).get_pixmap(matrix=mat, alpha=False)
                imagenes.append(Image.open(io.BytesIO(pix.tobytes("jpeg"))))
        return imagenes
    except Exception as e_fitz:  # noqa: BLE001
        raise RuntimeError(f"No fue posible rasterizar '{ruta_pdf.name}': {e_fitz}") from e_fitz


def imagen_a_jpeg(img: Image.Image, calidad: int = 85, lado_max: int = LADO_MAXIMO_IMAGEN) -> bytes:
    img = img.convert("RGB")
    if max(img.size) > lado_max:
        img.thumbnail((lado_max, lado_max), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=calidad, optimize=True)
    return buf.getvalue()


def rasterizar_expediente(rutas_pdf: Sequence[Path], dpi: int = 150,
                          max_paginas: int = 20) -> Tuple[List[bytes], int, List[str]]:
    """Rasteriza varios PDF del mismo rank como un solo expediente (en orden).

    Devuelve (jpeg por página, total de páginas del expediente, descripción por archivo)."""
    jpegs: List[bytes] = []
    total = 0
    detalle = []
    for ruta in rutas_pdf:
        n = contar_paginas_pdf(ruta) or 0
        total += n
        restante = max_paginas - len(jpegs)
        if restante <= 0:
            detalle.append(f"{ruta.name}: {n} pág. (no enviadas, límite {max_paginas})")
            continue
        imgs = rasterizar_pdf_a_imagenes(ruta, dpi=dpi, max_paginas=restante)
        jpegs.extend(imagen_a_jpeg(i) for i in imgs)
        detalle.append(f"{ruta.name}: {n or len(imgs)} pág., enviadas {len(imgs)}")
    return jpegs, max(total, len(jpegs)), detalle


# ---------------------------------------------------------------------------
# Esquema estricto de hechos verificables (sin clasificación)
# ---------------------------------------------------------------------------

TIPOS_DOCUMENTO = [
    "orden_operacion_fiduciaria", "factura", "cuenta_de_cobro", "nota_credito", "acta",
    "poliza_o_certificado_seguro", "comprobante_de_pago", "contrato_u_orden_de_compra",
    "correo_o_carta", "extracto_o_soporte_bancario", "otro",
]
TIPOS_VALOR = [
    "subtotal_base", "iva", "total_bruto", "retefuente", "reteiva", "reteica", "retegarantia",
    "amortizacion_anticipo", "descuento", "nota_credito", "neto_a_pagar", "valor_girado", "otro",
]
MONEDAS = ["COP", "USD", "EUR", "DESCONOCIDA"]
ROLES_PARTE = ["emisor", "beneficiario_del_pago", "contratante", "tomador_poliza", "asegurado_poliza",
               "beneficiario_poliza", "otro"]


def _obj(propiedades: Dict[str, Any], descripcion: str = "") -> Dict[str, Any]:
    d = {"type": "object", "properties": propiedades, "required": list(propiedades), "additionalProperties": False}
    if descripcion:
        d["description"] = descripcion
    return d


def _txt(descripcion: str) -> Dict[str, Any]:
    return {"type": "string", "description": descripcion}


ESQUEMA_HECHOS: Dict[str, Any] = _obj({
    "numero_op_visible": _txt("Número de orden de pago/operación impreso en el expediente; '' si no se ve."),
    "paginas_legibles": {"type": "integer", "description": "Cuántas de las páginas recibidas se pudieron leer."},
    "documento_legible": {"type": "boolean",
                          "description": "false si el escaneo impide leer con confianza las cifras principales."},
    "nota_calidad": _txt("Calidad del escaneo: páginas ilegibles, giradas, cortadas, texto especular, etc."),
    "documentos": {"type": "array", "items": _obj({
        "tipo": {"type": "string", "enum": TIPOS_DOCUMENTO},
        "pagina_inicio": {"type": "integer"},
        "pagina_fin": {"type": "integer"},
        "legible": {"type": "boolean"},
        "numero_documento": _txt("Número de factura, cuenta de cobro, orden, acta o póliza; '' si no se ve."),
        "fecha": _txt("Fecha del documento YYYY-MM-DD (o YYYY-MM); '' si no se ve."),
        "emisor_nombre": _txt("Razón social de quien emite el documento; ''."),
        "emisor_nit": _txt("NIT/CC del emisor tal como aparece; ''."),
        "beneficiario_nombre": _txt("A quién se ordena girar el pago (BENEF en la orden de operación); ''."),
        "beneficiario_nit": _txt("NIT/CC del beneficiario del giro; ''."),
        "concepto": _txt("Concepto/descripción textual del bien o servicio, copiado del documento; ''."),
        "moneda": {"type": "string", "enum": MONEDAS},
        "trm_impresa": {"anyOf": [{"type": "number"}, {"type": "null"}],
                        "description": "TRM SOLO si está impresa en el documento; null si no aparece."},
        "valores": {"type": "array", "items": _obj({
            "tipo": {"type": "string", "enum": TIPOS_VALOR},
            "monto": {"type": "number", "description": "Cifra impresa, sin separadores ni símbolo."},
            "moneda": {"type": "string", "enum": MONEDAS},
            "pagina": {"type": "integer"},
            "texto_literal": _txt("Fragmento del documento donde aparece la cifra."),
        })},
    })},
    "partes_relacionadas": {"type": "array", "items": _obj({
        "entidad": {"type": "string", "enum": ["NAB", "GAC", "OAC"]},
        "rol": {"type": "string", "enum": ROLES_PARTE},
        "pagina": {"type": "integer"},
        "texto_literal": _txt("Texto donde aparece la mención."),
    })},
    "endoso_o_cesion": _obj({
        "presente": {"type": "boolean"},
        "detalle": _txt("Texto de endoso/cesión/factoring y a favor de quién; ''."),
        "pagina": {"type": "integer", "description": "0 si no aplica."},
    }),
    "menciona_servicio_deuda": _obj({
        "presente": {"type": "boolean"},
        "detalle": _txt("Texto impreso sobre abono a capital, amortización de crédito o intereses; ''."),
        "pagina": {"type": "integer", "description": "0 si no aplica."},
    }),
    "observaciones_objetivas": _txt("Anomalías objetivas: tachaduras, firmas faltantes, documento incompleto; ''."),
})

PROMPT_EXTRACCION = """\
Eres un asistente de transcripción documental para una auditoría financiera del Patrimonio \
Autónomo del Aeropuerto Ernesto Cortissoz (Barranquilla). Las imágenes son, en orden, las \
páginas de UN expediente de orden de pago (página 1 = primera imagen). Puede contener: orden \
de operación / solicitud de operación de la fiduciaria, facturas, cuentas de cobro, notas \
crédito, actas, pólizas, comprobantes de pago, correos. Son escaneos: algunas páginas pueden \
venir giradas 90°/270° o con mala calidad; léelas igualmente si es posible.

TU ÚNICA TAREA ES TRANSCRIBIR HECHOS VERIFICABLES:
1. No clasifiques el gasto ni opines si procede; no menciones el contrato. Otro sistema evalúa.
2. No calcules ni deduzcas cifras: transcribe solo montos IMPRESOS. Si un total no aparece \
impreso, no lo incluyas. Nunca inventes ni completes cifras ilegibles.
3. Cada cifra lleva su página y el texto literal donde aparece.
4. Si un dato de texto no aparece, usa '' (cadena vacía). Si la TRM no está impresa, null.
5. Separa el expediente en documentos por rango de páginas. En la orden de operación de la \
fiduciaria, el "BENEF"/beneficiario es a quién se gira el pago: transcríbelo en \
beneficiario_nombre/beneficiario_nit. En facturas, el emisor es el proveedor.
6. Transcribe con exactitud: subtotal/base, IVA, total bruto, retefuente, reteIVA, reteICA, \
retegarantía, amortización de anticipos, descuentos, neto a pagar / valor girado, moneda y TRM.
7. partes_relacionadas: reporta SOLO menciones explícitas de NUEVO AEROPUERTO DE BARRANQUILLA \
SAS (NAB, NIT 900.913.341), GRUPO AEROPORTUARIO DEL CARIBE SAS (GAC, NIT 900.817.115) u \
OPERADORA AEROPORTUARIA DEL CARIBE SAS (OAC, NIT 900.849.079), con su rol en el documento \
(emisor, beneficiario del pago, contratante —p. ej. contratos "GAC-00X-YY"—, tomador, \
asegurado o beneficiario de póliza).
8. endoso_o_cesion: si la factura dice que fue endosada/cedida, o el pago se dirige a un tercero \
distinto del proveedor (factoring, cartera colectiva), descríbelo.
9. menciona_servicio_deuda: si el documento menciona abono a capital, amortización de crédito \
o intereses de un préstamo, transcríbelo.
10. Evalúa honestamente la legibilidad: documento_legible=false si no puedes leer con confianza \
las cifras principales.
Responde solo con el JSON del esquema.
"""


def _texto_usuario_extraccion(op: Any, paginas_enviadas: int, paginas_totales: int) -> str:
    extra = (f" Se envían las primeras {paginas_enviadas} de {paginas_totales} páginas."
             if paginas_totales > paginas_enviadas else "")
    return (f"Expediente de la orden de pago {normalizar_op(op) or '(sin número)'}: "
            f"{paginas_enviadas} página(s).{extra} Transcribe los hechos según el esquema.")


def _limpiar(v: Any) -> Any:
    if isinstance(v, str):
        v = v.strip()
        return v or None
    return v


def normalizar_hechos(h: Dict[str, Any]) -> Dict[str, Any]:
    """Tipos seguros y '' -> None; no altera cifras."""
    salida = {k: _limpiar(v) for k, v in h.items() if not isinstance(v, (list, dict))}
    docs = []
    for d in h.get("documentos") or []:
        dd = {k: _limpiar(v) for k, v in d.items() if k != "valores"}
        dd["valores"] = []
        for v in d.get("valores") or []:
            monto = a_float(v.get("monto"))
            if monto is None:
                continue
            dd["valores"].append({"tipo": v.get("tipo") or "otro", "monto": abs(monto),
                                  "moneda": v.get("moneda") or dd.get("moneda") or "DESCONOCIDA",
                                  "pagina": v.get("pagina"), "texto_literal": _limpiar(v.get("texto_literal"))})
        docs.append(dd)
    salida["documentos"] = docs
    salida["partes_relacionadas"] = list(h.get("partes_relacionadas") or [])
    for clave in ("endoso_o_cesion", "menciona_servicio_deuda"):
        sub = h.get(clave) or {}
        salida[clave] = {"presente": bool(sub.get("presente")), "detalle": _limpiar(sub.get("detalle")),
                         "pagina": sub.get("pagina") or None}
    salida["documento_legible"] = bool(h.get("documento_legible"))
    return salida


def extraer_hechos(rutas_pdf: Sequence[Path], cliente: ClienteLLM, op: Any = None,
                   dpi: int = 150, max_paginas: int = 20) -> Dict[str, Any]:
    jpegs, total, detalle = rasterizar_expediente(rutas_pdf, dpi=dpi, max_paginas=max_paginas)
    if not jpegs:
        raise RuntimeError("El PDF no tiene páginas rasterizables.")
    crudo = cliente.generar_json(PROMPT_EXTRACCION, _texto_usuario_extraccion(op, len(jpegs), total),
                                 jpegs, ESQUEMA_HECHOS)
    return {"modelo": cliente.descripcion(), "fecha": ahora(), "archivos": [p.name for p in rutas_pdf],
            "detalle_archivos": detalle, "paginas_totales": total, "paginas_enviadas": len(jpegs),
            "hechos": normalizar_hechos(crudo)}


# ---------------------------------------------------------------------------
# Contexto UNIVERSALIDAD de una fila (sin la columna de validación previa)
# ---------------------------------------------------------------------------

# Columnas que NUNCA se pasan al modelo: juicios previos (independencia del auditor).
COLUMNAS_EXCLUIDAS = {"Validación", "Validacion", "LINK AL SOPORTE VALIDADO"}


def _fecha_iso(v: Any) -> Optional[str]:
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    txt = str(v).strip()
    return txt[:10] if re.match(r"^\d{4}-\d{2}-\d{2}", txt) else (txt or None)


def _val(r: Any, clave: str) -> Any:
    """Valor de una fila de pandas con NaN/NA/NaT convertidos a None."""
    v = r.get(clave) if hasattr(r, "get") else None
    try:
        if v is None or pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    return v


def _fila_a_dict(r: Any) -> Dict[str, Any]:
    g = lambda k: _val(r, k)  # noqa: E731
    return {
        "fila_universalidad": int(g("FILA_UNIVERSALIDAD")) if g("FILA_UNIVERSALIDAD") is not None else None,
        "fila_origen_excel": int(g("FILA_ORIGEN_EXCEL")) if g("FILA_ORIGEN_EXCEL") is not None else None,
        "rank": int(g("RANK")) if g("RANK") is not None else None,
        "op": g("OP_NORM"),
        "tercero": _limpiar(str(g("TERCERO"))) if g("TERCERO") is not None else None,
        "nit_tercero": normalizar_nit(g("IDENTIFICACION TERCERO")),
        "cuenta_contable_numero": normalizar_op(g("NUMERO DE CUENTA CONTABLE")),
        "cuenta_contable": g("NOMBRE DE CUENTA CONTABLE"),
        "fecha_pago": _fecha_iso(g("FECHA DE PAGO O DESEMBOLSO")),
        "valor_registrado": a_float(g("VALOR DEBITADO O ACREDITADO")),
        "valor_base": a_float(g("VALOR BASE")),
        "valor_iva": a_float(g("VALOR IVA")),
        "cuenta_bancaria": g("NOMBRE DE LA CUENTA BANCARIA"),
        "tipo_soporte": g("TIPO DE SOPORTE"),
        "etiqueta_soporte": g("SOPORTE VALIDADO"),
        "moneda_carpeta": g("MONEDA_CARPETA"),
        "url": g("URL"),
    }


def contexto_universalidad(fila_ranking: Any, ranking: Any) -> Dict[str, Any]:
    """Datos de la fila + resto de filas de UNIVERSALIDAD con la misma OP."""
    ctx = _fila_a_dict(fila_ranking)
    ctx["es_reteg"] = es_traslado_reteg(ctx["tercero"], ctx["cuenta_contable"], ctx["etiqueta_soporte"])
    otras = []
    if ctx["op"] is not None and "OP_NORM" in ranking.columns:
        mismas = ranking[(ranking["OP_NORM"] == ctx["op"]) & (ranking["FILA_UNIVERSALIDAD"] != ctx["fila_universalidad"])]
        for _, r in mismas.iterrows():
            d = _fila_a_dict(r)
            d["es_reteg"] = es_traslado_reteg(d["tercero"], d["cuenta_contable"], d["etiqueta_soporte"])
            otras.append({k: d[k] for k in ("fila_universalidad", "rank", "tercero", "nit_tercero", "cuenta_contable",
                                            "fecha_pago", "valor_registrado", "valor_base", "valor_iva",
                                            "etiqueta_soporte", "es_reteg")})
    ctx["filas_misma_op"] = otras  # COLUMNAS_EXCLUIDAS nunca se copian al contexto
    return ctx


# ---------------------------------------------------------------------------
# Cruce determinístico
# ---------------------------------------------------------------------------

TIPOS_SOPORTE_FACTURA = {"factura", "cuenta_de_cobro"}
TIPOS_ORDEN = {"orden_operacion_fiduciaria", "comprobante_de_pago"}
RETENCIONES = ("retefuente", "reteiva", "reteica", "retegarantia", "amortizacion_anticipo", "descuento", "nota_credito")


def _es_cop(v: Dict[str, Any]) -> bool:
    return v.get("moneda") in ("COP", "DESCONOCIDA", None)


def resumir_cifras(hechos: Dict[str, Any]) -> Dict[str, Any]:
    """Cifras COP candidatas del expediente (sin inventar: solo sumas de montos impresos)."""
    docs = hechos.get("documentos") or []
    vistos = set()
    fact_base = fact_iva = fact_bruto = 0.0
    n_fact = hay_base = hay_iva = hay_bruto = 0
    for d in docs:
        if d.get("tipo") not in TIPOS_SOPORTE_FACTURA:
            continue
        clave = (d.get("tipo"), d.get("numero_documento") or f"p{d.get('pagina_inicio')}")
        if d.get("numero_documento") and clave in vistos:
            continue  # original y copia del mismo documento
        vistos.add(clave)
        vals = [v for v in d["valores"] if _es_cop(v)]
        base = next((v["monto"] for v in vals if v["tipo"] == "subtotal_base"), None)
        iva = next((v["monto"] for v in vals if v["tipo"] == "iva"), None)
        bruto = next((v["monto"] for v in vals if v["tipo"] == "total_bruto"), None)
        if base is None and iva is None and bruto is None:
            continue
        n_fact += 1
        if base is not None:
            fact_base += base
            hay_base += 1
        if iva is not None:
            fact_iva += iva
            hay_iva += 1
        if bruto is not None or base is not None:
            fact_bruto += bruto if bruto is not None else base + (iva or 0.0)
            hay_bruto += 1

    orden_bruto = orden_neto = None
    retenciones: Dict[str, float] = {}
    for d in docs:
        if d.get("tipo") not in TIPOS_ORDEN:
            continue
        for v in d["valores"]:
            if not _es_cop(v):
                continue
            if v["tipo"] == "total_bruto" and orden_bruto is None:
                orden_bruto = v["monto"]
            elif v["tipo"] in ("neto_a_pagar", "valor_girado") and orden_neto is None:
                orden_neto = v["monto"]
            elif v["tipo"] in RETENCIONES:
                retenciones[v["tipo"]] = retenciones.get(v["tipo"], 0.0) + v["monto"]
        if orden_bruto is not None or orden_neto is not None:
            break  # la primera orden con cifras es el documento primario
    if not retenciones:  # retenciones impresas en facturas
        for d in docs:
            if d.get("tipo") in TIPOS_SOPORTE_FACTURA:
                for v in d["valores"]:
                    if _es_cop(v) and v["tipo"] in RETENCIONES:
                        retenciones[v["tipo"]] = retenciones.get(v["tipo"], 0.0) + v["monto"]

    montos_usd = [v["monto"] for d in docs for v in d["valores"] if v.get("moneda") == "USD"]
    return {
        "facturas_con_cifras": n_fact,
        "facturas_base": fact_base if hay_base else None,
        "facturas_iva": fact_iva if hay_iva else None,
        "facturas_bruto": fact_bruto if hay_bruto else None,
        "orden_bruto": orden_bruto,
        "orden_neto": orden_neto,
        "retenciones_impresas": retenciones,
        "montos_usd": montos_usd,
        "trm_impresa": [d["trm_impresa"] for d in docs if a_float(d.get("trm_impresa"))],
    }


def _cuadre_contra(cifras: Dict[str, Any], registrado: Optional[float]) -> Dict[str, Any]:
    """Cuadre de las cifras del documento contra un valor registrado."""
    salida: Dict[str, Any] = {"registrado": abs(registrado) if registrado is not None else None}
    neto = cifras["orden_neto"]
    salida["neto_documento_coincide"] = rec.coincide(neto, registrado)
    candidatos = []
    if cifras["facturas_base"] is not None:
        candidatos.append(("facturas (base + IVA)", cifras["facturas_base"], cifras["facturas_iva"]))
    elif cifras["facturas_bruto"] is not None:
        candidatos.append(("facturas (total bruto, sin desglose IVA)", cifras["facturas_bruto"], None))
    if cifras["orden_bruto"] is not None:
        candidatos.append(("orden de operación (bruto)", cifras["orden_bruto"], None))
    resultados = []
    for fuente, base, iva in candidatos:
        resultados.append(dict(rec.reconciliar(base, iva, registrado).to_dict(), fuente=fuente))
    # prioridad: exacto > explicado por retenciones > primer candidato (el más directo)
    mejor = next((r for estado in (rec.EXACTO, rec.COMBINACION, rec.COMBINACION_ATIPICA)
                  for r in resultados if r["estado"] == estado), resultados[0] if resultados else None)
    salida["cuadre_documento"] = mejor
    bruto_ref = cifras["facturas_bruto"] if cifras["facturas_bruto"] is not None else cifras["orden_bruto"]
    salida["retenciones_impresas_cuadran"] = rec.cuadrar_retenciones_impresas(
        bruto_ref, cifras["retenciones_impresas"], registrado)
    return salida


def _cuadra(c: Dict[str, Any]) -> Tuple[bool, bool, bool]:
    """(cuadra_exacto, cuadra_con_retenciones_estandar, cuadra_solo_con_tarifas_atipicas)."""
    doc = c.get("cuadre_documento") or {}
    exacto = bool(c.get("neto_documento_coincide")) or doc.get("estado") == rec.EXACTO
    con_ret = doc.get("estado") == rec.COMBINACION or bool(c.get("retenciones_impresas_cuadran"))
    atipica = doc.get("estado") == rec.COMBINACION_ATIPICA and not con_ret
    return exacto, con_ret, atipica


def _dias(a: Optional[str], b: Optional[str]) -> Optional[int]:
    try:
        fa = datetime.strptime(a[:10], "%Y-%m-%d")
        fb = datetime.strptime(b[:10], "%Y-%m-%d")
        return (fa - fb).days
    except (TypeError, ValueError):
        return None


def cruzar_con_universalidad(hechos: Dict[str, Any], ctx: Dict[str, Any],
                             paginas_totales: Optional[int] = None,
                             paginas_enviadas: Optional[int] = None) -> Dict[str, Any]:
    """Señales determinísticas y reproducibles del cruce documento vs UNIVERSALIDAD."""
    cifras = resumir_cifras(hechos)
    registrado = ctx.get("valor_registrado")
    senales: Dict[str, Any] = {"reteg_interno": bool(ctx.get("es_reteg")), "cifras_documento": cifras}

    # --- Valor: fila y OP --------------------------------------------------
    fila = _cuadre_contra(cifras, registrado)
    fila["universalidad_base_iva"] = rec.reconciliar(ctx.get("valor_base"), ctx.get("valor_iva"), registrado).to_dict()
    senales["cuadre_fila"] = fila
    otras = [f for f in ctx.get("filas_misma_op", []) if not f.get("es_reteg") and f.get("valor_registrado") is not None]
    cuadre_op = None
    if otras and registrado is not None:
        total_op = abs(registrado) + sum(abs(f["valor_registrado"]) for f in otras)
        cuadre_op = _cuadre_contra(cifras, total_op)
        cuadre_op["filas_sumadas"] = 1 + len(otras)
    senales["cuadre_op"] = cuadre_op

    hay_cifras = any(cifras[k] is not None for k in ("facturas_bruto", "orden_bruto", "orden_neto"))
    exacto, con_ret, atipica = _cuadra(fila)
    op_exacto, op_ret, _ = _cuadra(cuadre_op) if cuadre_op else (False, False, False)
    if exacto:
        veredicto = "EXACTO_DOCUMENTO"
    elif con_ret:
        veredicto = "EXPLICADO_POR_RETENCIONES"
    elif atipica:
        veredicto = "EXPLICADO_POR_RETENCIONES_ATIPICAS"
    elif op_exacto or op_ret:
        veredicto = "CUADRA_A_NIVEL_OP"
    elif not hay_cifras:
        veredicto = ("SOLO_UNIVERSALIDAD" if fila["universalidad_base_iva"]["estado"] in (rec.EXACTO, rec.COMBINACION)
                     else "SIN_CIFRAS_DOCUMENTO")
    else:
        veredicto = "SIN_COMBINACION_ESTANDAR"
    doc = fila.get("cuadre_documento") or {}
    senales["veredicto_valor"] = veredicto
    senales["resumen_valor"] = {
        "EXACTO_DOCUMENTO": "El documento concilia exacto con el valor registrado.",
        "EXPLICADO_POR_RETENCIONES": doc.get("descripcion") or "Retenciones impresas explican la diferencia.",
        "EXPLICADO_POR_RETENCIONES_ATIPICAS": doc.get("descripcion"),
        "CUADRA_A_NIVEL_OP": (f"No cuadra contra la fila sola, pero sí contra la suma de las "
                              f"{(cuadre_op or {}).get('filas_sumadas')} filas de la OP en UNIVERSALIDAD "
                              f"({fmt_cop((cuadre_op or {}).get('registrado'))})."),
        "SOLO_UNIVERSALIDAD": ("El documento no deja leer cifras; solo UNIVERSALIDAD (base+IVA) concilia con el "
                               "registrado: " + fila["universalidad_base_iva"]["descripcion"]),
        "SIN_CIFRAS_DOCUMENTO": "No se pudieron leer cifras del documento ni conciliar UNIVERSALIDAD.",
        "SIN_COMBINACION_ESTANDAR": doc.get("descripcion") or "Diferencia sin explicar.",
    }[veredicto]

    # --- Tercero -------------------------------------------------------------
    candidatos = []
    for d in hechos.get("documentos") or []:
        for rol, nom, nit in (("beneficiario_del_giro", d.get("beneficiario_nombre"), d.get("beneficiario_nit")),
                              ("emisor", d.get("emisor_nombre"), d.get("emisor_nit"))):
            if nom or nit:
                sim = similitud_nombres(nom, ctx.get("tercero"))
                nit_ok = nits_coinciden(nit, ctx.get("nit_tercero"))
                candidatos.append({"tipo_documento": d.get("tipo"), "rol": rol, "nombre": nom, "nit": nit,
                                   "similitud": round(sim, 2), "nit_coincide": nit_ok,
                                   "coincide": bool(nit_ok) or (nit_ok is None and sim >= UMBRAL_SIMILITUD_TERCERO)})
    beneficiarios_orden = [c for c in candidatos if c["rol"] == "beneficiario_del_giro" and c["tipo_documento"] in TIPOS_ORDEN]
    senales["tercero"] = {
        "registrado": ctx.get("tercero"), "nit_registrado": ctx.get("nit_tercero"),
        "coincide_algun_documento": any(c["coincide"] for c in candidatos) if candidatos else None,
        "beneficiario_orden_coincide": (any(c["coincide"] for c in beneficiarios_orden)
                                        if beneficiarios_orden else None),
        "candidatos": candidatos[:12],
    }

    # --- Fecha ----------------------------------------------------------------
    fechas = sorted({d.get("fecha") for d in hechos.get("documentos") or [] if d.get("fecha")})
    difs = [x for x in (_dias(f, ctx.get("fecha_pago")) for f in fechas) if x is not None]
    senales["fecha"] = {"fecha_pago": ctx.get("fecha_pago"), "fechas_documento": fechas,
                        "dias_min_abs": min((abs(x) for x in difs), default=None),
                        "documento_posterior_al_pago": any(x > 5 for x in difs) if difs else None}

    # --- Moneda / TRM -------------------------------------------------------------
    usd = ctx.get("moneda_carpeta") == "USD" or bool(cifras["montos_usd"])
    senales["moneda"] = {"usd": usd, "moneda_carpeta": ctx.get("moneda_carpeta"), "montos_usd": cifras["montos_usd"][:10],
                         "trm_impresa": cifras["trm_impresa"], "trm_confirmable": bool(cifras["trm_impresa"]) if usd else None}

    # --- Partes relacionadas, endoso, servicio de deuda -----------------------------
    senales["parte_relacionada_universalidad"] = detectar_parte_relacionada(ctx.get("tercero"), ctx.get("nit_tercero"))
    senales["partes_relacionadas_documento"] = hechos.get("partes_relacionadas") or []
    senales["endoso_o_cesion"] = hechos.get("endoso_o_cesion")
    cuenta_txt = quitar_tildes(f"{ctx.get('cuenta_contable') or ''} {ctx.get('tipo_soporte') or ''}").upper()
    senales["servicio_deuda"] = {
        "documento": hechos.get("menciona_servicio_deuda"),
        "indicio_universalidad": bool(re.search(r"INTERES|CAPITAL|OBL(IGACION)?\.? FINANCIERA|AMORTIZACION CREDITO",
                                                cuenta_txt)),
    }

    # --- Calidad -------------------------------------------------------------------
    senales["calidad"] = {"documento_legible": hechos.get("documento_legible"),
                          "paginas_legibles": hechos.get("paginas_legibles"),
                          "paginas_enviadas": paginas_enviadas, "paginas_totales": paginas_totales,
                          "nota": hechos.get("nota_calidad")}

    # --- Anomalías de datos UNIVERSALIDAD (para "Control de calidad") ----------------
    anomalias = []
    base_u, iva_u = ctx.get("valor_base"), ctx.get("valor_iva")
    if base_u and iva_u and iva_u > 0.20 * base_u:
        anomalias.append(f"IVA atípico en UNIVERSALIDAD: IVA {fmt_cop(iva_u)} sobre base {fmt_cop(base_u)} "
                         f"({100 * iva_u / base_u:.1f}%).")
    bruto_doc = cifras["facturas_bruto"] if cifras["facturas_bruto"] is not None else cifras["orden_bruto"]
    if (bruto_doc is not None and base_u is not None and (exacto or con_ret or atipica)
            and not rec.coincide(bruto_doc, base_u + (iva_u or 0.0))):
        anomalias.append(f"Base+IVA de UNIVERSALIDAD ({fmt_cop(base_u + (iva_u or 0.0))}) no corresponde al bruto "
                         f"del documento ({fmt_cop(bruto_doc)}), aunque el valor registrado sí concilia.")
    senales["anomalias_datos"] = anomalias
    return senales


# ---------------------------------------------------------------------------
# CLI independiente
# ---------------------------------------------------------------------------

def main():
    from auditoria.config import cargar_env, configurar_logging, importar_script
    configurar_logging()
    cargar_env()
    orq = importar_script("00_orquestar_auditoria")
    p = argparse.ArgumentParser(description="Fase 3: extracción documental de un lote existente.")
    p.add_argument("--excel", required=True)
    p.add_argument("--modelo-vision", default=None, choices=["gemini", "claude"])
    p.add_argument("--limite", type=int, default=None)
    a = p.parse_args()
    args = ["--excel", a.excel, "--fases", "3"] + (["--modelo-vision", a.modelo_vision] if a.modelo_vision else [])
    args += ["--limite", str(a.limite)] if a.limite else []
    sys.exit(orq.main(args))


if __name__ == "__main__":
    main()
