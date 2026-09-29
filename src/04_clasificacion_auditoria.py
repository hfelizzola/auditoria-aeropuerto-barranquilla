"""
04_clasificacion_auditoria.py
=============================
Fase 4 — Clasificación final de cada transacción con la metodología real de
auditoría del numeral 22.3 c) (Contrato ANI 003 de 2015), en exactamente 5 etiquetas:

    COHERENTE | COHERENTE — VER NOTA | INCONCLUSO | NO CORRESPONDE | TERCERO NO COINCIDE

Entrada por fila: datos de UNIVERSALIDAD (sin ninguna validación previa: independencia
del auditor), hechos extraídos del soporte (fase 3) y señales determinísticas del cruce.
Salida: etiqueta + observación (1-3 líneas en español) + tags para "Hallazgos detallados".

Reglas duras aplicadas en Python (no dependen del modelo):
- Traslados RETEG entre cuentas propias del P.A. -> NO CORRESPONDE (sin llamar al modelo).
- La etiqueta debe ser una de las 5 exactas; si el modelo devuelve otra cosa, la fila queda
  pendiente (en blanco), nunca se inventa un resultado.
- No inventar cifras: si el documento no deja leer cifras y nada concilia, no puede quedar
  COHERENTE; un COHERENTE con diferencia sin explicar, TRM no confirmable o tercero que no
  coincide baja a COHERENTE — VER NOTA.
- INCONCLUSO siempre dice qué haría falta para cerrarlo.

Uso independiente (normalmente lo invoca 00_orquestar_auditoria.py):
    python src/04_clasificacion_auditoria.py --excel "data/output/Validacion_...xlsx"
"""

from __future__ import annotations

import argparse
import json
import logging
import re
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parent))
from auditoria.config import (  # noqa: E402
    COHERENTE, COHERENTE_VER_NOTA, ETIQUETAS, INCONCLUSO, NO_CORRESPONDE, TERCERO_NO_COINCIDE,
    fmt_cop, normalizar_etiqueta, normalizar_nombre, normalizar_op, quitar_tildes,
)
from auditoria.estado import ahora  # noqa: E402
from auditoria.modelos_llm import ClienteLLM, ErrorLLM  # noqa: E402

logger = logging.getLogger(__name__)

CATEGORIAS_AR = [
    "a_seguros_garantias", "b_aportes_subcuentas_ani", "c_comision_exito_consultor", "d_estudios_disenos",
    "e_gestion_social_ambiental", "f_gestion_predial", "g_intervenciones_obras",
    "h_operacion_administracion_impuestos", "i_comisiones_prestamistas",
    "servicio_deuda_excluido", "traslado_interno", "ninguna", "incierta",
]

ESQUEMA_CLASIFICACION: Dict[str, Any] = {
    "type": "object",
    "properties": {
        "etiqueta": {"type": "string", "enum": list(ETIQUETAS)},
        "categoria_ar": {"type": "string", "enum": CATEGORIAS_AR},
        "observacion": {"type": "string", "description": "1 a 3 líneas en español."},
        "requisito_cierre": {"type": "string", "description": "Solo si INCONCLUSO: qué haría falta para cerrarlo; si no, ''."},
        "tags": {"type": "array", "items": {"type": "string"}},
        "anomalia_datos": {"type": "string", "description": "Anomalía de los datos de UNIVERSALIDAD detectada; ''."},
    },
    "required": ["etiqueta", "categoria_ar", "observacion", "requisito_cierre", "tags", "anomalia_datos"],
    "additionalProperties": False,
}

PROMPT_CLASIFICACION = """\
Eres auditor técnico-financiero de la liquidación de la concesión del Aeropuerto Ernesto \
Cortissoz (Barranquilla), Contrato de Concesión ANI 003 de 2015, numeral 22.3 c). Recibes, \
para UNA transacción de la hoja UNIVERSALIDAD: (1) los datos contables registrados, (2) los \
hechos transcritos de su soporte PDF (ya extraídos, con página y texto literal) y (3) señales \
calculadas en Python de forma determinística (cuadre de valor y retenciones, coincidencia de \
tercero, fechas, moneda/TRM, partes relacionadas, RETEG). Las señales son exactas: no rehagas \
sus cálculos ni las contradigas; úsalas.

VERIFICA, en este orden:
1. Valor del documento (bruto, IVA, retenciones) vs. "VALOR DEBITADO O ACREDITADO". Una misma \
OP puede tener varias filas en UNIVERSALIDAD: el cuadre puede darse contra la SUMA de las filas \
de la OP (señal cuadre_op / veredicto CUADRA_A_NIVEL_OP).
2. Tercero/beneficiario/emisor del documento vs. el TERCERO registrado.
3. Concepto del documento vs. la cuenta contable registrada.
4. Fecha del documento vs. la fecha de pago registrada.
5. Reconocibilidad contractual: el concepto debe encuadrar en la lista TAXATIVA de AR_i del \
numeral 22.3 c): (a) seguros/garantías del contrato; (b) aportes a subcuentas ANI; (c) comisión \
de éxito al Consultor estructurador; (d) estudios y diseños de intervenciones; (e) gestión \
social/ambiental; (f) gestión predial; (g) intervenciones/obras verificadas por el Interventor; \
(h) operación/administración/impuestos; (i) comisiones a prestamistas DISTINTAS del servicio de \
deuda. El SERVICIO DE LA DEUDA (abono a capital o intereses de crédito) está EXCLUIDO \
expresamente: si aparece, dilo de forma explícita (categoria_ar = servicio_deuda_excluido).
6. Si el documento está en USD: registra el valor en USD y di si la TRM aplicada por la \
fiduciaria es o no confirmable en el propio soporte. Nunca inventes una TRM.

ETIQUETA (exactamente una de estas 5 cadenas):
- "COHERENTE": todo concilia sin reparos.
- "COHERENTE — VER NOTA": concilia razonablemente pero hay algo que anotar (diferencia de \
retención no explicada con combinación estándar, OCR deficiente que aun así permite confirmar \
lo esencial, patrón recurrente de un proveedor, posible costo no imputable a una parte \
relacionada, encuadre AR_i incierto, TRM no confirmable, cuadre solo a nivel de OP, etc.).
- "INCONCLUSO": el documento no permite confirmar con claridad (ilegible, cifras que no \
reconcilian de ninguna forma plausible). Llena requisito_cierre con QUÉ haría falta para cerrarlo.
- "NO CORRESPONDE": el pago no encaja en ninguna categoría AR_i taxativa (servicio de deuda, \
traslado interno entre cuentas propias del P.A. —patrón RETEG—, donación corporativa u otro \
concepto ajeno a la lista), aunque el documento esté perfectamente soportado.
- "TERCERO NO COINCIDE": el beneficiario real del soporte (según el documento / orden de \
operación) es distinto del tercero registrado en UNIVERSALIDAD para esa fila.

REGLAS DURAS:
- No inventes cifras. Cita solo cifras presentes en los hechos o en las señales. Si el \
documento no se puede leer con claridad y nada concilia, la etiqueta es INCONCLUSO.
- Si el veredicto de valor es SIN_COMBINACION_ESTANDAR, di explícitamente "diferencia sin \
explicar" con el monto; no fuerces el cuadre. Si es EXPLICADO_POR_RETENCIONES_ATIPICAS, la \
combinación usa tarifas no estándar: menciónala como posible coincidencia (VER NOTA).
- Partes relacionadas: NUEVO AEROPUERTO DE BARRANQUILLA SAS (NAB), GRUPO AEROPORTUARIO DEL \
CARIBE SAS (GAC), OPERADORA AEROPORTUARIA DEL CARIBE SAS (OAC). GAC como simple beneficiario \
nominal de una garantía de anticipo/cumplimiento de un contratista externo es NORMAL: no lo \
señales. GAC (o NAB/OAC) como beneficiario final del gasto, como parte contratante (contratos \
"GAC-00X-YY"), o como tomador + asegurado + beneficiario de una póliza a su propio nombre: \
marca "COHERENTE — VER NOTA" señalando el posible costo no imputable al P.A. (tag \
posible-costo-no-imputable-gac o posible-costo-no-imputable-nab / -oac).
- Endoso, cesión o pago a un tercero distinto del proveedor (factoring, cartera colectiva): \
anótalo (tag endoso-cesion); si el beneficiario del giro no es el tercero registrado, la \
etiqueta es TERCERO NO COINCIDE.

OBSERVACIÓN: 1 a 3 líneas en español, al estilo de estos ejemplos reales:
  "Orden de Operación General #11745 a B&S INGENIERIA SAS, MANTENIMIENTO, bruto $13.709.871 -> \
neto pagado $12.978.293. Retención concilia con retefuente 2,5% + reteIVA 15% = $731.577,15. \
Concepto: suministro de repuestos bombas HVAC -- gasto operativo reconocible."
  "OCR muy deficiente (29 de 37 páginas ilegibles); solo se distingue un ítem parcial de \
$3.833.104, que no permite reconciliar con la base UNIVERSALIDAD ni con el valor registrado. \
Se requiere documento más legible."
Menciona tipo de documento, página si aplica, la cifra exacta observada y el razonamiento de \
la clasificación. Formato de cifras colombiano ($12.345.678,90).

TAGS: etiquetas cortas en minúsculas con guiones para agrupar hallazgos del lote, p. ej. \
reteg-interno, tercero-no-coincide, servicio-deuda-excluido, no-corresponde-ar, \
posible-costo-no-imputable-gac, verificar-usd-trm, ocr-deficiente, univ-mismatch, \
diferencia-no-explicada, endoso-cesion, encuadre-ar-incierto, comisiones-aerolineas, \
donacion, publicidad, cuotas-gremiales. Si el proveedor muestra un patrón que valga la pena \
agrupar, agrega "proveedor:<RAZÓN SOCIAL>".

anomalia_datos: solo anomalías de los datos de UNIVERSALIDAD (OP duplicada, base/IVA que no \
corresponde al documento, IVA corrupto, varias filas de proveedores ajenos bajo la misma OP); \
'' si no hay.
Responde solo con el JSON del esquema.
"""


# ---------------------------------------------------------------------------
# Construcción de la entrada y ajustes determinísticos
# ---------------------------------------------------------------------------

def construir_entrada(ctx: Dict[str, Any], hechos: Dict[str, Any], senales: Dict[str, Any]) -> str:
    """JSON de entrada del modelo. Nunca incluye 'Resultado de validación', 'Observación'
    ni la columna 'Validación' de UNIVERSALIDAD (independencia del auditor)."""
    fila = {k: v for k, v in ctx.items() if k not in ("filas_misma_op", "url")}
    return json.dumps({
        "fila_universalidad": fila,
        "otras_filas_misma_op": ctx.get("filas_misma_op", [])[:25],
        "hechos_documento": hechos,
        "senales_deterministicas": senales,
    }, ensure_ascii=False, default=str, indent=1)


def normalizar_tag(tag: Any) -> Optional[str]:
    txt = str(tag or "").strip()
    if not txt:
        return None
    if txt.lower().startswith("proveedor:"):
        nombre = re.sub(r"\s+", " ", txt.split(":", 1)[1]).strip().upper()
        return f"proveedor:{nombre}" if nombre else None
    txt = quitar_tildes(txt).lower()
    txt = re.sub(r"[^a-z0-9]+", "-", txt).strip("-")
    return txt or None


def tags_deterministicos(etiqueta: str, categoria: str, senales: Dict[str, Any]) -> List[str]:
    tags = []
    if senales.get("reteg_interno"):
        tags.append("reteg-interno")
    moneda = senales.get("moneda") or {}
    if moneda.get("usd") and not moneda.get("trm_confirmable"):
        tags.append("verificar-usd-trm")
    if (senales.get("calidad") or {}).get("documento_legible") is False:
        tags.append("ocr-deficiente")
    veredicto = senales.get("veredicto_valor")
    if veredicto == "SIN_COMBINACION_ESTANDAR":
        tags.append("diferencia-no-explicada")
    elif veredicto == "EXPLICADO_POR_RETENCIONES_ATIPICAS":
        tags.append("retencion-atipica")
    elif veredicto == "CUADRA_A_NIVEL_OP":
        tags.append("cuadre-nivel-op")
    siglas = {senales.get("parte_relacionada_universalidad")} | {
        p.get("entidad") for p in senales.get("partes_relacionadas_documento") or []}
    tags += [f"parte-relacionada-{s.lower()}" for s in sorted(x for x in siglas if x)]
    if (senales.get("endoso_o_cesion") or {}).get("presente"):
        tags.append("endoso-cesion")
    if any("no corresponde al bruto" in a for a in senales.get("anomalias_datos") or []):
        tags.append("univ-mismatch")
    tags.append({COHERENTE: "coherente", COHERENTE_VER_NOTA: "ver-nota", INCONCLUSO: "inconcluso",
                 NO_CORRESPONDE: "no-corresponde", TERCERO_NO_COINCIDE: "tercero-no-coincide"}[etiqueta])
    if categoria == "servicio_deuda_excluido":
        tags.append("servicio-deuda-excluido")
    if categoria == "incierta":
        tags.append("encuadre-ar-incierto")
    return tags


_PALABRAS_CIERRE = re.compile(r"(se requiere|requiere|har[ií]a falta|falta|se necesita|para cerrar)", re.IGNORECASE)


def ajustar_clasificacion(resp: Dict[str, Any], senales: Dict[str, Any]) -> Dict[str, Any]:
    """Valida la respuesta del modelo y aplica las reglas duras. Lanza ErrorLLM si no es usable."""
    etiqueta = normalizar_etiqueta(resp.get("etiqueta"))
    if etiqueta is None:
        raise ErrorLLM(f"Etiqueta fuera de la taxonomía: {resp.get('etiqueta')!r}")
    observacion = re.sub(r"\s+", " ", str(resp.get("observacion") or "")).strip()
    if not observacion:
        raise ErrorLLM("El modelo no devolvió observación.")
    categoria = resp.get("categoria_ar") if resp.get("categoria_ar") in CATEGORIAS_AR else "incierta"
    requisito = str(resp.get("requisito_cierre") or "").strip()
    ajustes = []

    veredicto = senales.get("veredicto_valor")
    legible = (senales.get("calidad") or {}).get("documento_legible")
    moneda = senales.get("moneda") or {}
    tercero = senales.get("tercero") or {}

    if etiqueta in (COHERENTE, COHERENTE_VER_NOTA) and legible is False and veredicto == "SIN_CIFRAS_DOCUMENTO":
        etiqueta = INCONCLUSO
        ajustes.append("el documento no permite leer cifras y UNIVERSALIDAD no concilia (regla: no inventar cifras)")
    elif etiqueta == COHERENTE:
        motivos = []
        if veredicto in ("SIN_COMBINACION_ESTANDAR", "SIN_CIFRAS_DOCUMENTO", "SOLO_UNIVERSALIDAD", "CUADRA_A_NIVEL_OP",
                         "EXPLICADO_POR_RETENCIONES_ATIPICAS"):
            motivos.append({"SIN_COMBINACION_ESTANDAR": "diferencia sin explicar en el cuadre",
                            "EXPLICADO_POR_RETENCIONES_ATIPICAS": "el cuadre solo cierra con tarifas de retención atípicas",
                            "SIN_CIFRAS_DOCUMENTO": "sin cifras legibles en el documento",
                            "SOLO_UNIVERSALIDAD": "cifras confirmadas solo con UNIVERSALIDAD",
                            "CUADRA_A_NIVEL_OP": "cuadre solo a nivel de OP"}[veredicto])
        if moneda.get("usd") and not moneda.get("trm_confirmable"):
            motivos.append("TRM no confirmable en el soporte")
        if tercero.get("coincide_algun_documento") is False:
            motivos.append("el tercero registrado no aparece en el soporte")
        if motivos:
            etiqueta = COHERENTE_VER_NOTA
            ajustes.append("; ".join(motivos))

    if etiqueta == INCONCLUSO:
        if requisito and requisito.lower() not in observacion.lower():
            observacion += f" Para cerrarlo: {requisito.rstrip('.')}."
        elif not requisito and not _PALABRAS_CIERRE.search(observacion):
            observacion += (" Para cerrarlo se requiere un soporte legible con el desglose de valores "
                            "(bruto, IVA, retenciones) y el beneficiario del giro.")
    if ajustes:
        observacion += f" [Ajuste automático: {'; '.join(ajustes)}.]"

    tags = []
    for t in list(resp.get("tags") or []) + tags_deterministicos(etiqueta, categoria, senales):
        n = normalizar_tag(t)
        if n and n not in tags:
            tags.append(n)
    return {"etiqueta": etiqueta, "observacion": observacion, "categoria_ar": categoria,
            "requisito_cierre": requisito, "tags": tags,
            "anomalia_datos": str(resp.get("anomalia_datos") or "").strip(), "ajustes": ajustes}


def clasificar_reteg(ctx: Dict[str, Any]) -> Dict[str, Any]:
    """Regla dura: traslado interno entre cuentas propias del P.A. -> NO CORRESPONDE."""
    etiqueta_soporte = str(ctx.get("etiqueta_soporte") or "").strip()
    motivo = (f"etiquetado '{etiqueta_soporte}' en UNIVERSALIDAD" if "RETEG" in etiqueta_soporte.upper()
              else "cuenta con patrón de cuenta de ahorros propia del P.A.")
    obs = (f"Tercero {ctx.get('tercero')}, cuenta {ctx.get('cuenta_contable')}, {motivo} "
           f"({fmt_cop(ctx.get('valor_registrado'))}): traslado interno entre cuentas propias del P.A. "
           f"(retención en garantía), no es un pago a un tercero; no encuadra en ninguna categoría AR_i "
           f"del numeral 22.3 c).")
    return {"etiqueta": NO_CORRESPONDE, "observacion": obs, "categoria_ar": "traslado_interno",
            "requisito_cierre": "", "tags": ["reteg-interno", "no-corresponde"], "anomalia_datos": "",
            "ajustes": [], "modelo": "regla-determinística", "fecha": ahora()}


def clasificar_fila(ctx: Dict[str, Any], hechos: Dict[str, Any], senales: Dict[str, Any],
                    cliente: ClienteLLM) -> Dict[str, Any]:
    if ctx.get("es_reteg"):
        return clasificar_reteg(ctx)
    resp = cliente.generar_json(PROMPT_CLASIFICACION, construir_entrada(ctx, hechos, senales), [],
                                ESQUEMA_CLASIFICACION, max_tokens=16000)
    res = ajustar_clasificacion(resp, senales)
    res.update(modelo=cliente.descripcion(), fecha=ahora())
    return res


# ---------------------------------------------------------------------------
# Hallazgos detallados (agrupados por tag) y Control de calidad
# ---------------------------------------------------------------------------

PREFIJO_AUTO = "[Automático] "
TITULOS_TAG = {
    "servicio-deuda-excluido": "Servicio de deuda EXCLUIDO expresamente (numeral 22.3 c)",
    "reteg-interno": "RETEG — traslados internos entre cuentas propias del P.A. (no son pagos a terceros)",
    "tercero-no-coincide": "TERCERO NO COINCIDE — el beneficiario real del soporte difiere del tercero registrado en UNIVERSALIDAD",
    "no-corresponde-ar": "NO CORRESPONDE por no encuadrar en las categorías taxativas del AR_i (más allá de RETEG y servicio de deuda)",
    "posible-costo-no-imputable": "Posibles costos no imputables al P.A. (GAC/NAB/OAC como beneficiario final, parte contratante o tomador+asegurado)",
    "parte-relacionada": "Transacciones con participación de partes relacionadas (NAB / GAC / OAC)",
    "verificar-usd-trm": "Pagos en USD sin TRM confirmable en el expediente",
    "endoso-cesion": "Redirección de pago / posible cesión de cartera (factoring) a un tercero no proveedor",
    "encuadre-ar-incierto": "Categorías con encuadre AR_i incierto que ameritan criterio expreso de la Interventoría",
    "univ-mismatch": "Inconsistencias UNIVERSALIDAD vs. documento (base/IVA que no corresponden al soporte real)",
    "diferencia-no-explicada": "Reconciliación de retención sin combinación estándar (residuo sin explicar)",
    "retencion-atipica": "Cuadres que solo cierran con tarifas de retención atípicas (verificar; pueden ser coincidencia)",
    "cuadre-nivel-op": "Pagos que solo concilian a nivel de OP (varias imputaciones contables del mismo pago)",
    "ocr-deficiente": "Soportes con escaneo/OCR deficiente",
    "inconcluso": "Documentos clasificados INCONCLUSO — qué haría falta para cerrarlos",
    "pendiente-sin-pdf": "Ranks pendientes: sin soporte PDF disponible o sin procesar (resultado en blanco)",
}
# tags que no generan entrada propia (son la etiqueta misma o ya se agrupan en otra)
TAGS_SIN_ENTRADA = {"coherente", "ver-nota", "no-corresponde"}


def _clave_grupo(tag: str) -> Optional[str]:
    if tag in TAGS_SIN_ENTRADA:
        return None
    if tag.startswith("parte-relacionada-"):
        return "parte-relacionada"
    if tag.startswith("posible-costo-no-imputable"):
        return "posible-costo-no-imputable"
    return tag


def construir_hallazgos(registros: Sequence[Dict[str, Any]], pendientes: Sequence[Dict[str, Any]] = (),
                        min_casos_patron: int = 2) -> List[Tuple[str, str]]:
    """Entradas (título, detalle) agrupadas por tag, citando ranks/OP concretos."""
    grupos: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for r in registros:
        c = r["clasificacion"]
        claves = {_clave_grupo(t) for t in c.get("tags", [])}
        if c["etiqueta"] == NO_CORRESPONDE and not ({"reteg-interno", "servicio-deuda-excluido"} & set(c.get("tags", []))):
            claves.add("no-corresponde-ar")
        for k in claves - {None}:
            grupos[k].append(r)

    # Proveedor recurrente con notas (sin tag explícito): >= 3 ranks no "COHERENTE" del mismo tercero
    por_tercero: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for r in registros:
        if r["clasificacion"]["etiqueta"] != COHERENTE and r.get("tercero"):
            por_tercero[normalizar_nombre(r["tercero"])].append(r)
    for nombre, regs in por_tercero.items():
        clave = f"proveedor:{str(regs[0]['tercero']).strip().upper()}"
        if len(regs) >= 3 and clave not in grupos:
            grupos[clave] = regs

    entradas = []
    orden = list(TITULOS_TAG)
    for clave in sorted(grupos, key=lambda k: (orden.index(k) if k in orden else len(orden), k)):
        regs = sorted(grupos[clave], key=lambda r: r["rank"])
        if clave.startswith("proveedor:"):
            titulo = f"Patrón recurrente: {clave.split(':', 1)[1]}"
        elif clave in TITULOS_TAG:
            titulo = TITULOS_TAG[clave]
        else:
            titulo = f"Patrón: {clave}"
        if (clave not in TITULOS_TAG) and len(regs) < min_casos_patron:
            continue
        total = sum(abs(r.get("valor") or 0) for r in regs)
        lista = ", ".join(f"#{r['rank']} / OP {r.get('op')} ({r['clasificacion']['etiqueta']})" for r in regs)
        if clave == "inconcluso":
            ejemplos = " | ".join(
                f"#{r['rank']}: {(r['clasificacion'].get('requisito_cierre') or r['clasificacion']['observacion'])[:220]}"
                for r in regs[:8])
        else:
            ejemplos = " | ".join(f"#{r['rank']}: {r['clasificacion']['observacion'][:220]}" for r in regs[:3])
        detalle = (f"{len(regs)} caso(s), valor registrado total {fmt_cop(total)}: {lista}. "
                   f"{'Qué falta' if clave == 'inconcluso' else 'Ejemplos'}: {ejemplos}")
        entradas.append((PREFIJO_AUTO + titulo, detalle))

    if pendientes:
        lista = "; ".join(f"#{p['rank']} / OP {p.get('op')} ({p.get('motivo')})" for p in pendientes)
        entradas.append((PREFIJO_AUTO + TITULOS_TAG["pendiente-sin-pdf"],
                         f"{len(pendientes)} rank(s) quedan con 'Resultado de validación' EN BLANCO: {lista}. "
                         f"Reanude con el mismo comando o repróceselos con --modo lista."))
    return entradas


def construir_control_calidad(filas_lote: Iterable[Any], registros: Sequence[Dict[str, Any]],
                              discordancias: Sequence[Dict[str, Any]] = ()) -> List[List[Any]]:
    """Filas [Rank(s), OP, Tipo de anomalía, Detalle, Detectado] para la hoja Control de calidad."""
    hoy = ahora()[:10]
    filas_lote = list(filas_lote)
    salida: List[List[Any]] = []

    por_op: Dict[str, List[Any]] = defaultdict(list)
    por_rank: Dict[int, List[Any]] = defaultdict(list)
    for f in filas_lote:
        if normalizar_op(f.op):
            por_op[normalizar_op(f.op)].append(f)
        if f.rank is not None:
            por_rank[f.rank].append(f)
    for op, fs in sorted(por_op.items(), key=lambda x: min(f.rank or 0 for f in x[1])):
        ranks = sorted({f.rank for f in fs if f.rank is not None})
        if len(ranks) > 1:
            salida.append([", ".join(map(str, ranks)), op, "OP validada en más de un rank del lote",
                           "Líneas contables distintas del mismo OP/PDF; terceros: "
                           + " / ".join(sorted({str(f.tercero) for f in fs})), hoy])
    for rank, fs in sorted(por_rank.items()):
        ops = sorted({normalizar_op(f.op) for f in fs if normalizar_op(f.op)})
        if len(ops) > 1:
            salida.append([rank, ", ".join(ops), "Colisión de numeración (mismo # usado por OP distintas)",
                           f"Filas Excel {', '.join(str(f.fila_excel) for f in fs)}", hoy])
    for d in discordancias:
        salida.append([d.get("rank"), d.get("op"), "Rank/fila discordante con el ranking recalculado", d.get("detalle"), hoy])
    for r in sorted(registros, key=lambda x: x["rank"]):
        for a in (r.get("senales") or {}).get("anomalias_datos") or []:
            salida.append([r["rank"], r.get("op"), "Dato de UNIVERSALIDAD vs soporte (determinístico)", a, hoy])
        a = (r.get("clasificacion") or {}).get("anomalia_datos")
        if a:
            salida.append([r["rank"], r.get("op"), "Dato de UNIVERSALIDAD (señalado en la clasificación)", a, hoy])
    return salida


ENCABEZADOS_CONTROL = ["Rank(s)", "OP", "Tipo de anomalía", "Detalle", "Detectado"]


# ---------------------------------------------------------------------------
# CLI independiente
# ---------------------------------------------------------------------------

def main():
    from auditoria.config import cargar_env, configurar_logging, importar_script
    configurar_logging()
    cargar_env()
    orq = importar_script("00_orquestar_auditoria")
    p = argparse.ArgumentParser(description="Fase 4: clasificación de un lote con hechos ya extraídos.")
    p.add_argument("--excel", required=True)
    p.add_argument("--modelo-vision", default=None, choices=["gemini", "claude"])
    p.add_argument("--limite", type=int, default=None)
    a = p.parse_args()
    args = ["--excel", a.excel, "--fases", "4"] + (["--modelo-vision", a.modelo_vision] if a.modelo_vision else [])
    args += ["--limite", str(a.limite)] if a.limite else []
    sys.exit(orq.main(args))


if __name__ == "__main__":
    main()
