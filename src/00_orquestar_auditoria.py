"""
00_orquestar_auditoria.py
=========================
Punto de entrada único de la auditoría documental del numeral 22.3 c) (Contrato ANI
003 de 2015, Aeropuerto Ernesto Cortissoz). Encadena por sub-lotes:

  Fase 1 — Selección de filas (ranking Pareto o lista explícita) y creación/extensión del
           Excel de lote (hojas Resumen / Hallazgos detallados / Control de calidad).
  Fase 2 — Descarga de soportes PDF (02_descargar_soportes_excel.py).
  Fase 3 — Extracción documental con visión + cruce determinístico (03_extraccion_documental.py).
  Fase 4 — Clasificación en 5 etiquetas y observación (04_clasificacion_auditoria.py).

El trabajo pendiente se procesa en sub-lotes de --tamano-lote filas. Cada fila clasificada
se escribe y GUARDA en el Excel de inmediato; al cerrar cada sub-lote se actualizan
"Hallazgos detallados" y "Control de calidad". Reanudar es repetir el mismo comando (o
correrlo sin argumentos: toma el lote más reciente de data/output): las filas con
"Resultado de validación" se omiten y los PDF/extracciones ya hechos se reutilizan.

Ejemplos:
    # Ranks 4001 a 4500 del ranking Pareto, en sub-lotes de 20, con Gemini
    python src/00_orquestar_auditoria.py --modo pareto --desde 4001 --hasta 4500 --tamano-lote 20

    # Reprocesar ranks pendientes de rondas anteriores con Claude
    python src/00_orquestar_auditoria.py --modo lista --ops "1677,3015" --modelo-vision claude

    # Seleccionar por número de OP en vez de rank
    python src/00_orquestar_auditoria.py --modo lista --ops "1956,8598" --tipo-id op

    # Continuar un lote existente (o sin argumentos: el más reciente)
    python src/00_orquestar_auditoria.py --excel data/output/Validacion_Soportes_Rango4001-4500_UNIVERSALIDAD_ABAS1.xlsx
    python src/00_orquestar_auditoria.py

    # Solo crear el Excel del lote y descargar, sin llamar a ningún modelo
    python src/00_orquestar_auditoria.py --modo pareto --desde 4001 --hasta 4100 --fases 1,2
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import logging
import math
import os
import re
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parent))
from auditoria.config import (  # noqa: E402
    BASE_DIR, DIR_OUTPUT, DIR_SOPORTES, RUTAS_CANDIDATAS_UNIVERSALIDAD, a_float, cargar_env,
    configurar_logging, importar_script, normalizar_op,
)
from auditoria.estado import EstadoLote, ahora  # noqa: E402
from auditoria.lote_excel import LoteExcel, buscar_lote_mas_reciente, slug_lote  # noqa: E402
from auditoria.modelos_llm import ClienteLLM, ErrorCuotaLLM, ErrorLLM, crear_cliente  # noqa: E402

logger = logging.getLogger("orquestador")

SUFIJO_LOTE = "_UNIVERSALIDAD_ABAS1.xlsx"


# ---------------------------------------------------------------------------
# Argumentos
# ---------------------------------------------------------------------------

def parsear_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Orquestador de auditoría documental por lotes (numeral 22.3 c).",
        formatter_class=argparse.RawDescriptionHelpFormatter, epilog=__doc__.split("Ejemplos:")[1])
    g = p.add_argument_group("Fase 1 — selección")
    g.add_argument("--modo", choices=["pareto", "lista"], help="Cómo seleccionar filas nuevas para el lote.")
    g.add_argument("--desde", type=int, help="Primer rank (modo pareto).")
    g.add_argument("--hasta", type=int, help="Último rank (modo pareto).")
    g.add_argument("--ops", help="Lista separada por comas de ranks u OP (modo lista). Prefijos: 'r1677' rank, 'op1956' OP.")
    g.add_argument("--archivo-ops", help="CSV/TXT con columna 'rank'/'#' u 'op'/'OP' (o una columna sin encabezado).")
    g.add_argument("--tipo-id", choices=["rank", "op"], default="rank",
                   help="Cómo interpretar números sin prefijo en --ops/--archivo-ops (default: rank).")
    g.add_argument("--nombre-lote", help="Nombre del lote en modo lista (default: derivado de la lista).")
    g.add_argument("--excel", help="Excel de lote a crear/continuar (default: derivado del modo, o el más reciente).")
    g.add_argument("--universalidad", help="Ruta del .xlsx con la hoja UNIVERSALIDAD (default: data/raw/(A)BAS~1.xlsx).")

    g = p.add_argument_group("Ejecución")
    g.add_argument("--fases", default="1,2,3,4", help="Fases a ejecutar, p. ej. '1,2' o '3,4' (default: 1,2,3,4).")
    g.add_argument("--tamano-lote", type=int, default=20, help="Filas por sub-lote (default 20).")
    g.add_argument("--limite", type=int, default=None, help="Máximo de filas pendientes a procesar en esta corrida.")

    g = p.add_argument_group("Fase 2 — descarga")
    g.add_argument("--dir-soportes", help="Carpeta de PDFs del lote (default: data/soportes/<nombre del Excel>).")
    g.add_argument("--pausa", type=float, default=1.0, help="Pausa entre descargas en segundos (default 1.0).")
    g.add_argument("--reintentos", type=int, default=3, help="Intentos por descarga con backoff (default 3).")
    g.add_argument("--usuario")
    g.add_argument("--password")
    g.add_argument("--fedauth")
    g.add_argument("--rtfa")
    g.add_argument("--respaldo", action="append", help="Carpeta local sincronizada (OneDrive) donde buscar PDFs.")

    g = p.add_argument_group("Fases 3 y 4 — modelos")
    g.add_argument("--modelo-vision", choices=["gemini", "claude"],
                   default=os.getenv("MODELO_VISION", "gemini") or "gemini", help="Proveedor del modelo (default gemini).")
    g.add_argument("--modelo", help="ID del modelo (default: GEMINI_MODEL / CLAUDE_MODEL del .env).")
    g.add_argument("--max-paginas", type=int, default=20, help="Máximo de páginas por expediente enviadas al modelo.")
    g.add_argument("--dpi", type=int, default=150)
    g.add_argument("--reextraer", action="store_true", help="Ignora extracciones guardadas y vuelve a llamar al modelo de visión.")
    a = p.parse_args(argv)
    try:
        a.fases = {int(x) for x in str(a.fases).replace(" ", "").split(",") if x}
    except ValueError:
        p.error("--fases debe ser una lista como 1,2,3,4")
    if not a.fases <= {1, 2, 3, 4}:
        p.error("--fases solo admite 1, 2, 3 y 4")
    if a.modo == "pareto" and (a.desde is None or a.hasta is None or a.desde > a.hasta or a.desde < 1):
        p.error("--modo pareto requiere --desde N --hasta M con 1 <= N <= M")
    if a.modo == "lista" and not (a.ops or a.archivo_ops):
        p.error("--modo lista requiere --ops o --archivo-ops")
    if a.tamano_lote < 1:
        p.error("--tamano-lote debe ser >= 1")
    return a


# ---------------------------------------------------------------------------
# Fase 1 — selección y Excel del lote
# ---------------------------------------------------------------------------

def parsear_ids(tokens: Sequence[str], tipo_defecto: str) -> List[Tuple[str, str]]:
    ids = []
    for t in tokens:
        t = str(t).strip().lstrip("#")
        if not t:
            continue
        m = re.match(r"^(r|rank|op)[:\s_-]?(.+)$", t, re.IGNORECASE)
        if m:
            ids.append(("rank" if m.group(1).lower().startswith("r") else "op", normalizar_op(m.group(2))))
        else:
            ids.append((tipo_defecto, normalizar_op(t)))
    return [(k, v) for k, v in ids if v]


def leer_archivo_ids(ruta: Path, tipo_defecto: str) -> List[Tuple[str, str]]:
    with open(ruta, newline="", encoding="utf-8-sig") as fh:
        filas = [r for r in csv.reader(fh) if r and any(c.strip() for c in r)]
    if not filas:
        return []
    enc = [c.strip().lower() for c in filas[0]]
    for nombres, tipo in ((("rank", "#"), "rank"), (("op", "numero op"), "op")):
        for n in nombres:
            if n in enc:
                i = enc.index(n)
                return parsear_ids([f"{'r' if tipo == 'rank' else 'op'}{r[i]}" for r in filas[1:] if len(r) > i and r[i].strip()],
                                   tipo_defecto)
    return parsear_ids([r[0] for r in filas], tipo_defecto)


def seleccionar(args: argparse.Namespace, ranking) -> Tuple[Any, str, List[Tuple[str, str]]]:
    """Filas del ranking a incluir en el lote (DataFrame), descripción y ids pedidos."""
    rankeadas = ranking[ranking["RANK"].notna()]
    if args.modo == "pareto":
        sel = rankeadas[(rankeadas["RANK"] >= args.desde) & (rankeadas["RANK"] <= args.hasta)]
        return sel, f"Rango {args.desde}–{args.hasta}", []
    ids: List[Tuple[str, str]] = []
    if args.ops:
        ids += parsear_ids(args.ops.split(","), args.tipo_id)
    if args.archivo_ops:
        ids += leer_archivo_ids(Path(args.archivo_ops), args.tipo_id)
    indices, no_encontrados = [], []
    for tipo, valor in ids:
        if tipo == "rank":
            m = rankeadas[rankeadas["RANK"] == int(valor)] if valor.isdigit() else rankeadas.iloc[0:0]
        else:
            m = rankeadas[rankeadas["OP_NORM"] == valor]
        if m.empty:
            no_encontrados.append(f"{tipo} {valor}")
        indices += [i for i in m.index if i not in indices]
    if no_encontrados:
        logger.warning(f"No encontrados en el ranking (OP+URL): {', '.join(no_encontrados)}")
    sel = ranking.loc[indices].sort_values("RANK")
    return sel, f"Lista explícita ({len(sel)} filas)", ids


def ruta_lote_por_defecto(args: argparse.Namespace, ids: List[Tuple[str, str]]) -> Path:
    if args.modo == "pareto":
        return DIR_OUTPUT / f"Validacion_Soportes_Rango{args.desde}-{args.hasta}{SUFIJO_LOTE}"
    if args.nombre_lote:
        nombre = re.sub(r"[^\w.-]+", "_", args.nombre_lote)
        return DIR_OUTPUT / f"Validacion_Soportes_{nombre}{SUFIJO_LOTE}"
    firma = hashlib.sha1(",".join(sorted(f"{k}:{v}" for k, v in ids)).encode()).hexdigest()[:8]
    return DIR_OUTPUT / f"Validacion_Soportes_Lista{len(ids)}_{firma}{SUFIJO_LOTE}"


def _limpio(v: Any) -> Any:
    try:
        import pandas as pd
        if v is None or pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    return v


def registro_resumen(r: Any) -> Dict[str, Any]:
    """Columnas de datos de la hoja Resumen para una fila del ranking."""
    op = _limpio(r.get("OP_NORM"))
    valor = a_float(_limpio(r.get("VALOR DEBITADO O ACREDITADO")))
    fecha = _limpio(r.get("FECHA DE PAGO O DESEMBOLSO"))
    return {
        "rank": int(r["RANK"]),
        "fila_universalidad": int(r["FILA_UNIVERSALIDAD"]),
        "op": int(op) if op and str(op).isdigit() else op,
        "tercero": _limpio(r.get("TERCERO")),
        "cuenta": _limpio(r.get("NOMBRE DE CUENTA CONTABLE")),
        "fecha": fecha.strftime("%Y-%m-%d") if hasattr(fecha, "strftime") else fecha,
        "valor": int(valor) if valor is not None and float(valor).is_integer() else valor,
        "moneda": r.get("MONEDA_CARPETA"),
        "url": _limpio(r.get("URL")),
    }


def _millones(x: float) -> str:
    return "$" + f"{x / 1e6:,.0f}".replace(",", ".") + " millones"


def notas_lote(descripcion: str, ruta_univ: Path, df_total, ranking, seleccion) -> Tuple[str, List[str]]:
    col_v = "VALOR DEBITADO O ACREDITADO"
    total_univ = ranking["VALOR_ABSOLUTO"].sum()
    rankeadas = ranking[ranking["RANK"].notna()]
    total_opurl = rankeadas["VALOR_ABSOLUTO"].sum()
    rank_max = int(seleccion["RANK"].max()) if len(seleccion) else 0
    cubierto = rankeadas[rankeadas["RANK"] <= rank_max]["VALOR_ABSOLUTO"].sum()
    def pct(x: float) -> str:
        return f"{x:.1f}".replace(".", ",") + "%"

    titulo = f"Validación de soportes — {descripcion} — UNIVERSALIDAD"
    nota1 = (f"Archivo fuente: {ruta_univ.name} — pestaña UNIVERSALIDAD ({len(df_total):,} filas".replace(",", ".")
             + f"), columna URL. Metodología: ranking por valor absoluto de \"{col_v}\" sobre todas las filas; "
             f"se numera secuencialmente (#) solo entre las filas que tienen NUMERO OP y URL a la vez. "
             f"Lote generado por 00_orquestar_auditoria.py el {ahora()[:10]} con {len(seleccion)} transacciones.")
    nota2 = (f"Cobertura acumulada hasta el rank #{rank_max}: {_millones(cubierto)} de {_millones(total_univ)} del "
             f"universo total UNIVERSALIDAD ({pct(100 * cubierto / total_univ)}); {_millones(cubierto)} de "
             f"{_millones(total_opurl)} del universo con OP+URL ({pct(100 * cubierto / total_opurl)}). "
             f"Clasificación: COHERENTE / COHERENTE — VER NOTA / INCONCLUSO / NO CORRESPONDE / TERCERO NO COINCIDE.")
    return titulo, [nota1, nota2]


def fase1(args: argparse.Namespace, ranking, df_total, ruta_univ: Path) -> LoteExcel:
    seleccion, descripcion, ids = (None, "", [])
    if args.modo:
        seleccion, descripcion, ids = seleccionar(args, ranking)
    ruta = Path(args.excel) if args.excel else (ruta_lote_por_defecto(args, ids) if args.modo else None)
    if ruta is None:
        ruta = buscar_lote_mas_reciente(DIR_OUTPUT)
        if ruta is None:
            raise SystemExit("No hay lotes en data/output. Indique --modo pareto|lista (o --excel).")
        logger.info(f"Sin argumentos de selección: se reanuda el lote más reciente '{ruta.name}'.")
    if not ruta.is_absolute():
        ruta = (Path.cwd() / ruta) if (Path.cwd() / ruta).exists() or not (BASE_DIR / ruta).exists() else BASE_DIR / ruta

    cambio = not ruta.exists()
    if ruta.exists():
        lote = LoteExcel.abrir(ruta)
        logger.info(f"[FASE 1] Lote existente '{ruta.name}': {len(lote.filas())} filas; se toma como punto de partida.")
    else:
        if seleccion is None:
            raise SystemExit(f"No existe '{ruta}' y no se indicó --modo para crearlo.")
        titulo, notas = notas_lote(descripcion, ruta_univ, df_total, ranking, seleccion)
        lote = LoteExcel.crear(ruta, titulo, notas)
        logger.info(f"[FASE 1] Creando lote nuevo '{ruta.name}'.")

    if seleccion is not None:
        presentes = lote.ranks_presentes()
        nuevas = [registro_resumen(r) for _, r in seleccion.iterrows() if int(r["RANK"]) not in presentes]
        if nuevas:
            cambio = True
            lote.agregar_filas(nuevas)
            logger.info(f"[FASE 1] {len(nuevas)} fila(s) nuevas agregadas ({len(seleccion) - len(nuevas)} ya estaban).")
        elif len(seleccion):
            logger.info("[FASE 1] Todas las filas seleccionadas ya estaban en el lote.")
    if cambio:  # no reescribir un Excel existente si no hubo cambios
        lote.guardar()
    return lote


# ---------------------------------------------------------------------------
# Fases 2-4 por sub-lote
# ---------------------------------------------------------------------------

class Orquestador:
    def __init__(self, args: argparse.Namespace, lote: LoteExcel, ranking, modulos: Dict[str, Any]):
        self.args = args
        self.lote = lote
        self.ranking = ranking
        self.m02, self.m03, self.m04 = modulos["02"], modulos["03"], modulos["04"]
        self.estado = EstadoLote(slug_lote(lote.ruta))
        self.dir_soportes = Path(args.dir_soportes) if args.dir_soportes else DIR_SOPORTES / lote.ruta.stem
        self._cliente: Optional[ClienteLLM] = None
        self._descargador = None
        self.pendientes: Dict[int, Dict[str, Any]] = {}
        self.discordancias: Dict[int, Dict[str, Any]] = {}
        self.conteo: Dict[str, int] = {}
        rk = ranking["RANK"].notna()
        self._pos_rank = dict(zip(ranking.loc[rk, "RANK"].astype(int), ranking.index[rk]))
        self._pos_fila = dict(zip(ranking["FILA_UNIVERSALIDAD"].astype(int), ranking.index))

    # ----- utilidades -------------------------------------------------------

    def cliente(self) -> ClienteLLM:
        if self._cliente is None:
            self._cliente = crear_cliente(self.args.modelo_vision, self.args.modelo)
            logger.info(f"Modelo: {self._cliente.descripcion()}")
        return self._cliente

    def _contar(self, clave: str) -> None:
        self.conteo[clave] = self.conteo.get(clave, 0) + 1

    def _pendiente(self, f, motivo: str) -> None:
        self.pendientes[f.rank] = {"rank": f.rank, "op": f.op, "tercero": f.tercero, "valor": f.valor,
                                   "motivo": motivo, "url": f.url}
        self._contar("pendientes")
        logger.warning(f"   rank #{f.rank} OP {f.op}: queda EN BLANCO — {motivo}")

    def ubicar_en_ranking(self, f) -> Optional[Any]:
        """Fila del ranking para la fila del Excel (por rank; si no cuadra, por Fila UNIVERSALIDAD/OP)."""
        op = normalizar_op(f.op)
        pos = self._pos_rank.get(f.rank)
        if pos is not None and self.ranking.at[pos, "OP_NORM"] == op:
            return self.ranking.loc[pos]
        alterna = None
        pos_f = self._pos_fila.get(f.fila_universalidad)
        if pos_f is not None and self.ranking.at[pos_f, "OP_NORM"] == op:
            alterna = self.ranking.loc[pos_f]
        else:
            cands = self.ranking[self.ranking["OP_NORM"] == op]
            if f.valor is not None and len(cands) > 1:
                cands = cands[(cands["VALOR_ABSOLUTO"] - abs(a_float(f.valor) or 0)).abs() <= 1.5]
            if len(cands) == 1:
                alterna = cands.iloc[0]
        self.discordancias[f.rank] = {
            "rank": f.rank, "op": op,
            "detalle": (f"El rank #{f.rank} del Excel corresponde en el ranking recalculado a OP "
                        f"{self.ranking.at[pos, 'OP_NORM'] if pos is not None else '(ninguna)'}; "
                        + (f"se cruzó con la fila UNIVERSALIDAD {int(alterna['FILA_UNIVERSALIDAD'])}."
                           if alterna is not None else "no se pudo ubicar la fila de forma unívoca."))}
        return alterna

    # ----- fases --------------------------------------------------------------

    def fase2(self, filas: List[Any]) -> None:
        a_descargar = [f for f in filas if f.rank is not None and not self.m02.pdfs_de_rank(self.dir_soportes, f.rank, f.op)]
        if not a_descargar:
            return
        if self._descargador is None:
            self._descargador = self.m02.Descargador(
                self.dir_soportes, usuario=self.args.usuario, contrasena=self.args.password,
                fedauth=self.args.fedauth, rtfa=self.args.rtfa,
                respaldo_extra=[Path(p) for p in self.args.respaldo or []], reintentos=self.args.reintentos)
        logger.info(f"[FASE 2] Descargando {len(a_descargar)} soporte(s) a '{self.dir_soportes}'...")
        dicts = [{"rank": f.rank, "fila_universalidad": f.fila_universalidad, "op": f.op, "tercero": f.tercero,
                  "fecha": f.fecha, "valor": f.valor, "moneda": f.moneda, "url": f.url} for f in a_descargar]
        resultados = self.m02.descargar_filas(dicts, self.dir_soportes, pausa_segundos=self.args.pausa,
                                              descargador=self._descargador)
        for rank, r in resultados.items():
            self._contar(f"descarga_{r['estado']}")
            self.estado.actualizar(rank, descarga={"estado": r["estado"], "archivos": [p.name for p in r["archivos"]],
                                                   "detalle": r["detalle"], "fecha": ahora()})

    def procesar_fila(self, f) -> None:
        a = self.args
        pdfs = self.m02.pdfs_de_rank(self.dir_soportes, f.rank, f.op) if self.dir_soportes.exists() else []
        if not pdfs:
            det = (self.estado.leer(f.rank).get("descarga") or {}).get("detalle")
            self._pendiente(f, "sin PDF disponible tras la descarga" + (f" ({det})" if det else ""))
            return
        fila_rk = self.ubicar_en_ranking(f)
        if fila_rk is None:
            self._pendiente(f, "no se pudo ubicar la fila en UNIVERSALIDAD (ver Control de calidad)")
            return
        ctx = self.m03.contexto_universalidad(fila_rk, self.ranking)
        ident = {"op": normalizar_op(f.op), "tercero": f.tercero, "valor": a_float(f.valor)}

        if ctx["es_reteg"]:  # regla dura, no requiere modelo
            if 4 in a.fases:
                self._escribir(f, self.m04.clasificar_reteg(ctx), None, ident)
            return

        st = self.estado.leer(f.rank)
        extraccion = st.get("extraccion")
        if a.reextraer and extraccion and extraccion.get("fecha", "") < self._inicio:
            extraccion = None
        if extraccion is None and 3 in a.fases:
            try:
                logger.info(f"   [FASE 3] rank #{f.rank}: extrayendo hechos de {len(pdfs)} PDF...")
                extraccion = self.m03.extraer_hechos(pdfs, self.cliente(), op=f.op, dpi=a.dpi, max_paginas=a.max_paginas)
            except ErrorCuotaLLM:
                raise
            except (ErrorLLM, RuntimeError, OSError) as e:
                self.estado.actualizar(f.rank, error={"fase": 3, "mensaje": str(e), "fecha": ahora()})
                self._pendiente(f, f"falló la extracción: {e}")
                return
            senales = self.m03.cruzar_con_universalidad(extraccion["hechos"], ctx, extraccion["paginas_totales"],
                                                        extraccion["paginas_enviadas"])
            self.estado.actualizar(f.rank, extraccion=extraccion, senales=senales, contexto=ctx, error=None, **ident)
            self._contar("extraidas")
        if 4 not in a.fases:
            return
        if extraccion is None:
            self._pendiente(f, "sin hechos extraídos (ejecute la fase 3)")
            return
        # Las señales son determinísticas: se recalculan siempre con el código vigente.
        senales = self.m03.cruzar_con_universalidad(extraccion["hechos"], ctx, extraccion.get("paginas_totales"),
                                                    extraccion.get("paginas_enviadas"))
        try:
            logger.info(f"   [FASE 4] rank #{f.rank}: clasificando...")
            clasif = self.m04.clasificar_fila(ctx, extraccion["hechos"], senales, self.cliente())
        except ErrorCuotaLLM:
            raise
        except ErrorLLM as e:
            self.estado.actualizar(f.rank, error={"fase": 4, "mensaje": str(e), "fecha": ahora()})
            self._pendiente(f, f"falló la clasificación: {e}")
            return
        self._escribir(f, clasif, senales, ident)

    def _escribir(self, f, clasif: Dict[str, Any], senales: Optional[Dict[str, Any]], ident: Dict[str, Any]) -> None:
        self.lote.escribir_resultado(f.fila_excel, clasif["etiqueta"], clasif["observacion"])
        self.lote.guardar()  # persistir YA: una interrupción no pierde esta fila
        self.estado.actualizar(f.rank, clasificacion=dict(clasif, escrito_en_excel=ahora()),
                               error=None, **({"senales": senales} if senales else {}), **ident)
        self.pendientes.pop(f.rank, None)
        self._contar(clasif["etiqueta"])
        logger.info(f"   rank #{f.rank} OP {f.op}: {clasif['etiqueta']} — {clasif['observacion'][:140]}")

    # ----- hojas de resumen ---------------------------------------------------

    def actualizar_hojas(self) -> None:
        filas = self.lote.filas()
        por_rank = {f.rank: f for f in filas}
        registros = []
        for rank, d in self.estado.todos().items():
            f = por_rank.get(rank)
            if f is None or not d.get("clasificacion") or not f.validada:
                continue
            clasif = dict(d["clasificacion"])
            clasif["etiqueta"] = str(f.resultado).strip()  # el Excel manda (puede haber ediciones humanas)
            clasif["observacion"] = str(f.observacion or "")
            registros.append({"rank": rank, "op": normalizar_op(f.op), "tercero": f.tercero,
                              "valor": a_float(f.valor), "clasificacion": clasif, "senales": d.get("senales")})
        pendientes = [p for r, p in sorted(self.pendientes.items()) if r in por_rank and not por_rank[r].validada]
        self.lote.actualizar_hallazgos(self.m04.construir_hallazgos(registros, pendientes))
        self.lote.reescribir_control_automatico(
            self.m04.ENCABEZADOS_CONTROL,
            self.m04.construir_control_calidad(filas, registros, list(self.discordancias.values())),
            subtitulo=f"Actualizado {ahora()} — {len(registros)} filas clasificadas automáticamente en este lote.")
        self.lote.guardar()

    # ----- ciclo principal ---------------------------------------------------------

    def ejecutar(self) -> int:
        self._inicio = ahora()
        a = self.args
        pendientes = [f for f in self.lote.filas() if not f.validada and f.rank is not None]
        total_filas = len(self.lote.filas())
        logger.info(f"Lote '{self.lote.ruta.name}': {total_filas} filas, {total_filas - len(pendientes)} ya con "
                    f"resultado, {len(pendientes)} pendientes. Fases: {sorted(a.fases)}.")
        if a.limite:
            pendientes = pendientes[: a.limite]
        if not pendientes or not (a.fases & {2, 3, 4}):
            return 0
        n_sub = math.ceil(len(pendientes) / a.tamano_lote)
        codigo = 0
        t0 = time.time()
        try:
            for k in range(n_sub):
                sub = pendientes[k * a.tamano_lote:(k + 1) * a.tamano_lote]
                logger.info("=" * 78)
                logger.info(f"SUB-LOTE {k + 1}/{n_sub}: ranks #{sub[0].rank}–#{sub[-1].rank} ({len(sub)} filas)")
                logger.info("=" * 78)
                if 2 in a.fases:
                    self.fase2(sub)
                if a.fases & {3, 4}:
                    for i, f in enumerate(sub, 1):
                        logger.info(f"[{k * a.tamano_lote + i}/{len(pendientes)}] rank #{f.rank} OP {f.op} — {f.tercero}")
                        self.procesar_fila(f)
                self.actualizar_hojas()
                logger.info(f"Sub-lote {k + 1}/{n_sub} cerrado y guardado ({time.time() - t0:.0f}s). Conteo: {self.conteo}")
        except ErrorCuotaLLM as e:
            logger.error(f"[DETENIDO] {e}")
            logger.error("Todo lo procesado está guardado. Reanude con el mismo comando cuando haya cuota.")
            codigo = 2
        except KeyboardInterrupt:
            logger.warning("[INTERRUMPIDO] Se guardó todo lo procesado hasta la última fila. Reanude con el mismo comando.")
            codigo = 130
        finally:
            self._cerrar()
        return codigo

    def _cerrar(self) -> None:
        try:
            self.actualizar_hojas()
        except Exception as e:  # noqa: BLE001 - el cierre nunca debe ocultar el error original
            logger.error(f"No se pudieron actualizar las hojas de hallazgos: {e}")
        blancas = []
        for f in self.lote.filas():
            if f.validada or f.rank is None:
                continue
            p = self.pendientes.get(f.rank)
            err = self.estado.leer(f.rank).get("error")
            motivo = p["motivo"] if p else (f"error fase {err['fase']}: {err['mensaje']}" if err else "no procesada aún")
            blancas.append({"rank": f.rank, "op": f.op, "tercero": f.tercero, "valor": f.valor, "motivo": motivo, "url": f.url})
        ruta = self.estado.escribir_pendientes(blancas)
        logger.info("=" * 78)
        logger.info(f"RESUMEN: {self.conteo}")
        logger.info(f"Excel del lote: {self.lote.ruta}")
        if blancas:
            con_motivo = [b for b in blancas if b["motivo"] != "no procesada aún"]
            logger.info(f"Filas con resultado EN BLANCO: {len(blancas)} ({len(con_motivo)} con incidencia). Listado: {ruta}")
            for b in con_motivo[:30]:
                logger.info(f"   #{b['rank']} OP {b['op']}: {b['motivo']}")
        logger.info("=" * 78)


def main(argv: Optional[Sequence[str]] = None) -> int:
    configurar_logging()
    cargar_env()
    args = parsear_args(argv)
    m01 = importar_script("01_pareto_universalidad")
    modulos = {"02": importar_script("02_descargar_soportes_excel"),
               "03": importar_script("03_extraccion_documental"),
               "04": importar_script("04_clasificacion_auditoria")}

    ruta_univ = Path(args.universalidad) if args.universalidad else next(
        (p for p in RUTAS_CANDIDATAS_UNIVERSALIDAD if p.exists()), None)
    if ruta_univ is None or not ruta_univ.exists():
        raise SystemExit("No se encontró el archivo de UNIVERSALIDAD (use --universalidad).")
    df = m01.cargar_universalidad_con_cache(ruta_univ)
    ranking = m01.rankear_op_url(df)
    logger.info(f"UNIVERSALIDAD: {len(df):,} filas; {int(ranking['RANK'].notna().sum()):,} con OP+URL rankeadas.")

    lote = fase1(args, ranking, df, ruta_univ)
    estado_dir = EstadoLote(slug_lote(lote.ruta)).dir
    configurar_logging(archivo=estado_dir / "orquestador.log")
    if args.fases == {1}:
        logger.info(f"Fase 1 completada: {lote.ruta}")
        return 0
    return Orquestador(args, lote, ranking, modulos).ejecutar()


if __name__ == "__main__":
    sys.exit(main())
