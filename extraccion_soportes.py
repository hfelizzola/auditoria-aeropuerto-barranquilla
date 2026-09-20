"""
extraccion_soportes.py
======================
Pipeline de extracción de soportes del P.A. Aeropuerto Ernesto Cortissoz.

Diferencias frente al script piloto:
  * El grano es el PAQUETE (orden de pago), no el archivo: agrupa `1-1.pdf`,
    `1-2.pdf`, ... en una sola llamada, porque la factura y su acta suelen ir
    en archivos distintos.
  * Rasteriza con pymupdf (sin poppler), honrando /Rotate 270, en escala de
    grises y con DPI configurable por densidad del documento.
  * Esquema estricto vía response_schema; sin prosa que el modelo pueda ignorar.
  * Caché idempotente en SQLite: un 429 o una caída no pierde el trabajo hecho.
  * Normaliza el JSON anidado a cinco tablas relacionales listas para el cruce.
  * Backend intercambiable: Gemini o Anthropic.

Requisitos:
    pip install pymupdf pydantic pandas google-generativeai
    # opcional: pip install anthropic

Uso típico:
    paquetes = descubrir_paquetes("Soportes_2015")
    procesar_lote(paquetes, db="soportes.sqlite", backend="gemini")
    tablas = exportar_tablas("soportes.sqlite")
"""

from __future__ import annotations

import base64
import hashlib
import io
import json
import os
import re
import sqlite3
import time
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import fitz  # pymupdf
import pandas as pd

from esquema_soportes import PROMPT_EXTRACCION, PaqueteSoporte

# ---------------------------------------------------------------------------
# 1. Descubrimiento y agrupación por orden de pago
# ---------------------------------------------------------------------------

# Captura el número de OP en los nombres reales del fideicomiso:
#   "7862 OP 0002 SERVIPARAMO.pdf" -> 2
#   "7862 OP 3560  NUEVO AEROPUERTO.pdf" -> 3560
#   "1-2.pdf" -> 1   (expediente partido: <op>-<parte>)
_RE_OP_EXPLICITA = re.compile(r"\bOP\s*[-_ ]?0*(\d{1,6})\b", re.IGNORECASE)
_RE_OP_PARTIDA = re.compile(r"^0*(\d{1,6})\s*-\s*(\d{1,3})\s*$")


def inferir_op(ruta: Path) -> Optional[str]:
    """Deduce el número de OP a partir del nombre del archivo."""
    tallo = ruta.stem
    m = _RE_OP_EXPLICITA.search(tallo)
    if m:
        return str(int(m.group(1)))
    m = _RE_OP_PARTIDA.match(tallo)
    if m:
        return str(int(m.group(1)))
    return None


@dataclass
class Paquete:
    """Expediente de una orden de pago: uno o más PDF."""

    op: Optional[str]
    anio: Optional[int]
    mes: Optional[int]
    archivos: List[Path] = field(default_factory=list)

    @property
    def clave(self) -> str:
        """Llave estable. El número de OP se repite entre años: se cualifica."""
        if self.op is None:
            return f"SINOP::{self.archivos[0].as_posix()}"
        return f"{self.anio or '____'}-{self.mes or '__':0>2}::OP{self.op}"


def descubrir_paquetes(raiz: str | Path, patron: str = "**/*.pdf") -> List[Paquete]:
    """Recorre Soportes_<anio>/<mes>/ y agrupa los archivos por OP."""
    raiz = Path(raiz)
    grupos: Dict[tuple, Paquete] = {}

    for ruta in sorted(raiz.glob(patron)):
        anio, mes = _anio_mes_desde_ruta(ruta)
        op = inferir_op(ruta)
        llave = (anio, mes, op) if op else (anio, mes, ruta.as_posix())
        if llave not in grupos:
            grupos[llave] = Paquete(op=op, anio=anio, mes=mes)
        grupos[llave].archivos.append(ruta)

    for p in grupos.values():
        # "1-1.pdf" antes que "1-2.pdf": el orden de las páginas importa.
        p.archivos.sort(key=lambda r: (len(r.stem), r.stem))
    return list(grupos.values())


def _anio_mes_desde_ruta(ruta: Path) -> tuple[Optional[int], Optional[int]]:
    anio = mes = None
    for parte in ruta.parts:
        m = re.search(r"(?:Soportes[_ ]?)?((?:19|20)\d{2})", parte)
        if m and anio is None:
            anio = int(m.group(1))
        m = re.fullmatch(r"0?(\d{1,2})(?:[-_](?:19|20)\d{2})?", parte)
        if m and 1 <= int(m.group(1)) <= 12 and mes is None:
            mes = int(m.group(1))
    return anio, mes


# ---------------------------------------------------------------------------
# 2. Rasterización
# ---------------------------------------------------------------------------

DPI_BASE = 170          # legible para facturas y solicitudes fiduciarias
DPI_DENSO = 230         # tablas densas: legalizaciones de fondo rotatorio
MAX_PAGINAS = 60        # tope por paquete; por encima se trocea


def rasterizar(
    archivos: Iterable[Path],
    dpi: int = DPI_BASE,
    gris: bool = True,
    calidad: int = 80,
) -> List[dict]:
    """Convierte las páginas de uno o varios PDF en JPEG en memoria.

    pymupdf honra /Rotate, por lo que las páginas con /Rotate 270 salen
    derechas sin intervención adicional.
    """
    from PIL import Image

    partes: List[dict] = []
    for ruta in archivos:
        with fitz.open(ruta) as doc:
            zoom = dpi / 72.0
            matriz = fitz.Matrix(zoom, zoom)
            espacio = fitz.csGRAY if gris else fitz.csRGB
            for pagina in doc:
                pix = pagina.get_pixmap(matrix=matriz, colorspace=espacio)
                img = Image.frombytes("L" if gris else "RGB", (pix.width, pix.height), pix.samples)
                buf = io.BytesIO()
                img.save(buf, format="JPEG", quality=calidad, optimize=True)
                partes.append({"mime_type": "image/jpeg", "data": buf.getvalue()})
    return partes


def contar_paginas(archivos: Iterable[Path]) -> int:
    total = 0
    for ruta in archivos:
        with fitz.open(ruta) as doc:
            total += doc.page_count
    return total


def parece_denso(archivos: Iterable[Path], umbral_paginas: int = 4) -> bool:
    """Heurística barata: expedientes largos suelen traer tablas de legalización."""
    return contar_paginas(archivos) >= umbral_paginas


# ---------------------------------------------------------------------------
# 3. Backends de extracción
# ---------------------------------------------------------------------------


def extraer_gemini(partes: List[dict], modelo: str = "gemini-3.6-flash") -> dict:
    import google.generativeai as genai

    model = genai.GenerativeModel(
        model_name=modelo,
        generation_config={
            "response_mime_type": "application/json",
            "response_schema": PaqueteSoporte,   # esquema estricto, no prosa
            "temperature": 0.0,
        },
    )
    resp = model.generate_content([PROMPT_EXTRACCION] + partes)
    return json.loads(resp.text)


def extraer_anthropic(partes: List[dict], modelo: str = "claude-opus-5") -> dict:
    """Alternativa con Claude en modo visión, usando tool_choice como esquema."""
    import anthropic

    cliente = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    contenido = [
        {
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": p["mime_type"],
                "data": base64.b64encode(p["data"]).decode(),
            },
        }
        for p in partes
    ]
    contenido.append({"type": "text", "text": PROMPT_EXTRACCION})

    herramienta = {
        "name": "registrar_paquete",
        "description": "Registra los hechos transcritos del expediente.",
        "input_schema": PaqueteSoporte.model_json_schema(),
    }
    resp = cliente.messages.create(
        model=modelo,
        max_tokens=16000,
        temperature=0.0,
        tools=[herramienta],
        tool_choice={"type": "tool", "name": "registrar_paquete"},
        messages=[{"role": "user", "content": contenido}],
    )
    for bloque in resp.content:
        if bloque.type == "tool_use":
            return bloque.input
    raise RuntimeError("El modelo no devolvió el bloque estructurado.")


BACKENDS = {"gemini": extraer_gemini, "anthropic": extraer_anthropic}


# ---------------------------------------------------------------------------
# 4. Caché y trazabilidad
# ---------------------------------------------------------------------------

_DDL = """
CREATE TABLE IF NOT EXISTS paquetes (
    clave           TEXT PRIMARY KEY,
    op              TEXT,
    anio            INTEGER,
    mes             INTEGER,
    archivos        TEXT,
    hash_contenido  TEXT,
    n_paginas       INTEGER,
    backend         TEXT,
    modelo          TEXT,
    dpi             INTEGER,
    estado          TEXT,          -- OK | ERROR | ILEGIBLE
    mensaje_error   TEXT,
    json_crudo      TEXT,
    segundos        REAL,
    procesado_en    TEXT
);
CREATE INDEX IF NOT EXISTS ix_paq_op ON paquetes(op);
"""


def _abrir_db(db: str) -> sqlite3.Connection:
    con = sqlite3.connect(db)
    con.executescript(_DDL)
    return con


def _hash_archivos(archivos: Iterable[Path]) -> str:
    h = hashlib.sha256()
    for ruta in archivos:
        h.update(ruta.name.encode())
        h.update(str(ruta.stat().st_size).encode())
    return h.hexdigest()[:16]


# ---------------------------------------------------------------------------
# 5. Orquestación
# ---------------------------------------------------------------------------


def procesar_paquete(
    paquete: Paquete,
    con: sqlite3.Connection,
    backend: str = "gemini",
    modelo: Optional[str] = None,
    reintentos: int = 4,
    forzar: bool = False,
) -> str:
    """Procesa un expediente y persiste el resultado. Devuelve el estado."""
    clave = paquete.clave
    huella = _hash_archivos(paquete.archivos)

    if not forzar:
        fila = con.execute(
            "SELECT estado FROM paquetes WHERE clave=? AND hash_contenido=? AND estado='OK'",
            (clave, huella),
        ).fetchone()
        if fila:
            return "CACHE"

    dpi = DPI_DENSO if parece_denso(paquete.archivos) else DPI_BASE
    n_pag = contar_paginas(paquete.archivos)
    if n_pag > MAX_PAGINAS:
        # Un expediente muy largo se marca para revisión manual antes que
        # truncarlo en silencio y perder páginas con firmas de interventoría.
        _guardar(con, paquete, huella, n_pag, backend, modelo, dpi,
                 "ERROR", f"Expediente de {n_pag} páginas: excede MAX_PAGINAS.", None, 0.0)
        return "ERROR"

    partes = rasterizar(paquete.archivos, dpi=dpi)
    fn = BACKENDS[backend]
    modelo_usado = modelo or ("gemini-3.6-flash" if backend == "gemini" else "claude-opus-5")

    t0 = time.time()
    espera = 8
    for intento in range(1, reintentos + 1):
        try:
            datos = fn(partes, modelo_usado) if modelo else fn(partes)
            datos.setdefault("total_paginas", n_pag)
            # Validación estructural antes de persistir.
            PaqueteSoporte.model_validate(datos)
            _guardar(con, paquete, huella, n_pag, backend, modelo_usado, dpi,
                     "OK", None, json.dumps(datos, ensure_ascii=False),
                     round(time.time() - t0, 2))
            return "OK"
        except Exception as exc:  # cuota, 5xx, JSON inválido, esquema inválido
            texto = str(exc)
            transitorio = any(s in texto for s in ("429", "503", "500", "overloaded", "Resource"))
            if intento == reintentos or not transitorio:
                _guardar(con, paquete, huella, n_pag, backend, modelo_usado, dpi,
                         "ERROR", f"{type(exc).__name__}: {texto[:500]}", None,
                         round(time.time() - t0, 2))
                return "ERROR"
            time.sleep(espera)
            espera = min(espera * 2, 120)   # retroceso exponencial
    return "ERROR"


def _guardar(con, paquete, huella, n_pag, backend, modelo, dpi,
             estado, mensaje, json_crudo, segundos) -> None:
    con.execute(
        """INSERT OR REPLACE INTO paquetes
           (clave, op, anio, mes, archivos, hash_contenido, n_paginas, backend,
            modelo, dpi, estado, mensaje_error, json_crudo, segundos, procesado_en)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?, datetime('now'))""",
        (paquete.clave, paquete.op, paquete.anio, paquete.mes,
         json.dumps([r.as_posix() for r in paquete.archivos], ensure_ascii=False),
         huella, n_pag, backend, modelo, dpi, estado, mensaje, json_crudo, segundos),
    )
    con.commit()


def procesar_lote(
    paquetes: List[Paquete],
    db: str = "soportes.sqlite",
    backend: str = "gemini",
    modelo: Optional[str] = None,
    pausa: float = 2.0,
    limite: Optional[int] = None,
) -> pd.DataFrame:
    """Procesa una lista de expedientes con caché y reporte de avance."""
    con = _abrir_db(db)
    seleccion = paquetes[:limite] if limite else paquetes
    resumen = []

    for i, paquete in enumerate(seleccion, start=1):
        estado = procesar_paquete(paquete, con, backend=backend, modelo=modelo)
        resumen.append({"clave": paquete.clave, "op": paquete.op,
                        "archivos": len(paquete.archivos), "estado": estado})
        print(f"[{i}/{len(seleccion)}] {paquete.clave} -> {estado}")
        if estado != "CACHE" and i < len(seleccion):
            time.sleep(pausa)

    con.close()
    return pd.DataFrame(resumen)


# ---------------------------------------------------------------------------
# 6. Normalización a tablas relacionales
# ---------------------------------------------------------------------------


def exportar_tablas(db: str = "soportes.sqlite") -> Dict[str, pd.DataFrame]:
    """Aplana el JSON anidado en cinco tablas con llaves de cruce.

    Devuelve: paquetes, documentos, conceptos, valores, evidencias.
    """
    con = _abrir_db(db)
    filas = con.execute(
        "SELECT clave, op, anio, mes, archivos, n_paginas, estado, mensaje_error, json_crudo "
        "FROM paquetes"
    ).fetchall()
    con.close()

    t_paq, t_doc, t_con, t_val, t_evi = [], [], [], [], []

    for clave, op, anio, mes, archivos, n_pag, estado, err, crudo in filas:
        datos = json.loads(crudo) if crudo else {}
        t_paq.append({
            "clave": clave, "op": op, "anio": anio, "mes": mes,
            "archivos": archivos, "n_paginas": n_pag, "estado": estado,
            "mensaje_error": err,
            "op_detectada_en_documento": datos.get("numero_op_detectado"),
            "n_documentos": len(datos.get("documentos", [])),
        })
        for doc in datos.get("documentos", []):
            doc_id = f"{clave}#{doc.get('doc_indice')}"
            emisor = doc.get("emisor") or {}
            receptor = doc.get("receptor") or {}
            benef = doc.get("beneficiario_del_giro") or {}
            t_doc.append({
                "doc_id": doc_id, "clave": clave, "op": op,
                "tipo": doc.get("tipo"),
                "pagina_inicio": doc.get("pagina_inicio"),
                "pagina_fin": doc.get("pagina_fin"),
                "legible": doc.get("legible"),
                "numero_documento": doc.get("numero_documento"),
                "fecha_documento": doc.get("fecha_documento"),
                "emisor_nombre": emisor.get("nombre"),
                "emisor_nit": _solo_digitos(emisor.get("nit_o_cc")),
                "receptor_nombre": receptor.get("nombre"),
                "receptor_nit": _solo_digitos(receptor.get("nit_o_cc")),
                "beneficiario_nombre": benef.get("nombre"),
                "beneficiario_nit": _solo_digitos(benef.get("nit_o_cc")),
                "cuenta_o_patrimonio": doc.get("cuenta_bancaria_o_patrimonio"),
                "op_mencionada": doc.get("numero_op_mencionado"),
                "observaciones": doc.get("observaciones"),
            })
            for j, c in enumerate(doc.get("conceptos", []), start=1):
                t_con.append({"linea_id": f"{doc_id}:C{j}", "doc_id": doc_id,
                              "clave": clave, "op": op, **c})
            for j, v in enumerate(doc.get("valores", []), start=1):
                t_val.append({"valor_id": f"{doc_id}:V{j}", "doc_id": doc_id,
                              "clave": clave, "op": op, **v})
            for j, e in enumerate(doc.get("evidencias", []), start=1):
                t_evi.append({"evidencia_id": f"{doc_id}:E{j}", "doc_id": doc_id,
                              "clave_paquete": clave, "op": op, **e})

    return {
        "paquetes": pd.DataFrame(t_paq),
        "documentos": pd.DataFrame(t_doc),
        "conceptos": pd.DataFrame(t_con),
        "valores": pd.DataFrame(t_val),
        "evidencias": pd.DataFrame(t_evi),
    }


def _solo_digitos(x: Optional[str]) -> Optional[str]:
    if not x:
        return None
    d = re.sub(r"\D", "", str(x))
    return d[:9] if len(d) == 10 else d or None   # descarta dígito de verificación


def guardar_excel(tablas: Dict[str, pd.DataFrame], ruta: str = "soportes_estructurados.xlsx") -> None:
    with pd.ExcelWriter(ruta, engine="openpyxl") as w:
        for nombre, df in tablas.items():
            df.to_excel(w, sheet_name=nombre[:31], index=False)
    print(f"[OK] {ruta}")


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    paquetes = descubrir_paquetes("Soportes_2015")
    print(f"{len(paquetes)} expedientes descubiertos "
          f"({sum(len(p.archivos) for p in paquetes)} archivos).")
    procesar_lote(paquetes, db="soportes.sqlite", backend="gemini", limite=20)
    guardar_excel(exportar_tablas("soportes.sqlite"))
