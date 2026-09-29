"""
config.py
=========
Constantes y utilidades compartidas: rutas del proyecto, carga de .env, logging,
taxonomía de clasificación (5 etiquetas exactas), partes relacionadas y
normalizadores de OP / NIT / nombres de tercero.
"""

from __future__ import annotations

import difflib
import importlib
import logging
import os
import re
import sys
import unicodedata
from pathlib import Path
from typing import Any, Optional

# ---------------------------------------------------------------------------
# Rutas
# ---------------------------------------------------------------------------

DIR_SRC = Path(__file__).resolve().parent.parent
BASE_DIR = DIR_SRC.parent
DIR_DATA = BASE_DIR / "data"
DIR_RAW = DIR_DATA / "raw"
DIR_OUTPUT = DIR_DATA / "output"
DIR_SOPORTES = DIR_DATA / "soportes"
DIR_ESTADO = DIR_OUTPUT / "estado"
DIR_CACHE = DIR_OUTPUT / ".cache"

RUTAS_CANDIDATAS_UNIVERSALIDAD = [
    DIR_RAW / "(A)BAS~1.xlsx",
    BASE_DIR / "(A)BAS~1.xlsx",
]

# ---------------------------------------------------------------------------
# Taxonomía de clasificación (exactamente estas 5 cadenas, sin variantes)
# ---------------------------------------------------------------------------

COHERENTE = "COHERENTE"
COHERENTE_VER_NOTA = "COHERENTE — VER NOTA"  # guion largo (U+2014), como en los lotes manuales
INCONCLUSO = "INCONCLUSO"
NO_CORRESPONDE = "NO CORRESPONDE"
TERCERO_NO_COINCIDE = "TERCERO NO COINCIDE"

ETIQUETAS = (COHERENTE, COHERENTE_VER_NOTA, INCONCLUSO, NO_CORRESPONDE, TERCERO_NO_COINCIDE)


def normalizar_etiqueta(valor: Any) -> Optional[str]:
    """Devuelve una de las 5 etiquetas exactas o None si el texto no es reconocible.

    Tolera variantes de guion y espacios ("COHERENTE - VER NOTA", "coherente–ver nota").
    """
    if valor is None:
        return None
    txt = re.sub(r"\s+", " ", str(valor).strip().upper())
    txt = re.sub(r"\s*[-–—]+\s*", " — ", txt)
    for etiqueta in ETIQUETAS:
        if txt == etiqueta:
            return etiqueta
    return None


# ---------------------------------------------------------------------------
# Partes relacionadas (se vigilan, no se descartan automáticamente)
# ---------------------------------------------------------------------------

PARTES_RELACIONADAS = {
    "NAB": {"nombre": "NUEVO AEROPUERTO DE BARRANQUILLA SAS", "nit": "900913341",
            "patron": r"NUEVO\s+AEROPUERTO\s+DE\s+BARRANQUILLA"},
    "GAC": {"nombre": "GRUPO AEROPORTUARIO DEL CARIBE SAS", "nit": "900817115",
            "patron": r"GRUPO\s+AEROPORTUARIO\s+DEL\s+CARIBE"},
    "OAC": {"nombre": "OPERADORA AEROPORTUARIA DEL CARIBE SAS", "nit": "900849079",
            "patron": r"OPERADORA\s+AEROPORTUARIA\s+DEL\s+CARIBE"},
}


def detectar_parte_relacionada(nombre: Any = None, nit: Any = None) -> Optional[str]:
    """Devuelve 'NAB' / 'GAC' / 'OAC' si el nombre o NIT corresponde a una parte relacionada."""
    nit_norm = normalizar_nit(nit)
    nombre_txt = quitar_tildes(str(nombre or "")).upper()
    for sigla, datos in PARTES_RELACIONADAS.items():
        if nit_norm and nit_norm == datos["nit"]:
            return sigla
        if nombre_txt and re.search(datos["patron"], nombre_txt):
            return sigla
    return None


# ---------------------------------------------------------------------------
# Traslados internos RETEG (cuentas propias del P.A.)
# ---------------------------------------------------------------------------

# Cuenta "contable" que en realidad es una cuenta de ahorros propia del P.A.
PATRON_CUENTA_PROPIA_PA = re.compile(
    r"^\s*AHO\s+\d+.*\b(P\.?\s?A\.?|PA)\b.*(ERNESTO|AEROPUERTO|CORT)", re.IGNORECASE
)
PATRON_RETEG = re.compile(r"\bRETEG\b", re.IGNORECASE)


def es_traslado_reteg(tercero: Any, cuenta_contable: Any, etiqueta_soporte: Any) -> bool:
    """Regla dura: traslado interno entre cuentas propias del P.A. (patrón RETEG)."""
    if etiqueta_soporte is not None and PATRON_RETEG.search(str(etiqueta_soporte)):
        return True
    tercero_txt = quitar_tildes(str(tercero or "")).upper()
    return "BANCOLOMBIA" in tercero_txt and bool(PATRON_CUENTA_PROPIA_PA.search(str(cuenta_contable or "")))


# ---------------------------------------------------------------------------
# Normalizadores
# ---------------------------------------------------------------------------

def quitar_tildes(txt: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", txt) if unicodedata.category(c) != "Mn")


def normalizar_op(valor: Any) -> Optional[str]:
    """'1956', 1956, 1956.0, ' 1956 ' -> '1956'. Textos no numéricos se devuelven limpios."""
    if valor is None:
        return None
    txt = str(valor).strip()
    if txt.lower() in ("", "nan", "none", "nat", "<na>"):
        return None
    m = re.match(r"^0*(\d+)(?:\.0+)?$", txt)
    if m:
        return m.group(1) or "0"
    return txt


def normalizar_nit(valor: Any) -> Optional[str]:
    """Deja solo dígitos y quita el dígito de verificación cuando viene separado por guion."""
    if valor is None:
        return None
    txt = str(valor).strip()
    if txt.lower() in ("", "nan", "none"):
        return None
    txt = re.sub(r"\.0+$", "", txt)
    txt = txt.split("-")[0]
    digitos = re.sub(r"\D", "", txt)
    return digitos or None


def nits_coinciden(nit_a: Any, nit_b: Any) -> Optional[bool]:
    """True/False si ambos NIT existen; None si falta alguno.

    Tolera el dígito de verificación pegado (9005258861 vs 900525886).
    """
    a, b = normalizar_nit(nit_a), normalizar_nit(nit_b)
    if not a or not b:
        return None
    if a == b:
        return True
    corto, largo = sorted((a, b), key=len)
    return len(largo) == len(corto) + 1 and largo.startswith(corto)


_FORMAS_SOCIETARIAS = {
    "SAS", "SA", "S", "A", "LTDA", "LIMITADA", "LIMIT", "ESP", "E", "P", "SCA", "SC", "EU",
    "CIA", "Y", "DE", "DEL", "LA", "EL", "LOS", "LAS", "0", "SUCURSAL", "COLOMBIA",
}


def normalizar_nombre(nombre: Any) -> str:
    txt = quitar_tildes(str(nombre or "")).upper()
    txt = re.sub(r"[^A-Z0-9 ]", " ", txt)
    tokens = [t for t in txt.split() if t not in _FORMAS_SOCIETARIAS]
    return " ".join(tokens)


def similitud_nombres(a: Any, b: Any) -> float:
    """Similitud 0-1 entre dos razones sociales, ignorando formas societarias y tildes."""
    na, nb = normalizar_nombre(a), normalizar_nombre(b)
    if not na or not nb:
        return 0.0
    if na == nb or na in nb or nb in na:
        return 1.0
    ratio = difflib.SequenceMatcher(None, na, nb).ratio()
    ta, tb = set(na.split()), set(nb.split())
    jaccard = len(ta & tb) / len(ta | tb) if ta | tb else 0.0
    return max(ratio, jaccard)


def a_float(valor: Any) -> Optional[float]:
    try:
        if valor is None:
            return None
        f = float(valor)
        if f != f:  # NaN
            return None
        return f
    except (TypeError, ValueError):
        return None


def fmt_cop(valor: Optional[float]) -> str:
    """$12.345.678,90 (formato colombiano, como en las observaciones manuales)."""
    if valor is None:
        return "N/D"
    signo = "-" if valor < 0 else ""
    entero, dec = f"{abs(valor):,.2f}".split(".")
    entero = entero.replace(",", ".")
    return f"{signo}${entero}" + ("" if dec == "00" else f",{dec}")


# ---------------------------------------------------------------------------
# Entorno, logging e importación de scripts numerados
# ---------------------------------------------------------------------------

def cargar_env() -> None:
    """Carga .env de la raíz (python-dotenv si existe; si no, lectura nativa)."""
    ruta = BASE_DIR / ".env"
    try:
        from dotenv import load_dotenv
        load_dotenv(ruta)
        return
    except ImportError:
        pass
    if ruta.exists():
        for linea in ruta.read_text(encoding="utf-8").splitlines():
            linea = linea.strip()
            if linea and not linea.startswith("#") and "=" in linea:
                k, v = linea.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())


def configurar_logging(nivel: int = logging.INFO, archivo: Optional[Path] = None) -> None:
    handlers: list = [logging.StreamHandler(sys.stdout)]
    if archivo is not None:
        archivo.parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(archivo, encoding="utf-8"))
    logging.basicConfig(level=nivel, format="%(asctime)s [%(levelname)s] %(message)s",
                        handlers=handlers, force=True)


def importar_script(nombre_modulo: str):
    """Importa un script numerado de src/ (p. ej. '01_pareto_universalidad') como módulo."""
    if str(DIR_SRC) not in sys.path:
        sys.path.insert(0, str(DIR_SRC))
    return importlib.import_module(nombre_modulo)
