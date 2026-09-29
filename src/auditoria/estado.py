"""
estado.py
=========
Checkpoints ligeros del lote, para reanudar sin repetir trabajo ni llamadas al LLM.

- Un JSON por rank en data/output/estado/<lote>/ranks/rNNNNN.json con las secciones
  `descarga`, `extraccion` (hechos del modelo de visión), `senales` (cruce
  determinístico), `clasificacion` y `error`. Se escriben de forma atómica.
- Manifiesto de descargas `_manifiesto_descarga.csv` en la carpeta de soportes
  (rank -> archivo(s) local(es) -> estado), fusionado por rank.

El Excel de lote sigue siendo la fuente de verdad de qué filas están revisadas:
estos archivos solo evitan repetir descargas y llamadas al modelo.
"""

from __future__ import annotations

import csv
import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

from .config import DIR_ESTADO


def ahora() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _escribir_atomico(ruta: Path, texto: str) -> None:
    ruta.parent.mkdir(parents=True, exist_ok=True)
    tmp = ruta.with_name(ruta.name + ".tmp")
    tmp.write_text(texto, encoding="utf-8")
    os.replace(tmp, ruta)


class EstadoLote:
    def __init__(self, slug: str, base: Optional[Path] = None):
        # AUDITORIA_DIR_ESTADO permite redirigir los checkpoints (p. ej. en pruebas)
        base = base or Path(os.getenv("AUDITORIA_DIR_ESTADO") or DIR_ESTADO)
        self.dir = Path(base) / slug
        self.dir_ranks = self.dir / "ranks"
        self.dir_ranks.mkdir(parents=True, exist_ok=True)

    def ruta_rank(self, rank: int) -> Path:
        return self.dir_ranks / f"r{int(rank):05d}.json"

    def leer(self, rank: int) -> Dict[str, Any]:
        ruta = self.ruta_rank(rank)
        if not ruta.exists():
            return {}
        try:
            return json.loads(ruta.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return {}

    def actualizar(self, rank: int, **secciones: Any) -> Dict[str, Any]:
        datos = self.leer(rank)
        datos["rank"] = int(rank)
        for k, v in secciones.items():
            if v is None:
                datos.pop(k, None)
            else:
                datos[k] = v
        _escribir_atomico(self.ruta_rank(rank), json.dumps(datos, ensure_ascii=False, indent=2, default=str))
        return datos

    def todos(self) -> Dict[int, Dict[str, Any]]:
        salida = {}
        for p in sorted(self.dir_ranks.glob("r*.json")):
            try:
                d = json.loads(p.read_text(encoding="utf-8"))
                salida[int(d["rank"])] = d
            except (json.JSONDecodeError, KeyError, ValueError):
                continue
        return salida

    def escribir_pendientes(self, filas: List[Dict[str, Any]]) -> Path:
        ruta = self.dir / "pendientes_sin_resultado.csv"
        campos = ["rank", "op", "tercero", "valor", "motivo", "url"]
        ruta.parent.mkdir(parents=True, exist_ok=True)
        with open(ruta, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=campos, extrasaction="ignore")
            w.writeheader()
            w.writerows(filas)
        return ruta


# ---------------------------------------------------------------------------
# Manifiesto de descargas
# ---------------------------------------------------------------------------

CAMPOS_MANIFIESTO = ["rank", "fila_universalidad", "op", "tercero", "fecha", "valor", "moneda", "url",
                     "archivo_local", "estado", "detalle", "actualizado"]


def leer_manifiesto(ruta: Path) -> Dict[str, Dict[str, Any]]:
    if not ruta.exists():
        return {}
    with open(ruta, newline="", encoding="utf-8-sig") as fh:
        return {str(r.get("rank")): r for r in csv.DictReader(fh) if r.get("rank")}


def actualizar_manifiesto(ruta: Path, registros: Iterable[Dict[str, Any]]) -> None:
    """Fusiona por rank (el último registro de cada rank prevalece) y reescribe el CSV."""
    actual = leer_manifiesto(ruta)
    for reg in registros:
        reg = dict(reg, actualizado=reg.get("actualizado") or ahora())
        actual[str(reg["rank"])] = {**actual.get(str(reg["rank"]), {}), **reg}

    def _orden(k: str) -> Any:
        return (0, int(k)) if str(k).isdigit() else (1, str(k))

    ruta.parent.mkdir(parents=True, exist_ok=True)
    tmp = ruta.with_name(ruta.name + ".tmp")
    with open(tmp, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=CAMPOS_MANIFIESTO, extrasaction="ignore")
        w.writeheader()
        for k in sorted(actual, key=_orden):
            w.writerow(actual[k])
    os.replace(tmp, ruta)
