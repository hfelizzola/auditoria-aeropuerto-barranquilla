"""
reconciliacion.py
=================
Cuadre determinístico entre el valor bruto de un soporte (base + IVA) y el valor
registrado en UNIVERSALIDAD ("VALOR DEBITADO O ACREDITADO").

Si el bruto no coincide exactamente con lo registrado, se buscan combinaciones
estándar de retenciones colombianas que expliquen la diferencia:

    registrado = base + IVA - retefuente(% de base) - reteIVA(% del IVA) - reteICA(‰ de base)

dentro de una tolerancia de $1,50 COP (redondeo al peso de hasta tres retenciones).
Si ninguna combinación estándar cierra, el resultado es SIN_COMBINACION y la
observación debe decir "diferencia sin explicar": nunca se fuerza el cuadre.

Módulo sin dependencias externas para poder ajustarlo y probarlo con casos reales
(ver tests/test_reconciliacion.py). Las tarifas son parámetros editables.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Dict, List, Optional, Sequence

TOLERANCIA_COP = 1.50

# Tarifas en dos niveles. Primero se buscan combinaciones con tarifas ESTÁNDAR; solo si
# ninguna cierra se prueban las ATÍPICAS, y el resultado se marca COMBINACION_ATIPICA
# (un cuadre con tarifas raras puede ser casual: con ~1.500 combinaciones y ±$1,50 de
# tolerancia hay coincidencias espurias, p. ej. rank #3504 de la ronda 3501-4000).

# Retención en la fuente, % sobre la base (rango típico 0%-11%).
TARIFAS_RETEFUENTE: Sequence[float] = (0.0, 1.0, 1.5, 2.0, 2.5, 3.5, 4.0, 6.0, 10.0, 11.0)
TARIFAS_RETEFUENTE_ATIPICAS: Sequence[float] = (0.1, 0.5, 3.0, 5.0, 7.0, 8.0)
# Retención de IVA, % sobre el IVA.
TARIFAS_RETEIVA: Sequence[float] = (0.0, 15.0, 30.0, 100.0)
# Retención de ICA, por mil sobre la base (0‰-14‰) y tarifas especiales con decimales.
TARIFAS_RETEICA: Sequence[float] = tuple(float(x) for x in range(0, 15)) + (3.5,)
TARIFAS_RETEICA_ATIPICAS: Sequence[float] = (4.14, 4.8, 6.6, 6.9, 7.2, 8.28, 9.66, 11.04, 13.8)

EXACTO = "EXACTO"
COMBINACION = "COMBINACION"
COMBINACION_ATIPICA = "COMBINACION_ATIPICA"
SIN_COMBINACION = "SIN_COMBINACION"
SIN_DATOS = "SIN_DATOS"


def _pct(x: float) -> str:
    return f"{x:g}".replace(".", ",")


def _cop(x: float) -> str:
    entero, dec = f"{abs(x):,.2f}".split(".")
    return ("-" if x < 0 else "") + "$" + entero.replace(",", ".") + ("" if dec == "00" else "," + dec)


@dataclass
class Combinacion:
    retefuente_pct: float
    reteiva_pct: float
    reteica_por_mil: float
    retefuente: float
    reteiva: float
    reteica: float
    total_retenido: float
    residuo: float  # diferencia real - total_retenido (dentro de la tolerancia)

    @property
    def componentes(self) -> int:
        return sum(1 for t in (self.retefuente_pct, self.reteiva_pct, self.reteica_por_mil) if t)

    def describir(self) -> str:
        partes = []
        if self.retefuente_pct:
            partes.append(f"retefuente {_pct(self.retefuente_pct)}%")
        if self.reteiva_pct:
            partes.append(f"reteIVA {_pct(self.reteiva_pct)}%")
        if self.reteica_por_mil:
            partes.append(f"reteICA {_pct(self.reteica_por_mil)}‰")
        return (" + ".join(partes) or "sin retenciones") + f" = {_cop(self.total_retenido)}"


@dataclass
class ResultadoCuadre:
    estado: str
    base: Optional[float]
    iva: Optional[float]
    bruto: Optional[float]
    registrado: Optional[float]
    diferencia: Optional[float]
    sin_desglose_iva: bool = False
    combinaciones: List[Combinacion] = field(default_factory=list)
    nota: str = ""

    @property
    def cuadra(self) -> bool:
        return self.estado in (EXACTO, COMBINACION, COMBINACION_ATIPICA)

    def describir(self) -> str:
        if self.estado == SIN_DATOS:
            return f"Sin datos suficientes para cuadrar ({self.nota or 'falta bruto o registrado'})."
        if self.estado == EXACTO:
            return f"Bruto {_cop(self.bruto)} coincide exacto con el registrado {_cop(self.registrado)}."
        if self.estado in (COMBINACION, COMBINACION_ATIPICA):
            mejor = self.combinaciones[0]
            extra = f" ({len(self.combinaciones)} combinaciones posibles)" if len(self.combinaciones) > 1 else ""
            atipica = (" — combinación ATÍPICA (tarifas no estándar; puede ser coincidencia, verificar)"
                       if self.estado == COMBINACION_ATIPICA else "")
            return (f"Bruto {_cop(self.bruto)} → registrado {_cop(self.registrado)}: diferencia "
                    f"{_cop(self.diferencia)} explicada por {mejor.describir()}{extra}{atipica}.")
        return (f"Bruto {_cop(self.bruto)} vs registrado {_cop(self.registrado)}: diferencia "
                f"{_cop(self.diferencia)} sin combinación estándar de retenciones (diferencia sin explicar)"
                + (f"; {self.nota}" if self.nota else "") + ".")

    def to_dict(self) -> Dict:
        d = asdict(self)
        d["combinaciones"] = [dict(asdict(c), descripcion=c.describir()) for c in self.combinaciones]
        d["descripcion"] = self.describir()
        return d


def buscar_combinaciones(
    base: float,
    iva: float,
    diferencia: float,
    tolerancia: float = TOLERANCIA_COP,
    tarifas_retefuente: Sequence[float] = TARIFAS_RETEFUENTE,
    tarifas_reteiva: Sequence[float] = TARIFAS_RETEIVA,
    tarifas_reteica: Sequence[float] = TARIFAS_RETEICA,
) -> List[Combinacion]:
    """Todas las combinaciones (retefuente, reteIVA, reteICA) cuya suma explica `diferencia`
    dentro de la tolerancia, ordenadas de la más simple (menos componentes) a la más compleja."""
    encontradas: List[Combinacion] = []
    for rf in tarifas_retefuente:
        v_rf = base * rf / 100.0
        for riva in tarifas_reteiva:
            v_riva = iva * riva / 100.0
            for rica in tarifas_reteica:
                v_rica = base * rica / 1000.0
                total = v_rf + v_riva + v_rica
                residuo = diferencia - total
                if abs(residuo) <= tolerancia:
                    encontradas.append(Combinacion(rf, riva, rica, round(v_rf, 2), round(v_riva, 2),
                                                   round(v_rica, 2), round(total, 2), round(residuo, 2)))
    encontradas.sort(key=lambda c: (c.componentes, abs(c.residuo)))
    return encontradas


def reconciliar(
    base: Optional[float],
    iva: Optional[float],
    registrado: Optional[float],
    tolerancia: float = TOLERANCIA_COP,
    max_combinaciones: int = 5,
) -> ResultadoCuadre:
    """Cuadra base + IVA contra el valor registrado (se usa su valor absoluto)."""
    if base is None or registrado is None:
        return ResultadoCuadre(SIN_DATOS, base, iva, None, registrado, None,
                               nota="falta la base del documento" if base is None else "falta el valor registrado")
    registrado = abs(registrado)
    sin_desglose = iva is None
    iva_v = iva or 0.0
    bruto = base + iva_v
    diferencia = round(bruto - registrado, 2)

    if abs(diferencia) <= tolerancia:
        return ResultadoCuadre(EXACTO, base, iva, bruto, registrado, diferencia, sin_desglose)
    if diferencia < 0:
        return ResultadoCuadre(SIN_COMBINACION, base, iva, bruto, registrado, diferencia, sin_desglose,
                               nota="el registrado es MAYOR que el bruto del documento; ninguna retención lo explica")

    combos = buscar_combinaciones(base, iva_v, diferencia, tolerancia)
    if combos:
        return ResultadoCuadre(COMBINACION, base, iva, bruto, registrado, diferencia, sin_desglose,
                               combos[:max_combinaciones])
    combos = buscar_combinaciones(
        base, iva_v, diferencia, tolerancia,
        tarifas_retefuente=tuple(TARIFAS_RETEFUENTE) + tuple(TARIFAS_RETEFUENTE_ATIPICAS),
        tarifas_reteica=tuple(TARIFAS_RETEICA) + tuple(TARIFAS_RETEICA_ATIPICAS))
    if combos:
        return ResultadoCuadre(COMBINACION_ATIPICA, base, iva, bruto, registrado, diferencia, sin_desglose,
                               combos[:max_combinaciones])
    nota = "sin desglose de IVA en el documento (se probó con IVA = 0)" if sin_desglose else ""
    return ResultadoCuadre(SIN_COMBINACION, base, iva, bruto, registrado, diferencia, sin_desglose, nota=nota)


def cuadrar_retenciones_impresas(
    bruto: Optional[float],
    retenciones_impresas: Dict[str, float],
    registrado: Optional[float],
    tolerancia: float = TOLERANCIA_COP,
) -> Optional[bool]:
    """¿bruto - Σ(retenciones y descuentos impresos) == registrado? None si no hay datos."""
    if bruto is None or registrado is None or not retenciones_impresas:
        return None
    neto = bruto - sum(abs(v) for v in retenciones_impresas.values())
    return abs(neto - abs(registrado)) <= tolerancia


def coincide(a: Optional[float], b: Optional[float], tolerancia: float = TOLERANCIA_COP) -> Optional[bool]:
    if a is None or b is None:
        return None
    return abs(abs(a) - abs(b)) <= tolerancia
