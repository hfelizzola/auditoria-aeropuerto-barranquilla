"""
validacion_22_3.py
==================
Capa determinista de validación: cruza los soportes extraídos contra la hoja
UNIVERSALIDAD y aplica la lista taxativa de costos reconocibles del
Numeral 22.3 c) del Contrato de Concesión ANI 003 de 2015.

Aquí NO interviene ningún modelo de lenguaje. Toda conclusión es trazable a:
  (a) una regla identificada por código, y
  (b) una evidencia con archivo, página y texto literal.

Dos pruebas independientes por orden de pago:
  1. CUADRE DE VALOR  — ¿el valor girado coincide con el soporte?
  2. RECONOCIMIENTO   — ¿el concepto está en la lista taxativa, y se acreditó
                        la condición que el Contrato le exige?

Requisitos: pandas, openpyxl
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Set

import pandas as pd

# ---------------------------------------------------------------------------
# Parámetros de auditoría — editables y documentados
# ---------------------------------------------------------------------------

TOLERANCIA_ABS = 1_000.0     # COP: redondeos de centavos y ajustes al peso
TOLERANCIA_REL = 0.005       # 0,5% sobre el valor girado

# NIT de los receptores cuya calidad de "Consultor" está acreditada (AR3).
# Se completa con el concepto jurídico; por ahora, la UT estructuradora.
NITS_CONSULTOR: Set[str] = set()

# Prestamistas identificados, para distinguir AR9 de la exclusión de deuda.
NITS_PRESTAMISTA: Set[str] = set()


# ---------------------------------------------------------------------------
# Catálogo de componentes de la fórmula
# ---------------------------------------------------------------------------

COMPONENTES = {
    "AR1_SEGUROS":  "Primas y comisiones de las garantías y seguros del Contrato",
    "AR2_SUBCTA":   "Aportes a las Subcuentas de la Cuenta ANI",
    "AR3_EXITO":    "Comisión de Éxito al Consultor",
    "AR4_ESTUDIOS": "Estudios y Diseños de las Intervenciones y su Cronograma",
    "AR5_SOCAMB":   "Gestión Social y Ambiental",
    "AR6_PREDIAL":  "Gestión Predial",
    "AR7_INTERV":   "Actuaciones en las Intervenciones verificadas por el Interventor",
    "AR8_OYM":      "Operación y mantenimiento, administración e impuestos",
    "AR9_PRESTAM":  "Comisiones y otros pagos a Prestamistas distintos del servicio de deuda",
    "PL_ANTICIPO":  "Anticipo o fondo rotatorio: costo solo en la porción legalizada",
    "RA_RETRIB":    "Retribución o giro al Concesionario (resta en la fórmula)",
    "DYM_MORA":     "Deducciones, multas y sanciones (resta en la fórmula)",
    "EX_DEUDA":     "Servicio de la deuda: principal e intereses (exclusión expresa)",
    "EX_LISTA":     "Fuera de la lista taxativa del 22.3 c)",
    "NC_NOCOSTO":   "No constituye costo realizado",
    "SIN_DETERMINAR": "Evidencia insuficiente para clasificar",
}

ESTADO_RECONOCIBLE = "RECONOCIBLE"
ESTADO_CONDICIONADO = "RECONOCIBLE — CONDICIÓN NO ACREDITADA"
ESTADO_NO_RECONOCIBLE = "NO RECONOCIBLE"
ESTADO_OTRO_COMPONENTE = "OTRO COMPONENTE DE LA FÓRMULA"
ESTADO_INDETERMINADO = "INDETERMINADO"


@dataclass
class Regla:
    """Una regla del 22.3 c) expresada como condición sobre las evidencias."""

    codigo: str
    componente: str
    # Evidencias que activan la regla (basta una).
    dispara_si: List[str]
    # Evidencias que deben estar presentes para que el costo sea certificable.
    exige: List[str]
    # Evidencias cuya presencia anula el reconocimiento.
    excluye_si: List[str]
    estado_si_cumple: str
    estado_si_no_cumple: str
    fundamento: str


# Orden de precedencia: la primera regla que dispara gana. Las exclusiones
# expresas van primero para que no las absorba una categoría general.
REGLAS: List[Regla] = [
    Regla(
        "R01", "EX_DEUDA",
        dispara_si=["abono_a_capital_o_principal_de_credito", "intereses_de_credito"],
        exige=[], excluye_si=[],
        estado_si_cumple=ESTADO_NO_RECONOCIBLE,
        estado_si_no_cumple=ESTADO_NO_RECONOCIBLE,
        fundamento="22.3 c) AR9 excluye expresamente el servicio de la deuda "
                   "(intereses y principal).",
    ),
    Regla(
        "R02", "DYM_MORA",
        dispara_si=["interes_de_mora_multa_o_sancion"],
        exige=[], excluye_si=[],
        estado_si_cumple=ESTADO_OTRO_COMPONENTE,
        estado_si_no_cumple=ESTADO_OTRO_COMPONENTE,
        fundamento="Corresponde al componente DyM, que resta en la fórmula.",
    ),
    Regla(
        "R03", "EX_LISTA",
        dispara_si=["depreciacion_amortizacion_o_valorizacion"],
        exige=[], excluye_si=[],
        estado_si_cumple=ESTADO_NO_RECONOCIBLE,
        estado_si_no_cumple=ESTADO_NO_RECONOCIBLE,
        fundamento="El valor se define sobre valores brutos, sin depreciaciones, "
                   "amortizaciones ni valorizaciones.",
    ),
    Regla(
        "R04", "PL_ANTICIPO",
        dispara_si=["entrega_de_anticipo_o_fondo_rotatorio"],
        exige=["acta_o_relacion_de_legalizacion"],
        excluye_si=[],
        estado_si_cumple=ESTADO_RECONOCIBLE,
        estado_si_no_cumple=ESTADO_CONDICIONADO,
        fundamento="ARᵢ son costos realizados. Un anticipo es entrega de recursos "
                   "hasta que se legaliza con facturas y actas.",
    ),
    Regla(
        "R05", "AR3_EXITO",
        dispara_si=["pago_denominado_comision_de_exito"],
        exige=["receptor_identificado_como_consultor_estructurador"],
        excluye_si=[],
        estado_si_cumple=ESTADO_RECONOCIBLE,
        estado_si_no_cumple=ESTADO_CONDICIONADO,
        fundamento="22.3 c): Comisión de Éxito efectivamente desembolsada AL CONSULTOR.",
    ),
    Regla(
        "R06", "AR1_SEGUROS",
        dispara_si=["poliza_garantia_unica_de_cumplimiento",
                    "poliza_responsabilidad_civil_aeropuertos",
                    "poliza_todo_riesgo_danos_materiales"],
        exige=[], excluye_si=[],
        estado_si_cumple=ESTADO_RECONOCIBLE,
        estado_si_no_cumple=ESTADO_RECONOCIBLE,
        fundamento="22.3 c): primas y comisiones de la Garantía Única, la póliza de "
                   "RC Aeropuertos y el seguro de daños contra todo riesgo.",
    ),
    Regla(
        "R07", "EX_LISTA",
        dispara_si=["otra_poliza_no_prevista_en_el_contrato"],
        exige=[], excluye_si=["poliza_garantia_unica_de_cumplimiento",
                              "poliza_responsabilidad_civil_aeropuertos",
                              "poliza_todo_riesgo_danos_materiales"],
        estado_si_cumple=ESTADO_NO_RECONOCIBLE,
        estado_si_no_cumple=ESTADO_INDETERMINADO,
        fundamento="La lista de seguros del 22.3 c) es taxativa: solo las tres pólizas "
                   "señaladas en el Apéndice 1 — Parte Especial.",
    ),
    Regla(
        "R08", "AR2_SUBCTA",
        dispara_si=["subcuenta_ani_destino_identificada"],
        exige=[], excluye_si=[],
        estado_si_cumple=ESTADO_RECONOCIBLE,
        estado_si_no_cumple=ESTADO_CONDICIONADO,
        fundamento="22.3 c): aportes a las Subcuentas de la Cuenta ANI, Predios, "
                   "Compensaciones Ambientales, Redes Mayores, Redes Menores y "
                   "Promoción del Proyecto.",
    ),
    Regla(
        "R09", "AR6_PREDIAL",
        dispara_si=["gestion_predial"],
        exige=[], excluye_si=["cargo_a_subcuenta_predios"],
        estado_si_cumple=ESTADO_RECONOCIBLE,
        estado_si_no_cumple=ESTADO_NO_RECONOCIBLE,
        fundamento="22.3 c): Gestión Predial, exceptuando lo hecho contra la "
                   "Subcuenta Predios.",
    ),
    Regla(
        "R10", "AR5_SOCAMB",
        dispara_si=["gestion_social_o_ambiental"],
        exige=[], excluye_si=["cargo_a_subcuenta_compensaciones_ambientales"],
        estado_si_cumple=ESTADO_RECONOCIBLE,
        estado_si_no_cumple=ESTADO_NO_RECONOCIBLE,
        fundamento="22.3 c): Gestión Social y Ambiental, exceptuando lo hecho contra "
                   "la Subcuenta Compensaciones Ambientales.",
    ),
    Regla(
        "R11", "AR4_ESTUDIOS",
        dispara_si=["estudio_o_diseno_de_intervencion"],
        exige=[], excluye_si=[],
        estado_si_cumple=ESTADO_RECONOCIBLE,
        estado_si_no_cumple=ESTADO_RECONOCIBLE,
        fundamento="22.3 c): Estudios y Diseños de las Intervenciones y del "
                   "Cronograma para su ejecución.",
    ),
    Regla(
        "R12", "AR7_INTERV",
        dispara_si=["acta_de_obra_o_avance_suscrita", "intervencion_identificada"],
        exige=["firma_o_visto_bueno_del_interventor"],
        excluye_si=[],
        estado_si_cumple=ESTADO_RECONOCIBLE,
        estado_si_no_cumple=ESTADO_CONDICIONADO,
        fundamento="22.3 c): actuaciones ejecutadas en las Intervenciones por el "
                   "Concesionario Y VERIFICADAS POR EL INTERVENTOR.",
    ),
    Regla(
        "R13", "AR9_PRESTAM",
        dispara_si=["comision_u_honorario_a_prestamista"],
        exige=[],
        excluye_si=["abono_a_capital_o_principal_de_credito", "intereses_de_credito"],
        estado_si_cumple=ESTADO_RECONOCIBLE,
        estado_si_no_cumple=ESTADO_NO_RECONOCIBLE,
        fundamento="22.3 c): comisiones y otros pagos a los Prestamistas, DISTINTOS "
                   "del servicio de la deuda.",
    ),
    Regla(
        "R14", "AR8_OYM",
        dispara_si=["operacion_o_mantenimiento", "gasto_de_administracion",
                    "impuesto_tasa_o_contribucion"],
        exige=[], excluye_si=[],
        estado_si_cumple=ESTADO_RECONOCIBLE,
        estado_si_no_cumple=ESTADO_RECONOCIBLE,
        fundamento="22.3 c): costos de Operación y Mantenimiento, gastos de "
                   "administración e impuestos.",
    ),
]


# ---------------------------------------------------------------------------
# Prueba 1 — Cuadre de valor
# ---------------------------------------------------------------------------


def agregar_universalidad(ruta_xlsx: str, hoja: str = "UNIVERSALIDAD") -> pd.DataFrame:
    """Suma las imputaciones contables por orden de pago.

    Una OP puede tener varias filas en UNIVERSALIDAD, así que el cuadre se hace
    OP contra la SUMA de sus filas, nunca fila a fila.
    """
    df = pd.read_excel(ruta_xlsx, sheet_name=hoja)
    df.columns = [c.strip() for c in df.columns]
    df["FECHA DE PAGO O DESEMBOLSO"] = pd.to_datetime(
        df["FECHA DE PAGO O DESEMBOLSO"], errors="coerce"
    )
    df["mes_pago"] = df["FECHA DE PAGO O DESEMBOLSO"].dt.month
    df["op_norm"] = df["NUMERO OP"].apply(_normalizar_op)

    con_op = df[df["op_norm"].notna()].copy()
    # El número de OP se repite entre años y terceros: se cualifica con el año.
    con_op["clave_op"] = con_op["AÑO"].astype("Int64").astype(str) + "::OP" + con_op["op_norm"]

    agg = con_op.groupby("clave_op").agg(
        op=("op_norm", "first"),
        anio=("AÑO", "first"),
        n_filas=("op_norm", "size"),
        n_terceros=("IDENTIFICACION TERCERO", "nunique"),
        terceros=("TERCERO", lambda s: " | ".join(sorted(set(map(str, s))))),
        cuentas=("NOMBRE DE CUENTA CONTABLE", lambda s: " | ".join(sorted(set(map(str, s))))),
        valor_base=("VALOR BASE", "sum"),
        valor_iva=("VALOR IVA", "sum"),
        valor_girado=("VALOR  DEBITADO O ACREDITADO", "sum"),
        fecha_min=("FECHA DE PAGO O DESEMBOLSO", "min"),
        fecha_max=("FECHA DE PAGO O DESEMBOLSO", "max"),
        tiene_url=("URL", lambda s: s.notna().any()),
    ).reset_index()

    agg["valor_girado_abs"] = agg["valor_girado"].abs()
    # Bandera del hallazgo de soporte ambiguo: misma OP, varios terceros o años.
    agg["op_ambigua"] = agg["n_terceros"] > 1
    return agg


def _normalizar_op(x) -> Optional[str]:
    if pd.isna(x):
        return None
    s = str(x).strip()
    if re.fullmatch(r"\d{1,2}/\d{4}", s):     # "05/2015" es periodo de GMF, no OP
        return None
    m = re.search(r"(\d{1,6})", s)
    return str(int(m.group(1))) if m else None


def cuadrar_valores(
    agg_op: pd.DataFrame,
    valores: pd.DataFrame,
    paquetes: pd.DataFrame,
) -> pd.DataFrame:
    """Compara el valor girado contable contra las cifras del soporte.

    Reconstruye el puente documental:
        total_bruto − retenciones − amortización de anticipo = neto a pagar
    y lo contrasta con |VALOR DEBITADO| agregado de la OP.
    """
    if valores.empty:
        return agg_op.assign(estado_cuadre="SIN SOPORTE PROCESADO")

    v = valores.copy()
    v["op"] = v["op"].astype(str)
    v["clave_op"] = v.apply(
        lambda r: f"{str(r['clave']).split('-')[0]}::OP{r['op']}", axis=1
    )

    piv = v.pivot_table(index="clave_op", columns="tipo", values="monto",
                        aggfunc="sum", fill_value=0.0)
    for col in ["total_bruto", "subtotal", "iva", "retefuente", "reteica", "reteiva",
                "retegarantia", "amortizacion_de_anticipo", "descuento",
                "neto_a_pagar", "valor_girado_o_pagado", "valor_legalizado"]:
        if col not in piv.columns:
            piv[col] = 0.0

    piv["bruto_doc"] = piv["total_bruto"].where(
        piv["total_bruto"] != 0, piv["subtotal"] + piv["iva"]
    )
    piv["retenciones_doc"] = (piv["retefuente"] + piv["reteica"] +
                              piv["reteiva"] + piv["retegarantia"])
    piv["neto_calculado"] = (piv["bruto_doc"] - piv["retenciones_doc"]
                             - piv["amortizacion_de_anticipo"] - piv["descuento"])
    # Referencia preferida: la cifra impresa de giro; si no, el neto reconstruido.
    piv["neto_doc"] = piv["valor_girado_o_pagado"].where(
        piv["valor_girado_o_pagado"] != 0,
        piv["neto_a_pagar"].where(piv["neto_a_pagar"] != 0, piv["neto_calculado"]),
    )

    out = agg_op.merge(piv.reset_index(), on="clave_op", how="left")
    out["diferencia"] = out["valor_girado_abs"] - out["neto_doc"].fillna(0)
    out["diferencia_rel"] = out["diferencia"] / out["valor_girado_abs"].replace(0, pd.NA)

    def _estado(r):
        if pd.isna(r.get("neto_doc")) or r.get("neto_doc") == 0:
            return "SIN CIFRA EN EL SOPORTE"
        if abs(r["diferencia"]) <= TOLERANCIA_ABS:
            return "CUADRA"
        if abs(r["diferencia_rel"] or 1) <= TOLERANCIA_REL:
            return "CUADRA (dentro de tolerancia relativa)"
        return "DIFERENCIA"

    out["estado_cuadre"] = out.apply(_estado, axis=1)
    return out


# ---------------------------------------------------------------------------
# Prueba 2 — Reconocimiento del concepto
# ---------------------------------------------------------------------------


def clasificar_paquete(evidencias_op: pd.DataFrame) -> Dict:
    """Aplica las reglas en orden de precedencia sobre las evidencias de una OP."""
    presentes = set(evidencias_op.loc[evidencias_op["presente"] == True, "clave"])   # noqa: E712
    negadas = set(evidencias_op.loc[evidencias_op["presente"] == False, "clave"])    # noqa: E712

    for regla in REGLAS:
        if not (presentes & set(regla.dispara_si)):
            continue
        if presentes & set(regla.excluye_si):
            return _veredicto(regla, regla.estado_si_no_cumple,
                              "Presente una evidencia que anula el reconocimiento.",
                              presentes, negadas)
        faltantes = [e for e in regla.exige if e not in presentes]
        if faltantes:
            return _veredicto(regla, regla.estado_si_no_cumple,
                              "Condición contractual no acreditada: " + ", ".join(faltantes),
                              presentes, negadas)
        motivo = ("Condiciones acreditadas en el soporte."
                  if regla.estado_si_cumple == ESTADO_RECONOCIBLE
                  else "El soporte muestra la causal que determina esta clasificación.")
        return _veredicto(regla, regla.estado_si_cumple, motivo, presentes, negadas)

    return {
        "regla": None, "componente": "SIN_DETERMINAR",
        "estado": ESTADO_INDETERMINADO,
        "motivo": "Ninguna evidencia del soporte activa una categoría del 22.3 c).",
        "fundamento": None,
        "evidencias_presentes": ";".join(sorted(presentes)),
        "evidencias_negadas": ";".join(sorted(negadas)),
    }


def _veredicto(regla: Regla, estado: str, motivo: str,
               presentes: Set[str], negadas: Set[str]) -> Dict:
    return {
        "regla": regla.codigo,
        "componente": regla.componente,
        "estado": estado,
        "motivo": motivo,
        "fundamento": regla.fundamento,
        "evidencias_presentes": ";".join(sorted(presentes)),
        "evidencias_negadas": ";".join(sorted(negadas)),
    }


def clasificar_todos(evidencias: pd.DataFrame) -> pd.DataFrame:
    """Veredicto por orden de pago."""
    if evidencias.empty:
        return pd.DataFrame()
    filas = []
    for clave, grupo in evidencias.groupby("clave_paquete"):
        v = clasificar_paquete(grupo)
        v["clave_paquete"] = clave
        v["op"] = grupo["op"].iloc[0]
        filas.append(v)
    return pd.DataFrame(filas)


# ---------------------------------------------------------------------------
# Reporte consolidado
# ---------------------------------------------------------------------------


def construir_reporte(
    ruta_universalidad: str,
    tablas: Dict[str, pd.DataFrame],
    salida: str = "Validacion_Soportes_22_3.xlsx",
) -> Dict[str, pd.DataFrame]:
    """Une cuadre de valor, veredicto de concepto y pista de evidencia."""
    agg = agregar_universalidad(ruta_universalidad)
    cuadre = cuadrar_valores(agg, tablas.get("valores", pd.DataFrame()),
                             tablas.get("paquetes", pd.DataFrame()))
    veredictos = clasificar_todos(tablas.get("evidencias", pd.DataFrame()))

    if not veredictos.empty:
        veredictos["clave_op"] = veredictos["clave_paquete"].str.replace(
            r"^(\d{4})-\d{2}::", r"\1::", regex=True
        )
        consolidado = cuadre.merge(
            veredictos.drop(columns=["op"]), on="clave_op", how="left"
        )
    else:
        consolidado = cuadre.assign(estado=ESTADO_INDETERMINADO)

    consolidado["prioridad"] = consolidado.apply(_prioridad, axis=1)
    consolidado = consolidado.sort_values(
        ["prioridad", "valor_girado_abs"], ascending=[True, False]
    )

    resumen = (consolidado.groupby(["componente", "estado"], dropna=False)
               .agg(n_ops=("clave_op", "nunique"),
                    valor=("valor_girado_abs", "sum"))
               .reset_index().sort_values("valor", ascending=False))

    with pd.ExcelWriter(salida, engine="openpyxl") as w:
        consolidado.to_excel(w, sheet_name="Consolidado_OP", index=False)
        resumen.to_excel(w, sheet_name="Resumen_componente", index=False)
        for nombre in ["documentos", "conceptos", "valores", "evidencias"]:
            if nombre in tablas and not tablas[nombre].empty:
                tablas[nombre].to_excel(w, sheet_name=nombre[:31], index=False)
    print(f"[OK] {salida}")
    return {"consolidado": consolidado, "resumen": resumen}


def _prioridad(r) -> int:
    """1 es la prioridad más alta de revisión humana."""
    estado = r.get("estado")
    cuadre = r.get("estado_cuadre", "")
    if estado == ESTADO_NO_RECONOCIBLE:
        return 1
    if cuadre == "DIFERENCIA":
        return 2
    if estado == ESTADO_CONDICIONADO:
        return 3
    if estado in (ESTADO_INDETERMINADO, None) or cuadre.startswith("SIN"):
        return 4
    if estado == ESTADO_OTRO_COMPONENTE:
        return 5
    return 6


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    from extraccion_soportes import exportar_tablas

    tablas = exportar_tablas("soportes.sqlite")
    construir_reporte("_A_BAS_1.xlsx", tablas)
