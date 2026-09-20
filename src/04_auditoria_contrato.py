"""
04_auditoria_contrato.py
========================
Módulo de consolidación y dictamen contractual según el Numeral 22.3 c)
del Contrato de Concesión ANI 003 de 2015 para el Aeropuerto Ernesto Cortissoz.

Objetivo:
- Unir (merge) la base muestral contable `data/output/base_auditoria_pareto.csv`
  con las extracciones de Gemini `data/output/resultados_gemini.csv` por número de OP.
- Aplicar las reglas contractuales taxativas para emitir el "DICTAMEN_AUDITORIA":
  * Si la categoría es "EXCLUIDO": dictamen "NO RECONOCIBLE" (exclusión expresa del Numeral 22.3 c).
  * Si "es_parte_relacionada" es True: dictamen "POR VERIFICAR" (exige prueba de mercado / vinculados).
  * Si la categoría es "AR7" o "requiere_acta_interventoria" es True: dictamen "CONDICIONADO" (a interventoría).
  * Si es "AR1", "AR4", o "AR8": dictamen "RECONOCIBLE".
  * Otras categorías de la fórmula (AR2, AR3, AR5, AR6, AR9): dictamen "RECONOCIBLE".
  * Si no cuenta con soporte procesado en datalake: "PENDIENTE DE SOPORTE".
- Exportar la base auditada completa a `data/output/reporte_auditoria_final.csv`.
- Generar tablero resumen de control de costos y porcentajes de reconocimiento.
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path
from typing import Optional

import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)


# Definición formal de los dictámenes de auditoría
DICTAMEN_NO_RECONOCIBLE = "NO RECONOCIBLE"
DICTAMEN_POR_VERIFICAR = "POR VERIFICAR"
DICTAMEN_CONDICIONADO = "CONDICIONADO"
DICTAMEN_RECONOCIBLE = "RECONOCIBLE"
DICTAMEN_PENDIENTE = "PENDIENTE DE SOPORTE"


def evaluar_dictamen(fila: pd.Series) -> str:
    """
    Aplica el árbol de decisión contractual del Numeral 22.3 c)
    en estricto orden de prelación jurídica.
    """
    categoria = str(fila.get("CATEGORIA_CONTRACTUAL", "")).strip().upper()
    es_relacionada = bool(fila.get("ES_PARTE_RELACIONADA", False))
    requiere_acta = bool(fila.get("REQUIERE_ACTA_INTERVENTORIA", False))

    # Si no se tiene información procesada por Gemini para esta OP
    if not categoria or categoria in ["NAN", "NONE", ""]:
        return DICTAMEN_PENDIENTE

    # 1. Exclusión expresa contractual (servicio a la deuda, multas, fuera de lista taxativa)
    if categoria == "EXCLUIDO":
        return DICTAMEN_NO_RECONOCIBLE

    # 2. Operaciones con partes relacionadas (sociedades vinculadas al concesionario)
    if es_relacionada:
        return DICTAMEN_POR_VERIFICAR

    # 3. Intervenciones de obra (AR7) o anticipos que requieran acta de interventoría
    if categoria == "AR7" or requiere_acta:
        return DICTAMEN_CONDICIONADO

    # 4. Costos legalmente autorizados y reconocibles de la fórmula
    if categoria in ["AR1", "AR2", "AR3", "AR4", "AR5", "AR6", "AR8", "AR9"]:
        return DICTAMEN_RECONOCIBLE

    # 5. Valor por defecto ante cualquier otra categoría no clasificada
    return DICTAMEN_POR_VERIFICAR


def explicar_motivo_dictamen(fila: pd.Series) -> str:
    """Retorna la motivación contractual auditable de la calificación."""
    dictamen = fila.get("DICTAMEN_AUDITORIA")
    categoria = str(fila.get("CATEGORIA_CONTRACTUAL", "")).strip().upper()
    emisor = str(fila.get("EMISOR_NOMBRE", ""))

    if dictamen == DICTAMEN_NO_RECONOCIBLE:
        return "Exclusión expresa Numeral 22.3 c) (Servicio de deuda: capital/intereses o costo no taxativo)."
    elif dictamen == DICTAMEN_POR_VERIFICAR:
        return f"Transacción con parte relacionada ({emisor}). Exige acreditar prueba de precios de mercado."
    elif dictamen == DICTAMEN_CONDICIONADO:
        return "Costo asociado a Intervenciones (AR7) o anticipo. Requiere verificación y visto bueno del Interventor."
    elif dictamen == DICTAMEN_RECONOCIBLE:
        return f"Costo admisible clasificado taxativamente en componente {categoria} del Numeral 22.3 c)."
    else:
        return "Soporte documental pendiente de descarga o procesamiento en datalake."


def ejecutar_auditoria(
    ruta_pareto: Optional[Path] = None,
    ruta_gemini: Optional[Path] = None,
    ruta_salida: Optional[Path] = None,
) -> pd.DataFrame:
    """
    Ejecuta el cruce contable, calcula dictámenes y exporta el reporte final de auditoría.
    """
    base_dir = Path(__file__).resolve().parent.parent

    if ruta_pareto is None:
        ruta_pareto = base_dir / "data" / "output" / "base_auditoria_pareto.csv"

    if ruta_gemini is None:
        ruta_gemini = base_dir / "data" / "output" / "resultados_gemini.csv"

    if ruta_salida is None:
        ruta_salida = base_dir / "data" / "output" / "reporte_auditoria_final.csv"

    ruta_salida.parent.mkdir(parents=True, exist_ok=True)

    if not ruta_pareto.exists():
        raise FileNotFoundError(
            f"No existe la base Pareto en '{ruta_pareto}'. Ejecute '01_pareto_universalidad.py'."
        )

    logger.info(f"Cargando base muestral Pareto desde '{ruta_pareto.name}'...")
    df_pareto = pd.read_csv(ruta_pareto)

    # Cargar resultados de Gemini si existen
    if ruta_gemini.exists():
        logger.info(f"Cargando resultados de Gemini desde '{ruta_gemini.name}'...")
        df_gemini = pd.read_csv(ruta_gemini)
    else:
        logger.warning(
            f"No se encontró '{ruta_gemini.name}'. Se creará estructura vacía para cruce."
        )
        df_gemini = pd.DataFrame(
            columns=[
                "NUMERO_OP",
                "ARCHIVO_PDF",
                "NUMERO_FACTURA_OP",
                "EMISOR_NOMBRE",
                "EMISOR_NIT",
                "CONCEPTO_PAGO",
                "VALOR_TOTAL_EXTRAIDO",
                "CATEGORIA_CONTRACTUAL",
                "REQUIERE_ACTA_INTERVENTORIA",
                "ES_PARTE_RELACIONADA",
                "JUSTIFICACION",
            ]
        )

    def _normalizar_llave(val):
        if pd.isna(val) or val is None:
            return ""
        s = str(val).strip()
        if s in ["nan", "None", "", "NaN"]:
            return ""
        import re
        m = re.match(r"^(\d+)\.0+$", s)
        if m:
            return m.group(1)
        try:
            num = float(s)
            if num.is_integer():
                return str(int(num))
        except (ValueError, TypeError):
            pass
        return s

    # Asegurar llaves de cruce estandarizadas como cadenas limpias
    col_op_pareto = "OP_LIMPIA" if "OP_LIMPIA" in df_pareto.columns else "NUMERO OP"
    df_pareto["LLAVE_OP"] = df_pareto[col_op_pareto].apply(_normalizar_llave)
    df_pareto.loc[df_pareto["LLAVE_OP"] == "", "LLAVE_OP"] = None

    df_gemini["LLAVE_OP"] = df_gemini["NUMERO_OP"].apply(_normalizar_llave)
    df_gemini.loc[df_gemini["LLAVE_OP"] == "", "LLAVE_OP"] = None

    # Cruce por izquierda (Left Join) para mantener la totalidad de la muestra Pareto
    logger.info("Realizando cruce contable de órdenes de pago...")
    df_final = pd.merge(
        df_pareto,
        df_gemini.drop(columns=["NUMERO_OP"], errors="ignore"),
        on="LLAVE_OP",
        how="left",
        suffixes=("_CONTABLE", "_GEMINI"),
    )

    # 1. Aplicar evaluación de dictamen
    df_final["DICTAMEN_AUDITORIA"] = df_final.apply(evaluar_dictamen, axis=1)

    # 2. Asignar motivación jurídica contractual
    df_final["MOTIVO_DICTAMEN"] = df_final.apply(explicar_motivo_dictamen, axis=1)

    # 3. Cuadre de valores entre contabilidad y soporte facturado
    if "VALOR_TOTAL_EXTRAIDO" in df_final.columns:
        df_final["VALOR_TOTAL_EXTRAIDO"] = pd.to_numeric(
            df_final["VALOR_TOTAL_EXTRAIDO"], errors="coerce"
        ).fillna(0.0)
        df_final["DIFERENCIA_VALOR"] = (
            df_final["VALOR_ABSOLUTO"] - df_final["VALOR_TOTAL_EXTRAIDO"]
        )
    else:
        df_final["VALOR_TOTAL_EXTRAIDO"] = 0.0
        df_final["DIFERENCIA_VALOR"] = df_final["VALOR_ABSOLUTO"]

    # 4. Guardar resultado final
    df_final.to_csv(ruta_salida, index=False, encoding="utf-8-sig")

    # 5. Generar tablero de control y estadísticas ejecutivas
    total_auditado = df_final["VALOR_ABSOLUTO"].sum()
    resumen = (
        df_final.groupby("DICTAMEN_AUDITORIA")
        .agg(
            Cantidad_Transacciones=("VALOR_ABSOLUTO", "count"),
            Monto_Total_COP=("VALOR_ABSOLUTO", "sum"),
        )
        .reset_index()
    )
    resumen["Participacion_Pct"] = (resumen["Monto_Total_COP"] / total_auditado) * 100.0
    resumen = resumen.sort_values(by="Monto_Total_COP", ascending=False)

    logger.info("=" * 75)
    logger.info("INFORME EJECUTIVO DE AUDITORÍA CONTRACTUAL (NUMERAL 22.3 c):")
    logger.info(f"Total registros auditados: {len(df_final):,}")
    logger.info(f"Monto total auditado:     ${total_auditado:,.2f} COP")
    logger.info("-" * 75)
    for _, row in resumen.iterrows():
        dictamen = row["DICTAMEN_AUDITORIA"]
        cant = int(row["Cantidad_Transacciones"])
        monto = row["Monto_Total_COP"]
        pct = row["Participacion_Pct"]
        logger.info(
            f" * {dictamen:<22} | {cant:>5} reg | ${monto:>18,.2f} COP | {pct:>6.2f}%"
        )
    logger.info("=" * 75)
    logger.info(f"Reporte final guardado exitosamente en: '{ruta_salida}'")

    return df_final


def main():
    parser = argparse.ArgumentParser(description="Consolidación y dictamen de auditoría contractual.")
    parser.add_argument("--salida", type=str, default=None, help="Ruta personalizada de salida para el reporte final.")
    args = parser.parse_args()

    ejecutar_auditoria(ruta_salida=Path(args.salida) if args.salida else None)


if __name__ == "__main__":
    main()
