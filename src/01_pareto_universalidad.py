"""
01_pareto_universalidad.py
==========================
Módulo para la lectura de la base contable del fideicomiso y selección
muestral bajo el principio de Pareto (80/20) para el Numeral 22.3 c).

Objetivo:
- Leer el archivo Excel `data/raw/(A)BAS~1.xlsx` en la hoja "UNIVERSALIDAD".
- Limpiar y normalizar los nombres de columnas y tipos de datos.
- Convertir "VALOR DEBITADO O ACREDITADO" a su valor absoluto.
- Ordenar de mayor a menor y calcular el porcentaje individual y acumulado.
- Filtrar las transacciones de mayor materialidad (hasta el 80% acumulado / top 1000).
- Exportar la base resultante a `data/output/base_auditoria_pareto.csv`.
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd

# Configuración de registro / logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)


def buscar_archivo_origen(rutas_candidatas: list[Path]) -> Path:
    """Busca y retorna la primera ruta existente de la lista candidata."""
    for ruta in rutas_candidatas:
        if ruta.exists():
            return ruta
    raise FileNotFoundError(
        f"No se encontró el archivo Excel en ninguna de las siguientes rutas: "
        f"{[str(r) for r in rutas_candidatas]}"
    )


def cargar_universalidad(ruta_excel: Path, hoja: str = "UNIVERSALIDAD") -> pd.DataFrame:
    """
    Carga la hoja UNIVERSALIDAD del archivo Excel y estandariza los nombres de columnas.
    """
    logger.info(f"Cargando hoja '{hoja}' desde '{ruta_excel}'...")
    df = pd.read_excel(ruta_excel, sheet_name=hoja)
    logger.info(f"Registros leídos inicialmente: {len(df):,}")

    # Limpiar espacios en blanco al inicio/final de los nombres de columnas
    df.columns = [c.strip() if isinstance(c, str) else str(c) for c in df.columns]

    return df


def aplicar_pareto(
    df: pd.DataFrame,
    col_valor: str = "VALOR DEBITADO O ACREDITADO",
    pareto_umbral: float = 80.0,
    top_max: Optional[int] = 1000,
    filtrar_solo_80: bool = False,
) -> Tuple[pd.DataFrame, dict]:
    """
    Calcula el valor absoluto, ordena de mayor a menor y extrae la muestra Pareto
    hasta alcanzar el umbral porcentual o un número máximo de registros.
    """
    if col_valor not in df.columns:
        # Búsqueda difusa por si hay variaciones en el nombre
        candidatos = [c for c in df.columns if "DEBITADO" in c.upper() or "ACREDITADO" in c.upper()]
        if candidatos:
            col_valor = candidatos[0]
            logger.warning(f"Columna exacta no encontrada. Usando candidata: '{col_valor}'")
        else:
            raise KeyError(f"No se encontró la columna de valor ('{col_valor}') en el archivo.")

    df_proc = df.copy()

    # Convertir a numérico y obtener valor absoluto
    df_proc["VALOR_ABSOLUTO"] = (
        pd.to_numeric(df_proc[col_valor], errors="coerce").fillna(0.0).abs()
    )

    # Filtrar registros con valor cero
    df_proc = df_proc[df_proc["VALOR_ABSOLUTO"] > 0].copy()

    # Ordenar de mayor a menor valor (Pareto)
    df_proc = df_proc.sort_values(by="VALOR_ABSOLUTO", ascending=False).reset_index(drop=True)

    # Cálculo de métricas de Pareto
    valor_total_poblacion = df_proc["VALOR_ABSOLUTO"].sum()
    df_proc["ORDEN_PARETO"] = range(1, len(df_proc) + 1)
    df_proc["PORCENTAJE_INDIVIDUAL"] = (df_proc["VALOR_ABSOLUTO"] / valor_total_poblacion) * 100.0
    df_proc["PORCENTAJE_ACUMULADO"] = (
        df_proc["VALOR_ABSOLUTO"].cumsum() / valor_total_poblacion
    ) * 100.0

    # Marcar transacciones pertenecientes al núcleo 80% (Pareto de alta materialidad)
    df_proc["ES_PARETO_80"] = (
        df_proc["PORCENTAJE_ACUMULADO"].shift(1, fill_value=0.0) < pareto_umbral
    )

    # Si se solicitó explícitamente recortar solo al 80%:
    if filtrar_solo_80:
        df_salida = df_proc[df_proc["ES_PARETO_80"]].copy()
        if top_max and len(df_salida) > top_max:
            df_salida = df_proc.head(top_max).copy()
    else:
        # Por defecto: TODO el universo contable ordenado por Pareto de mayor a menor
        df_salida = df_proc

    valor_muestra = df_salida["VALOR_ABSOLUTO"].sum()
    pct_cobertura = (valor_muestra / valor_total_poblacion) * 100.0

    estadisticas = {
        "total_registros_poblacion": len(df),
        "total_registros_base": len(df_salida),
        "total_registros_pareto_80": int(df_proc["ES_PARETO_80"].sum()),
        "valor_total_poblacion": valor_total_poblacion,
        "valor_total_base": valor_muestra,
        "porcentaje_cobertura": pct_cobertura,
    }

    return df_salida, estadisticas


def normalizar_columnas_clave(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza identificadores y columnas clave para etapas posteriores."""
    df_res = df.copy()

    # Normalizar NUMERO OP
    col_op = None
    for c in ["NUMERO OP", "NUMERO OP ", "OP", "N° OP", "NUMERO_OP"]:
        if c in df_res.columns:
            col_op = c
            break

    if col_op:
        op_numeric = pd.to_numeric(df_res[col_op], errors="coerce").astype("Int64")
        df_res["OP_LIMPIA"] = op_numeric.astype(str).replace("<NA>", None)
        # Si alguno no era numérico, rescatar valor original
        mask_faltante = df_res["OP_LIMPIA"].isna() & df_res[col_op].notna()
        df_res.loc[mask_faltante, "OP_LIMPIA"] = df_res.loc[mask_faltante, col_op].astype(str).str.strip()
        df_res.loc[df_res["OP_LIMPIA"].isin(["nan", "None", "", "NaN"]), "OP_LIMPIA"] = None
    else:
        df_res["OP_LIMPIA"] = None

    # Normalizar enlace al PDF
    col_url = None
    for c in ["URL", "LINK AL SOPORTE VALIDADO", "LINK_SOPORTE", "LINK"]:
        if c in df_res.columns:
            # Comprobar si tiene urls que empiecen por http
            tiene_http = df_res[c].astype(str).str.contains(r"^https?://", case=False).any()
            if tiene_http:
                col_url = c
                break
            elif col_url is None:
                col_url = c

    if col_url:
        df_res["LINK_SOPORTE"] = df_res[col_url].astype(str).str.strip()
        df_res.loc[df_res["LINK_SOPORTE"].isin(["nan", "None", ""]), "LINK_SOPORTE"] = None
    else:
        df_res["LINK_SOPORTE"] = None

    return df_res


def ejecutar_analisis_pareto(
    ruta_entrada: Optional[Path] = None,
    ruta_salida: Optional[Path] = None,
    umbral_pct: float = 80.0,
    top_max: Optional[int] = 1000,
    filtrar_solo_80: bool = False,
) -> Path:
    """Ejecuta el flujo completo de ordenación Pareto sobre toda la universalidad."""
    base_dir = Path(__file__).resolve().parent.parent

    if ruta_entrada is None:
        candidatos = [
            base_dir / "data" / "raw" / "(A)BAS~1.xlsx",
            base_dir / "(A)BAS~1.xlsx",
            Path("data/raw/(A)BAS~1.xlsx"),
            Path("(A)BAS~1.xlsx"),
        ]
        ruta_entrada = buscar_archivo_origen(candidatos)

    if ruta_salida is None:
        ruta_salida = base_dir / "data" / "output" / "base_auditoria_pareto.csv"

    ruta_salida.parent.mkdir(parents=True, exist_ok=True)

    # 1. Cargar datos
    df_raw = cargar_universalidad(ruta_entrada)

    # 2. Aplicar Pareto (toda la población ordenada de mayor a menor)
    df_pareto, stats = aplicar_pareto(
        df_raw,
        pareto_umbral=umbral_pct,
        top_max=top_max,
        filtrar_solo_80=filtrar_solo_80,
    )

    # 3. Normalizar columnas clave para downstream
    df_final = normalizar_columnas_clave(df_pareto)

    # 4. Guardar archivo
    df_final.to_csv(ruta_salida, index=False, encoding="utf-8-sig")

    logger.info("=" * 65)
    logger.info("RESUMEN DE POBLACIÓN UNIVERSALIDAD ORDENADA POR PARETO:")
    logger.info(f" - Registros totales en población:  {stats['total_registros_poblacion']:,}")
    logger.info(f" - Registros incluidos en base:     {stats['total_registros_base']:,}")
    logger.info(f" - Registros en núcleo Pareto 80%:  {stats['total_registros_pareto_80']:,}")
    logger.info(f" - Monto total de la base:          ${stats['valor_total_base']:,.2f} COP")
    logger.info(f" - Archivo generado en:             {ruta_salida}")
    logger.info("=" * 65)

    return ruta_salida


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Generación de base contable universal ordenada por Pareto.")
    parser.add_argument("--solo-80", action="store_true", help="Filtrar únicamente el 80%% de materialidad (Pareto clásico).")
    args = parser.parse_args()

    ejecutar_analisis_pareto(filtrar_solo_80=args.solo_80)


if __name__ == "__main__":
    main()
