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
    top_max: int = 1000,
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

    # Ordenar de mayor a menor valor
    df_proc = df_proc.sort_values(by="VALOR_ABSOLUTO", ascending=False).reset_index(drop=True)

    # Cálculo de métricas de Pareto
    valor_total_poblacion = df_proc["VALOR_ABSOLUTO"].sum()
    df_proc["PORCENTAJE_INDIVIDUAL"] = (df_proc["VALOR_ABSOLUTO"] / valor_total_poblacion) * 100.0
    df_proc["PORCENTAJE_ACUMULADO"] = (
        df_proc["VALOR_ABSOLUTO"].cumsum() / valor_total_poblacion
    ) * 100.0

    # Condición de corte Pareto (el registro que cruza el 80% se incluye)
    filtro_80 = df_proc["PORCENTAJE_ACUMULADO"].shift(1, fill_value=0.0) < pareto_umbral
    df_pareto = df_proc[filtro_80].copy()

    # Si la cantidad sobrepasa top_max o se requiere limitar a top 1000:
    if len(df_pareto) > top_max:
        logger.info(f"Limitando muestra a los primeros {top_max} registros.")
        df_pareto = df_proc.head(top_max).copy()
    elif len(df_pareto) < top_max and len(df_proc) >= top_max:
        # Si el 80% fue alcanzado en menos de top_max registros, se mantiene el 80%
        # pero se deja constancia
        pass

    valor_muestra = df_pareto["VALOR_ABSOLUTO"].sum()
    pct_cobertura = (valor_muestra / valor_total_poblacion) * 100.0

    estadisticas = {
        "total_registros_poblacion": len(df),
        "total_registros_muestra": len(df_pareto),
        "valor_total_poblacion": valor_total_poblacion,
        "valor_total_muestra": valor_muestra,
        "porcentaje_cobertura": pct_cobertura,
    }

    return df_pareto, estadisticas


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
    top_max: int = 1000,
) -> Path:
    """Ejecuta el flujo completo de selección Pareto y guarda el archivo CSV."""
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

    # 2. Aplicar Pareto
    df_pareto, stats = aplicar_pareto(df_raw, pareto_umbral=umbral_pct, top_max=top_max)

    # 3. Normalizar columnas clave para downstream
    df_final = normalizar_columnas_clave(df_pareto)

    # 4. Guardar archivo
    df_final.to_csv(ruta_salida, index=False, encoding="utf-8-sig")

    logger.info("=" * 60)
    logger.info("RESUMEN SELECCIÓN MUESTRAL PARETO:")
    logger.info(f" - Registros en población: {stats['total_registros_poblacion']:,}")
    logger.info(f" - Registros seleccionados: {stats['total_registros_muestra']:,}")
    logger.info(f" - Valor total auditado:   ${stats['valor_total_muestra']:,.2f} COP")
    logger.info(f" - Cobertura sobre total:  {stats['porcentaje_cobertura']:.2f}%")
    logger.info(f" - Archivo generado en:    {ruta_salida}")
    logger.info("=" * 60)

    return ruta_salida


if __name__ == "__main__":
    ejecutar_analisis_pareto()
