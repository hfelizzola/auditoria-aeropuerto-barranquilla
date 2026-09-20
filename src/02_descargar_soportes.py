"""
02_descargar_soportes.py
========================
Módulo para la descarga secuencial y autenticada de los soportes documentales
(archivos PDF) correspondientes a las órdenes de pago seleccionadas en la muestra.

Soporte de Autenticación para SharePoint Online (aerobaq.sharepoint.com):
1. Autenticación directa por Usuario y Contraseña (vía protocolo WS-Trust SAML 1.1 de Microsoft Online).
2. Autenticación mediante Cookies de Sesión ('FedAuth' y 'rtFa') para cuentas corporativas con
   Doble Factor de Autenticación (MFA / 2FA / Microsoft Authenticator).
3. Autenticación básica HTTP (Basic Auth) para repositorios web protegidos estándar.
4. Respaldo local inteligente si el archivo ya fue descargado o sincronizado previamente.
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
import time
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional, Tuple
from urllib.parse import unquote, urlparse

import pandas as pd
import requests

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)

# Cargar variables de entorno desde .env con fallback nativo
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    env_file = Path(__file__).resolve().parent.parent / ".env"
    if env_file.exists():
        with open(env_file, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip())

HEADERS_HTTP = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/pdf,*/*",
}


def sanitizar_nombre_op(op_valor: any) -> Optional[str]:
    """Limpia el identificador de orden de pago para usarlo como nombre de archivo."""
    if pd.isna(op_valor) or op_valor is None:
        return None
    val_str = str(op_valor).strip()
    if val_str in ["nan", "None", "", "0", "NaN"]:
        return None

    # Si viene como flotante con .0 (ej. '1953.0')
    m = re.match(r"^(\d+)\.0+$", val_str)
    if m:
        return m.group(1)

    try:
        num = float(val_str)
        if num.is_integer():
            return str(int(num))
    except ValueError:
        pass

    # Limpiar caracteres prohibidos en nombres de archivo
    return re.sub(r'[\\/*?:"<>|]', "_", val_str).strip()


def autenticar_sharepoint_saml(usuario: str, contrasena: str, dominio_sitio: str) -> Optional[requests.Session]:
    """
    Ejecuta el flujo WS-Trust SAML 1.1 contra Microsoft Online Security Token Service (STS)
    para adquirir las cookies de sesión 'FedAuth' y 'rtFa' de SharePoint Online.
    """
    logger.info(f"Iniciando autenticación en Microsoft Online para el usuario '{usuario}'...")

    soap_request = f"""<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"
      xmlns:a="http://www.w3.org/2005/08/addressing"
      xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
    <s:Header>
      <a:Action s:mustUnderstand="1">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action>
      <a:ReplyTo>
        <a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address>
      </a:ReplyTo>
      <a:To s:mustUnderstand="1">https://login.microsoftonline.com/extSTS.srf</a:To>
      <o:Security s:mustUnderstand="1"
         xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
        <o:UsernameToken>
          <o:Username>{usuario}</o:Username>
          <o:Password>{contrasena}</o:Password>
        </o:UsernameToken>
      </o:Security>
    </s:Header>
    <s:Body>
      <t:RequestSecurityToken xmlns:t="http://schemas.xmlsoap.org/ws/2005/02/trust">
        <wsp:AppliesTo xmlns:wsp="http://schemas.xmlsoap.org/ws/2004/09/policy">
          <a:EndpointReference>
            <a:Address>{dominio_sitio}</a:Address>
          </a:EndpointReference>
        </wsp:AppliesTo>
        <t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType>
        <t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType>
        <t:TokenType>urn:oasis:names:tc:SAML:1.0:assertion</t:TokenType>
      </t:RequestSecurityToken>
    </s:Body>
    </s:Envelope>"""

    try:
        resp_sts = requests.post(
            "https://login.microsoftonline.com/extSTS.srf",
            data=soap_request.encode("utf-8"),
            headers={"Content-Type": "application/soap+xml; charset=utf-8"},
            timeout=20,
        )

        root = ET.fromstring(resp_sts.text)

        # Comprobar si hubo fallo SOAP
        fault = root.find(".//{http://www.w3.org/2003/05/soap-envelope}Fault")
        if fault is not None:
            reason = fault.find(".//{http://www.w3.org/2003/05/soap-envelope}Text")
            texto_razon = reason.text if reason is not None else "Error desconocido"
            logger.error(f"[ERROR LOGIN MICROSOFT] {texto_razon}")
            if "Authentication Failure" in texto_razon:
                logger.error("Verifique que el usuario y la contraseña en .env sean correctos.")
            elif "AADSTS50076" in resp_sts.text or "multi-factor" in resp_sts.text.lower():
                logger.warning(
                    "[MFA DETECTADO] La cuenta exige Doble Factor (MFA/Microsoft Authenticator). "
                    "Para cuentas con MFA, use la opción de cookie 'SHAREPOINT_FEDAUTH' en el archivo .env."
                )
            return None

        # Extraer BinarySecurityToken
        token_elem = root.find(
            ".//{http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd}BinarySecurityToken"
        )
        if token_elem is None or not token_elem.text:
            logger.error("No se encontró BinarySecurityToken en la respuesta de Microsoft STS.")
            return None

        security_token = token_elem.text

        # Enviar token al endpoint de inicio de sesión de SharePoint
        login_url = f"{dominio_sitio}/_forms/default.aspx?wa=wsignin1.0"
        session = requests.Session()
        session.headers.update(HEADERS_HTTP)

        resp_sp = session.post(login_url, data=security_token, timeout=20, allow_redirects=False)

        # Verificar que se hayan establecido las cookies de autenticación
        cookies = session.cookies.get_dict()
        if "FedAuth" in cookies:
            logger.info("[AUTH OK] Autenticación SAML exitosa con SharePoint Online (Cookie FedAuth adquirida).")
            return session
        else:
            logger.warning(
                f"SharePoint respondió (código {resp_sp.status_code}) pero no emitió cookie 'FedAuth'. "
                f"Headers: {dict(resp_sp.headers)}"
            )
            return None

    except Exception as e:
        logger.error(f"Excepción durante la autenticación de SharePoint: {e}")
        return None


def crear_sesion_autenticada(
    usuario: Optional[str] = None,
    contrasena: Optional[str] = None,
    fedauth_cookie: Optional[str] = None,
    rtfa_cookie: Optional[str] = None,
    dominio_sitio: str = "https://aerobaq.sharepoint.com",
) -> requests.Session:
    """
    Crea y devuelve una sesión de requests autenticada utilizando:
    1. Cookie FedAuth (si se provee o está en .env).
    2. Usuario y Contraseña vía SAML (si están disponibles).
    3. Sesión no autenticada estándar como fallback.
    """
    session = requests.Session()
    session.headers.update(HEADERS_HTTP)

    # 1. Prioridad: Cookie FedAuth directa (óptima para cuentas con MFA)
    cookie_fedauth = fedauth_cookie or os.getenv("SHAREPOINT_FEDAUTH", "").strip()
    cookie_rtfa = rtfa_cookie or os.getenv("SHAREPOINT_RTFA", "").strip()

    if cookie_fedauth:
        logger.info("[AUTH] Configurando sesión mediante Cookie FedAuth provista...")
        parsed = urlparse(dominio_sitio)
        host = parsed.netloc

        session.cookies.set("FedAuth", cookie_fedauth, domain=host, path="/")
        if cookie_rtfa:
            session.cookies.set("rtFa", cookie_rtfa, domain=host, path="/")
        logger.info(f"[AUTH OK] Cookies de sesión SharePoint inyectadas para '{host}'.")
        return session

    # 2. Autenticación con Usuario y Contraseña
    user = usuario or os.getenv("SHAREPOINT_USER", "").strip()
    pwd = contrasena or os.getenv("SHAREPOINT_PASSWORD", "").strip()

    if user and pwd:
        sesion_saml = autenticar_sharepoint_saml(user, pwd, dominio_sitio)
        if sesion_saml is not None:
            return sesion_saml
        else:
            logger.warning("[AVISO AUTH] La autenticación automática con usuario/contraseña falló.")
            logger.warning(
                "Sugerencia: Si su cuenta tiene autenticación de dos factores (MFA), "
                "inicie sesión en https://aerobaq.sharepoint.com en su navegador, presione F12, "
                "copie la cookie 'FedAuth' y péguela en SHAREPOINT_FEDAUTH en su archivo .env."
            )

    # 3. Fallback: sesión sin autenticación previa
    logger.info("[INFO] Continuando con sesión HTTP estándar (sin credenciales autenticadas).")
    return session


def buscar_respaldo_local(nombre_archivo: str, directorios_base: list[Path]) -> Optional[Path]:
    """Busca si el archivo PDF existe en directorios locales de respaldo."""
    for base in directorios_base:
        if not base.exists():
            continue
        candidatos = list(base.rglob(nombre_archivo))
        if candidatos:
            return candidatos[0]
    return None


def descargar_soportes(
    ruta_csv: Optional[Path] = None,
    directorio_destino: Optional[Path] = None,
    pausa_segundos: float = 1.0,
    limite: Optional[int] = None,
    usuario: Optional[str] = None,
    contrasena: Optional[str] = None,
    fedauth: Optional[str] = None,
) -> dict:
    """
    Descarga los PDFs de las órdenes de pago de la muestra Pareto con soporte de login.
    """
    base_dir = Path(__file__).resolve().parent.parent

    if ruta_csv is None:
        ruta_csv = base_dir / "data" / "output" / "base_auditoria_pareto.csv"

    if directorio_destino is None:
        directorio_destino = base_dir / "data" / "datalake_pdfs"

    directorio_destino.mkdir(parents=True, exist_ok=True)

    rutas_respaldo = [
        base_dir.parent / "CUARTO DE DATOS GAC",
        base_dir.parent / "Auditoria",
        base_dir / "Soportes",
    ]

    if not ruta_csv.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo de entrada '{ruta_csv}'. "
            f"Ejecute primero '01_pareto_universalidad.py'."
        )

    logger.info(f"Leyendo universo contable desde '{ruta_csv}'...")
    df = pd.read_csv(ruta_csv, dtype={"OP_LIMPIA": str, "NUMERO OP": str})

    # Identificar columnas
    col_enlace = "LINK_SOPORTE" if "LINK_SOPORTE" in df.columns else "URL"
    col_op = "OP_LIMPIA" if "OP_LIMPIA" in df.columns else "NUMERO OP"

    logger.info(f"Usando columna de enlace: '{col_enlace}' y columna de OP: '{col_op}'.")

    # Limpiar y deduplicar en estricto orden Pareto (conservando la primera aparición = mayor valor)
    df_validos = df.dropna(subset=[col_enlace]).copy()
    df_validos = df_validos[df_validos[col_enlace].astype(str).str.strip().str.startswith(("http://", "https://"))]
    df_validos["OP_ID"] = df_validos[col_op].apply(sanitizar_nombre_op)
    df_validos = df_validos.dropna(subset=["OP_ID"]).drop_duplicates(subset=["OP_ID"], keep="first")

    total_ops_unicas = len(df_validos)
    logger.info(f"Total de Órdenes de Pago únicas con enlace identificadas: {total_ops_unicas:,}")

    # Clasificar entre las ya descargadas en datalake y las pendientes
    ops_pendientes = []
    ops_existentes = 0

    for _, fila in df_validos.iterrows():
        op_id = fila["OP_ID"]
        archivo_esperado = directorio_destino / f"OP_{op_id}.pdf"
        if archivo_esperado.exists() and archivo_esperado.stat().st_size > 1024:
            ops_existentes += 1
        else:
            ops_pendientes.append(fila)

    logger.info("=" * 65)
    logger.info("ESTADO ACTUAL DEL DATALAKE (ORDEN PARETO):")
    logger.info(f" - Soportes ya descargados en disco: {ops_existentes:,}")
    logger.info(f" - Soportes pendientes de descarga:   {len(ops_pendientes):,}")
    logger.info("=" * 65)

    if not ops_pendientes:
        logger.info("[TODO AL DÍA] ¡Todos los soportes documentales ya están descargados en el datalake local!")
        return {
            "total": total_ops_unicas,
            "ya_existian": ops_existentes,
            "descargados": 0,
            "recuperados_local": 0,
            "fallidos": 0,
        }

    # Aplicar tamaño de lote definido por el usuario
    if limite is not None and limite > 0:
        lote_a_procesar = ops_pendientes[:limite]
        logger.info(f"[LOTE SELECCIONADO] Procesando las siguientes {len(lote_a_procesar)} OPs pendientes de mayor valor.")
    else:
        lote_a_procesar = ops_pendientes
        logger.info(f"[LOTE COMPLETO] Procesando la totalidad de las {len(lote_a_procesar)} OPs pendientes.")

    # Obtener el dominio del sitio SharePoint del primer enlace disponible
    primer_enlace = lote_a_procesar[0][col_enlace]
    parsed_url = urlparse(primer_enlace)
    dominio_sitio = f"{parsed_url.scheme}://{parsed_url.netloc}"

    # Inicializar sesión autenticada
    session = crear_sesion_autenticada(
        usuario=usuario,
        contrasena=contrasena,
        fedauth_cookie=fedauth,
        dominio_sitio=dominio_sitio,
    )

    stats = {
        "total_lote": len(lote_a_procesar),
        "descargados": 0,
        "ya_existian": ops_existentes,
        "recuperados_local": 0,
        "fallidos": 0,
    }

    for pos, fila in enumerate(lote_a_procesar, 1):
        op_id = fila["OP_ID"]
        url = str(fila[col_enlace]).strip()
        monto_str = f"${fila.get('VALOR_ABSOLUTO', 0):,.0f}" if 'VALOR_ABSOLUTO' in fila else ""
        orden_pareto = fila.get("ORDEN_PARETO", pos)

        nombre_archivo_destino = f"OP_{op_id}.pdf"
        ruta_archivo_destino = directorio_destino / nombre_archivo_destino

        logger.info(
            f"[{pos}/{len(lote_a_procesar)}] (Pareto #{orden_pareto} {monto_str}) "
            f"Descargando OP {op_id}..."
        )
        descarga_exitosa = False

        try:
            resp = session.get(url, timeout=35, allow_redirects=True)
            if resp.status_code == 200:
                contenido = resp.content
                # Verificar encabezado de archivo PDF
                if contenido.startswith(b"%PDF"):
                    with open(ruta_archivo_destino, "wb") as f:
                        f.write(contenido)
                    logger.info(f"[OK] OP {op_id}: Guardado exitosamente ({len(contenido):,} bytes).")
                    stats["descargados"] += 1
                    descarga_exitosa = True
                else:
                    logger.warning(
                        f"[AVISO] OP {op_id}: Código 200 recibido pero el contenido es HTML/Login "
                        f"(tipo: {resp.headers.get('Content-Type')}). "
                        f"Se requiere autenticación activa (configure usuario/contraseña o FedAuth en .env)."
                    )
            elif resp.status_code in [401, 403]:
                logger.warning(f"[ACCESO DENEGADO HTTP {resp.status_code}] OP {op_id}: Autenticación requerida.")
            else:
                logger.error(f"[ERROR HTTP {resp.status_code}] OP {op_id} al descargar.")

        except requests.exceptions.RequestException as e:
            logger.error(f"[EXCEPCIÓN DE RED] OP {op_id}: {e}")

        # 3. Respaldo local si la descarga web falló
        if not descarga_exitosa:
            nombre_origen = Path(unquote(urlparse(url).path)).name
            if nombre_origen:
                archivo_local = buscar_respaldo_local(nombre_origen, rutas_respaldo)
                if archivo_local and archivo_local.exists():
                    try:
                        import shutil
                        shutil.copy2(archivo_local, ruta_archivo_destino)
                        logger.info(
                            f"[RECUPERADO LOCAL] OP {op_id}: Copiado desde '{archivo_local.name}' "
                            f"({archivo_local.stat().st_size:,} bytes)."
                        )
                        stats["recuperados_local"] += 1
                        descarga_exitosa = True
                    except Exception as e_copy:
                        logger.error(f"Error copiando respaldo local para OP {op_id}: {e_copy}")

        if not descarga_exitosa:
            stats["fallidos"] += 1

        time.sleep(pausa_segundos)

    logger.info("=" * 65)
    logger.info("RESUMEN DE DESCARGA DE SOPORTES:")
    logger.info(f" - Total OPs en este lote:    {len(lote_a_procesar):,}")
    logger.info(f" - Descargadas vía Web:       {stats['descargados']:,}")
    logger.info(f" - Ya existentes en disco:    {stats['ya_existian']:,}")
    logger.info(f" - Recuperadas de respaldo:   {stats['recuperados_local']:,}")
    logger.info(f" - Pendientes / Fallidas:     {stats['fallidos']:,}")
    logger.info(f" - Directorio destino:        {directorio_destino}")
    logger.info("=" * 65)

    return stats


def main():
    parser = argparse.ArgumentParser(description="Descarga autenticada de soportes PDF desde SharePoint por lotes en orden Pareto.")
    parser.add_argument("--usuario", type=str, default=None, help="Usuario o correo de SharePoint / Microsoft 365.")
    parser.add_argument("--password", type=str, default=None, help="Contraseña de SharePoint.")
    parser.add_argument("--fedauth", type=str, default=None, help="Cookie de sesión FedAuth (para cuentas con MFA).")
    parser.add_argument("--lote", "--limite", dest="lote", type=int, default=None, help="Cantidad de OPs pendientes a descargar en este lote (en orden Pareto).")
    parser.add_argument("--pausa", type=float, default=1.0, help="Pausa en segundos entre descargas (default 1.0).")
    args = parser.parse_args()

    descargar_soportes(
        limite=args.lote,
        pausa_segundos=args.pausa,
        usuario=args.usuario,
        contrasena=args.password,
        fedauth=args.fedauth,
    )


if __name__ == "__main__":
    main()
