"""
02_descargar_soportes.py
========================
Módulo para la descarga secuencial y autenticada de los soportes documentales
(archivos PDF) a partir de un archivo Excel de VALIDACIÓN POR LOTE (los archivos
"Validacion_Soportes_RangoXXXX-YYYY_UNIVERSALIDAD_ABAS1.xlsx" generados para cada
tramo del ranking, hoja "Resumen").

A diferencia de la versión anterior (que leía el CSV completo `base_auditoria_pareto.csv`
en orden Pareto), esta versión descarga EXACTAMENTE el lote que está definido en el
Excel de validación que le indiques con --excel, respetando su columna "#" (rank) y
"OP", y usando la columna "URL soporte" como enlace de descarga.

Soporte de Autenticación para SharePoint Online (aerobaq.sharepoint.com):
1. Autenticación directa por Usuario y Contraseña (vía protocolo WS-Trust SAML 1.1 de Microsoft Online).
2. Autenticación mediante Cookies de Sesión ('FedAuth' y 'rtFa') para cuentas corporativas con
   Doble Factor de Autenticación (MFA / 2FA / Microsoft Authenticator).
3. Autenticación básica HTTP (Basic Auth) para repositorios web protegidos estándar.
4. Respaldo local inteligente si el archivo ya fue descargado o sincronizado previamente
   (por ejemplo, si ya sincronizaste la biblioteca de SharePoint con OneDrive: apunta
   --respaldo a esa carpeta local y el script copia de ahí en vez de pedir red).

Uso típico:
    python 02_descargar_soportes.py --excel "Validacion_Soportes_Rango2301-2600_UNIVERSALIDAD_ABAS1.xlsx"

    # Solo probar con los primeros 10 pendientes del lote:
    python 02_descargar_soportes.py --excel "...xlsx" --lote 10

    # Reintentar TODO el lote aunque ya tenga "Resultado de validación" (por defecto se saltan):
    python 02_descargar_soportes.py --excel "...xlsx" --incluir-validadas

    # Usando cookie FedAuth (cuentas con MFA):
    python 02_descargar_soportes.py --excel "...xlsx" --fedauth "AAMkAGI2...."

    # Usando una carpeta ya sincronizada con OneDrive como respaldo primero (evita red):
    python 02_descargar_soportes.py --excel "...xlsx" --respaldo "C:\\Users\\tú\\OneDrive - ANI\\FTAPA"
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
from typing import Optional
from urllib.parse import unquote, urlparse

import openpyxl
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

# Encabezados esperados en la hoja "Resumen" de los archivos de validación por lote.
# Se busca esta fila automáticamente porque cada lote trae un número distinto de
# líneas de nota/aviso antes de la tabla.
COLUMNAS_ESPERADAS = {
    "rank": "#",
    "fila": "Fila UNIVERSALIDAD",
    "op": "OP",
    "tercero": "Tercero (según UNIVERSALIDAD)",
    "cuenta": "Cuenta contable",
    "fecha": "Fecha",
    "valor": "Valor (COP)",
    "moneda": "Moneda carpeta",
    "url": "URL soporte",
    "resultado": "Resultado de validación",
    "observacion": "Observación",
}


def sanitizar_nombre(valor: any, largo_max: int = 40) -> str:
    """Limpia un texto para usarlo como parte de un nombre de archivo."""
    if valor is None:
        return ""
    txt = str(valor).strip()
    txt = re.sub(r'[\\/*?:"<>|]', "_", txt)
    txt = re.sub(r"\s+", "_", txt)
    return txt[:largo_max].rstrip("_")


def sanitizar_op(op_valor: any) -> Optional[str]:
    """Limpia el número de OP para usarlo como nombre de archivo."""
    if op_valor is None:
        return None
    val_str = str(op_valor).strip()
    if val_str in ["nan", "None", "", "0", "NaN"]:
        return None
    m = re.match(r"^(\d+)\.0+$", val_str)
    if m:
        return m.group(1)
    try:
        num = float(val_str)
        if num.is_integer():
            return str(int(num))
    except ValueError:
        pass
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
                    "Para cuentas con MFA, use la opción de cookie 'SHAREPOINT_FEDAUTH' en el archivo .env "
                    "o el argumento --fedauth."
                )
            return None

        token_elem = root.find(
            ".//{http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd}BinarySecurityToken"
        )
        if token_elem is None or not token_elem.text:
            logger.error("No se encontró BinarySecurityToken en la respuesta de Microsoft STS.")
            return None

        security_token = token_elem.text

        login_url = f"{dominio_sitio}/_forms/default.aspx?wa=wsignin1.0"
        session = requests.Session()
        session.headers.update(HEADERS_HTTP)

        resp_sp = session.post(login_url, data=security_token, timeout=20, allow_redirects=False)

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
    1. Cookie FedAuth (si se provee o está en .env) — recomendado si tu cuenta tiene MFA.
    2. Usuario y Contraseña vía SAML (si están disponibles).
    3. Sesión no autenticada estándar como fallback.
    """
    session = requests.Session()
    session.headers.update(HEADERS_HTTP)

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

    user = usuario or os.getenv("SHAREPOINT_USER", "").strip()
    pwd = contrasena or os.getenv("SHAREPOINT_PASSWORD", "").strip()

    if user and pwd:
        sesion_saml = autenticar_sharepoint_saml(user, pwd, dominio_sitio)
        if sesion_saml is not None:
            return sesion_saml
        else:
            logger.warning("[AVISO AUTH] La autenticación automática con usuario/contraseña falló.")
            logger.warning(
                "Sugerencia: si su cuenta tiene autenticación de dos factores (MFA), "
                "inicie sesión en https://aerobaq.sharepoint.com en su navegador, presione F12 "
                "(pestaña Application/Almacenamiento > Cookies), copie el valor de la cookie 'FedAuth' "
                "y páselo con --fedauth o en SHAREPOINT_FEDAUTH en su archivo .env."
            )

    logger.info("[INFO] Continuando con sesión HTTP estándar (sin credenciales autenticadas).")
    return session


def buscar_respaldo_local(nombre_archivo: str, directorios_base: list[Path]) -> Optional[Path]:
    """Busca si el archivo PDF ya existe en directorios locales de respaldo
    (por ejemplo, una biblioteca de SharePoint ya sincronizada con OneDrive)."""
    for base in directorios_base:
        if not base.exists():
            continue
        candidatos = list(base.rglob(nombre_archivo))
        if candidatos:
            return candidatos[0]
    return None


def _encontrar_fila_encabezado(ws) -> Optional[int]:
    """Ubica la fila de encabezado en la hoja 'Resumen' buscando la fila que
    contiene '#' y 'URL soporte' entre sus celdas. Necesario porque cada lote
    trae un número distinto de líneas de nota/aviso antes de la tabla."""
    for r in range(1, min(ws.max_row, 30) + 1):
        valores = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if COLUMNAS_ESPERADAS["rank"] in valores and COLUMNAS_ESPERADAS["url"] in valores:
            return r
    return None


def leer_lote_excel(ruta_excel: Path, hoja: str = "Resumen") -> list[dict]:
    """Lee el Excel de validación por lote y devuelve una lista de diccionarios,
    uno por fila de transacción a descargar."""
    wb = openpyxl.load_workbook(ruta_excel, data_only=True)
    if hoja not in wb.sheetnames:
        raise ValueError(f"La hoja '{hoja}' no existe en '{ruta_excel.name}'. Hojas disponibles: {wb.sheetnames}")
    ws = wb[hoja]

    fila_hdr = _encontrar_fila_encabezado(ws)
    if fila_hdr is None:
        raise ValueError(
            f"No se encontró la fila de encabezado (columnas '#' y 'URL soporte') en "
            f"'{ruta_excel.name}' / hoja '{hoja}'. Verifique que sea un archivo de validación por lote."
        )

    # Mapear nombre de columna -> índice de columna en esa fila de encabezado
    col_idx = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(row=fila_hdr, column=c).value
        for clave, etiqueta in COLUMNAS_ESPERADAS.items():
            if val == etiqueta:
                col_idx[clave] = c

    faltantes = set(["rank", "op", "url"]) - set(col_idx.keys())
    if faltantes:
        raise ValueError(f"Faltan columnas obligatorias en el encabezado: {faltantes}")

    filas = []
    for r in range(fila_hdr + 1, ws.max_row + 1):
        op_val = ws.cell(row=r, column=col_idx["op"]).value
        url_val = ws.cell(row=r, column=col_idx.get("url")).value
        if op_val is None or url_val is None:
            continue  # fila vacía (p.ej. separador entre secciones)
        url_val = str(url_val).strip()
        if not url_val.startswith(("http://", "https://")):
            continue

        fila_dict = {
            "rank": ws.cell(row=r, column=col_idx["rank"]).value,
            "fila_universalidad": ws.cell(row=r, column=col_idx.get("fila")).value if "fila" in col_idx else None,
            "op": op_val,
            "tercero": ws.cell(row=r, column=col_idx.get("tercero")).value if "tercero" in col_idx else None,
            "cuenta": ws.cell(row=r, column=col_idx.get("cuenta")).value if "cuenta" in col_idx else None,
            "fecha": ws.cell(row=r, column=col_idx.get("fecha")).value if "fecha" in col_idx else None,
            "valor": ws.cell(row=r, column=col_idx.get("valor")).value if "valor" in col_idx else None,
            "moneda": ws.cell(row=r, column=col_idx.get("moneda")).value if "moneda" in col_idx else None,
            "url": url_val,
            "resultado": ws.cell(row=r, column=col_idx.get("resultado")).value if "resultado" in col_idx else None,
        }
        filas.append(fila_dict)

    return filas


def descargar_soportes_desde_excel(
    ruta_excel: Path,
    hoja: str = "Resumen",
    directorio_destino: Optional[Path] = None,
    pausa_segundos: float = 1.0,
    limite: Optional[int] = None,
    incluir_validadas: bool = False,
    usuario: Optional[str] = None,
    contrasena: Optional[str] = None,
    fedauth: Optional[str] = None,
    rtfa: Optional[str] = None,
    respaldo_extra: Optional[list[Path]] = None,
) -> dict:
    """
    Descarga los PDFs listados en un Excel de validación por lote (hoja 'Resumen').

    Por defecto, omite las filas que ya tengan algo escrito en 'Resultado de
    validación' (para no repetir descargas de un lote parcialmente trabajado);
    use incluir_validadas=True para forzar la descarga de todo el lote de todas
    formas.
    """
    ruta_excel = Path(ruta_excel)
    if not ruta_excel.exists():
        raise FileNotFoundError(f"No se encontró el archivo Excel '{ruta_excel}'.")

    if directorio_destino is None:
        directorio_destino = Path.cwd() / "data" / "soportes" / ruta_excel.stem
    directorio_destino.mkdir(parents=True, exist_ok=True)

    rutas_respaldo = list(respaldo_extra or [])
    # Rutas de respaldo típicas del proyecto (se ignoran silenciosamente si no existen)
    rutas_respaldo += [
        Path.cwd().parent / "CUARTO DE DATOS GAC",
        Path.cwd().parent / "Auditoria",
        Path.cwd() / "Soportes",
    ]

    logger.info(f"Leyendo lote de validación desde '{ruta_excel.name}' (hoja '{hoja}')...")
    filas = leer_lote_excel(ruta_excel, hoja=hoja)
    logger.info(f"Total de filas con URL válida en el lote: {len(filas):,}")

    if not incluir_validadas:
        antes = len(filas)
        filas = [f for f in filas if not (f["resultado"] and str(f["resultado"]).strip())]
        omitidas = antes - len(filas)
        if omitidas:
            logger.info(
                f"[OMITIDAS] {omitidas:,} filas ya tienen 'Resultado de validación' y se omiten "
                f"(use --incluir-validadas para forzar su descarga de todas formas)."
            )

    # Clasificar entre ya descargadas en el destino y pendientes
    pendientes = []
    ya_existian = 0
    for f in filas:
        op_id = sanitizar_op(f["op"]) or "SINOP"
        rank = f["rank"] if f["rank"] is not None else 0
        tercero_slug = sanitizar_nombre(f["tercero"])
        nombre_archivo = f"{int(rank):05d}_OP{op_id}_{tercero_slug}.pdf".replace("__", "_")
        f["_nombre_archivo"] = nombre_archivo
        ruta_destino = directorio_destino / nombre_archivo
        if ruta_destino.exists() and ruta_destino.stat().st_size > 1024:
            ya_existian += 1
        else:
            pendientes.append(f)

    logger.info("=" * 65)
    logger.info(f"LOTE: {ruta_excel.name}")
    logger.info(f" - Ya descargados en '{directorio_destino}': {ya_existian:,}")
    logger.info(f" - Pendientes de descarga:                  {len(pendientes):,}")
    logger.info("=" * 65)

    if not pendientes:
        logger.info("[TODO AL DÍA] Todos los soportes de este lote ya están en disco.")
        return {"total_lote": len(filas), "ya_existian": ya_existian, "descargados": 0,
                "recuperados_local": 0, "fallidos": 0, "directorio": str(directorio_destino)}

    if limite is not None and limite > 0:
        lote_a_procesar = pendientes[:limite]
        logger.info(f"[LOTE PARCIAL] Procesando {len(lote_a_procesar)} de {len(pendientes)} filas pendientes.")
    else:
        lote_a_procesar = pendientes
        logger.info(f"[LOTE COMPLETO] Procesando las {len(lote_a_procesar)} filas pendientes.")

    primer_enlace = lote_a_procesar[0]["url"]
    parsed_url = urlparse(primer_enlace)
    dominio_sitio = f"{parsed_url.scheme}://{parsed_url.netloc}"

    session = crear_sesion_autenticada(
        usuario=usuario, contrasena=contrasena,
        fedauth_cookie=fedauth, rtfa_cookie=rtfa,
        dominio_sitio=dominio_sitio,
    )

    stats = {"total_lote": len(filas), "ya_existian": ya_existian, "descargados": 0,
              "recuperados_local": 0, "fallidos": 0, "directorio": str(directorio_destino)}

    manifiesto = []

    for pos, f in enumerate(lote_a_procesar, 1):
        rank = f["rank"]
        op_id = sanitizar_op(f["op"]) or "SINOP"
        url = f["url"]
        nombre_archivo = f["_nombre_archivo"]
        ruta_archivo_destino = directorio_destino / nombre_archivo
        monto = f.get("valor")
        monto_str = f"${abs(monto):,.0f}" if isinstance(monto, (int, float)) else ""

        logger.info(f"[{pos}/{len(lote_a_procesar)}] (rank #{rank} {monto_str}) Descargando OP {op_id}...")

        descarga_exitosa = False
        estado = "fallido"

        try:
            resp = session.get(url, timeout=35, allow_redirects=True)
            if resp.status_code == 200:
                contenido = resp.content
                if contenido.startswith(b"%PDF"):
                    with open(ruta_archivo_destino, "wb") as fh:
                        fh.write(contenido)
                    logger.info(f"[OK] rank #{rank} / OP {op_id}: guardado ({len(contenido):,} bytes).")
                    stats["descargados"] += 1
                    descarga_exitosa = True
                    estado = "descargado_web"
                else:
                    logger.warning(
                        f"[AVISO] rank #{rank} / OP {op_id}: código 200 pero el contenido es HTML/Login "
                        f"(tipo: {resp.headers.get('Content-Type')}). Se requiere autenticación activa."
                    )
            elif resp.status_code in (401, 403):
                logger.warning(f"[ACCESO DENEGADO HTTP {resp.status_code}] rank #{rank} / OP {op_id}: autenticación requerida.")
            else:
                logger.error(f"[ERROR HTTP {resp.status_code}] rank #{rank} / OP {op_id} al descargar.")

        except requests.exceptions.RequestException as e:
            logger.error(f"[EXCEPCIÓN DE RED] rank #{rank} / OP {op_id}: {e}")

        if not descarga_exitosa:
            nombre_origen = Path(unquote(urlparse(url).path)).name
            if nombre_origen:
                archivo_local = buscar_respaldo_local(nombre_origen, rutas_respaldo)
                if archivo_local and archivo_local.exists():
                    try:
                        import shutil
                        shutil.copy2(archivo_local, ruta_archivo_destino)
                        logger.info(f"[RECUPERADO LOCAL] rank #{rank} / OP {op_id}: copiado desde '{archivo_local.name}'.")
                        stats["recuperados_local"] += 1
                        descarga_exitosa = True
                        estado = "recuperado_local"
                    except Exception as e_copy:
                        logger.error(f"Error copiando respaldo local para OP {op_id}: {e_copy}")

        if not descarga_exitosa:
            stats["fallidos"] += 1

        manifiesto.append({
            "rank": rank, "fila_universalidad": f.get("fila_universalidad"), "op": f.get("op"),
            "tercero": f.get("tercero"), "fecha": f.get("fecha"), "valor": f.get("valor"),
            "moneda": f.get("moneda"), "url": url, "archivo_local": nombre_archivo if descarga_exitosa else "",
            "estado": estado,
        })

        time.sleep(pausa_segundos)

    # Escribir manifiesto CSV para trazabilidad (rank/OP -> archivo local)
    ruta_manifiesto = directorio_destino / "_manifiesto_descarga.csv"
    modo_escritura = "a" if ruta_manifiesto.exists() else "w"
    import csv
    with open(ruta_manifiesto, modo_escritura, newline="", encoding="utf-8-sig") as fh:
        campos = ["rank", "fila_universalidad", "op", "tercero", "fecha", "valor", "moneda", "url", "archivo_local", "estado"]
        writer = csv.DictWriter(fh, fieldnames=campos)
        if modo_escritura == "w":
            writer.writeheader()
        writer.writerows(manifiesto)

    logger.info("=" * 65)
    logger.info("RESUMEN DE DESCARGA DE SOPORTES:")
    logger.info(f" - Total filas en el lote (tras filtros):  {len(filas):,}")
    logger.info(f" - Descargadas vía Web:                    {stats['descargados']:,}")
    logger.info(f" - Ya existentes en disco:                 {stats['ya_existian']:,}")
    logger.info(f" - Recuperadas de respaldo local:           {stats['recuperados_local']:,}")
    logger.info(f" - Pendientes / Fallidas:                  {stats['fallidos']:,}")
    logger.info(f" - Directorio destino:                     {directorio_destino}")
    logger.info(f" - Manifiesto:                              {ruta_manifiesto}")
    logger.info("=" * 65)

    return stats


def main():
    parser = argparse.ArgumentParser(
        description="Descarga autenticada de soportes PDF a partir de un Excel de validación por lote."
    )
    parser.add_argument("--excel", type=str, required=True,
                         help="Ruta al archivo Excel de validación por lote (hoja 'Resumen').")
    parser.add_argument("--hoja", type=str, default="Resumen", help="Nombre de la hoja con la tabla (default: 'Resumen').")
    parser.add_argument("--destino", type=str, default=None,
                         help="Carpeta destino de los PDFs (default: ./data/soportes/<nombre_del_excel>/).")
    parser.add_argument("--usuario", type=str, default=None, help="Usuario o correo de SharePoint / Microsoft 365.")
    parser.add_argument("--password", type=str, default=None, help="Contraseña de SharePoint.")
    parser.add_argument("--fedauth", type=str, default=None, help="Cookie de sesión FedAuth (para cuentas con MFA).")
    parser.add_argument("--rtfa", type=str, default=None, help="Cookie de sesión rtFa (opcional, junto con --fedauth).")
    parser.add_argument("--lote", "--limite", dest="lote", type=int, default=None,
                         help="Cantidad máxima de filas PENDIENTES a descargar en esta corrida (default: todas).")
    parser.add_argument("--pausa", type=float, default=1.0, help="Pausa en segundos entre descargas (default 1.0).")
    parser.add_argument("--incluir-validadas", action="store_true",
                         help="Descarga también las filas que ya tienen 'Resultado de validación' (por defecto se omiten).")
    parser.add_argument("--respaldo", action="append", default=None,
                         help="Carpeta local adicional donde buscar el PDF antes de ir a la red "
                              "(por ejemplo, tu biblioteca de SharePoint ya sincronizada con OneDrive). "
                              "Puede repetirse varias veces.")
    args = parser.parse_args()

    respaldo_extra = [Path(p) for p in args.respaldo] if args.respaldo else None

    descargar_soportes_desde_excel(
        ruta_excel=Path(args.excel),
        hoja=args.hoja,
        directorio_destino=Path(args.destino) if args.destino else None,
        pausa_segundos=args.pausa,
        limite=args.lote,
        incluir_validadas=args.incluir_validadas,
        usuario=args.usuario,
        contrasena=args.password,
        fedauth=args.fedauth,
        rtfa=args.rtfa,
        respaldo_extra=respaldo_extra,
    )


if __name__ == "__main__":
    main()
