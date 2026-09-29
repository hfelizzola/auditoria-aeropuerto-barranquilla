"""
02_descargar_soportes_excel.py
==============================
Descarga secuencial y autenticada de los soportes documentales (PDF) de un lote de
validación: el Excel "Validacion_Soportes_RangoXXXX-YYYY_UNIVERSALIDAD_ABAS1.xlsx"
(hoja "Resumen"), respetando su columna "#" (rank) y "OP", y usando "URL soporte"
como enlace de descarga.

Se usa de dos formas:
- como script CLI independiente (--excel ...), o
- como módulo desde el orquestador (00_orquestar_auditoria.py), vía `descargar_filas()`.

Orden de intentos por fila (idempotente):
1. Si el PDF ya existe en disco (> 1 KB) en la carpeta del lote o en otra carpeta de
   data/soportes/ (lotes anteriores), no se vuelve a descargar.
2. GET autenticado a SharePoint (cookie FedAuth/rtFa o usuario/contraseña vía SAML),
   con reintentos y backoff exponencial ante fallas de red / 429 / 5xx.
3. Microsoft Graph API (`/v1.0/shares/{share_id}/driveItem`, scope Files.Read.All) cuando
   el GET directo no devuelve un PDF — típico de los enlaces de compartición `:b:`.
   Si el enlace apunta a una carpeta, se descargan todos sus PDF como partes del mismo rank.
   Token: GRAPH_ACCESS_TOKEN en .env, o flujo de código de dispositivo con MSAL si hay
   GRAPH_CLIENT_ID (y opcionalmente GRAPH_TENANT_ID).
4. Respaldo local por nombre de archivo (biblioteca de SharePoint sincronizada con OneDrive).
5. Si nada funciona: estado "fallido" (el rank queda SIN resultado y se lista como pendiente).

Autenticación SharePoint Online (aerobaq.sharepoint.com):
1. Usuario y contraseña (WS-Trust SAML 1.1 de Microsoft Online).
2. Cookies de sesión 'FedAuth' y 'rtFa' para cuentas con MFA.

Uso típico:
    python src/02_descargar_soportes_excel.py --excel "Validacion_Soportes_Rango2301-2600_UNIVERSALIDAD_ABAS1.xlsx"
    python src/02_descargar_soportes_excel.py --excel "...xlsx" --lote 10
    python src/02_descargar_soportes_excel.py --excel "...xlsx" --fedauth "AAMkAGI2...."
    python src/02_descargar_soportes_excel.py --excel "...xlsx" --respaldo "C:\\Users\\tú\\OneDrive - ANI\\FTAPA"
"""

from __future__ import annotations

import argparse
import base64
import glob
import logging
import os
import re
import shutil
import sys
import time
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import unquote, urlparse

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent))
from auditoria.config import DIR_CACHE, DIR_SOPORTES, cargar_env  # noqa: E402
from auditoria.estado import actualizar_manifiesto, ahora  # noqa: E402
from auditoria.lote_excel import LoteExcel  # noqa: E402


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)

cargar_env()

HEADERS_HTTP = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/pdf,*/*",
}


def sanitizar_nombre(valor: Any, largo_max: int = 40) -> str:
    """Limpia un texto para usarlo como parte de un nombre de archivo."""
    if valor is None:
        return ""
    txt = str(valor).strip()
    txt = re.sub(r'[\\/*?:"<>|]', "_", txt)
    txt = re.sub(r"\s+", "_", txt)
    return txt[:largo_max].rstrip("_")


def sanitizar_op(op_valor: Any) -> Optional[str]:
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


# ---------------------------------------------------------------------------
# Nombres de archivo, respaldo local y existentes
# ---------------------------------------------------------------------------

TAMANO_MINIMO_PDF = 1024


def nombre_archivo_soporte(rank: Any, op: Any, tercero: Any, parte: int = 1) -> str:
    """'03505_OP11745_B&S_INGENIERIA_SAS.pdf' (parte 2+ -> '..._parte2.pdf')."""
    op_id = sanitizar_op(op) or "SINOP"
    rank_i = int(float(rank)) if rank is not None else 0
    base = f"{rank_i:05d}_OP{op_id}_{sanitizar_nombre(tercero)}".replace("__", "_").rstrip("_")
    return f"{base}.pdf" if parte == 1 else f"{base}_parte{parte}.pdf"


def pdfs_de_rank(directorio: Path, rank: Any, op: Any) -> List[Path]:
    """PDFs válidos (> 1 KB) ya presentes para un rank/OP, ordenados (parte 1, 2, ...)."""
    op_id = sanitizar_op(op) or "SINOP"
    patron = glob.escape(f"{int(float(rank)):05d}_OP{op_id}_") + "*.pdf"
    return sorted((p for p in Path(directorio).glob(patron) if p.stat().st_size > TAMANO_MINIMO_PDF),
                  key=lambda p: (len(p.name), p.name))


def es_pdf_valido(ruta: Path) -> bool:
    try:
        with open(ruta, "rb") as fh:
            return ruta.stat().st_size > TAMANO_MINIMO_PDF and fh.read(5).startswith(b"%PDF")
    except OSError:
        return False


class IndiceRespaldo:
    """Índice nombre_de_archivo -> ruta de las carpetas de respaldo (se construye una sola vez)."""

    def __init__(self, directorios: List[Path]):
        self.directorios = [Path(d) for d in directorios if d and Path(d).exists()]
        self._indice: Optional[Dict[str, Path]] = None

    def buscar(self, nombre_archivo: str) -> Optional[Path]:
        if not self.directorios or not nombre_archivo:
            return None
        if self._indice is None:
            logger.info(f"Indexando carpetas de respaldo local: {[str(d) for d in self.directorios]}...")
            self._indice = {}
            for base in self.directorios:
                for raiz, _, archivos in os.walk(base):
                    for a in archivos:
                        if a.lower().endswith(".pdf"):
                            self._indice.setdefault(a.lower(), Path(raiz) / a)
            logger.info(f"Respaldo local indexado: {len(self._indice):,} PDF.")
        return self._indice.get(nombre_archivo.lower())


def buscar_respaldo_local(nombre_archivo: str, directorios_base: List[Path]) -> Optional[Path]:
    """Compatibilidad: búsqueda puntual en carpetas de respaldo."""
    return IndiceRespaldo(directorios_base).buscar(nombre_archivo)


# ---------------------------------------------------------------------------
# Microsoft Graph (respaldo para enlaces de compartición ':b:')
# ---------------------------------------------------------------------------

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = "https://graph.microsoft.com/Files.Read.All"


def codificar_share_id(url: str) -> str:
    """Codificación de URL de compartición para /shares/{id}: 'u!' + base64url sin relleno."""
    b64 = base64.b64encode(url.encode("utf-8")).decode("ascii")
    return "u!" + b64.rstrip("=").replace("/", "_").replace("+", "-")


class ClienteGraph:
    """Descarga vía Graph API. Token de GRAPH_ACCESS_TOKEN o MSAL (device code)."""

    def __init__(self, token: Optional[str] = None):
        self._token = token or os.getenv("GRAPH_ACCESS_TOKEN", "").strip() or None
        self._client_id = os.getenv("GRAPH_CLIENT_ID", "").strip()
        self._tenant = os.getenv("GRAPH_TENANT_ID", "").strip() or "organizations"
        self._msal_fallo = False

    @property
    def disponible(self) -> bool:
        return bool(self._token or (self._client_id and not self._msal_fallo))

    def _obtener_token(self) -> Optional[str]:
        if self._token:
            return self._token
        if not self._client_id or self._msal_fallo:
            return None
        try:
            import msal
        except ImportError:
            logger.warning("GRAPH_CLIENT_ID configurado pero falta 'msal' (pip install msal).")
            self._msal_fallo = True
            return None
        ruta_cache = DIR_CACHE / "msal_token_cache.bin"
        cache = msal.SerializableTokenCache()
        if ruta_cache.exists():
            cache.deserialize(ruta_cache.read_text(encoding="utf-8"))
        app = msal.PublicClientApplication(self._client_id, token_cache=cache,
                                           authority=f"https://login.microsoftonline.com/{self._tenant}")
        cuentas = app.get_accounts()
        resultado = app.acquire_token_silent([GRAPH_SCOPE], account=cuentas[0]) if cuentas else None
        if not resultado:
            flujo = app.initiate_device_flow(scopes=[GRAPH_SCOPE])
            if "user_code" not in flujo:
                logger.error(f"No se pudo iniciar el flujo de dispositivo de Graph: {flujo}")
                self._msal_fallo = True
                return None
            logger.warning(f"[GRAPH] {flujo['message']}")
            resultado = app.acquire_token_by_device_flow(flujo)
        if "access_token" not in resultado:
            logger.error(f"[GRAPH] No se obtuvo token: {resultado.get('error_description')}")
            self._msal_fallo = True
            return None
        if cache.has_state_changed:
            DIR_CACHE.mkdir(parents=True, exist_ok=True)
            ruta_cache.write_text(cache.serialize(), encoding="utf-8")
        self._token = resultado["access_token"]
        return self._token

    def _get(self, url: str, **kw) -> requests.Response:
        token = self._obtener_token()
        if not token:
            raise RuntimeError("Sin token de Graph")
        headers = {"Authorization": f"Bearer {token}", "Prefer": "redeemSharingLinkIfNecessary"}
        return requests.get(url, headers=headers, timeout=60, **kw)

    def descargar(self, url_compartida: str, destinos: List[Path]) -> List[Path]:
        """Descarga el archivo (o los PDF de la carpeta) apuntado por el enlace.

        `destinos` es la lista de rutas a usar en orden (parte 1, 2, ...)."""
        share_id = codificar_share_id(url_compartida)
        r = self._get(f"{GRAPH_BASE}/shares/{share_id}/driveItem")
        if r.status_code != 200:
            raise RuntimeError(f"Graph driveItem HTTP {r.status_code}: {r.text[:200]}")
        item = r.json()
        if "folder" in item:
            hijos, siguiente = [], f"{GRAPH_BASE}/shares/{share_id}/driveItem/children"
            while siguiente:
                rr = self._get(siguiente)
                rr.raise_for_status()
                datos = rr.json()
                hijos += [h for h in datos.get("value", []) if h.get("name", "").lower().endswith(".pdf")]
                siguiente = datos.get("@odata.nextLink")
            items = sorted(hijos, key=lambda h: h.get("name", ""))
        else:
            items = [item]
        guardados = []
        for item_pdf, destino in zip(items, destinos):
            url_descarga = item_pdf.get("@microsoft.graph.downloadUrl")
            resp = requests.get(url_descarga, timeout=120) if url_descarga else None
            if resp is None or resp.status_code != 200 or not resp.content.startswith(b"%PDF"):
                raise RuntimeError(f"Graph no entregó un PDF para '{item_pdf.get('name')}'")
            destino.write_bytes(resp.content)
            guardados.append(destino)
        if len(items) > len(destinos):
            logger.warning(f"La carpeta compartida tiene {len(items)} PDF; se descargaron {len(destinos)}.")
        return guardados


# ---------------------------------------------------------------------------
# Descarga de una fila y de un conjunto de filas
# ---------------------------------------------------------------------------

def _get_con_reintentos(session: requests.Session, url: str, reintentos: int, espera_base: float) -> Optional[requests.Response]:
    """GET con backoff exponencial ante fallas de red, 429 y 5xx. Devuelve la última respuesta."""
    resp = None
    for intento in range(1, reintentos + 1):
        try:
            resp = session.get(url, timeout=60, allow_redirects=True)
            if resp.status_code not in (429, 500, 502, 503, 504):
                return resp
            motivo = f"HTTP {resp.status_code}"
        except requests.exceptions.RequestException as e:
            motivo = f"{type(e).__name__}: {e}"
        if intento < reintentos:
            espera = espera_base * (2 ** (intento - 1))
            if resp is not None and resp.headers.get("Retry-After", "").isdigit():
                espera = max(espera, float(resp.headers["Retry-After"]))
            logger.warning(f"   reintento {intento}/{reintentos - 1} en {espera:.0f}s ({motivo})")
            time.sleep(espera)
    return resp


class Descargador:
    """Estado compartido de una corrida de descargas (sesión, Graph, índice de respaldo)."""

    def __init__(self, directorio_destino: Path, usuario=None, contrasena=None, fedauth=None, rtfa=None,
                 respaldo_extra: Optional[List[Path]] = None, reintentos: int = 3, espera_base: float = 2.0,
                 buscar_en_otros_lotes: bool = True):
        self.directorio = Path(directorio_destino)
        self.directorio.mkdir(parents=True, exist_ok=True)
        self._cred = dict(usuario=usuario, contrasena=contrasena, fedauth_cookie=fedauth, rtfa_cookie=rtfa)
        self._session: Optional[requests.Session] = None
        self.graph = ClienteGraph()
        base = Path(__file__).resolve().parent.parent
        self.respaldo = IndiceRespaldo(list(respaldo_extra or []) + [
            base.parent / "CUARTO DE DATOS GAC", base.parent / "Auditoria", base / "Soportes"])
        self.reintentos = reintentos
        self.espera_base = espera_base
        self.buscar_en_otros_lotes = buscar_en_otros_lotes

    def _sesion(self, url: str) -> requests.Session:
        if self._session is None:
            p = urlparse(url)
            self._session = crear_sesion_autenticada(dominio_sitio=f"{p.scheme}://{p.netloc}", **self._cred)
        return self._session

    def _en_otros_lotes(self, rank, op) -> List[Path]:
        if not self.buscar_en_otros_lotes or not DIR_SOPORTES.exists():
            return []
        for carpeta in DIR_SOPORTES.iterdir():
            if carpeta.is_dir() and carpeta.resolve() != self.directorio.resolve():
                encontrados = pdfs_de_rank(carpeta, rank, op)
                if encontrados:
                    return encontrados
        return []

    def descargar_fila(self, fila: Dict[str, Any]) -> Dict[str, Any]:
        """Descarga el soporte de una fila del lote. Devuelve {estado, archivos, detalle}."""
        rank, op, url = fila["rank"], fila["op"], fila.get("url")
        existentes = pdfs_de_rank(self.directorio, rank, op)
        if existentes:
            return {"estado": "ya_existia", "archivos": existentes, "detalle": ""}

        de_otro_lote = self._en_otros_lotes(rank, op)
        if de_otro_lote:
            copias = []
            for p in de_otro_lote:
                destino = self.directorio / p.name
                shutil.copy2(p, destino)
                copias.append(destino)
            return {"estado": "recuperado_local", "archivos": copias,
                    "detalle": f"copiado de data/soportes/{de_otro_lote[0].parent.name}"}

        if not url or not str(url).lower().startswith(("http://", "https://")):
            return {"estado": "fallido", "archivos": [], "detalle": "sin URL de soporte"}

        destinos = [self.directorio / nombre_archivo_soporte(rank, op, fila.get("tercero"), parte=i)
                    for i in range(1, 21)]
        detalles = []

        # 1) GET directo autenticado, con reintentos
        resp = _get_con_reintentos(self._sesion(url), url, self.reintentos, self.espera_base)
        if resp is not None and resp.status_code == 200 and resp.content.startswith(b"%PDF"):
            destinos[0].write_bytes(resp.content)
            return {"estado": "descargado_web", "archivos": [destinos[0]], "detalle": f"{len(resp.content):,} bytes"}
        if resp is None:
            detalles.append("GET sin respuesta")
        elif resp.status_code == 200:
            detalles.append(f"GET 200 pero no es PDF ({resp.headers.get('Content-Type')}): requiere sesión activa")
        else:
            detalles.append(f"GET HTTP {resp.status_code}")

        # 2) Graph API (enlaces ':b:' o GET directo que no entrega el PDF)
        if self.graph.disponible:
            for intento in range(1, self.reintentos + 1):
                try:
                    guardados = self.graph.descargar(url, destinos)
                    if guardados:
                        return {"estado": "descargado_graph", "archivos": guardados,
                                "detalle": "; ".join(detalles + [f"Graph OK ({len(guardados)} archivo(s))"])}
                    detalles.append("Graph: sin PDF")
                    break
                except Exception as e:  # noqa: BLE001
                    if intento == self.reintentos:
                        detalles.append(f"Graph: {e}")
                    else:
                        time.sleep(self.espera_base * (2 ** (intento - 1)))
        elif "/:b:/" in url:
            detalles.append("enlace ':b:' y Graph no configurado (GRAPH_ACCESS_TOKEN o GRAPH_CLIENT_ID)")

        # 3) Respaldo local (biblioteca sincronizada con OneDrive)
        archivo_local = self.respaldo.buscar(Path(unquote(urlparse(url).path)).name)
        if archivo_local is not None:
            shutil.copy2(archivo_local, destinos[0])
            return {"estado": "recuperado_local", "archivos": [destinos[0]],
                    "detalle": "; ".join(detalles + [f"copiado de '{archivo_local}'"])}

        return {"estado": "fallido", "archivos": [], "detalle": "; ".join(detalles)}


def descargar_filas(filas: List[Dict[str, Any]], directorio_destino: Path, pausa_segundos: float = 1.0,
                    descargador: Optional[Descargador] = None, **kwargs_descargador) -> Dict[Any, Dict[str, Any]]:
    """Descarga una lista de filas del lote y actualiza el manifiesto. {rank: resultado}."""
    descargador = descargador or Descargador(directorio_destino, **kwargs_descargador)
    ruta_manifiesto = Path(directorio_destino) / "_manifiesto_descarga.csv"
    resultados: Dict[Any, Dict[str, Any]] = {}
    for pos, fila in enumerate(filas, 1):
        monto = fila.get("valor")
        monto_str = f" ${abs(monto):,.0f}" if isinstance(monto, (int, float)) else ""
        res = descargador.descargar_fila(fila)
        resultados[fila["rank"]] = res
        nivel = logging.WARNING if res["estado"] == "fallido" else logging.INFO
        logger.log(nivel, f"[{pos}/{len(filas)}] rank #{fila['rank']} OP {fila['op']}{monto_str}: "
                          f"{res['estado']}" + (f" — {res['detalle']}" if res["detalle"] else ""))
        actualizar_manifiesto(ruta_manifiesto, [{
            "rank": fila["rank"], "fila_universalidad": fila.get("fila_universalidad"), "op": fila.get("op"),
            "tercero": fila.get("tercero"), "fecha": fila.get("fecha"), "valor": fila.get("valor"),
            "moneda": fila.get("moneda"), "url": fila.get("url"),
            "archivo_local": ";".join(p.name for p in res["archivos"]), "estado": res["estado"],
            "detalle": res["detalle"], "actualizado": ahora(),
        }])
        if res["estado"] in ("descargado_web", "descargado_graph") and pos < len(filas):
            time.sleep(pausa_segundos)
    return resultados


# ---------------------------------------------------------------------------
# Lectura del lote y CLI independiente
# ---------------------------------------------------------------------------

def leer_lote_excel(ruta_excel: Path, hoja: str = "Resumen") -> List[Dict[str, Any]]:
    """Filas del Excel de lote con URL (texto plano o hipervínculo embebido)."""
    if hoja != "Resumen":
        raise ValueError("Solo se soporta la hoja 'Resumen' del layout de lote.")
    return [
        {"rank": f.rank, "fila_universalidad": f.fila_universalidad, "op": f.op, "tercero": f.tercero,
         "cuenta": f.cuenta, "fecha": f.fecha, "valor": f.valor, "moneda": f.moneda, "url": f.url,
         "resultado": f.resultado, "fila_excel": f.fila_excel}
        for f in LoteExcel.abrir(Path(ruta_excel)).filas()
        if f.op is not None and f.rank is not None and f.url
    ]


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
    respaldo_extra: Optional[List[Path]] = None,
) -> dict:
    """Descarga los PDFs listados en un Excel de validación por lote (hoja 'Resumen').

    Por defecto omite las filas que ya tienen 'Resultado de validación'.
    """
    ruta_excel = Path(ruta_excel)
    if not ruta_excel.exists():
        raise FileNotFoundError(f"No se encontró el archivo Excel '{ruta_excel}'.")
    directorio_destino = Path(directorio_destino or DIR_SOPORTES / ruta_excel.stem)

    filas = leer_lote_excel(ruta_excel, hoja=hoja)
    logger.info(f"Total de filas con URL en el lote '{ruta_excel.name}': {len(filas):,}")
    if not incluir_validadas:
        filas = [f for f in filas if not (f["resultado"] and str(f["resultado"]).strip())]
    if limite:
        filas = filas[:limite]
    logger.info(f"Filas a procesar en esta corrida: {len(filas):,} -> '{directorio_destino}'")

    resultados = descargar_filas(filas, directorio_destino, pausa_segundos=pausa_segundos,
                                 usuario=usuario, contrasena=contrasena, fedauth=fedauth, rtfa=rtfa,
                                 respaldo_extra=respaldo_extra)
    stats: Dict[str, Any] = {"total_lote": len(filas), "directorio": str(directorio_destino)}
    for r in resultados.values():
        stats[r["estado"]] = stats.get(r["estado"], 0) + 1
    fallidos = [rank for rank, r in resultados.items() if r["estado"] == "fallido"]
    logger.info("=" * 65)
    logger.info(f"RESUMEN DE DESCARGA: {stats}")
    if fallidos:
        logger.warning(f"Ranks SIN PDF tras agotar reintentos ({len(fallidos)}): {fallidos}")
    logger.info("=" * 65)
    stats["ranks_fallidos"] = fallidos
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
                        help="Carpeta local adicional donde buscar el PDF si la red falla "
                             "(por ejemplo, tu biblioteca de SharePoint ya sincronizada con OneDrive). "
                             "Puede repetirse varias veces.")
    args = parser.parse_args()

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
        respaldo_extra=[Path(p) for p in args.respaldo] if args.respaldo else None,
    )


if __name__ == "__main__":
    main()
