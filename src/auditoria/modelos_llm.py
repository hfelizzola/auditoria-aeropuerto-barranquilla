"""
modelos_llm.py
==============
Clientes de modelo (Gemini o Claude) con una única interfaz:

    cliente.generar_json(instrucciones, texto, imagenes_jpeg, esquema) -> dict

- Salida JSON restringida por esquema (Gemini: response_json_schema; Claude:
  output_config.format json_schema).
- Reintentos con backoff exponencial ante errores transitorios (429/5xx/red).
- Si la cuota se agota tras los reintentos se lanza `ErrorCuotaLLM`, para que el
  orquestador detenga el lote de forma limpia (todo lo hecho ya está en disco).

Variables de entorno:
    GEMINI_API_KEY, GEMINI_MODEL (default gemini-2.5-flash)
    ANTHROPIC_API_KEY, CLAUDE_MODEL (default claude-opus-5), CLAUDE_EFFORT (default high),
    CLAUDE_FALLBACKS (default 1: activa el respaldo de modelo del servidor ante rechazos)
"""

from __future__ import annotations

import base64
import json
import logging
import os
import random
import re
import time
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)

MODELO_GEMINI_DEFECTO = "gemini-2.5-flash"
MODELO_CLAUDE_DEFECTO = "claude-opus-5"


class ErrorLLM(RuntimeError):
    """Error no recuperable para un documento concreto (se deja pendiente y se continúa)."""


class ErrorCuotaLLM(ErrorLLM):
    """Cuota agotada o servicio caído tras todos los reintentos: detener el lote."""


def _parsear_json(texto: str) -> Dict[str, Any]:
    limpio = (texto or "").strip()
    if limpio.startswith("```"):
        limpio = re.sub(r"^```(?:json)?\s*", "", limpio)
        limpio = re.sub(r"\s*```$", "", limpio)
    try:
        datos = json.loads(limpio)
    except json.JSONDecodeError as e:
        raise ErrorLLM(f"El modelo no devolvió JSON válido: {e}; inicio: {limpio[:200]!r}") from e
    if not isinstance(datos, dict):
        raise ErrorLLM("El modelo devolvió JSON que no es un objeto.")
    return datos


def _esperar(intento: int, base: float = 4.0, tope: float = 120.0) -> float:
    return min(tope, base * (2 ** (intento - 1))) + random.uniform(0, 1)


class ClienteLLM:
    proveedor = "base"
    modelo = ""

    def generar_json(self, instrucciones: str, texto: str, imagenes_jpeg: List[bytes],
                     esquema: Dict[str, Any], max_tokens: int = 16000) -> Dict[str, Any]:
        raise NotImplementedError

    def descripcion(self) -> str:
        return f"{self.proveedor}:{self.modelo}"


class ClienteGemini(ClienteLLM):
    proveedor = "gemini"

    def __init__(self, modelo: Optional[str] = None, max_reintentos: int = 6):
        try:
            from google import genai
            from google.genai import types
        except ImportError as e:
            raise ErrorLLM("Falta la librería google-genai. Ejecute: pip install google-genai") from e
        api_key = os.getenv("GEMINI_API_KEY", "").strip()
        if not api_key or api_key.startswith("tu_clave"):
            raise ErrorLLM("GEMINI_API_KEY no está configurada en .env.")
        self._types = types
        self._client = genai.Client(api_key=api_key)
        self.modelo = modelo or os.getenv("GEMINI_MODEL", "").strip() or MODELO_GEMINI_DEFECTO
        self.max_reintentos = max_reintentos

    def generar_json(self, instrucciones, texto, imagenes_jpeg, esquema, max_tokens=16000):
        from google.genai import errors as genai_errors

        types = self._types
        partes = [types.Part.from_bytes(data=img, mime_type="image/jpeg") for img in imagenes_jpeg]
        partes.append(types.Part.from_text(text=texto))
        config = types.GenerateContentConfig(
            system_instruction=instrucciones,
            temperature=0.1,
            # En Gemini 2.5 el razonamiento consume del mismo presupuesto de salida
            max_output_tokens=max(max_tokens, 32768),
            response_mime_type="application/json",
            response_json_schema=esquema,
        )
        for intento in range(1, self.max_reintentos + 1):
            try:
                resp = self._client.models.generate_content(model=self.modelo, contents=partes, config=config)
                if not resp.text:
                    razon = getattr(resp.candidates[0], "finish_reason", None) if resp.candidates else None
                    raise ErrorLLM(f"Gemini devolvió respuesta vacía (finish_reason={razon}).")
                return _parsear_json(resp.text)
            except genai_errors.APIError as e:
                if e.code in (401, 403):
                    raise ErrorCuotaLLM(f"Gemini rechazó la credencial ({e.code}): {e}") from e
                reintentable = e.code in (408, 429, 500, 502, 503, 504)
                if not reintentable:
                    raise ErrorLLM(f"Gemini rechazó la solicitud ({e.code}): {e}") from e
                if intento == self.max_reintentos:
                    raise ErrorCuotaLLM(f"Gemini sigue fallando tras {intento} intentos ({e.code}): {e}") from e
                espera = _esperar(intento)
                logger.warning(f"[REINTENTO {intento}/{self.max_reintentos}] Gemini {e.code}; espero {espera:.0f}s.")
                time.sleep(espera)
            except ErrorLLM:
                raise
            except Exception as e:  # errores de red de httpx u OSError
                if not (isinstance(e, OSError) or type(e).__module__.startswith("httpx")):
                    raise
                if intento == self.max_reintentos:
                    raise ErrorCuotaLLM(f"Sin conexión con Gemini tras {intento} intentos: {e}") from e
                time.sleep(_esperar(intento))
        raise ErrorCuotaLLM("Gemini: reintentos agotados.")


class ClienteClaude(ClienteLLM):
    proveedor = "claude"

    def __init__(self, modelo: Optional[str] = None, max_reintentos: int = 4):
        try:
            import anthropic
        except ImportError as e:
            raise ErrorLLM("Falta la librería anthropic. Ejecute: pip install anthropic") from e
        self._anthropic = anthropic
        # El SDK ya reintenta 429/5xx/red; aquí se suman reintentos largos para cuotas.
        self._client = anthropic.Anthropic(max_retries=3)
        self.modelo = modelo or os.getenv("CLAUDE_MODEL", "").strip() or MODELO_CLAUDE_DEFECTO
        self.esfuerzo = os.getenv("CLAUDE_EFFORT", "high").strip() or "high"
        self.usar_fallbacks = os.getenv("CLAUDE_FALLBACKS", "1").strip() not in ("0", "false", "no")
        self.max_reintentos = max_reintentos

    def generar_json(self, instrucciones, texto, imagenes_jpeg, esquema, max_tokens=16000):
        anthropic = self._anthropic
        contenido: List[Dict[str, Any]] = [
            {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                         "data": base64.standard_b64encode(img).decode("ascii")}}
            for img in imagenes_jpeg
        ]
        contenido.append({"type": "text", "text": texto})
        kwargs: Dict[str, Any] = dict(
            model=self.modelo,
            max_tokens=max(max_tokens, 32000),
            # Instrucciones estables: se cachean entre documentos del lote.
            system=[{"type": "text", "text": instrucciones, "cache_control": {"type": "ephemeral"}}],
            messages=[{"role": "user", "content": contenido}],
            thinking={"type": "adaptive"},
            output_config={"effort": self.esfuerzo, "format": {"type": "json_schema", "schema": esquema}},
        )
        if self.usar_fallbacks:
            # Respaldo del lado del servidor si el modelo rechaza la solicitud (beta).
            kwargs["extra_headers"] = {"anthropic-beta": "server-side-fallback-2026-07-01"}
            kwargs["extra_body"] = {"fallbacks": "default"}

        for intento in range(1, self.max_reintentos + 1):
            try:
                # streaming: solicitudes con ~20 imágenes pueden tardar; evita timeouts HTTP
                with self._client.messages.stream(**kwargs) as stream:
                    resp = stream.get_final_message()
            except anthropic.RateLimitError as e:
                if intento == self.max_reintentos:
                    raise ErrorCuotaLLM(f"Claude: límite de tasa/cuota tras {intento} intentos: {e}") from e
                espera = float(e.response.headers.get("retry-after", 0) or 0) or _esperar(intento, base=15.0)
                logger.warning(f"[REINTENTO {intento}/{self.max_reintentos}] Claude 429; espero {espera:.0f}s.")
                time.sleep(espera)
                continue
            except anthropic.APIConnectionError as e:
                if intento == self.max_reintentos:
                    raise ErrorCuotaLLM(f"Sin conexión con Claude tras {intento} intentos: {e}") from e
                time.sleep(_esperar(intento))
                continue
            except anthropic.BadRequestError as e:
                raise ErrorLLM(f"Claude rechazó la solicitud (400): {e.message}") from e
            except anthropic.APIStatusError as e:
                if e.status_code >= 500 and intento < self.max_reintentos:
                    time.sleep(_esperar(intento))
                    continue
                if e.status_code >= 500 or e.status_code in (401, 403):
                    raise ErrorCuotaLLM(f"Claude no disponible ({e.status_code}): {e.message}") from e
                raise ErrorLLM(f"Error de API de Claude ({e.status_code}): {e.message}") from e

            if resp.stop_reason == "refusal":
                raise ErrorLLM("Claude declinó procesar el documento (stop_reason=refusal).")
            if resp.stop_reason == "max_tokens":
                raise ErrorLLM("Respuesta de Claude truncada (max_tokens); aumente max_tokens.")
            texto_resp = next((b.text for b in resp.content if b.type == "text"), "")
            return _parsear_json(texto_resp)
        raise ErrorCuotaLLM("Claude: reintentos agotados.")


def crear_cliente(proveedor: str, modelo: Optional[str] = None) -> ClienteLLM:
    proveedor = (proveedor or "gemini").strip().lower()
    if proveedor == "gemini":
        return ClienteGemini(modelo)
    if proveedor == "claude":
        return ClienteClaude(modelo)
    raise ValueError(f"Proveedor de modelo no soportado: {proveedor!r} (use gemini o claude)")
