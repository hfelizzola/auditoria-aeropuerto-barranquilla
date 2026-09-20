"""
esquema_soportes.py
===================
Esquema estricto de extracción de soportes del P.A. Aeropuerto Ernesto Cortissoz.

PRINCIPIO RECTOR
----------------
El modelo de lenguaje SOLO transcribe hechos verificables, con página y texto
literal. NO clasifica según el Contrato, NO calcula, NO emite veredictos.
La clasificación 22.3 y el cuadre de valores viven en `validacion_22_3.py`,
en código determinista y auditable.

Grano de extracción: un PAQUETE = una orden de pago (OP), que puede venir en
varios archivos PDF y contener varios documentos internos.

Requisitos: pydantic>=2
"""

from __future__ import annotations

from typing import List, Literal, Optional

from pydantic import BaseModel, Field

# ---------------------------------------------------------------------------
# Vocabularios cerrados
# ---------------------------------------------------------------------------

TipoDocumento = Literal[
    "solicitud_fiduciaria",
    "factura_venta",
    "cuenta_de_cobro",
    "nota_credito",
    "acta_de_obra_o_avance",
    "acta_de_legalizacion_anticipo",
    "acta_de_liquidacion_contrato",
    "informe_interventoria",
    "certificacion_interventor",
    "comprobante_de_egreso_o_pago",
    "extracto_o_soporte_bancario",
    "correo_electronico",
    "contrato_orden_de_compra_u_otrosi",
    "poliza_o_certificado_de_seguro",
    "declaracion_o_recibo_tributario",
    "planilla_seguridad_social",
    "cotizacion_o_propuesta",
    "informe_auditoria_o_revisoria",
    "caratula_o_indice",
    "otro",
]

TipoValor = Literal[
    "subtotal",
    "iva",
    "impuesto_consumo_u_otro",
    "total_bruto",
    "retefuente",
    "reteica",
    "reteiva",
    "retegarantia",
    "amortizacion_de_anticipo",
    "descuento",
    "neto_a_pagar",
    "valor_girado_o_pagado",
    "valor_legalizado",
    "saldo_pendiente",
    "otro",
]

# Cada clave corresponde a una condición que el Numeral 22.3 c) exige acreditar,
# o a una causal expresa de exclusión. El modelo solo dice si aparece o no.
ClaveEvidencia = Literal[
    # --- AR7: actuaciones en las Intervenciones, verificadas por el Interventor
    "firma_o_visto_bueno_del_interventor",
    "acta_de_obra_o_avance_suscrita",
    "intervencion_identificada",
    # --- Anticipos y fondo rotatorio: solo la porción legalizada es costo
    "entrega_de_anticipo_o_fondo_rotatorio",
    "acta_o_relacion_de_legalizacion",
    # --- AR1: primas y comisiones de garantías del Contrato
    "poliza_garantia_unica_de_cumplimiento",
    "poliza_responsabilidad_civil_aeropuertos",
    "poliza_todo_riesgo_danos_materiales",
    "otra_poliza_no_prevista_en_el_contrato",
    # --- AR2: aportes a Subcuentas de la Cuenta ANI
    "subcuenta_ani_destino_identificada",
    # --- AR3: Comisión de Éxito al Consultor
    "pago_denominado_comision_de_exito",
    "receptor_identificado_como_consultor_estructurador",
    # --- AR4: estudios y diseños de las Intervenciones
    "estudio_o_diseno_de_intervencion",
    # --- AR5 / AR6: gestión social-ambiental y predial, con sus excepciones
    "gestion_social_o_ambiental",
    "gestion_predial",
    "cargo_a_subcuenta_compensaciones_ambientales",
    "cargo_a_subcuenta_predios",
    # --- AR8: O&M, administración e impuestos
    "operacion_o_mantenimiento",
    "gasto_de_administracion",
    "impuesto_tasa_o_contribucion",
    # --- AR9 y exclusión expresa del servicio de la deuda
    "comision_u_honorario_a_prestamista",
    "abono_a_capital_o_principal_de_credito",
    "intereses_de_credito",
    # --- Exclusiones generales (valores brutos, costos directos)
    "depreciacion_amortizacion_o_valorizacion",
    "interes_de_mora_multa_o_sancion",
    # --- Soporte del giro y del período
    "periodo_de_prestacion_del_servicio_explicito",
    "comprobante_de_pago_o_egreso_presente",
    "solicitud_fiduciaria_firmada",
]


# ---------------------------------------------------------------------------
# Piezas atómicas
# ---------------------------------------------------------------------------


class Valor(BaseModel):
    """Una cifra TRANSCRITA del documento. Nunca calculada por el modelo."""

    tipo: TipoValor
    monto: float = Field(description="Número sin separadores de miles ni símbolo.")
    moneda: Literal["COP", "USD", "EUR", "DESCONOCIDA"] = "COP"
    trm: Optional[float] = Field(
        default=None, description="Solo si el documento la expresa explícitamente."
    )
    pagina: int
    texto_literal: str = Field(
        description="Fragmento textual del documento donde aparece la cifra."
    )


class ConceptoLinea(BaseModel):
    """Una línea o ítem facturado/cobrado."""

    descripcion: str
    cantidad: Optional[float] = None
    valor_unitario: Optional[float] = None
    valor_linea: Optional[float] = None
    moneda: Literal["COP", "USD", "EUR", "DESCONOCIDA"] = "COP"
    periodo_servicio_inicio: Optional[str] = Field(
        default=None, description="YYYY-MM-DD o YYYY-MM si el documento lo indica."
    )
    periodo_servicio_fin: Optional[str] = None
    contrato_u_orden_referenciado: Optional[str] = None
    acta_referenciada: Optional[str] = None
    intervencion_referenciada: Optional[str] = Field(
        default=None,
        description="Nombre o número de Intervención textual, p. ej. 'Terminal', "
        "'Pista', 'Intervención 3'. Null si no se menciona.",
    )
    pagina: int
    texto_literal: str


class Evidencia(BaseModel):
    """Presencia o ausencia de una condición contractual. Sin interpretación."""

    clave: ClaveEvidencia
    presente: bool
    detalle: Optional[str] = Field(
        default=None,
        description="Qué se observó: nombre del firmante, número de acta, "
        "nombre de la subcuenta, número de póliza, etc.",
    )
    pagina: Optional[int] = None
    texto_literal: Optional[str] = None


class Parte(BaseModel):
    nombre: Optional[str] = None
    nit_o_cc: Optional[str] = Field(
        default=None, description="Solo dígitos, sin puntos ni dígito de verificación."
    )
    rol_declarado: Optional[str] = Field(
        default=None, description="Como aparece: 'contratista', 'beneficiario', etc."
    )


class Documento(BaseModel):
    """Un documento identificable dentro del PDF."""

    doc_indice: int = Field(description="Consecutivo dentro del paquete, desde 1.")
    tipo: TipoDocumento
    pagina_inicio: int
    pagina_fin: int
    legible: bool = Field(
        description="False si el escaneo impide leer las cifras con confianza."
    )
    numero_documento: Optional[str] = Field(
        default=None, description="Número de factura, cuenta de cobro, acta o póliza."
    )
    fecha_documento: Optional[str] = Field(default=None, description="YYYY-MM-DD.")
    emisor: Optional[Parte] = None
    receptor: Optional[Parte] = None
    beneficiario_del_giro: Optional[Parte] = Field(
        default=None,
        description="A quién se ordena girar, cuando difiere del emisor.",
    )
    cuenta_bancaria_o_patrimonio: Optional[str] = Field(
        default=None,
        description="Cuenta o subcuenta de origen/destino tal como aparece.",
    )
    numero_op_mencionado: Optional[str] = None
    conceptos: List[ConceptoLinea] = Field(default_factory=list)
    valores: List[Valor] = Field(default_factory=list)
    evidencias: List[Evidencia] = Field(default_factory=list)
    observaciones: Optional[str] = Field(
        default=None,
        description="Anomalías objetivas: tachaduras, páginas ilegibles, "
        "documento incompleto, firmas faltantes.",
    )


class PaqueteSoporte(BaseModel):
    """Resultado de extracción de un expediente de orden de pago completo."""

    numero_op_detectado: Optional[str] = None
    total_paginas: int
    documentos: List[Documento]
    notas_de_extraccion: Optional[str] = None


# ---------------------------------------------------------------------------
# Prompt
# ---------------------------------------------------------------------------

PROMPT_EXTRACCION = """\
Eres un asistente de transcripción documental para una auditoría financiera de \
un patrimonio autónomo en liquidación. Las imágenes que recibes son las páginas, \
en orden, de UN expediente de orden de pago (OP) de un fideicomiso. Un expediente \
puede contener varios documentos distintos: solicitud fiduciaria, facturas, cuentas \
de cobro, actas de obra, informes de interventoría, comprobantes de egreso, pólizas, \
correos.

TU ÚNICA TAREA ES TRANSCRIBIR HECHOS VERIFICABLES. Reglas estrictas:

1. NO clasifiques el gasto, no opines si procede o no, no menciones contratos ni \
normas. Otro sistema hace esa evaluación.
2. NO calcules ni deduzcas cifras. Transcribe solo montos impresos en la página. \
Si un total no aparece impreso, no lo inventes: omítelo.
3. Cada dato debe traer el número de página (1 = primera imagen) y el `texto_literal` \
en que se apoya, copiado tal cual del documento.
4. Si un dato no aparece, usa null. Nunca rellenes con suposiciones.
5. Separa el expediente en documentos por su rango de páginas. Un documento que se \
repite (original y copia) se reporta dos veces, con sus páginas.
6. Si una página es ilegible, marca `legible: false` en ese documento y descríbelo \
en `observaciones`.
7. En `evidencias`, reporta ÚNICAMENTE las claves que puedas afirmar o negar con lo \
que ves. Si no hay elementos para pronunciarte sobre una clave, no la incluyas. \
`presente: false` significa "busqué y el documento muestra que no está" \
(por ejemplo, un acta sin firma del interventor), no "no sé".
8. Presta atención especial y transcribe con exactitud: número de OP, período de \
prestación del servicio, retenciones (retefuente, reteICA, reteIVA, retegarantía), \
amortización de anticipos, moneda y TRM, y firmas o sellos de interventoría.
9. Los NIT van solo con dígitos, sin puntos ni dígito de verificación.
10. Las fechas en formato YYYY-MM-DD; si solo hay mes y año, YYYY-MM.

Responde exclusivamente con el JSON del esquema solicitado, sin texto adicional \
ni marcadores de código.
"""
