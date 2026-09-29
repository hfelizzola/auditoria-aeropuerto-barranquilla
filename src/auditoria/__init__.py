"""
Paquete `auditoria`: piezas compartidas por el orquestador (00) y las fases 01-04.

- config:          rutas, .env, logging, taxonomía de 5 etiquetas y constantes del proyecto.
- reconciliacion:  cuadre determinístico de valores y búsqueda de combinaciones de retención.
- lote_excel:      lectura/escritura del Excel de lote (hojas Resumen / Hallazgos / Control).
- modelos_llm:     clientes Gemini / Claude con salida JSON estricta y reintentos.
- estado:          checkpoints por rank (JSON) y manifiesto de descargas.
"""
