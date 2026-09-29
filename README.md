# Automatización de Auditoría Contable - Concesión Aeropuerto Ernesto Cortissoz
## Auditoría documental de la liquidación bajo el Numeral 22.3 c) (Contrato ANI 003 de 2015)

Este proyecto automatiza la revisión, fila por fila, de los pagos registrados en la hoja
`UNIVERSALIDAD` del Patrimonio Autónomo del Aeropuerto Ernesto Cortissoz contra su soporte PDF
en SharePoint, con la misma metodología y el mismo Excel de entrega que se usan en la revisión
manual. Un solo comando selecciona el lote, descarga los soportes, extrae los hechos de cada
documento con un modelo de visión (Gemini o Claude), cuadra valores y retenciones en Python y
clasifica cada transacción en 5 etiquetas.

---

## 1. Estructura del Proyecto

```text
Auditoria Automatizada/
├── data/
│   ├── raw/(A)BAS~1.xlsx                 # Base contable (hoja UNIVERSALIDAD)
│   ├── soportes/<lote>/                  # PDFs descargados por lote + _manifiesto_descarga.csv
│   └── output/
│       ├── Validacion_Soportes_Rango…_UNIVERSALIDAD_ABAS1.xlsx   # ENTREGABLE de cada lote
│       ├── estado/<lote>/ranks/rNNNNN.json   # checkpoints por rank (hechos, señales, clasificación)
│       ├── estado/<lote>/pendientes_sin_resultado.csv
│       ├── estado/<lote>/orquestador.log
│       └── .cache/                        # copia rápida de UNIVERSALIDAD (se regenera sola)
├── src/
│   ├── 00_orquestar_auditoria.py         # PUNTO DE ENTRADA ÚNICO (fases 1-4 por sub-lotes)
│   ├── 01_pareto_universalidad.py        # Ranking Pareto canónico (# y Fila UNIVERSALIDAD)
│   ├── 02_descargar_soportes_excel.py    # Descarga SharePoint / Graph / respaldo local
│   ├── 03_extraccion_documental.py       # Visión -> hechos verificables + cruce determinístico
│   ├── 04_clasificacion_auditoria.py     # 5 etiquetas + observación + hallazgos agrupados
│   ├── auditoria/                        # Módulos compartidos
│   │   ├── config.py                     #   taxonomía, partes relacionadas, RETEG, normalizadores
│   │   ├── reconciliacion.py             #   cuadre de retenciones (módulo aparte y testeable)
│   │   ├── lote_excel.py                 #   layout exacto del Excel de lote
│   │   ├── modelos_llm.py                #   clientes Gemini / Claude con JSON estricto
│   │   └── estado.py                     #   checkpoints y manifiesto
│   └── legacy/                           # Versión anterior de la descarga (CSV Pareto), archivada
├── tests/                                # python -m unittest discover -s tests
├── .env                                  # Claves y credenciales (NO se versiona)
└── requirements.txt
```

---

## 2. Instalación y configuración

Requisitos: Python 3.9 o superior (recomendado 3.11, como en CI).

```bash
pip install -r requirements.txt
cp .env.example .env      # y complete sus valores
```

Variables principales del `.env` (ver [`.env.example`](.env.example)):

| Variable | Uso |
| :--- | :--- |
| `MODELO_VISION` | `gemini` (defecto) o `claude`; se puede cambiar por corrida con `--modelo-vision`. |
| `GEMINI_API_KEY`, `GEMINI_MODEL` | Cliente Gemini (defecto `gemini-2.5-flash`). |
| `ANTHROPIC_API_KEY`, `CLAUDE_MODEL`, `CLAUDE_EFFORT` | Cliente Claude (defecto `claude-opus-5`, esfuerzo `high`). |
| `CLAUDE_FALLBACKS` | `1` (defecto): si Claude rechaza una solicitud, el servidor la reintenta con un modelo de respaldo. `0` lo desactiva. |
| `SHAREPOINT_USER` / `SHAREPOINT_PASSWORD` | Descarga autenticada (cuentas sin MFA). |
| `SHAREPOINT_FEDAUTH` / `SHAREPOINT_RTFA` | Cookies de sesión si la cuenta tiene MFA (F12 → Cookies de aerobaq.sharepoint.com). |
| `GRAPH_ACCESS_TOKEN` o `GRAPH_CLIENT_ID` (+`GRAPH_TENANT_ID`) | Respaldo por Microsoft Graph (`/v1.0/shares/{id}/driveItem`, scope `Files.Read.All`) para enlaces `:b:`. Con `GRAPH_CLIENT_ID` se usa el flujo de código de dispositivo (requiere `pip install msal`). |

---

## 3. Uso: un solo comando por lote

```bash
# Ranks 4001-4500 del ranking Pareto, en sub-lotes de 20 filas, con Gemini
python src/00_orquestar_auditoria.py --modo pareto --desde 4001 --hasta 4500 --tamano-lote 20

# El mismo lote con Claude
python src/00_orquestar_auditoria.py --modo pareto --desde 4001 --hasta 4500 --modelo-vision claude

# Reprocesar ranks pendientes de rondas anteriores (o elegidos a mano)
python src/00_orquestar_auditoria.py --modo lista --ops "1677,3015" --nombre-lote Pendientes_ronda5
python src/00_orquestar_auditoria.py --modo lista --archivo-ops ranks.csv
python src/00_orquestar_auditoria.py --modo lista --ops "1956,8598" --tipo-id op   # por número de OP

# Continuar un lote existente (p. ej. uno que se trabajó a mano)
python src/00_orquestar_auditoria.py --excel data/output/Validacion_Soportes_Rango4001-4500_UNIVERSALIDAD_ABAS1.xlsx

# Reanudar el lote más reciente de data/output, sin argumentos
python src/00_orquestar_auditoria.py
```

Opciones útiles:

| Opción | Efecto |
| :--- | :--- |
| `--tamano-lote N` | Filas por sub-lote (10, 20, 100, 500…). Tras cada sub-lote se actualizan las hojas de hallazgos y control. |
| `--limite N` | Máximo de filas pendientes a procesar en esta corrida (útil para pruebas). |
| `--fases 1,2` | Ejecuta solo algunas fases (p. ej. crear el Excel y descargar sin llamar a ningún modelo). |
| `--max-paginas 20` / `--dpi 150` | Páginas por expediente enviadas al modelo y resolución del rasterizado. |
| `--respaldo RUTA` | Carpeta local sincronizada con OneDrive donde buscar el PDF si la red falla (repetible). |
| `--fedauth …`, `--usuario …`, `--password …` | Credenciales de SharePoint por línea de comandos. |
| `--reextraer` | Vuelve a llamar al modelo de visión aunque exista una extracción guardada. |

**Reanudar es seguro.** Cada fila clasificada se escribe y se guarda en el Excel de inmediato
(guardado atómico). Si se corta la conexión, se agota la cuota del modelo (el proceso se detiene
limpiamente con código 2) o se cancela con Ctrl+C, basta repetir el mismo comando: las filas con
"Resultado de validación" se omiten, los PDF ya descargados no se vuelven a pedir y las
extracciones ya hechas no se vuelven a pagar. Si el Excel está abierto en Excel, el guardado
reintenta y avisa.

### Las 4 fases

1. **Selección.** `--modo pareto`: ordena TODA la población de UNIVERSALIDAD por valor absoluto de
   "VALOR DEBITADO O ACREDITADO" (orden estable), numera `#` solo entre las filas con NUMERO OP y
   URL, y toma las posiciones N a M. `Fila UNIVERSALIDAD` es la posición en la población ordenada
   + 1. (Verificado: reproduce exactamente los 804 ranks de los lotes 2301-2600 y 3501-4000.)
   `--modo lista`: ranks u OP explícitos. Si el Excel del lote ya existe, solo se agregan las
   filas que falten; nunca se regeneran las presentes.
2. **Descarga.** Para cada fila sin resultado: si el PDF ya está en disco (>1 KB) o en la carpeta
   de otro lote de `data/soportes/`, se reutiliza; si no, GET autenticado con reintentos y
   backoff → Microsoft Graph (enlaces `:b:`; si el enlace es una carpeta, todos sus PDF quedan como
   partes del mismo rank) → respaldo local. El manifiesto `_manifiesto_descarga.csv` registra
   rank → archivo(s) → estado (`descargado_web`, `descargado_graph`, `recuperado_local`,
   `ya_existia`, `fallido`). Si la URL del Excel viene como hipervínculo embebido y no como
   texto, se lee del paquete OOXML (`workbook.xml` → rels → hoja → rels de la hoja).
3. **Extracción documental.** Rasteriza el expediente (todas las partes del rank, en orden) a
   150 DPI en memoria —los soportes son escaneos sin capa de texto; no se usa OCR local— y el
   modelo de visión transcribe SOLO hechos verificables con un esquema JSON estricto: tercero/NIT
   emisor y beneficiario del giro, número de documento, fecha, concepto, base, IVA, retenciones
   impresas, neto, moneda y TRM impresa, menciones de NAB/GAC/OAC y su rol, endosos/cesiones,
   menciones de servicio de deuda y calidad del escaneo, cada cifra con página y texto literal.
   Luego Python (sin LLM) cruza contra UNIVERSALIDAD a nivel de fila y de OP (suma de sus filas),
   cuadra retenciones, compara tercero, fecha y moneda y detecta RETEG y partes relacionadas.
   Todo queda en `estado/<lote>/ranks/rNNNNN.json`.
4. **Clasificación.** Con los datos de UNIVERSALIDAD (nunca la columna "Validación" ni un
   resultado previo), los hechos y las señales determinísticas, el modelo emite etiqueta,
   observación y tags; Python valida la etiqueta y aplica las reglas duras antes de escribir en
   `Resultado de validación` y `Observación`.

---

## 4. Metodología y taxonomía de clasificación

Por cada transacción se verifica: (1) valor del documento (bruto, IVA, retenciones) vs.
"VALOR DEBITADO O ACREDITADO"; (2) tercero/beneficiario vs. TERCERO; (3) concepto vs. cuenta
contable; (4) fecha del documento vs. fecha de pago; (5) encuadre en la lista **taxativa** de
AR_i del numeral 22.3 c): (a) seguros/garantías del contrato, (b) aportes a subcuentas ANI,
(c) comisión de éxito al Consultor estructurador, (d) estudios y diseños, (e) gestión
social/ambiental, (f) gestión predial, (g) intervenciones verificadas por el Interventor,
(h) operación/administración/impuestos, (i) comisiones a prestamistas distintas del servicio de
deuda —el **servicio de la deuda** (capital o intereses) está excluido y se marca explícitamente—;
(6) en USD, el valor en USD y si la TRM es confirmable en el propio soporte (nunca se inventa).

| Etiqueta | Cuándo |
| :--- | :--- |
| `COHERENTE` | Todo concilia sin reparos. |
| `COHERENTE — VER NOTA` | Concilia razonablemente, pero hay algo que anotar: retención sin combinación estándar, OCR deficiente que aun así confirma lo esencial, patrón recurrente de proveedor, posible costo no imputable a parte relacionada, encuadre AR_i incierto, TRM no confirmable, cuadre solo a nivel de OP. |
| `INCONCLUSO` | El documento no permite confirmar (ilegible, cifras irreconciliables). La observación dice qué haría falta para cerrarlo. |
| `NO CORRESPONDE` | El pago no encaja en ninguna AR_i (servicio de deuda, traslado interno RETEG, donación u otro concepto ajeno), aunque esté bien soportado. |
| `TERCERO NO COINCIDE` | El beneficiario real del soporte es distinto del tercero registrado en UNIVERSALIDAD. |

### Reglas duras aplicadas en código
- **No inventar cifras**: el modelo solo transcribe montos impresos; si el documento no deja leer
  cifras y nada concilia, la fila no puede quedar COHERENTE.
- **Sin PDF → celda en blanco**: la fila nunca se marca como revisada; se lista en
  `pendientes_sin_resultado.csv` y en una entrada de "Hallazgos detallados".
- **RETEG** (tercero BANCOLOMBIA SA con cuenta "AHO … PA ERNESTO CO"/"P A AEROPUERTO" o etiqueta
  "RETEG OP n") → `NO CORRESPONDE` sin llamar al modelo.
- **Partes relacionadas** (NAB 900.913.341, GAC 900.817.115, OAC 900.849.079): se señalan y
  razonan, no se descartan. GAC como simple beneficiario de una garantía de anticipo de un
  contratista externo es normal; como beneficiario final, parte contratante ("GAC-00X-YY") o
  tomador+asegurado+beneficiario de su propia póliza → `COHERENTE — VER NOTA` (posible costo no
  imputable al P.A.).
- **Cuadre de retenciones** (`src/auditoria/reconciliacion.py`): si base+IVA no coincide con lo
  registrado, se prueban retefuente (0-11 %), reteIVA (0/15/30/100 %) y reteICA (0-14 ‰) con
  tolerancia de $1,50. Primero con tarifas estándar; si solo cierra con tarifas especiales se
  reporta como *combinación atípica* (puede ser coincidencia) y la fila queda en VER NOTA. Si no
  existe combinación: "diferencia sin explicar". Las tarifas son parámetros editables y
  `tests/test_reconciliacion.py` contiene casos reales para ajustarlas.
- **Independencia del auditor**: ni "Resultado de validación", ni "Observación", ni la columna
  "Validación" de UNIVERSALIDAD se envían al modelo.
- Un `COHERENTE` con diferencia sin explicar, TRM no confirmable o tercero ausente del soporte
  baja automáticamente a `COHERENTE — VER NOTA` (la observación lo indica con
  "[Ajuste automático: …]").

---

## 5. Entregable: el Excel del lote

Mismo layout que el entregable manual (nunca se cambia):

- **Resumen**: filas 1-5 de título y notas (descripción, cobertura acumulada); encabezado en la
  fila 6 desde la columna B: `#`, `Fila UNIVERSALIDAD`, `OP`, `Tercero (según UNIVERSALIDAD)`,
  `Cuenta contable`, `Fecha`, `Valor (COP)`, `Moneda carpeta`, `URL soporte`,
  `Resultado de validación`, `Observación`. El código solo escribe en las dos últimas.
- **Hallazgos detallados**: entradas agrupadas por patrón (no por rank) marcadas `[Automático]`,
  que se actualizan en su lugar en cada sub-lote: servicio de deuda, RETEG, TERCERO NO COINCIDE,
  NO CORRESPONDE, partes relacionadas / posibles costos no imputables, USD sin TRM, endosos y
  factoring, encuadre AR_i incierto, inconsistencias UNIVERSALIDAD vs. documento, diferencias
  sin explicar, OCR deficiente, INCONCLUSO (con qué falta), proveedores recurrentes y pendientes
  sin PDF. Las entradas escritas a mano no se tocan.
- **Control de calidad**: sección automática al final (se reescribe en cada sub-lote) con
  anomalías de los datos de UNIVERSALIDAD: OP presentes en más de un rank del lote, colisiones de
  numeración, ranks del Excel discordantes con el ranking recalculado, base/IVA que no
  corresponden al documento, IVA atípico. Se documentan, no se corrigen.

### Scripts por fase (uso avanzado)
Cada fase se puede correr sola sobre un lote existente:
```bash
python src/01_pareto_universalidad.py                       # exporta data/output/base_auditoria_pareto.csv
python src/02_descargar_soportes_excel.py --excel <lote.xlsx> [--lote 10] [--fedauth …]
python src/03_extraccion_documental.py --excel <lote.xlsx> [--modelo-vision claude] [--limite 5]
python src/04_clasificacion_auditoria.py --excel <lote.xlsx> [--limite 5]
```

### Pruebas
```bash
python -m unittest discover -s tests
```
Incluyen los casos reales de cuadre de la ronda 3501-4000, el layout del Excel, las reglas duras
y una corrida de punta a punta del orquestador con un modelo simulado (requiere PyMuPDF).

---

## 6. Control de Versiones y Despliegue en GitHub

El proyecto cuenta con configuración para versionado seguro y pipeline de Integración Continua (CI):

### Seguridad de datos y exclusiones (.gitignore)
El archivo `.gitignore` garantiza que **NUNCA** se suban a GitHub:
- Secretos o credenciales (`.env`).
- Archivos Excel contables del fideicomiso (`*.xlsx`, `*.xlsm`).
- Soportes documentales en PDF (`data/soportes/`, `*.pdf`) y checkpoints de `data/output/`.
- Informes institucionales (`*.docx`, `*.pptx`).

Se incluye la plantilla pública segura [`.env.example`](.env.example) para documentar las variables requeridas.

### Publicar en tu repositorio de GitHub

1. **Crea un repositorio vacío** en tu cuenta de GitHub (ej. `Auditoria-Aeropuerto-Barranquilla`) sin añadir README ni .gitignore.
2. **Ejecuta el script interactivo** de publicación:
   ```powershell
   .\setup_github.ps1
   ```
   O realiza los comandos manualmente:
   ```powershell
   # 1. Vincular repositorio remoto
   git remote add origin https://github.com/TU_USUARIO/TU_REPOSITORIO.git

   # 2. Subir código a la rama main
   git push -u origin main
   ```

### Integración Continua (GitHub Actions)
El flujo en [`.github/workflows/ci.yml`](.github/workflows/ci.yml) se ejecuta automáticamente ante cada `push` o `pull_request`:
- Verifica la compilación y sintaxis de todos los scripts en `src/` y `src/auditoria/`.
- Ejecuta las pruebas de `tests/` (reconciliación, layout del Excel, reglas duras, orquestador).
- Comprueba que `.env` no haya sido versionado por error.
- Valida la integridad estructural de directorios.


