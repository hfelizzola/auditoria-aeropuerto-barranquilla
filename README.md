# Automatización de Auditoría Contable - Concesión Aeropuerto Ernesto Cortissoz
## Fiscalización y Liquidación Contractual bajo el Numeral 22.3 c) (Contrato ANI 003 de 2015)

Este proyecto implementa una arquitectura automatizada de analítica contable y extracción documental asistida por inteligencia artificial (Google Gemini Multimodal) para la auditoría de costos, gastos e inversiones del Patrimonio Autónomo del Aeropuerto Ernesto Cortissoz de Barranquilla.

---

## 1. Estructura del Proyecto

```text
Auditoria Automatizada/
├── data/
│   ├── raw/
│   │   └── (A)BAS~1.xlsx             # Base contable matriz (hoja UNIVERSALIDAD)
│   ├── datalake_pdfs/                # Datalake local de PDFs descargados (OP_{id}.pdf)
│   └── output/
│       ├── base_auditoria_pareto.csv  # Muestra Pareto 80/20 de alta materialidad
│       ├── resultados_gemini.csv      # Extracciones y clasificaciones de Gemini
│       └── reporte_auditoria_final.csv # Dictamen contable-jurídico cruzado final
├── src/
│   ├── __init__.py
│   ├── 01_pareto_universalidad.py    # Filtro Pareto 80/20 y selección muestral
│   ├── 02_descargar_soportes.py      # Descarga controlada y resiliente de PDFs
│   ├── 03_extraccion_gemini.py       # Rasterizado en memoria (150 DPI) y análisis con Gemini
│   └── 04_auditoria_contrato.py      # Cruce contable y emisión de dictamen contractual
├── .env                              # Clave de API de Gemini y configuración de modelo
├── requirements.txt                  # Dependencias de Python
└── README.md                         # Documentación operativa
```

---

## 2. Requisitos e Instalación

### Requisitos previos
- Python 3.10 o superior.
- Clave de API de Google Gemini ([Google AI Studio](https://aistudio.google.com/)).

### Instalación de dependencias
```bash
pip install -r requirements.txt
```

### Configuración del entorno (`.env`)
Edite el archivo `.env` en la raíz del proyecto y configure sus credenciales:
```env
# Clave de Gemini
GEMINI_API_KEY=AIzaSy...tu_clave_real_aqui
GEMINI_MODEL=gemini-2.5-flash

# Opción A: Autenticación directa de SharePoint (si no tiene MFA obligatorio)
SHAREPOINT_USER=tu_usuario@aerobaq.com
SHAREPOINT_PASSWORD=tu_contraseña

# Opción B: Si su cuenta tiene Doble Factor (MFA / Microsoft Authenticator):
# En su navegador (Edge/Chrome), inicie sesión en https://aerobaq.sharepoint.com,
# presione F12 -> pestaña "Aplicación" o "Almacenamiento" -> Cookies -> "aerobaq.sharepoint.com",
# y copie el valor de la cookie 'FedAuth':
SHAREPOINT_FEDAUTH=
SHAREPOINT_RTFA=
```

---

## 3. Guía de Ejecución Paso a Paso

### Paso 1: Selección Muestral de Pareto (80/20)
Lee la base matriz `data/raw/(A)BAS~1.xlsx` (hoja `UNIVERSALIDAD`), normaliza las columnas contables, convierte a valor absoluto, ordena de mayor a menor y extrae el 80% del valor total de la población (equivalente con exactitud al top 1.000 de transacciones).
```bash
python src/01_pareto_universalidad.py
```
* **Salida generada**: `data/output/base_auditoria_pareto.csv`
* **Estadísticas obtenidas**: 999 registros seleccionados que concentran **$624.515.778.070,19 COP** (80.00% de la masa total debitada/acreditada).

### Paso 2: Descarga de Soportes Documentales (Datalake)
Itera sobre los enlaces a los PDFs (`LINK_SOPORTE` / `URL`) de SharePoint, autentica la sesión contra Microsoft Online o inyecta las cookies de sesión y descarga los archivos con pausas e idempotencia.
```bash
python src/02_descargar_soportes.py
```
* **Opciones disponibles**:
  * `--usuario usuario@aerobaq.com --password tu_clave`: Pasa credenciales por línea de comandos.
  * `--fedauth <valor>`: Pasa la cookie de sesión de SharePoint directamente.
  * `--limite 10`: Procesa únicamente un lote de prueba de 10 órdenes de pago.
  * `--pausa 1.0`: Configura el tiempo de espera en segundos entre descargas.
* **Salida generada**: Archivos PDF guardados en `data/datalake_pdfs/OP_{numero}.pdf`.

### Paso 3: Rasterizado y Extracción Multimodal con Gemini
Convierte cada soporte PDF a imágenes JPEG a 150 DPI en memoria (vía `pdf2image` / `PIL` y backend `pymupdf`) y lo somete al modelo de visión de Gemini con un esquema estricto de JSON. Implementa reintentos automáticos con backoff exponencial ante errores 429 (límite de cuota) o 503 (servicio no disponible).
```bash
python src/03_extraccion_gemini.py
```
* **Opciones disponibles**:
  * `--modelo gemini-2.5-flash`: Permite alternar entre modelos (`gemini-2.5-flash`, `gemini-1.5-flash`, etc.).
  * `--limite 20`: Limita el número de PDFs a procesar en la corrida.
* **Salida generada**: `data/output/resultados_gemini.csv`.

### Paso 4: Cruce y Dictamen de Auditoría Contractual
Cruza las órdenes de pago contables con las extracciones documentales y aplica las reglas taxativas de reconocimiento del Numeral 22.3 c).
```bash
python src/04_auditoria_contrato.py
```
* **Salida generada**: `data/output/reporte_auditoria_final.csv`.

---

## 4. Lógica Contractual de Dictamen (Numeral 22.3 c)

El Numeral 22.3 c) del Contrato de Concesión establece una **lista taxativa** de costos y gastos reconocibles para la liquidación. El script aplica la siguiente matriz de decisiones en estricto orden de prelación:

| Condición Documental / Factura | Categoría | Dictamen Emitido | Fundamento Jurídico-Contractual |
| :--- | :--- | :--- | :--- |
| Servicio de deuda (amortización a capital o pago de intereses de crédito) | `EXCLUIDO` | **`NO RECONOCIBLE`** | Exclusión expresa legal del Numeral 22.3 c) (AR9 prohíbe el servicio de deuda). |
| Emisor es parte vinculada (*NUEVO AEROPUERTO*, *GRUPO AEROPORTUARIO DEL CARIBE*, *OPERADORA AEROPORTUARIA*) | Cualquiera | **`POR VERIFICAR`** | Conflicto de interés y precios de transferencia. Exige aportar estudio y prueba de precios de mercado. |
| Inversión en Intervenciones de Obra (`AR7`) o desembolso de anticipos contractuales | `AR7` / Anticipo | **`CONDICIONADO`** | Sujeto a la existencia y verificación del Acta de Obra suscrita por la Interventoría técnica. |
| Pólizas y garantías contractuales | `AR1` | **`RECONOCIBLE`** | Primas y comisiones de pólizas de cumplimiento, RCE y todo riesgo autorizadas. |
| Diseños y estudios de ingeniería | `AR4` | **`RECONOCIBLE`** | Estudios y diseños de intervenciones previstos contractualmente. |
| Operación, Mantenimiento y Administración | `AR8` | **`RECONOCIBLE`** | Costos de OPEX, servicios generales, aseo, vigilancia e impuestos prediales/tasas. |
| Sin soporte descargado / procesado en datalake | N/A | **`PENDIENTE DE SOPORTE`** | Registro contable que requiere gestión de descarga física o autenticación en SharePoint. |

---

## 5. Cuadre Automático de Cifras
El reporte final computa la diferencia:
$$\text{DIFERENCIA\_VALOR} = \text{VALOR\_ABSOLUTO (Contabilidad)} - \text{VALOR\_TOTAL\_EXTRAIDO (Factura)}$$
Permitiendo a los auditores identificar inmediatamente glosas, retenciones tributarias no conciliadas o desviaciones entre el giro bancario del fideicomiso y la factura comercial.

---

## 6. Control de Versiones y Despliegue en GitHub

El proyecto cuenta con configuración para versionado seguro y pipeline de Integración Continua (CI):

### Seguridad de datos y exclusiones (.gitignore)
El archivo `.gitignore` garantiza que **NUNCA** se suban a GitHub:
- Secretos o credenciales (`.env`).
- Archivos Excel contables del fideicomiso (`*.xlsx`, `*.xlsm`).
- Soportes documentales en PDF (`data/datalake_pdfs/`, `*.pdf`).
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
- Verifica la compilación y sintaxis de todos los scripts en `src/`.
- Comprueba que `.env` no haya sido versionado por error.
- Valida la integridad estructural de directorios.


