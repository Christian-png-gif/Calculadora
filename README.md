# Calculadora Tx

Aplicación web (Streamlit) para analizar facturas emitidas, facturas recibidas, movimientos bancarios y retenciones de ISR: calcula ingresos facturados vs. cobrados, gastos deducibles/no deducibles, IVA (con base en flujo de efectivo), retenciones de ISR, concilia facturas contra movimientos bancarios, genera alertas de validación y produce un reporte Excel descargable.

> ⚠️ **Los resultados de esta herramienta son auxiliares.** Deben ser revisados por un contador o especialista fiscal antes de presentar declaraciones o tomar decisiones fiscales. La aplicación no sustituye la asesoría profesional.

---

## 1. Instalación

Requisitos: Python 3.10 o superior.

```bash
# 1. Clona o descarga el proyecto
git clone <url-de-tu-repositorio>
cd calculadora-tx

# 2. Crea un entorno virtual (opcional pero recomendado)
python -m venv venv
source venv/bin/activate      # En Windows: venv\Scripts\activate

# 3. Instala las dependencias
pip install -r requirements.txt
```

## 2. Ejecución local

```bash
streamlit run app.py
```

Esto abre la aplicación en tu navegador (normalmente en `http://localhost:8501`). Desde el panel izquierdo carga los cuatro archivos Excel y presiona **Procesar información**.

---

## 3. Estructura de las cuatro bases de datos

Todos los archivos deben ser Excel (`.xlsx`). Los nombres de columna deben coincidir **exactamente** (minúsculas, sin acentos en los encabezados) con lo indicado abajo. Si falta alguna columna obligatoria, la aplicación te lo indicará antes de procesar nada.

### 3.1. Facturas emitidas

| Columna | Descripción |
|---|---|
| `folio` | Folio único de la factura (ej. `F-0001`) |
| `fecha` | Fecha de emisión |
| `cliente` | Nombre del cliente |
| `concepto` | Concepto facturado |
| `subtotal` | Importe antes de IVA |
| `iva` | IVA de la factura |
| `total` | Importe total (subtotal + iva) |

### 3.2. Facturas recibidas (gastos)

| Columna | Descripción |
|---|---|
| `folio_gasto` | Folio único del gasto (ej. `G-0001`) |
| `fecha` | Fecha del gasto |
| `proveedor` | Nombre del proveedor |
| `concepto` | Concepto del gasto |
| `subtotal` | Importe antes de IVA |
| `iva` | IVA del gasto |
| `total` | Importe total |
| `comprobante` | `Si` / `No` — indica si existe CFDI/comprobante fiscal |

### 3.3. Movimientos bancarios

| Columna | Descripción |
|---|---|
| `fecha` | Fecha del movimiento |
| `monto` | Importe (positivo para ingresos, negativo para egresos) |
| `tipo` | Tipo de movimiento: `Deposito`, `Ingreso`, `Cobro`, `Abono` (entradas) o `Egreso` (salidas) |
| `referencia` | Referencia bancaria — para ingresos suele coincidir con el folio de la factura emitida |

### 3.4. Retenciones de ISR

| Columna | Descripción |
|---|---|
| `fecha` | Fecha de la retención |
| `referencia` | Referencia de la retención |
| `concepto` | Concepto de la retención |
| `isr_retenido` | Monto de ISR retenido |

### Notas de crédito, devoluciones y cancelaciones

Se identifican porque el **importe (`total`) viene en negativo** dentro de la base de facturas emitidas o de gastos. La aplicación las separa automáticamente del resto de las operaciones (no se mezclan en la conciliación ni en el IVA), pero se conservan íntegras en el reporte Excel, en una hoja aparte ("Notas de crédito"), para trazabilidad.

---

## 4. Cómo funcionan los cálculos

### 4.1. Ingresos

- **Total facturado** = suma de `total` en facturas emitidas (sin notas de crédito).
- **Total cobrado** = suma de `total` de las facturas cuyo estado de conciliación es "Conciliado" o "Conciliado con diferencia dentro de tolerancia".
- **Pendiente de cobro** = facturado − cobrado.

### 4.2. Gastos y deducibilidad

Se aplican reglas fiscales generales sobre cada gasto:

- **No deducible** si no hay comprobante fiscal (`comprobante = No`), o si el concepto es `Multa`, `Donativo no autorizado` o `Gasto personal`.
- **Revisión manual** si el concepto es `Recargo` (los recargos fiscales normalmente no son deducibles, pero podrían corresponder a un recargo comercial deducible — no se asume sin más contexto) o si el concepto viene vacío.
- **Deducible** en cualquier otro caso, siempre que exista comprobante fiscal.

Estas reglas están centralizadas en las constantes `CONCEPTOS_NO_DEDUCIBLES` y `CONCEPTOS_REVISION_MANUAL` al inicio de `app.py`, por si es necesario ajustarlas.

### 4.3. IVA (con base en flujo de efectivo)

El cálculo de IVA se hace con base en **flujo de efectivo**, es decir, solo se considera:

- **IVA trasladado**: el de las facturas emitidas efectivamente **cobradas** (conciliadas contra el banco).
- **IVA acreditable**: el de los gastos efectivamente **pagados** (conciliados contra el banco) y clasificados como **deducibles**.
- **IVA neto** = IVA trasladado − IVA acreditable.

Se usa una tasa general del 16%, tomando el `iva` ya calculado en cada base (no se recalcula a partir del subtotal, para no introducir diferencias por redondeo respecto a la factura original).

### 4.4. Retenciones de ISR

En las bases con las que se probó la aplicación, las retenciones de ISR **no traen un folio o UUID** que las ligue de forma confiable a una factura o movimiento específico, y el monto retenido tampoco corresponde a una tasa fija y verificable sobre alguna factura. Por eso la aplicación **no intenta adivinar** esa relación: presenta el total retenido y marca cada renglón como pendiente de conciliación manual, en vez de arriesgar coincidencias falsas.

> Si tus bases reales llegan a incluir un folio o UUID que relacione cada retención con su factura, se puede extender fácilmente la función `procesar_isr()` en `app.py` para conciliar por folio, igual que se hace con ingresos y gastos.

### 4.5. Conciliación bancaria

Se concilian por separado:

- **Ingresos**: facturas emitidas vs. movimientos bancarios de tipo `Deposito`, `Ingreso`, `Cobro` o `Abono`.
- **Egresos**: gastos vs. movimientos bancarios de tipo `Egreso`.

Prioridad de coincidencia:

1. **Folio interno**: si la `referencia` del movimiento bancario coincide exactamente con el folio del documento. Esto solo aplica a ingresos, porque en los datos disponibles el banco no registra el folio del gasto en las referencias de pago — para egresos se usa directamente el criterio 2.
2. **Importe (± $10.00) + Fecha (± 5 días naturales)**: se usa como respaldo cuando no hay coincidencia de folio, y es el único criterio disponible para egresos.

Si hay **dos o más movimientos igualmente probables** para un mismo documento, la aplicación **no adivina**: lo marca como "Requiere revisión manual".

Cada documento queda clasificado como:

- **Conciliado** (coincidencia exacta)
- **Conciliado con diferencia dentro de tolerancia**
- **No conciliado**
- **Posible duplicado** (el folio se repite en el archivo original)
- **Requiere revisión manual**

Los pagos parciales o pagos que cubren varias facturas también se identifican como **"Conciliado con diferencia dentro de tolerancia"** o **"No conciliado"** según qué tan cerca esté el importe — no se intenta descomponer un pago en varias facturas automáticamente, para evitar asignaciones incorrectas.

---

## 5. Alertas

La aplicación genera alertas con nivel de prioridad (alta / media / baja) para:

- Facturas o gastos sin movimiento bancario relacionado.
- Movimientos bancarios sin documento relacionado.
- Diferencias de importe o de fecha dentro de tolerancia.
- Folios duplicados.
- Datos obligatorios faltantes.
- Importes en cero.
- Gastos sin clasificación clara de deducibilidad.
- Retenciones de ISR no conciliadas.
- Casos de revisión manual por coincidencias múltiples.

Todas se muestran en la pestaña **Alertas** (filtrable por prioridad) y en la hoja "Alertas y validaciones" del reporte Excel.

---

## 6. Reporte Excel

El botón **Descargar reporte Excel completo** genera un archivo con las siguientes hojas:

1. Resumen ejecutivo
2. Ingresos facturados y cobrados
3. Gastos deducibles
4. Gastos no deducibles
5. Cálculo de IVA
6. Retenciones de ISR
7. Conciliación bancaria
8. Partidas no conciliadas
9. Alertas y validaciones
10. Datos - Facturas emitidas / Gastos / Movimientos (datos originales + campos calculados, para trazabilidad)
11. Notas de crédito (si existen)

También hay un botón adicional para descargar **solo las excepciones** (partidas no conciliadas) en un archivo separado, pensado para revisión manual rápida.

---

## 7. Publicar en GitHub

```bash
git init
git add app.py requirements.txt README.md
git commit -m "Primera versión de Calculadora Tx"
git branch -M main
git remote add origin https://github.com/<tu-usuario>/<tu-repositorio>.git
git push -u origin main
```

No subas archivos con datos reales de clientes al repositorio si este va a ser público.

## 8. Desplegar en Streamlit Community Cloud

1. Entra a [share.streamlit.io](https://share.streamlit.io) e inicia sesión con tu cuenta de GitHub.
2. Haz clic en **"New app"**.
3. Selecciona el repositorio y la rama donde subiste el proyecto.
4. En **"Main file path"** escribe `app.py`.
5. Haz clic en **"Deploy"**.

Streamlit Community Cloud instalará automáticamente lo indicado en `requirements.txt` y publicará la aplicación en una URL pública tipo `https://<tu-app>.streamlit.app`.

---

## 9. Supuestos aplicados (autorizados explícitamente antes de construir la aplicación)

- La conciliación usa la prioridad **Folio interno → Importe → Fecha** (sin RFC, porque ninguna de las cuatro bases lo incluye).
- Los pagos a proveedores (egresos) se concilian únicamente por **importe + fecha**, ya que la referencia bancaria de esos pagos no corresponde al folio del gasto.
- Las notas de crédito, devoluciones y cancelaciones se identifican por **importe negativo**.
- Las retenciones de ISR se presentan como resumen y quedan marcadas para **revisión manual**, sin intentar ligarlas automáticamente a una factura.
- La clasificación de deducibilidad usa reglas fiscales generales (ver sección 4.2), no un catálogo específico del cliente.

Si cualquiera de estos supuestos cambia (por ejemplo, si las bases llegan a incluir RFC o UUID), estas reglas se pueden ajustar directamente en `app.py` sin rediseñar la aplicación.
