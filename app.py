# -*- coding: utf-8 -*-
"""
Calculadora Tx
==============
Aplicación web (Streamlit) para analizar facturas emitidas, facturas
recibidas, movimientos bancarios y retenciones de ISR: calcula ingresos
facturados vs. cobrados, gastos deducibles/no deducibles, IVA (base flujo
de efectivo), retenciones de ISR, concilia facturas contra movimientos
bancarios y genera alertas de validación + un reporte Excel descargable.

Cómo ejecutar localmente:
    streamlit run app.py

El código está organizado en dos grandes bloques:
  1. LÓGICA DE NEGOCIO (funciones puras de pandas, sin Streamlit) — validación,
     clasificación, conciliación, cálculo de IVA/ISR, alertas y reporte Excel.
  2. INTERFAZ (Streamlit) — carga de archivos, filtros, tablero y descargas.

Ver README.md para el detalle de columnas obligatorias, reglas de cálculo
y supuestos aplicados.
"""

import io
from datetime import timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# =============================================================================
# 1. LÓGICA DE NEGOCIO
# =============================================================================

TASA_IVA = 0.16
TOLERANCIA_IMPORTE = 10.00        # pesos
TOLERANCIA_DIAS = 5               # días naturales

# Columnas obligatorias por archivo de entrada. Si el usuario sube un archivo
# al que le falta alguna, se le avisa con claridad antes de procesar nada.
COLUMNAS_REQUERIDAS = {
    "facturas_emitidas": ["folio", "fecha", "cliente", "concepto", "subtotal", "iva", "total"],
    "facturas_recibidas": ["folio_gasto", "fecha", "proveedor", "concepto", "subtotal", "iva", "total", "comprobante"],
    "movimientos_bancarios": ["fecha", "monto", "tipo", "referencia"],
    "retenciones_isr": ["fecha", "referencia", "concepto", "isr_retenido"],
}

NOMBRES_BASE = {
    "facturas_emitidas": "Facturas emitidas",
    "facturas_recibidas": "Facturas recibidas (gastos)",
    "movimientos_bancarios": "Movimientos bancarios",
    "retenciones_isr": "Retenciones de ISR",
}

# Reglas fiscales generales de deducibilidad, confirmadas con el cliente
# sobre los conceptos reales que maneja en sus gastos. Ajustables sin tocar
# el resto del código.
CONCEPTOS_NO_DEDUCIBLES = {"Multa", "Donativo no autorizado", "Gasto personal"}
CONCEPTOS_REVISION_MANUAL = {"Recargo"}

# Tipos de movimiento bancario que representan salida de dinero (pago a
# proveedores). Cualquier otro tipo se trata como entrada de dinero (cobro).
TIPOS_EGRESO = {"Egreso"}

PRIORIDAD_ALERTA = {
    "duplicado": "alta",
    "dato_faltante": "alta",
    "importe_invalido": "alta",
    "archivo_formato": "alta",
    "no_conciliado": "media",
    "diferencia_importe": "media",
    "diferencia_fecha": "media",
    "gasto_sin_clasificar": "media",
    "iva_inconsistente": "media",
    "revision_manual": "media",
    "isr_no_conciliado": "baja",
}


def validar_columnas(df: pd.DataFrame, clave_base: str, nombre_archivo: str) -> list:
    """Verifica columnas obligatorias. Regresa lista de mensajes de error comprensibles."""
    requeridas = COLUMNAS_REQUERIDAS[clave_base]
    presentes = set(df.columns)
    errores = []
    for col in requeridas:
        if col not in presentes:
            errores.append(
                f"❌ Al archivo de **{NOMBRES_BASE[clave_base]}** ('{nombre_archivo}') le falta "
                f"la columna obligatoria **'{col}'**. Revisa que el encabezado esté escrito exactamente así."
            )
    return errores


def preparar_dataframe(df: pd.DataFrame, columna_fecha: str = "fecha") -> pd.DataFrame:
    """Normaliza la columna de fecha a datetime sin alterar el resto de los datos originales."""
    df = df.copy()
    if columna_fecha in df.columns:
        df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors="coerce")
    return df


def separar_notas_credito(df: pd.DataFrame, columna_total: str = "total") -> tuple:
    """
    Las notas de crédito, devoluciones y cancelaciones llegan identificadas en
    los datos como importes negativos. Se separan (no se descartan) para no
    mezclarlas con operaciones normales en los cálculos de conciliación, pero
    se conservan íntegras para el reporte y la trazabilidad.
    """
    normales = df[df[columna_total] > 0].copy()
    notas_credito = df[df[columna_total] < 0].copy()
    en_cero = df[df[columna_total] == 0].copy()
    return normales, notas_credito, en_cero


def clasificar_gasto(row: pd.Series) -> str:
    """Aplica reglas fiscales generales de deducibilidad sobre un renglón de gastos."""
    tiene_comprobante = str(row.get("comprobante", "")).strip().lower() in ("si", "sí")
    concepto = row.get("concepto")
    concepto = "" if pd.isna(concepto) else str(concepto).strip()

    if not tiene_comprobante:
        return "no_deducible"
    if concepto in CONCEPTOS_NO_DEDUCIBLES:
        return "no_deducible"
    if concepto in CONCEPTOS_REVISION_MANUAL or concepto == "":
        return "revision_manual"
    return "deducible"


def clasificar_gastos(gastos: pd.DataFrame) -> pd.DataFrame:
    gastos = gastos.copy()
    gastos["clasificacion_deducibilidad"] = gastos.apply(clasificar_gasto, axis=1)
    return gastos


def _dentro_de_tolerancia(importe_doc, importe_mov, fecha_doc, fecha_mov):
    diff_importe = abs(importe_doc - importe_mov)
    if pd.isna(fecha_doc) or pd.isna(fecha_mov):
        diff_dias = None
    else:
        diff_dias = abs((fecha_doc - fecha_mov).days)
    ok_importe = diff_importe <= TOLERANCIA_IMPORTE
    ok_fecha = diff_dias is not None and diff_dias <= TOLERANCIA_DIAS
    return ok_importe, ok_fecha, diff_importe, diff_dias


def conciliar(documentos: pd.DataFrame, movimientos: pd.DataFrame,
              col_folio_doc: str, col_importe_doc: str,
              col_referencia_mov: str, col_importe_mov: str,
              usar_referencia_como_folio: bool = True) -> pd.DataFrame:
    """
    Motor de conciliación. Prioridad de coincidencia:
      1) Folio interno (referencia bancaria == folio del documento) — solo
         para ingresos, ya que el banco sí registra el folio de la factura
         en la referencia del depósito/cobro.
      2) Importe (± $10) + Fecha (± 5 días) — se usa siempre como respaldo,
         y es el único criterio disponible para egresos (pagos a
         proveedores), porque el banco no registra el folio del gasto en
         la referencia de pago.
    Si hay dos o más movimientos igualmente probables, NO se adivina: se
    marca 'Requiere revisión manual'.
    """
    documentos = documentos.reset_index(drop=True).copy()
    movimientos = movimientos.reset_index(drop=True).copy()

    documentos["estado_conciliacion"] = "No conciliado"
    documentos["movimiento_relacionado"] = np.nan
    documentos["diferencia_importe"] = np.nan
    documentos["diferencia_dias"] = np.nan

    usados = set()

    if usar_referencia_como_folio:
        for i, doc in documentos.iterrows():
            candidatos = movimientos[
                (movimientos[col_referencia_mov] == doc[col_folio_doc])
                & (~movimientos.index.isin(usados))
            ]
            if len(candidatos) == 1:
                mov = candidatos.iloc[0]
                _, _, diff_imp, diff_dias = _dentro_de_tolerancia(
                    doc[col_importe_doc], mov[col_importe_mov], doc["fecha"], mov["fecha"]
                )
                documentos.at[i, "movimiento_relacionado"] = candidatos.index[0]
                documentos.at[i, "diferencia_importe"] = diff_imp
                documentos.at[i, "diferencia_dias"] = diff_dias
                documentos.at[i, "estado_conciliacion"] = (
                    "Conciliado" if diff_imp <= 0.001 else "Conciliado con diferencia dentro de tolerancia"
                )
                usados.add(candidatos.index[0])
            elif len(candidatos) > 1:
                documentos.at[i, "estado_conciliacion"] = "Requiere revisión manual"

    pendientes = documentos[documentos["estado_conciliacion"] == "No conciliado"]
    for i in pendientes.index:
        doc = documentos.loc[i]
        disponibles = movimientos[~movimientos.index.isin(usados)]
        matches = []
        for j, mov in disponibles.iterrows():
            ok_importe, ok_fecha, diff_imp, diff_dias = _dentro_de_tolerancia(
                doc[col_importe_doc], mov[col_importe_mov], doc["fecha"], mov["fecha"]
            )
            if ok_importe and ok_fecha:
                matches.append((j, diff_imp, diff_dias))

        if len(matches) == 1:
            j, diff_imp, diff_dias = matches[0]
            documentos.at[i, "movimiento_relacionado"] = j
            documentos.at[i, "diferencia_importe"] = diff_imp
            documentos.at[i, "diferencia_dias"] = diff_dias
            documentos.at[i, "estado_conciliacion"] = (
                "Conciliado" if diff_imp <= 0.001 else "Conciliado con diferencia dentro de tolerancia"
            )
            usados.add(j)
        elif len(matches) > 1:
            documentos.at[i, "estado_conciliacion"] = "Requiere revisión manual"

    # Folios duplicados (mismo folio repetido en el archivo original)
    dup_mask = documentos.duplicated(subset=[col_folio_doc], keep=False)
    ya_en_revision = documentos["estado_conciliacion"] == "Requiere revisión manual"
    documentos.loc[dup_mask & ~ya_en_revision, "estado_conciliacion"] = "Posible duplicado"

    return documentos


def marcar_movimientos_sin_documento(movimientos: pd.DataFrame, documentos_conciliados: pd.DataFrame) -> pd.DataFrame:
    usados = set(documentos_conciliados["movimiento_relacionado"].dropna().astype(int))
    movimientos = movimientos.reset_index(drop=True).copy()
    movimientos["tiene_documento"] = movimientos.index.isin(usados)
    return movimientos


def calcular_iva_flujo(facturas_conc: pd.DataFrame, gastos_conc: pd.DataFrame) -> dict:
    """IVA con base en flujo de efectivo: solo cuenta lo efectivamente cobrado/pagado."""
    conciliados = ["Conciliado", "Conciliado con diferencia dentro de tolerancia"]
    cobradas = facturas_conc[facturas_conc["estado_conciliacion"].isin(conciliados)]
    iva_trasladado = cobradas["iva"].sum()

    pagados_deducibles = gastos_conc[
        gastos_conc["estado_conciliacion"].isin(conciliados)
        & (gastos_conc["clasificacion_deducibilidad"] == "deducible")
    ]
    iva_acreditable = pagados_deducibles["iva"].sum()

    return {
        "iva_trasladado": round(float(iva_trasladado), 2),
        "iva_acreditable": round(float(iva_acreditable), 2),
        "iva_neto": round(float(iva_trasladado - iva_acreditable), 2),
    }


def procesar_isr(isr: pd.DataFrame) -> pd.DataFrame:
    """
    En las bases disponibles, las retenciones de ISR no traen folio ni UUID
    que las ligue de forma confiable a una factura o movimiento específico,
    y el monto retenido tampoco corresponde a una tasa fija sobre alguna
    factura (se validó contra los datos reales: la relación importe/factura
    no es constante). Por eso esta función no intenta adivinar relaciones —
    en vez de arriesgar falsos positivos, presenta el total retenido y marca
    cada renglón como pendiente de conciliación manual.

    Si las bases del cliente llegan a incluir un folio o UUID de relación,
    agregar esa columna en COLUMNAS_REQUERIDAS['retenciones_isr'] y construir
    aquí una conciliación por folio, igual que la de ingresos y gastos.
    """
    isr = isr.copy()
    isr["estado_conciliacion"] = "Requiere revisión manual"
    return isr


def generar_alertas(facturas_conc, gastos_conc, ingresos_mov, egresos_mov, isr_proc, resultado_iva) -> pd.DataFrame:
    filas = []

    def agregar(tipo, descripcion, origen, referencia=""):
        filas.append({
            "tipo": tipo, "descripcion": descripcion,
            "prioridad": PRIORIDAD_ALERTA.get(tipo, "media"),
            "origen": origen, "referencia": referencia,
        })

    for _, r in facturas_conc[facturas_conc["estado_conciliacion"] == "No conciliado"].iterrows():
        agregar("no_conciliado", f"Factura {r['folio']} sin movimiento bancario relacionado.", "Facturas emitidas", r["folio"])

    for _, r in gastos_conc[gastos_conc["estado_conciliacion"] == "No conciliado"].iterrows():
        agregar("no_conciliado", f"Gasto {r['folio_gasto']} sin movimiento bancario relacionado.", "Gastos", r["folio_gasto"])

    for _, r in ingresos_mov[~ingresos_mov["tiene_documento"]].iterrows():
        agregar("no_conciliado", f"Depósito de ${r['monto']:,.2f} ({r['referencia']}) sin factura relacionada.", "Movimientos bancarios", r["referencia"])

    for _, r in egresos_mov[~egresos_mov["tiene_documento"]].iterrows():
        agregar("no_conciliado", f"Pago de ${abs(r['monto']):,.2f} ({r['referencia']}) sin gasto relacionado.", "Movimientos bancarios", r["referencia"])

    for nombre_base, df, col_folio in [("Facturas emitidas", facturas_conc, "folio"), ("Gastos", gastos_conc, "folio_gasto")]:
        con_dif = df[df["estado_conciliacion"] == "Conciliado con diferencia dentro de tolerancia"]
        for _, r in con_dif.iterrows():
            if r["diferencia_importe"] and r["diferencia_importe"] > 0.01:
                agregar("diferencia_importe", f"{nombre_base} {r[col_folio]}: diferencia de ${r['diferencia_importe']:,.2f} respecto al movimiento bancario.", nombre_base, r[col_folio])
            if r["diferencia_dias"] and r["diferencia_dias"] > 0:
                agregar("diferencia_fecha", f"{nombre_base} {r[col_folio]}: diferencia de {int(r['diferencia_dias'])} día(s) respecto al movimiento bancario.", nombre_base, r[col_folio])

        dups = df[df["estado_conciliacion"] == "Posible duplicado"]
        for folio in dups[col_folio].unique():
            agregar("duplicado", f"El folio {folio} aparece repetido en {nombre_base.lower()}.", nombre_base, folio)

        revisar = df[df["estado_conciliacion"] == "Requiere revisión manual"]
        for _, r in revisar.iterrows():
            agregar("revision_manual", f"{nombre_base} {r[col_folio]}: hay más de un movimiento bancario igualmente probable. Se requiere revisión manual.", nombre_base, r[col_folio])

    for _, r in gastos_conc[gastos_conc["clasificacion_deducibilidad"] == "revision_manual"].iterrows():
        agregar("gasto_sin_clasificar", f"Gasto {r['folio_gasto']} (concepto: '{r['concepto']}') requiere revisión manual de deducibilidad.", "Gastos", r["folio_gasto"])

    for _, r in isr_proc.iterrows():
        agregar("isr_no_conciliado", f"Retención ISR {r['referencia']} de ${r['isr_retenido']:,.2f} no pudo ligarse automáticamente a una factura; requiere revisión manual.", "Retenciones ISR", r["referencia"])

    if resultado_iva["iva_neto"] < 0:
        agregar("iva_inconsistente", f"El IVA neto del periodo es a favor por ${abs(resultado_iva['iva_neto']):,.2f}.", "IVA")

    return pd.DataFrame(filas, columns=["tipo", "descripcion", "prioridad", "origen", "referencia"])


def validar_datos_faltantes(df: pd.DataFrame, columnas_clave: list, nombre_base: str, col_folio: str) -> pd.DataFrame:
    filas = []
    for col in columnas_clave:
        for _, r in df[df[col].isna()].iterrows():
            filas.append({
                "tipo": "dato_faltante",
                "descripcion": f"Falta el dato '{col}' en el registro {r.get(col_folio, '(sin folio)')} de {nombre_base}.",
                "prioridad": "alta", "origen": nombre_base, "referencia": r.get(col_folio, ""),
            })
    return pd.DataFrame(filas, columns=["tipo", "descripcion", "prioridad", "origen", "referencia"])


def validar_importes_invalidos(df: pd.DataFrame, columnas_importe: list, nombre_base: str, col_folio: str) -> pd.DataFrame:
    """Detecta importes en cero (los negativos ya se tratan aparte como notas de crédito)."""
    filas = []
    for col in columnas_importe:
        for _, r in df[df[col] == 0].iterrows():
            filas.append({
                "tipo": "importe_invalido",
                "descripcion": f"El campo '{col}' del registro {r.get(col_folio, '(sin folio)')} de {nombre_base} está en cero.",
                "prioridad": "alta", "origen": nombre_base, "referencia": r.get(col_folio, ""),
            })
    return pd.DataFrame(filas, columns=["tipo", "descripcion", "prioridad", "origen", "referencia"])


def generar_reporte_excel(resumen: dict, fe_conc, fr_conc, iva_res, isr_proc,
                           ingresos_mov, egresos_mov, alertas, no_conciliadas,
                           fe_nc, fr_nc) -> bytes:
    """Genera el reporte Excel descargable con todas las hojas requeridas."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        # --- Resumen ejecutivo ---
        resumen_df = pd.DataFrame(list(resumen.items()), columns=["Indicador", "Valor"])
        resumen_df.to_excel(writer, sheet_name="Resumen ejecutivo", index=False, startrow=2)
        ws = writer.sheets["Resumen ejecutivo"]
        ws["A1"] = "Calculadora Tx — Resumen ejecutivo"
        ws["A2"] = (
            "Advertencia: estos resultados son auxiliares y deben ser revisados por un "
            "contador o especialista fiscal antes de presentar declaraciones o tomar decisiones fiscales."
        )

        # --- Ingresos facturados y cobrados ---
        cols_ing = ["folio", "fecha", "cliente", "concepto", "subtotal", "iva", "total", "estado_conciliacion", "diferencia_importe", "diferencia_dias"]
        fe_conc[cols_ing].to_excel(writer, sheet_name="Ingresos fact. y cobrados", index=False)

        # --- Gastos deducibles / no deducibles ---
        cols_gas = ["folio_gasto", "fecha", "proveedor", "concepto", "subtotal", "iva", "total", "comprobante", "clasificacion_deducibilidad", "estado_conciliacion"]
        fr_conc[fr_conc["clasificacion_deducibilidad"] == "deducible"][cols_gas].to_excel(writer, sheet_name="Gastos deducibles", index=False)
        fr_conc[fr_conc["clasificacion_deducibilidad"] != "deducible"][cols_gas].to_excel(writer, sheet_name="Gastos no deducibles", index=False)

        # --- Cálculo de IVA ---
        iva_df = pd.DataFrame(list(iva_res.items()), columns=["Concepto", "Monto"])
        iva_df.to_excel(writer, sheet_name="Cálculo de IVA", index=False)

        # --- Retenciones de ISR ---
        isr_proc.to_excel(writer, sheet_name="Retenciones de ISR", index=False)

        # --- Conciliación bancaria (todo) ---
        conc_ing = fe_conc[cols_ing].copy()
        conc_ing["origen"] = "Ingreso"
        conc_gas = fr_conc[["folio_gasto", "fecha", "proveedor", "concepto", "subtotal", "iva", "total", "estado_conciliacion", "diferencia_importe", "diferencia_dias"]].copy()
        conc_gas = conc_gas.rename(columns={"folio_gasto": "folio", "proveedor": "cliente"})
        conc_gas["origen"] = "Egreso"
        conciliacion_completa = pd.concat([conc_ing, conc_gas], ignore_index=True)
        conciliacion_completa.to_excel(writer, sheet_name="Conciliación bancaria", index=False)

        # --- Partidas no conciliadas ---
        no_conciliadas.to_excel(writer, sheet_name="Partidas no conciliadas", index=False)

        # --- Alertas y validaciones ---
        alertas.to_excel(writer, sheet_name="Alertas y validaciones", index=False)

        # --- Datos procesados (originales + notas de crédito, trazabilidad completa) ---
        fe_conc.to_excel(writer, sheet_name="Datos - Facturas emitidas", index=False)
        fr_conc.to_excel(writer, sheet_name="Datos - Gastos", index=False)
        ingresos_mov.to_excel(writer, sheet_name="Datos - Mov. ingresos", index=False)
        egresos_mov.to_excel(writer, sheet_name="Datos - Mov. egresos", index=False)
        if len(fe_nc) or len(fr_nc):
            pd.concat([
                fe_nc.assign(origen="Facturas emitidas"),
                fr_nc.rename(columns={"folio_gasto": "folio", "proveedor": "cliente"}).assign(origen="Gastos"),
            ], ignore_index=True).to_excel(writer, sheet_name="Notas de crédito", index=False)

    return buffer.getvalue()


# =============================================================================
# 2. INTERFAZ (STREAMLIT)
# =============================================================================

st.set_page_config(page_title="Calculadora Tx", layout="wide", page_icon="📊")

st.title("📊 Calculadora Tx")
st.caption(
    "Analiza facturas emitidas, facturas recibidas, movimientos bancarios y retenciones "
    "de ISR: ingresos vs. cobrado, gastos deducibles, IVA, ISR y conciliación bancaria."
)
st.warning(
    "⚠️ Los resultados de esta herramienta son **auxiliares**. Deben ser revisados por un "
    "contador o especialista fiscal antes de presentar declaraciones o tomar decisiones fiscales."
)

with st.sidebar:
    st.header("1. Carga tus archivos")
    archivo_fe = st.file_uploader("Facturas emitidas (Excel)", type=["xlsx", "xls"], key="fe")
    archivo_fr = st.file_uploader("Facturas recibidas / gastos (Excel)", type=["xlsx", "xls"], key="fr")
    archivo_mb = st.file_uploader("Movimientos bancarios (Excel)", type=["xlsx", "xls"], key="mb")
    archivo_isr = st.file_uploader("Retenciones de ISR (Excel)", type=["xlsx", "xls"], key="isr")
    st.divider()
    procesar = st.button("Procesar información", type="primary", width="stretch")

if not procesar:
    st.info("Carga los cuatro archivos en el panel izquierdo y presiona **Procesar información** para comenzar.")
    st.stop()

# --- Carga y validación ---
faltan = [n for n, a in [("facturas_emitidas", archivo_fe), ("facturas_recibidas", archivo_fr),
                          ("movimientos_bancarios", archivo_mb), ("retenciones_isr", archivo_isr)] if a is None]
if faltan:
    st.error("Faltan archivos por cargar: " + ", ".join(NOMBRES_BASE[f] for f in faltan))
    st.stop()

try:
    fe_raw = pd.read_excel(archivo_fe)
    fr_raw = pd.read_excel(archivo_fr)
    mb_raw = pd.read_excel(archivo_mb)
    isr_raw = pd.read_excel(archivo_isr)
except Exception as e:
    st.error(f"No se pudo leer alguno de los archivos. Verifica que sean archivos Excel válidos. Detalle técnico: {e}")
    st.stop()

errores = []
errores += validar_columnas(fe_raw, "facturas_emitidas", archivo_fe.name)
errores += validar_columnas(fr_raw, "facturas_recibidas", archivo_fr.name)
errores += validar_columnas(mb_raw, "movimientos_bancarios", archivo_mb.name)
errores += validar_columnas(isr_raw, "retenciones_isr", archivo_isr.name)

if errores:
    st.error("No se pudo procesar la información porque faltan columnas obligatorias:")
    for e in errores:
        st.markdown(e)
    st.stop()

# --- Preparación ---
fe = preparar_dataframe(fe_raw)
fr = preparar_dataframe(fr_raw)
mb = preparar_dataframe(mb_raw)
isr = preparar_dataframe(isr_raw)

fe_norm, fe_nc, fe_cero = separar_notas_credito(fe, "total")
fr_norm, fr_nc, fr_cero = separar_notas_credito(fr, "total")

fr_clasif = clasificar_gastos(fr_norm)

ingresos_mb = mb[~mb["tipo"].isin(TIPOS_EGRESO)].copy()
egresos_mb = mb[mb["tipo"].isin(TIPOS_EGRESO)].copy()
egresos_mb["monto"] = egresos_mb["monto"].abs()

fe_conc = conciliar(fe_norm, ingresos_mb, "folio", "total", "referencia", "monto", usar_referencia_como_folio=True)
fr_conc = conciliar(fr_clasif, egresos_mb, "folio_gasto", "total", "referencia", "monto", usar_referencia_como_folio=False)

ingresos_marcados = marcar_movimientos_sin_documento(ingresos_mb, fe_conc)
egresos_marcados = marcar_movimientos_sin_documento(egresos_mb, fr_conc)

resultado_iva = calcular_iva_flujo(fe_conc, fr_conc)
isr_proc = procesar_isr(isr)

alertas = generar_alertas(fe_conc, fr_conc, ingresos_marcados, egresos_marcados, isr_proc, resultado_iva)
alertas = pd.concat([
    alertas,
    validar_datos_faltantes(fe, COLUMNAS_REQUERIDAS["facturas_emitidas"], "Facturas emitidas", "folio"),
    validar_datos_faltantes(fr, COLUMNAS_REQUERIDAS["facturas_recibidas"], "Gastos", "folio_gasto"),
    validar_importes_invalidos(fe, ["subtotal", "iva", "total"], "Facturas emitidas", "folio"),
    validar_importes_invalidos(fr, ["subtotal", "iva", "total"], "Gastos", "folio_gasto"),
], ignore_index=True)

# --- Indicadores del tablero ---
total_facturado = fe_norm["total"].sum()
conciliados_ok = ["Conciliado", "Conciliado con diferencia dentro de tolerancia"]
total_cobrado = fe_conc[fe_conc["estado_conciliacion"].isin(conciliados_ok)]["total"].sum()
total_pendiente_cobro = total_facturado - total_cobrado

gastos_deducibles = fr_clasif[fr_clasif["clasificacion_deducibilidad"] == "deducible"]["total"].sum()
gastos_no_deducibles = fr_clasif[fr_clasif["clasificacion_deducibilidad"] == "no_deducible"]["total"].sum()

partidas_conciliadas = pd.concat([fe_conc[fe_conc["estado_conciliacion"].isin(conciliados_ok)],
                                   fr_conc[fr_conc["estado_conciliacion"].isin(conciliados_ok)]])
partidas_no_conciliadas = pd.concat([fe_conc[~fe_conc["estado_conciliacion"].isin(conciliados_ok)],
                                      fr_conc[~fr_conc["estado_conciliacion"].isin(conciliados_ok)]])
total_partidas = len(fe_conc) + len(fr_conc)
pct_conciliacion = (len(partidas_conciliadas) / total_partidas * 100) if total_partidas else 0

isr_total = isr["isr_retenido"].sum()

alertas_por_prioridad = alertas["prioridad"].value_counts().to_dict() if len(alertas) else {}

resumen = {
    "Total facturado": round(float(total_facturado), 2),
    "Total cobrado": round(float(total_cobrado), 2),
    "Total pendiente de cobro": round(float(total_pendiente_cobro), 2),
    "Gastos deducibles": round(float(gastos_deducibles), 2),
    "Gastos no deducibles": round(float(gastos_no_deducibles), 2),
    "IVA trasladado": resultado_iva["iva_trasladado"],
    "IVA acreditable": resultado_iva["iva_acreditable"],
    "IVA neto": resultado_iva["iva_neto"],
    "ISR retenido total": round(float(isr_total), 2),
    "Partidas conciliadas (número)": int(len(partidas_conciliadas)),
    "Partidas conciliadas (monto)": round(float(partidas_conciliadas["total"].sum()), 2),
    "Partidas no conciliadas (número)": int(len(partidas_no_conciliadas)),
    "Partidas no conciliadas (monto)": round(float(partidas_no_conciliadas["total"].sum()), 2),
    "Porcentaje de conciliación": round(pct_conciliacion, 1),
    "Alertas alta prioridad": int(alertas_por_prioridad.get("alta", 0)),
    "Alertas media prioridad": int(alertas_por_prioridad.get("media", 0)),
    "Alertas baja prioridad": int(alertas_por_prioridad.get("baja", 0)),
}

# =============================================================================
# FILTROS
# =============================================================================
st.divider()
st.subheader("Filtros")
col_f1, col_f2, col_f3 = st.columns(3)
with col_f1:
    fecha_min = fe["fecha"].min()
    fecha_max = fe["fecha"].max()
    if pd.notna(fecha_min) and pd.notna(fecha_max):
        rango_fechas = st.date_input("Periodo (facturas emitidas)", value=(fecha_min.date(), fecha_max.date()))
    else:
        rango_fechas = None
with col_f2:
    clientes_disp = ["(Todos)"] + sorted(fe["cliente"].dropna().unique().tolist())
    cliente_sel = st.selectbox("Cliente", clientes_disp)
with col_f3:
    estados_disp = ["(Todos)"] + sorted(fe_conc["estado_conciliacion"].dropna().unique().tolist())
    estado_sel = st.selectbox("Estado de conciliación (ingresos)", estados_disp)

fe_filtrado = fe_conc.copy()
if rango_fechas and len(rango_fechas) == 2:
    ini, fin = pd.to_datetime(rango_fechas[0]), pd.to_datetime(rango_fechas[1])
    fe_filtrado = fe_filtrado[(fe_filtrado["fecha"] >= ini) & (fe_filtrado["fecha"] <= fin)]
if cliente_sel != "(Todos)":
    fe_filtrado = fe_filtrado[fe_filtrado["cliente"] == cliente_sel]
if estado_sel != "(Todos)":
    fe_filtrado = fe_filtrado[fe_filtrado["estado_conciliacion"] == estado_sel]

# =============================================================================
# TABLERO EJECUTIVO
# =============================================================================
st.divider()
st.subheader("Tablero ejecutivo")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total facturado", f"${resumen['Total facturado']:,.2f}")
c2.metric("Total cobrado", f"${resumen['Total cobrado']:,.2f}")
c3.metric("Pendiente de cobro", f"${resumen['Total pendiente de cobro']:,.2f}")
c4.metric("% Conciliación", f"{resumen['Porcentaje de conciliación']}%")

c5, c6, c7, c8 = st.columns(4)
c5.metric("Gastos deducibles", f"${resumen['Gastos deducibles']:,.2f}")
c6.metric("Gastos no deducibles", f"${resumen['Gastos no deducibles']:,.2f}")
c7.metric("IVA neto", f"${resumen['IVA neto']:,.2f}")
c8.metric("ISR retenido total", f"${resumen['ISR retenido total']:,.2f}")

c9, c10, c11 = st.columns(3)
c9.metric("Alertas de prioridad alta", resumen["Alertas alta prioridad"])
c10.metric("Alertas de prioridad media", resumen["Alertas media prioridad"])
c11.metric("Alertas de prioridad baja", resumen["Alertas baja prioridad"])

# --- Gráficas ---
g1, g2 = st.columns(2)
with g1:
    fig1 = px.bar(
        x=["Facturado", "Cobrado", "Pendiente"],
        y=[resumen["Total facturado"], resumen["Total cobrado"], resumen["Total pendiente de cobro"]],
        title="Facturado vs. cobrado", labels={"x": "", "y": "Monto ($)"},
    )
    st.plotly_chart(fig1, width="stretch")
with g2:
    fig2 = px.pie(
        names=["Deducibles", "No deducibles"],
        values=[resumen["Gastos deducibles"], resumen["Gastos no deducibles"]],
        title="Gastos deducibles vs. no deducibles",
    )
    st.plotly_chart(fig2, width="stretch")

g3, g4 = st.columns(2)
with g3:
    fig3 = px.bar(
        x=["IVA trasladado", "IVA acreditable", "IVA neto"],
        y=[resultado_iva["iva_trasladado"], resultado_iva["iva_acreditable"], resultado_iva["iva_neto"]],
        title="Integración del IVA", labels={"x": "", "y": "Monto ($)"},
    )
    st.plotly_chart(fig3, width="stretch")
with g4:
    estado_counts = pd.concat([fe_conc["estado_conciliacion"], fr_conc["estado_conciliacion"]]).value_counts()
    fig4 = px.pie(names=estado_counts.index, values=estado_counts.values, title="Estado de la conciliación bancaria")
    st.plotly_chart(fig4, width="stretch")

if fe["fecha"].notna().any():
    evol = fe_norm.copy()
    evol["mes"] = evol["fecha"].dt.to_period("M").astype(str)
    ev_ingresos = evol.groupby("mes")["total"].sum().rename("Ingresos")
    evol_g = fr_norm.copy()
    evol_g["mes"] = evol_g["fecha"].dt.to_period("M").astype(str)
    ev_gastos = evol_g.groupby("mes")["total"].sum().rename("Gastos")
    ev_df = pd.concat([ev_ingresos, ev_gastos], axis=1).fillna(0).reset_index()
    fig5 = px.line(ev_df, x="mes", y=["Ingresos", "Gastos"], title="Evolución mensual de ingresos y gastos", markers=True)
    st.plotly_chart(fig5, width="stretch")

fig6 = px.bar(alertas["prioridad"].value_counts().reset_index(), x="prioridad", y="count",
              title="Alertas y excepciones por prioridad") if len(alertas) else None
if fig6:
    st.plotly_chart(fig6, width="stretch")

# =============================================================================
# TABLAS FILTRABLES
# =============================================================================
st.divider()
st.subheader("Detalle")

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "Ingresos", "Gastos", "IVA", "ISR", "Conciliación bancaria", "Alertas"
])

with tab1:
    st.markdown(f"**{len(fe_filtrado)}** facturas emitidas según filtros aplicados.")
    st.dataframe(fe_filtrado, width="stretch")
    if len(fe_nc):
        st.markdown("**Notas de crédito / devoluciones detectadas (facturas emitidas):**")
        st.dataframe(fe_nc, width="stretch")

with tab2:
    st.markdown(f"**{len(fr_conc)}** gastos procesados.")
    st.dataframe(fr_conc, width="stretch")
    if len(fr_nc):
        st.markdown("**Notas de crédito / devoluciones detectadas (gastos):**")
        st.dataframe(fr_nc, width="stretch")

with tab3:
    st.dataframe(pd.DataFrame(list(resultado_iva.items()), columns=["Concepto", "Monto"]), width="stretch")
    st.caption("Cálculo con base en flujo de efectivo: solo se considera IVA de facturas cobradas y gastos pagados y deducibles.")

with tab4:
    st.markdown(f"**${isr_total:,.2f}** en retenciones de ISR ({len(isr)} registros).")
    st.dataframe(isr_proc, width="stretch")
    st.caption(
        "Las retenciones de ISR no traen un folio o UUID que las ligue a una factura específica en los "
        "datos cargados, por lo que se muestran para revisión manual en vez de conciliarse automáticamente."
    )

with tab5:
    st.markdown("**Ingresos (facturas vs. banco):**")
    st.dataframe(fe_conc, width="stretch")
    st.markdown("**Egresos (gastos vs. banco):**")
    st.dataframe(fr_conc, width="stretch")

with tab6:
    if len(alertas):
        prioridades = ["(Todas)"] + sorted(alertas["prioridad"].unique().tolist())
        prioridad_sel = st.selectbox("Filtrar por prioridad", prioridades)
        alertas_mostrar = alertas if prioridad_sel == "(Todas)" else alertas[alertas["prioridad"] == prioridad_sel]
        st.dataframe(alertas_mostrar, width="stretch")
    else:
        st.success("No se generaron alertas.")

# =============================================================================
# DESCARGA
# =============================================================================
st.divider()
st.subheader("Reporte descargable")

excel_bytes = generar_reporte_excel(
    resumen, fe_conc, fr_conc, resultado_iva, isr_proc,
    ingresos_marcados, egresos_marcados, alertas, partidas_no_conciliadas,
    fe_nc, fr_nc,
)
st.download_button(
    "⬇️ Descargar reporte Excel completo",
    data=excel_bytes,
    file_name="calculadora_tx_reporte.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

if len(partidas_no_conciliadas):
    excepciones_buffer = io.BytesIO()
    partidas_no_conciliadas.to_excel(excepciones_buffer, index=False, engine="openpyxl")
    st.download_button(
        "⬇️ Descargar solo excepciones (no conciliadas) para revisión manual",
        data=excepciones_buffer.getvalue(),
        file_name="calculadora_tx_excepciones.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
