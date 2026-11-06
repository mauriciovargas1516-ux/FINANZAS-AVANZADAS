import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

st.set_page_config(page_title="Calculadora de Flujos - Finanzas", layout="wide")

st.title("📊 Generador de Flujos de Caja y Excel con Fórmulas")
st.caption("Genera automáticamente FCL, Tabla de Amortización (PAGO) y FCI en pestañas separadas con fórmulas nativas de Excel.")

# --- PLANTILLAS PRECARGADAS ---
plantilla = st.sidebar.selectbox(
    "📂 Cargar Plantilla Base",
    [
        "1. Taller Foto - Estructura J (Planta Piloto $120M)",
        "2. Guía Problema 2 (Planta Piloto $130M)",
        "3. Ejercicio Práctico - Mariscos en Conserva",
        "4. Personalizado (Desde cero)"
    ]
)

IVA_RATE = 0.19

if "1. Taller Foto" in plantilla:
    def_horizonte = 6
    def_tax = 29.5
    def_ke = 13.0
    def_kd = 6.5
    def_pct_deuda = 30.0
    def_plazo_deuda = 5
    def_gasto_credito = 0.0
    def_inv = pd.DataFrame([
        {"Ítem": "Máquina nueva", "Monto": 120000000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": True, "Vida Útil": 6, "Valor Rescate ($)": 45000000.0, "Intangible": False, "Capital Trabajo": False},
        {"Ítem": "Puesta en marcha (instalaciones)", "Monto": 12000000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Capital Trabajo": False},
        {"Ítem": "Pruebas de producción (ajustes)", "Monto": 7000000.0, "Tipo Moneda": "Bruto (dividir 1.19)", "Afecto IVA (Créd)": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Capital Trabajo": False},
        {"Ítem": "Capacitación mano de obra", "Monto": 7000000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Capital Trabajo": False},
        {"Ítem": "Capital de trabajo", "Monto": 36000000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Capital Trabajo": True},
    ])
    def_modo_op = "Por Tasas de Crecimiento Anual (%)"
    def_ing_a1 = 190000000.0
    def_ing_tipo = "Bruto (dividir 1.19)"
    def_cv_a1 = 38000.0 * (700.0 + 300.0 + 450.0) # 55.100.000
    def_cv_tipo = "Neto"
    def_cf_a1 = 5300000.0
    def_cf_tipo = "Bruto (dividir 1.19)"
    def_tasas = pd.DataFrame({
        "Año": [1, 2, 3, 4, 5, 6],
        "Crec. Ingresos (%)": [0.0, 4.5, 4.5, 4.5, 4.5, 4.5],
        "Crec. Costos Variables (%)": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
        "Crec. Costos Fijos (%)": [0.0, 6.0, 6.0, 3.0, 3.0, 3.0]
    })

elif "2. Guía Problema 2" in plantilla:
    def_horizonte = 6
    def_tax = 29.5
    def_ke = 13.0
    def_kd = 7.5
    def_pct_deuda = 30.0
    def_plazo_deuda = 5
    def_gasto_credito = 0.0
    def_inv = pd.DataFrame([
        {"Ítem": "Máquina nueva", "Monto": 130000000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": True, "Vida Útil": 6, "Valor Rescate ($)": 55000000.0, "Intangible": False, "Capital Trabajo": False},
        {"Ítem": "Gastos puesta en marcha", "Monto": 14000000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Capital Trabajo": False},
        {"Ítem": "Pruebas de producción", "Monto": 6000000.0, "Tipo Moneda": "Bruto (dividir 1.19)", "Afecto IVA (Créd)": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Capital Trabajo": False},
        {"Ítem": "Capacitación mano de obra", "Monto": 7000000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Capital Trabajo": False},
        {"Ítem": "Capital de trabajo", "Monto": 35000000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Capital Trabajo": True},
    ])
    def_modo_op = "Por Tasas de Crecimiento Anual (%)"
    def_ing_a1 = 190000000.0
    def_ing_tipo = "Bruto (dividir 1.19)"
    def_cv_a1 = 0.0
    def_cv_tipo = "Neto"
    def_cf_a1 = 60400000.0
    def_cf_tipo = "Neto"
    def_tasas = pd.DataFrame({
        "Año": [1, 2, 3, 4, 5, 6],
        "Crec. Ingresos (%)": [0.0, 4.5, 4.5, 4.5, 4.5, 4.5],
        "Crec. Costos Variables (%)": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
        "Crec. Costos Fijos (%)": [0.0, 6.0, 6.0, 3.0, 3.0, 3.0]
    })

elif "3. Ejercicio Práctico" in plantilla:
    def_horizonte = 5
    def_tax = 29.5
    def_ke = 13.0
    def_kd = 8.5
    def_pct_deuda = 30.0
    def_plazo_deuda = 4
    def_gasto_credito = 2.0
    def_inv = pd.DataFrame([
        {"Ítem": "Maquinaria y equipos", "Monto": 13000000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": True, "Vida Útil": 5, "Valor Rescate ($)": 3250000.0, "Intangible": False, "Capital Trabajo": False},
        {"Ítem": "Implementos de prod.", "Monto": 1800000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": True, "Vida Útil": 5, "Valor Rescate ($)": 450000.0, "Intangible": False, "Capital Trabajo": False},
        {"Ítem": "Terrenos", "Monto": 3000000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": False, "Depreciable": False, "Vida Útil": 5, "Valor Rescate ($)": 750000.0, "Intangible": False, "Capital Trabajo": False},
        {"Ítem": "Implementos oficina", "Monto": 2000000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": True, "Vida Útil": 5, "Valor Rescate ($)": 500000.0, "Intangible": False, "Capital Trabajo": False},
        {"Ítem": "Vehículo", "Monto": 8000000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": True, "Vida Útil": 5, "Valor Rescate ($)": 2000000.0, "Intangible": False, "Capital Trabajo": False},
        {"Ítem": "Construcción", "Monto": 7500000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": True, "Vida Útil": 5, "Valor Rescate ($)": 1875000.0, "Intangible": False, "Capital Trabajo": False},
        {"Ítem": "Gastos org. y puesta en marcha", "Monto": 1000000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": False, "Vida Útil": 5, "Valor Rescate ($)": 0.0, "Intangible": True, "Capital Trabajo": False},
        {"Ítem": "Capital de trabajo", "Monto": 1100000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": False, "Vida Útil": 5, "Valor Rescate ($)": 0.0, "Intangible": False, "Capital Trabajo": True},
    ])
    def_modo_op = "Tabla Manual Año a Año"
    def_ing_a1 = 38000000.0
    def_ing_tipo = "Neto"
    def_cv_a1 = 17100840.34
    def_cv_tipo = "Neto"
    def_cf_a1 = 7460000.0
    def_cf_tipo = "Neto"
    def_tasas = pd.DataFrame({
        "Año": [1, 2, 3, 4, 5],
        "Crec. Ingresos (%)": [0.0]*5,
        "Crec. Costos Variables (%)": [0.0]*5,
        "Crec. Costos Fijos (%)": [0.0]*5
    })
else:
    def_horizonte = 6
    def_tax = 29.5
    def_ke = 13.0
    def_kd = 6.5
    def_pct_deuda = 30.0
    def_plazo_deuda = 5
    def_gasto_credito = 0.0
    def_inv = pd.DataFrame([
        {"Ítem": "Activo Fijo Principal", "Monto": 100000000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": True, "Vida Útil": 6, "Valor Rescate ($)": 20000000.0, "Intangible": False, "Capital Trabajo": False},
        {"Ítem": "Capital de Trabajo", "Monto": 20000000.0, "Tipo Moneda": "Neto", "Afecto IVA (Créd)": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Capital Trabajo": True},
    ])
    def_modo_op = "Por Tasas de Crecimiento Anual (%)"
    def_ing_a1 = 150000000.0
    def_ing_tipo = "Neto"
    def_cv_a1 = 50000000.0
    def_cv_tipo = "Neto"
    def_cf_a1 = 10000000.0
    def_cf_tipo = "Neto"
    def_tasas = pd.DataFrame({
        "Año": [1, 2, 3, 4, 5, 6],
        "Crec. Ingresos (%)": [0.0, 5.0, 5.0, 5.0, 5.0, 5.0],
        "Crec. Costos Variables (%)": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
        "Crec. Costos Fijos (%)": [0.0, 3.0, 3.0, 3.0, 3.0, 3.0]
    })

# --- SIDEBAR: PARÁMETROS GENERALES ---
st.sidebar.header("⚙️ Parámetros Generales")
horizonte = st.sidebar.number_input("Horizonte de Evaluación (años)", min_value=1, max_value=15, value=def_horizonte)
tasa_imp = st.sidebar.number_input("Tasa de Impuesto (%)", value=def_tax, step=0.5) / 100.0
tasa_ke = st.sidebar.number_input("Tasa Costo de Capital Ke (%)", value=def_ke, step=0.5) / 100.0

st.sidebar.header("🏦 Financiamiento (Deuda)")
base_credito = st.sidebar.selectbox("Base para % del Crédito", ["Inversión Bruta (con IVA 19%)", "Inversión Neta (sin IVA)"])
pct_deuda = st.sidebar.number_input("% Financiado con Crédito", value=def_pct_deuda, step=5.0) / 100.0
tasa_kd = st.sidebar.number_input("Tasa de Interés Crédito Kd (%)", value=def_kd, step=0.25) / 100.0
plazo_deuda = st.sidebar.number_input("Plazo del Crédito (años)", min_value=1, max_value=15, value=def_plazo_deuda)
gasto_credito_pct = st.sidebar.number_input("Gastos asociados al crédito (%)", value=def_gasto_credito, step=0.5) / 100.0

st.sidebar.header("📐 Criterio Docente")
descuento_fcl = st.sidebar.radio("Descontar FCL usando:", ["Ke (Costo Capital directo)", "WACC (Costo Promedio Ponderado)"])
incluir_prestamo_fci0 = st.sidebar.checkbox("Sumar +Préstamo en Año 0 del FCI (Estándar diapositivas)", value=True)

wacc_calc = tasa_ke * (1 - pct_deuda) + tasa_kd * (1 - tasa_imp) * pct_deuda
st.sidebar.info(f"**WACC Calculado:** {wacc_calc*100:.3f}%")

# --- SECCIÓN 1: INVERSIONES ---
st.subheader("1️⃣ Inversiones Iniciales, Depreciación y Ajuste de IVA (Año 0)")
inv_df = st.data_editor(
    def_inv,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Tipo Moneda": st.column_config.SelectboxColumn("Tipo Moneda", options=["Neto", "Bruto (dividir 1.19)"]),
    }
)

# Procesar inversiones
inv_processed = []
for _, row in inv_df.iterrows():
    monto_raw = float(row["Monto"] or 0.0)
    es_bruto = "Bruto" in str(row["Tipo Moneda"])
    val_neto = monto_raw / (1.0 + IVA_RATE) if es_bruto else monto_raw
    afecto_iva = bool(row["Afecto IVA (Créd)"])
    val_iva = val_neto * IVA_RATE if afecto_iva else 0.0
    val_bruto = val_neto + val_iva
    vida = max(1, int(row["Vida Útil"] or horizonte))
    rescate = float(row["Valor Rescate ($)"] or 0.0)
    dep_anual = (val_neto - rescate) / vida if bool(row["Depreciable"]) else 0.0
    amort_anual = val_neto / vida if bool(row["Intangible"]) else 0.0
    inv_processed.append({
        "Ítem": row["Ítem"],
        "Valor Neto": val_neto,
        "IVA (19%)": val_iva,
        "Valor Bruto": val_bruto,
        "Valor Rescate": rescate,
        "Depreciación Anual": dep_anual,
        "Amortización Anual": amort_anual,
        "Vida Útil": vida,
        "Es CT": bool(row["Capital Trabajo"])
    })
df_inv_calc = pd.DataFrame(inv_processed)

total_inv_neta = df_inv_calc["Valor Neto"].sum()
total_inv_bruta = df_inv_calc["Valor Bruto"].sum()
total_rescate_af = df_inv_calc[~df_inv_calc["Es CT"]]["Valor Rescate"].sum()
total_recup_ct = df_inv_calc[df_inv_calc["Es CT"]]["Valor Neto"].sum()
dep_anual_total = df_inv_calc["Depreciación Anual"].sum()
amort_anual_total = df_inv_calc["Amortización Anual"].sum()

base_prestamo_val = total_inv_bruta if "Bruta" in base_credito else total_inv_neta
monto_credito = base_prestamo_val * pct_deuda * (1.0 + gasto_credito_pct)
cuota_anual = monto_credito * (tasa_kd * (1 + tasa_kd)**plazo_deuda) / ((1 + tasa_kd)**plazo_deuda - 1) if monto_credito > 0 else 0.0

c1, c2, c3, c4 = st.columns(4)
c1.metric("Inversión Total Neta", f"${total_inv_neta:,.0f}")
c2.metric("Inversión Total Bruta", f"${total_inv_bruta:,.0f}")
c3.metric("Monto Crédito a Solicitar", f"${monto_credito:,.0f}")
c4.metric("Cuota Anual (PAGO)", f"${cuota_anual:,.0f}")

# --- SECCIÓN 2: INGRESOS Y COSTOS OPERACIONALES ---
st.subheader("2️⃣ Ingresos y Costos Operacionales")
modo_op = st.radio("Método de ingreso de datos operativos:", ["Por Tasas de Crecimiento Anual (%)", "Calculadora Unitaria (Q × P) + Tasas"], horizontal=True)

if modo_op == "Calculadora Unitaria (Q × P) + Tasas":
    col_u1, col_u2, col_u3, col_u4 = st.columns(4)
    q_base = col_u1.number_input("Unidades Año 1 (Q)", value=38000.0, step=1000.0)
    p_base = col_u2.number_input("Precio Unitario ($)", value=5000.0, step=100.0)
    p_tipo = col_u3.selectbox("Precio está en:", ["Bruto (dividir 1.19)", "Neto"])
    cv_unit = col_u4.number_input("Costo Variable Unitario Neto ($)", value=1450.0, step=50.0)
    ing_base_neto = (q_base * p_base) / 1.19 if "Bruto" in p_tipo else (q_base * p_base)
    cv_base_neto = q_base * cv_unit
    col_f1, col_f2 = st.columns(2)
    cf_raw = col_f1.number_input("Costos Fijos Año 1 ($)", value=5300000.0, step=100000.0)
    cf_tipo_sel = col_f2.selectbox("Costos Fijos están en:", ["Bruto (dividir 1.19)", "Neto"])
    cf_base_neto = cf_raw / 1.19 if "Bruto" in cf_tipo_sel else cf_raw
else:
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        ing_raw = st.number_input("Ingresos Año 1 ($)", value=def_ing_a1, step=1000000.0)
        ing_t = st.selectbox("Ingresos Año 1 en:", ["Bruto (dividir 1.19)", "Neto"], index=0 if "Bruto" in def_ing_tipo else 1)
        ing_base_neto = ing_raw / 1.19 if "Bruto" in ing_t else ing_raw
    with col_b:
        cv_raw = st.number_input("Costos Variables Totales Año 1 ($)", value=def_cv_a1, step=500000.0)
        cv_t = st.selectbox("Costos Variables en:", ["Neto", "Bruto (dividir 1.19)"], index=0 if "Neto" in def_cv_tipo else 1)
        cv_base_neto = cv_raw / 1.19 if "Bruto" in cv_t else cv_raw
    with col_c:
        cf_raw = st.number_input("Costos Fijos / Egresos Op. Año 1 ($)", value=def_cf_a1, step=500000.0)
        cf_t = st.selectbox("Costos Fijos en:", ["Bruto (dividir 1.19)", "Neto"], index=0 if "Bruto" in def_cf_tipo else 1)
        cf_base_neto = cf_raw / 1.19 if "Bruto" in cf_t else cf_raw

# Ajustar longitud de tabla de tasas según horizonte
if len(def_tasas) != horizonte:
    tasas_init = pd.DataFrame({
        "Año": list(range(1, horizonte + 1)),
        "Crec. Ingresos (%)": [0.0] + [4.5]*(horizonte - 1),
        "Crec. Costos Variables (%)": [0.0]*horizonte,
        "Crec. Costos Fijos (%)": [0.0] + [6.0]*min(2, horizonte-1) + [3.0]*max(0, horizonte-3)
    })
else:
    tasas_init = def_tasas

st.markdown("**Tasas de Crecimiento Anual por Año (%)** *(Año 1 es base 0%)*:")
tasas_df = st.data_editor(tasas_init, use_container_width=True, hide_index=True)

# Calcular series operacionales
ing_series, cv_series, cf_series = [], [], []
curr_i, curr_v, curr_f = ing_base_neto, cv_base_neto, cf_base_neto
for idx, r_row in tasas_df.iterrows():
    if idx > 0:
        curr_i *= (1.0 + float(r_row["Crec. Ingresos (%)"]) / 100.0)
        curr_v *= (1.0 + float(r_row["Crec. Costos Variables (%)"]) / 100.0)
        curr_f *= (1.0 + float(r_row["Crec. Costos Fijos (%)"]) / 100.0)
    ing_series.append(curr_i)
    cv_series.append(curr_v)
    cf_series.append(curr_f)

# --- TABLA DE AMORTIZACIÓN ---
amort_rows = []
saldo = monto_credito
amort_rows.append({"Período": 0, "Cuota": 0.0, "Interés": 0.0, "Amortización": 0.0, "Saldo": saldo})
intereses_anuales = []
amort_deuda_anuales = []

for t in range(1, horizonte + 1):
    if t <= plazo_deuda and monto_credito > 0:
        interes = saldo * tasa_kd
        amort_cap = cuota_anual - interes
        saldo = max(0.0, saldo - amort_cap)
        c_val = cuota_anual
    else:
        interes, amort_cap, c_val = 0.0, 0.0, 0.0
    intereses_anuales.append(interes)
    amort_deuda_anuales.append(amort_cap)
    amort_rows.append({"Período": t, "Cuota": c_val, "Interés": interes, "Amortización": amort_cap, "Saldo": saldo})

df_amort = pd.DataFrame(amort_rows)

# --- FUNCIONES DE EVALUACIÓN ---
def calc_metrics(flujos, tasa_desc):
    f_act = [flujos[t] / ((1.0 + tasa_desc)**t) for t in range(len(flujos))]
    van = sum(f_act)
    # TIR
    roots = np.roots(flujos)
    real_roots = roots[np.isreal(roots)].real
    pos_roots = [r for r in real_roots if r > 0]
    tir = (min(pos_roots, key=lambda r: abs((r - 1) - 0.15)) - 1.0) if pos_roots else 0.0
    # Payback tradicional
    acum = np.cumsum(flujos)
    pb_str = "No recupera en el horizonte"
    for t in range(1, len(flujos)):
        if acum[t] >= 0:
            prev_acum = abs(acum[t-1])
            frac_ano = prev_acum / flujos[t] if flujos[t] > 0 else 0
            anos = (t - 1) + frac_ano
            meses = int(round((anos - int(anos)) * 12))
            pb_str = f"{anos:.2f} años ({int(anos)} años y {meses} meses)"
            break
    return van, tir, pb_str, f_act

tasa_fcl_usada = wacc_calc if "WACC" in descuento_fcl else tasa_ke

# Construir vectores FCL y FCI para vista previa
fcl_vec = [-total_inv_neta]
fci_vec = [-total_inv_neta + (monto_credito if incluir_prestamo_fci0 else 0.0)]

for t in range(horizonte):
    dep_t = dep_anual_total
    amort_int_t = amort_anual_total
    # FCL
    bai_fcl = ing_series[t] - cv_series[t] - cf_series[t] - dep_t - amort_int_t
    imp_fcl = -bai_fcl * tasa_imp
    bdi_fcl = bai_fcl + imp_fcl
    rescate_t = total_rescate_af if t == horizonte - 1 else 0.0
    rec_ct_t = total_recup_ct if t == horizonte - 1 else 0.0
    fcl_t = bdi_fcl + dep_t + amort_int_t + rescate_t + rec_ct_t
    fcl_vec.append(fcl_t)
    # FCI
    bai_fci = ing_series[t] - cv_series[t] - cf_series[t] - dep_t - amort_int_t - intereses_anuales[t]
    imp_fci = -bai_fci * tasa_imp
    bdi_fci = bai_fci + imp_fci
    fci_t = bdi_fci + dep_t + amort_int_t - amort_deuda_anuales[t] + rescate_t + rec_ct_t
    fci_vec.append(fci_t)

van_fcl, tir_fcl, pb_fcl, _ = calc_metrics(fcl_vec, tasa_fcl_usada)
van_fci, tir_fci, pb_fci, _ = calc_metrics(fci_vec, tasa_ke)

st.subheader("3️⃣ Resultados y Evaluación (Vista Previa)")
col_r1, col_r2 = st.columns(2)
with col_r1:
    st.markdown(f"#### 📈 Flujo de Caja Libre (Descontado al {tasa_fcl_usada*100:.2f}%)")
    st.metric("VAN FCL", f"${van_fcl:,.0f}")
    st.metric("TIR FCL", f"{tir_fcl*100:.2f}%")
    st.write(f"**Periodo de Recuperación (Payback):** {pb_fcl}")
with col_r2:
    st.markdown(f"#### 💼 Flujo del Inversionista (Descontado a Ke = {tasa_ke*100:.2f}%)")
    st.metric("VAN Inversionista", f"${van_fci:,.0f}")
    st.metric("TIR Inversionista", f"{tir_fci*100:.2f}%")
    st.write(f"**Periodo de Recuperación (Payback):** {pb_fci}")

# --- GENERADOR DE EXCEL CON FÓRMULAS NATIVAS ---
def generar_excel_formulas():
    wb = openpyxl.Workbook()
    font_title = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    fill_title = PatternFill("solid", fgColor="1F4E78")
    font_bold = Font(name="Calibri", size=10, bold=True)
    fill_sub = PatternFill("solid", fgColor="D9E1F2")
    fill_total = PatternFill("solid", fgColor="EDEDED")
    thin_border = Border(bottom=Side(style="thin", color="BFBFBF"), top=Side(style="thin", color="BFBFBF"))
    num_fmt_curr = "#,##0;(#,##0);\"-\""
    num_fmt_pct = "0.00%"

    # --- PESTAÑA 1: SUPUESTOS E INVERSIÓN ---
    ws1 = wb.active
    ws1.title = "1_Supuestos_Inversion"
    ws1["B2"] = "PARÁMETROS Y SUPUESTOS DEL PROYECTO"
    ws1["B2"].font = font_title
    ws1["B2"].fill = fill_title

    params = [
        ("Horizonte (años)", horizonte, "0"),
        ("Tasa de Impuesto", tasa_imp, num_fmt_pct),
        ("Tasa Costo de Capital (Ke)", tasa_ke, num_fmt_pct),
        ("Tasa Crédito Bancario (Kd)", tasa_kd, num_fmt_pct),
        ("% Financiamiento Deuda", pct_deuda, num_fmt_pct),
        ("Plazo del Crédito (años)", plazo_deuda, "0"),
        ("Tasa IVA", IVA_RATE, num_fmt_pct),
    ]
    for i, (label, val, fmt) in enumerate(params, start=3):
        ws1[f"B{i}"] = label
        ws1[f"C{i}"] = val
        ws1[f"C{i}"].number_format = fmt
    ws1["B10"] = "WACC (Costo Promedio Ponderado)"
    ws1["B10"].font = font_bold
    ws1["C10"] = "=C5*(1-C7)+C6*(1-C4)*C7"
    ws1["C10"].number_format = num_fmt_pct
    ws1["C10"].font = font_bold

    # Tabla de Inversiones, IVA y Depreciación
    ws1["B12"] = "DETALLE DE INVERSIONES, AJUSTE DE IVA Y DEPRECIACIÓN"
    ws1["B12"].font = font_title
    ws1["B12"].fill = fill_title

    headers_inv = ["Ítem de Inversión", "Valor Neto", "IVA (19%)", "Valor Bruto", "Vida Útil", "Valor Rescate", "Depreciación Anual", "Amortización Anual"]
    for col_idx, h in enumerate(headers_inv, start=2):
        cell = ws1.cell(13, col_idx, h)
        cell.font = font_bold
        cell.fill = fill_sub

    start_inv_row = 14
    for idx, r in df_inv_calc.iterrows():
        curr_r = start_inv_row + idx
        ws1.cell(curr_r, 2, r["Ítem"])
        c_neto = ws1.cell(curr_r, 3, r["Valor Neto"])
        c_neto.number_format = num_fmt_curr
        # IVA formula
        c_iva = ws1.cell(curr_r, 4, f"=C{curr_r}*$C$9" if r["IVA (19%)"] > 0 else 0)
        c_iva.number_format = num_fmt_curr
        # Bruto formula
        c_bruto = ws1.cell(curr_r, 5, f"=C{curr_r}+D{curr_r}")
        c_bruto.number_format = num_fmt_curr
        ws1.cell(curr_r, 6, r["Vida Útil"])
        c_resc = ws1.cell(curr_r, 7, r["Valor Rescate"])
        c_resc.number_format = num_fmt_curr
        c_dep = ws1.cell(curr_r, 8, f"=(C{curr_r}-G{curr_r})/F{curr_r}" if r["Depreciación Anual"] > 0 else 0)
        c_dep.number_format = num_fmt_curr
        c_amo = ws1.cell(curr_r, 9, f"=C{curr_r}/F{curr_r}" if r["Amortización Anual"] > 0 else 0)
        c_amo.number_format = num_fmt_curr

    tot_inv_row = start_inv_row + len(df_inv_calc)
    ws1.cell(tot_inv_row, 2, "TOTALES").font = font_bold
    for col_i in [3, 4, 5, 7, 8, 9]:
        col_let = get_column_letter(col_i)
        c_tot = ws1.cell(tot_inv_row, col_i, f"=SUM({col_let}{start_inv_row}:{col_let}{tot_inv_row-1})")
        c_tot.font = font_bold
        c_tot.fill = fill_total
        c_tot.number_format = num_fmt_curr

    #Monto préstamo en hoja 1
    loan_row = tot_inv_row + 2
    ws1.cell(loan_row, 2, "Monto del Préstamo a Solicitar").font = font_bold
    base_col_let = "E" if "Bruta" in base_credito else "C"
    ws1.cell(loan_row, 3, f"={base_col_let}{tot_inv_row}*C7*(1+{gasto_credito_pct})").number_format = num_fmt_curr
    ws1.cell(loan_row, 3).font = font_bold

    # --- PESTAÑA 3: TABLA DE AMORTIZACIÓN (Se crea para referenciar desde FCI) ---
    ws3 = wb.create_sheet("3_Tabla_Amortizacion")
    ws3["B2"] = " CÁLCULO DE CRÉDITO Y TABLA DE AMORTIZACIÓN"
    ws3["B2"].font = font_title
    ws3["B2"].fill = fill_title

    ws3["B4"] = "Monto del Crédito"
    ws3["C4"] = f"='1_Supuestos_Inversion'!C{loan_row}"
    ws3["C4"].number_format = num_fmt_curr
    ws3["B5"] = "Tasa de Interés Anual"
    ws3["C5"] = "='1_Supuestos_Inversion'!C6"
    ws3["C5"].number_format = num_fmt_pct
    ws3["B6"] = "Plazo (Años)"
    ws3["C6"] = "='1_Supuestos_Inversion'!C8"
    ws3["B7"] = "Cuota Anual Fija (Función PAGO)"
    ws3["B7"].font = font_bold
    ws3["C7"] = "=PMT(C5,C6,-C4)"
    ws3["C7"].font = font_bold
    ws3["C7"].number_format = num_fmt_curr

    headers_amo = ["Período", "Cuota", "Interés", "Amortización", "Saldo Insoluto"]
    for idx, h in enumerate(headers_amo, start=2):
        c = ws3.cell(9, idx, h)
        c.font = font_bold
        c.fill = fill_sub

    ws3.cell(10, 2, 0)
    c_s0 = ws3.cell(10, 6, "=C4")
    c_s0.number_format = num_fmt_curr

    for p in range(1, horizonte + 1):
        r_idx = 10 + p
        ws3.cell(r_idx, 2, p)
        if p <= plazo_deuda:
            c_cuota = ws3.cell(r_idx, 3, "=$C$7")
            c_int = ws3.cell(r_idx, 4, f"=F{r_idx-1}*$C$5")
            c_am = ws3.cell(r_idx, 5, f"=C{r_idx}-D{r_idx}")
            c_sal = ws3.cell(r_idx, 6, f"=F{r_idx-1}-E{r_idx}")
        else:
            c_cuota = ws3.cell(r_idx, 3, 0)
            c_int = ws3.cell(r_idx, 4, 0)
            c_am = ws3.cell(r_idx, 5, 0)
            c_sal = ws3.cell(r_idx, 6, 0)
        for c_item in [c_cuota, c_int, c_am, c_sal]:
            c_item.number_format = num_fmt_curr

    # --- FUNCIÓN PARA CREAR HOJAS DE FLUJO (FCL Y FCI) ---
    def poblar_hoja_flujo(ws, es_inversionista=False):
        titulo = "FLUJO DE CAJA DEL INVERSIONISTA (FCI)" if es_inversionista else "FLUJO DE CAJA LIBRE (FCL)"
        ws["B2"] = titulo
        ws["B2"].font = font_title
        ws["B2"].fill = fill_title

        ws.cell(4, 2, "Concepto / Año").font = font_bold
        ws.cell(4, 2).fill = fill_sub
        for t in range(horizonte + 1):
            c = ws.cell(4, 3 + t, t)
            c.font = font_bold
            c.fill = fill_sub
            c.alignment = Alignment(horizontal="right")

        # Filas del Flujo
        filas_labels = [
            (5, "(+) Ingresos por Ventas"),
            (6, "(-) Costos Variables"),
            (7, "(-) Costos Fijos / Egresos Op."),
            (8, "(-) Depreciación"),
            (9, "(-) Amortización Intangibles"),
            (10, "(-) Intereses del Préstamo" if es_inversionista else "(-) Intereses (No aplica en FCL)"),
            (11, "(=) Utilidad Antes de Impuestos (BAI)"),
            (12, "(-) Impuesto a la Renta (29,5%)"),
            (13, "(=) Utilidad Después de Impuestos (BDI)"),
            (14, "(+) Depreciación"),
            (15, "(+) Amortización Intangibles"),
            (16, "(-) Inversión Inicial (Activos + Gastos + CT)"),
            (17, "(+) Préstamo Recibido" if es_inversionista else "(+) Préstamo (No aplica en FCL)"),
            (18, "(-) Amortización de la Deuda" if es_inversionista else "(-) Amortización Deuda (No aplica en FCL)"),
            (19, "(+) Valor de Rescate Activos Fijos"),
            (20, "(+) Recuperación Capital de Trabajo"),
            (21, "(=) FLUJO DE CAJA NETO"),
            (22, "Factor de Actualización"),
            (23, "Flujo de Caja Actualizado (VP)"),
            (24, "Flujo de Caja Acumulado (Para Payback)")
        ]
        for r_num, lbl in filas_labels:
            cell_lbl = ws.cell(r_num, 2, lbl)
            if r_num in [11, 13, 21]:
                cell_lbl.font = font_bold
                cell_lbl.fill = fill_total

        # Año 0 (Columna C = 3)
        ws.cell(16, 3, f"=-'1_Supuestos_Inversion'!C{tot_inv_row}").number_format = num_fmt_curr
        if es_inversionista and incluir_prestamo_fci0:
            ws.cell(17, 3, "='3_Tabla_Amortizacion'!C4").number_format = num_fmt_curr
        else:
            ws.cell(17, 3, 0).number_format = num_fmt_curr
        ws.cell(21, 3, "=SUM(C13:C20)").font = font_bold
        ws.cell(21, 3).fill = fill_total
        ws.cell(21, 3).number_format = num_fmt_curr

        tasa_ref = "'1_Supuestos_Inversion'!$C$5" if (es_inversionista or "Ke" in descuento_fcl) else "'1_Supuestos_Inversion'!$C$10"
        ws.cell(22, 3, f"=(1/(1+{tasa_ref}))^C4").number_format = "0.0000"
        ws.cell(23, 3, "=C21*C22").number_format = num_fmt_curr
        ws.cell(24, 3, "=C21").number_format = num_fmt_curr

        # Años 1 a N
        for t in range(1, horizonte + 1):
            col_i = 3 + t
            col_let = get_column_letter(col_i)
            prev_let = get_column_letter(col_i - 1)
            g_ing = float(tasas_df.iloc[t-1]["Crec. Ingresos (%)"]) / 100.0
            g_cv = float(tasas_df.iloc[t-1]["Crec. Costos Variables (%)"]) / 100.0
            g_cf = float(tasas_df.iloc[t-1]["Crec. Costos Fijos (%)"]) / 100.0

            if t == 1:
                ws.cell(5, col_i, ing_series[0])
                ws.cell(6, col_i, -cv_series[0])
                ws.cell(7, col_i, -cf_series[0])
            else:
                ws.cell(5, col_i, f"={prev_let}5*(1+{g_ing})")
                ws.cell(6, col_i, f"={prev_let}6*(1+{g_cv})")
                ws.cell(7, col_i, f"={prev_let}7*(1+{g_cf})")

            ws.cell(8, col_i, f"=-'1_Supuestos_Inversion'!$H${tot_inv_row}")
            ws.cell(9, col_i, f"=-'1_Supuestos_Inversion'!$I${tot_inv_row}")
            if es_inversionista:
                ws.cell(10, col_i, f"=-'3_Tabla_Amortizacion'!D{10+t}")
            else:
                ws.cell(10, col_i, 0)

            ws.cell(11, col_i, f"=SUM({col_let}5:{col_let}10)").font = font_bold
            ws.cell(11, col_i).fill = fill_total
            ws.cell(12, col_i, f"=-{col_let}11*'1_Supuestos_Inversion'!$C$4")
            ws.cell(13, col_i, f"=SUM({col_let}11:{col_let}12)").font = font_bold
            ws.cell(13, col_i).fill = fill_total
            ws.cell(14, col_i, f"=-{col_let}8")
            ws.cell(15, col_i, f"=-{col_let}9")
            ws.cell(16, col_i, 0)
            ws.cell(17, col_i, 0)
            if es_inversionista:
                ws.cell(18, col_i, f"=-'3_Tabla_Amortizacion'!E{10+t}")
            else:
                ws.cell(18, col_i, 0)

            if t == horizonte:
                ws.cell(19, col_i, total_rescate_af)
                ws.cell(20, col_i, total_recup_ct)
            else:
                ws.cell(19, col_i, 0)
                ws.cell(20, col_i, 0)

            ws.cell(21, col_i, f"=SUM({col_let}13:{col_let}20)").font = font_bold
            ws.cell(21, col_i).fill = fill_total
            ws.cell(22, col_i, f"=(1/(1+{tasa_ref}))^{col_let}4").number_format = "0.0000"
            ws.cell(23, col_i, f"={col_let}21*{col_let}22")
            ws.cell(24, col_i, f"={prev_let}24+{col_let}21")

            for r_curr in range(5, 22):
                ws.cell(r_curr, col_i).number_format = num_fmt_curr
            ws.cell(23, col_i).number_format = num_fmt_curr
            ws.cell(24, col_i).number_format = num_fmt_curr

        last_let = get_column_letter(3 + horizonte)
        ws.cell(26, 2, "VAN (Suma Flujos Actualizados)").font = font_bold
        ws.cell(26, 3, f"=SUM(C23:{last_let}23)").font = font_bold
        ws.cell(26, 3).number_format = num_fmt_curr

        ws.cell(27, 2, "VAN (Con Función VNA / NPV)").font = font_bold
        ws.cell(27, 3, f"=C21+NPV({tasa_ref},D21:{last_let}21)").font = font_bold
        ws.cell(27, 3).number_format = num_fmt_curr

        ws.cell(28, 2, "TIR (Con Función TIR / IRR)").font = font_bold
        ws.cell(28, 3, f"=IRR(C21:{last_let}21)").font = font_bold
        ws.cell(28, 3).number_format = num_fmt_pct

        ws.cell(29, 2, "Periodo de Recuperación (Payback)").font = font_bold
        ws.cell(29, 3, pb_fci if es_inversionista else pb_fcl).font = font_bold

    ws2 = wb.create_sheet("2_Flujo_Caja_Libre")
    poblar_hoja_flujo(ws2, es_inversionista=False)

    ws4 = wb.create_sheet("4_Flujo_Inversionista")
    poblar_hoja_flujo(ws4, es_inversionista=True)

    # Ordenar pestañas: 1, 2, 3, 4
    wb._sheets = [ws1, ws2, ws3, ws4]

    for sheet in wb.worksheets:
        sheet.column_dimensions["B"].width = 36
        for col_i in range(3, 15):
            sheet.column_dimensions[get_column_letter(col_i)].width = 16

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

st.divider()
excel_bytes = generar_excel_formulas()
st.download_button(
    label="📥 DESCARGAR EXCEL CON FÓRMULAS NATIVAS (.XLSX)",
    data=excel_bytes,
    file_name="Taller_Flujos_de_Caja_Formulas.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True
)
