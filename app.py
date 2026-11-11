import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

st.set_page_config(
    page_title="Libro1.xlsx - Excel",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Camuflaje CSS Verde Excel Online + Ocultamiento de insignias de Streamlit Cloud
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    .stDeployButton, [data-testid="stDeployButton"] {display: none !important;}
    [data-testid="stStatusWidget"] {display: none !important;}
    [data-testid="stHeader"] {display: none !important;}
    [data-testid="stToolbar"] {display: none !important;}
    [data-testid="stDecoration"] {display: none !important;}
    [class*="viewerBadge"], [class*="_profileContainer"], [class*="_profilePreview"] {display: none !important;}
    iframe[height="0"] {display: none !important;}
    .block-container {
        padding-top: 0.1rem !important;
        padding-bottom: 1rem !important;
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
        max-width: 100% !important;
    }
    .excel-top {
        background-color: #107c41;
        color: white;
        font-family: 'Segoe UI', Calibri, sans-serif;
        font-size: 12px;
        padding: 5px 12px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .excel-ribbon {
        font-family: 'Segoe UI', Calibri, sans-serif;
        font-size: 12px;
        color: #333;
        background-color: #f3f2f1;
        border-bottom: 1px solid #e1dfdd;
        padding: 4px 12px;
        display: flex;
        gap: 18px;
    }
    .excel-toolbar {
        font-family: 'Segoe UI', Calibri, sans-serif;
        font-size: 11.5px;
        color: #444;
        background-color: #faf9f8;
        border-bottom: 1px solid #d2d0ce;
        padding: 4px 12px;
        display: flex;
        gap: 14px;
        align-items: center;
    }
    .excel-box {
        border: 1px solid #c8c6c4;
        background: white;
        padding: 1px 8px;
        font-size: 11px;
    }
    .excel-fx {
        font-family: Consolas, monospace;
        font-size: 12px;
        background-color: #ffffff;
        border: 1px solid #d2d0ce;
        padding: 3px 10px;
        margin-bottom: 6px;
        color: #333;
    }
    button[data-baseweb="tab"] {
        font-family: 'Segoe UI', Calibri, sans-serif !important;
        font-size: 12.5px !important;
        padding: 4px 14px !important;
        color: #444444 !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #107c41 !important;
        font-weight: 600 !important;
    }
    div[data-baseweb="tab-highlight"] {
        background-color: #107c41 !important;
    }
    div[data-testid="stExpander"] summary p {
        font-family: 'Segoe UI', Calibri, sans-serif !important;
        font-size: 12.5px !important;
        font-weight: 400 !important;
    }
    .nota-celda {
        font-family: Calibri, sans-serif;
        font-size: 12.5px;
        color: #333;
        border: 1px solid #d2d0ce;
        padding: 8px 12px;
        margin-top: 6px;
        background-color: #faf9f8;
    }
</style>
<div class="excel-top">
    <span><b>Excel</b> &nbsp;&nbsp;|&nbsp;&nbsp; Libro1.xlsx &nbsp; <span style="opacity:0.85; font-size:11px;">✓ Guardado en OneDrive</span></span>
    <span>Compartir &nbsp;&nbsp; Comentarios</span>
</div>
<div class="excel-ribbon">
    <span>Archivo</span>
    <span style="border-bottom: 2px solid #107c41; font-weight: 600; color: #107c41;">Inicio</span>
    <span>Insertar</span>
    <span>Disposición de página</span>
    <span>Fórmulas</span>
    <span>Datos</span>
    <span>Revisar</span>
    <span>Vista</span>
</div>
<div class="excel-toolbar">
    <span>Deshacer</span> |
    <span class="excel-box">Calibri &nbsp; ▾</span>
    <span class="excel-box">11 &nbsp; ▾</span> |
    <span><b>N</b> &nbsp; <i>K</i> &nbsp; <u>S</u></span> |
    <span>Alinear ▾ &nbsp; Ajustar texto</span> |
    <span class="excel-box">Contabilidad &nbsp; ▾</span>
    <span><b>$</b> &nbsp; <b>%</b> &nbsp; <b>000</b></span> |
    <span>Autosuma ∑</span>
</div>
<div class="excel-fx">
    <b>C21</b> &nbsp;&nbsp;|&nbsp;&nbsp; <i>fx</i> &nbsp; =SUMA(C13:C20)
</div>
""", unsafe_allow_html=True)

IVA_RATE = 0.19

plantilla = st.sidebar.selectbox(
    "Plantilla",
    [
        "Caso 1 - Estructura J (120M)",
        "Caso 2 - Guia P2 (130M)",
        "Caso 3 - Conservas (13M)",
        "Caso 4 - En blanco"
    ]
)

if "Caso 1" in plantilla:
    def_params = pd.DataFrame([{
        "Horizonte (años)": 6, "Impuesto (%)": 29.5, "Ke (%)": 13.0,
        "% Deuda": 30.0, "Kd (%)": 6.5, "Plazo Créd.": 5, "Gasto Créd. (%)": 0.0
    }])
    def_inv = pd.DataFrame([
        {"Ítem": "Máquina nueva", "Monto": 120000000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": True, "Vida Útil": 6, "Valor Rescate ($)": 45000000.0, "Intangible": False, "Cap. Trabajo": False},
        {"Ítem": "Puesta en marcha (instalaciones)", "Monto": 12000000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Cap. Trabajo": False},
        {"Ítem": "Pruebas de producción (ajustes)", "Monto": 7000000.0, "Moneda": "Bruto (/1.19)", "IVA Créd": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Cap. Trabajo": False},
        {"Ítem": "Capacitación mano de obra", "Monto": 7000000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Cap. Trabajo": False},
        {"Ítem": "Capital de trabajo", "Monto": 36000000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Cap. Trabajo": True},
    ])
    def_op_base = pd.DataFrame([{
        "Unidades A1 (opcional)": 38000.0, "Precio A1 ($)": 5000.0, "Ingresos A1 ($)": 190000000.0, "Moneda Ing.": "Bruto (/1.19)",
        "CV Total A1 ($)": 55100000.0, "Moneda CV": "Neto", "CF Total A1 ($)": 5300000.0, "Moneda CF": "Bruto (/1.19)"
    }])
    def_tasas = pd.DataFrame({
        "Año": [1, 2, 3, 4, 5, 6],
        "Crec. Ingresos (%)": [0.0, 4.5, 4.5, 4.5, 4.5, 4.5],
        "Crec. Costos Var. (%)": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
        "Crec. Costos Fijos (%)": [0.0, 6.0, 6.0, 3.0, 3.0, 3.0]
    })

elif "Caso 2" in plantilla:
    def_params = pd.DataFrame([{
        "Horizonte (años)": 6, "Impuesto (%)": 29.5, "Ke (%)": 13.0,
        "% Deuda": 30.0, "Kd (%)": 7.5, "Plazo Créd.": 5, "Gasto Créd. (%)": 0.0
    }])
    def_inv = pd.DataFrame([
        {"Ítem": "Máquina nueva", "Monto": 130000000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": True, "Vida Útil": 6, "Valor Rescate ($)": 55000000.0, "Intangible": False, "Cap. Trabajo": False},
        {"Ítem": "Gastos puesta en marcha", "Monto": 14000000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Cap. Trabajo": False},
        {"Ítem": "Pruebas de producción", "Monto": 6000000.0, "Moneda": "Bruto (/1.19)", "IVA Créd": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Cap. Trabajo": False},
        {"Ítem": "Capacitación mano de obra", "Monto": 7000000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Cap. Trabajo": False},
        {"Ítem": "Capital de trabajo", "Monto": 35000000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Cap. Trabajo": True},
    ])
    def_op_base = pd.DataFrame([{
        "Unidades A1 (opcional)": 1.0, "Precio A1 ($)": 190000000.0, "Ingresos A1 ($)": 190000000.0, "Moneda Ing.": "Bruto (/1.19)",
        "CV Total A1 ($)": 0.0, "Moneda CV": "Neto", "CF Total A1 ($)": 60400000.0, "Moneda CF": "Neto"
    }])
    def_tasas = pd.DataFrame({
        "Año": [1, 2, 3, 4, 5, 6],
        "Crec. Ingresos (%)": [0.0, 4.5, 4.5, 4.5, 4.5, 4.5],
        "Crec. Costos Var. (%)": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
        "Crec. Costos Fijos (%)": [0.0, 6.0, 6.0, 3.0, 3.0, 3.0]
    })

elif "Caso 3" in plantilla:
    def_params = pd.DataFrame([{
        "Horizonte (años)": 5, "Impuesto (%)": 29.5, "Ke (%)": 13.0,
        "% Deuda": 30.0, "Kd (%)": 8.5, "Plazo Créd.": 4, "Gasto Créd. (%)": 2.0
    }])
    def_inv = pd.DataFrame([
        {"Ítem": "Maquinaria y equipos", "Monto": 13000000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": True, "Vida Útil": 5, "Valor Rescate ($)": 3250000.0, "Intangible": False, "Cap. Trabajo": False},
        {"Ítem": "Implementos de prod.", "Monto": 1800000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": True, "Vida Útil": 5, "Valor Rescate ($)": 450000.0, "Intangible": False, "Cap. Trabajo": False},
        {"Ítem": "Terrenos", "Monto": 3000000.0, "Moneda": "Neto", "IVA Créd": False, "Depreciable": False, "Vida Útil": 5, "Valor Rescate ($)": 750000.0, "Intangible": False, "Cap. Trabajo": False},
        {"Ítem": "Implementos oficina", "Monto": 2000000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": True, "Vida Útil": 5, "Valor Rescate ($)": 500000.0, "Intangible": False, "Cap. Trabajo": False},
        {"Ítem": "Vehículo", "Monto": 8000000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": True, "Vida Útil": 5, "Valor Rescate ($)": 2000000.0, "Intangible": False, "Cap. Trabajo": False},
        {"Ítem": "Construcción", "Monto": 7500000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": True, "Vida Útil": 5, "Valor Rescate ($)": 1875000.0, "Intangible": False, "Cap. Trabajo": False},
        {"Ítem": "Gastos org. y puesta en marcha", "Monto": 1000000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": False, "Vida Útil": 5, "Valor Rescate ($)": 0.0, "Intangible": True, "Cap. Trabajo": False},
        {"Ítem": "Capital de trabajo", "Monto": 1100000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": False, "Vida Útil": 5, "Valor Rescate ($)": 0.0, "Intangible": False, "Cap. Trabajo": True},
    ])
    def_op_base = pd.DataFrame([{
        "Unidades A1 (opcional)": 50000.0, "Precio A1 ($)": 760.0, "Ingresos A1 ($)": 38000000.0, "Moneda Ing.": "Neto",
        "CV Total A1 ($)": 17100840.34, "Moneda CV": "Neto", "CF Total A1 ($)": 7460000.0, "Moneda CF": "Neto"
    }])
    def_tasas = pd.DataFrame({
        "Año": [1, 2, 3, 4, 5],
        "Crec. Ingresos (%)": [0.0, 20.0, 13.3333, 2.9412, 2.8571],
        "Crec. Costos Var. (%)": [0.0, 20.0, 13.3333, 2.9412, 2.8571],
        "Crec. Costos Fijos (%)": [0.0, 1.0724, -5.0398, 20.3911, 0.0]
    })
else:
    def_params = pd.DataFrame([{
        "Horizonte (años)": 6, "Impuesto (%)": 29.5, "Ke (%)": 13.0,
        "% Deuda": 30.0, "Kd (%)": 6.5, "Plazo Créd.": 5, "Gasto Créd. (%)": 0.0
    }])
    def_inv = pd.DataFrame([
        {"Ítem": "Activo Principal", "Monto": 100000000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": True, "Vida Útil": 6, "Valor Rescate ($)": 20000000.0, "Intangible": False, "Cap. Trabajo": False},
        {"Ítem": "Capital de Trabajo", "Monto": 20000000.0, "Moneda": "Neto", "IVA Créd": True, "Depreciable": False, "Vida Útil": 6, "Valor Rescate ($)": 0.0, "Intangible": False, "Cap. Trabajo": True},
    ])
    def_op_base = pd.DataFrame([{
        "Unidades A1 (opcional)": 10000.0, "Precio A1 ($)": 15000.0, "Ingresos A1 ($)": 150000000.0, "Moneda Ing.": "Neto",
        "CV Total A1 ($)": 50000000.0, "Moneda CV": "Neto", "CF Total A1 ($)": 10000000.0, "Moneda CF": "Neto"
    }])
    def_tasas = pd.DataFrame({
        "Año": [1, 2, 3, 4, 5, 6],
        "Crec. Ingresos (%)": [0.0, 5.0, 5.0, 5.0, 5.0, 5.0],
        "Crec. Costos Var. (%)": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
        "Crec. Costos Fijos (%)": [0.0, 3.0, 3.0, 3.0, 3.0, 3.0]
    })

base_credito = st.sidebar.selectbox("Base Crédito", ["Inv. Bruta", "Inv. Neta"])
descuento_fcl = st.sidebar.selectbox("Tasa Descuento FCL", ["Ke", "WACC"])
incluir_prestamo_fci0 = st.sidebar.checkbox("Sumar Préstamo en Año 0 FCI", value=True)
usar_q_por_p = st.sidebar.checkbox("Calcular Ingresos A1 como Unidades x Precio", value=False)
restar_rescate_dep = st.sidebar.checkbox("Restar Valor Rescate al depreciar", value=True)

with st.sidebar.expander("Calculadora CAPM (Opcional)"):
    rf_val = st.number_input("Tasa Libre Riesgo Rf (%)", value=6.0, step=0.5)
    beta_val = st.number_input("Beta", value=1.25, step=0.05)
    rm_val = st.number_input("Retorno Mercado Rm (%)", value=12.0, step=0.5)
    capm_res = rf_val + beta_val * (rm_val - rf_val)
    st.caption(f"Ke (CAPM) = {capm_res:.2f}%")

with st.expander("Celdas de Entrada y Supuestos (Clic para mostrar/ocultar)", expanded=False):
    st.caption("1. Parámetros Financieros")
    params_df = st.data_editor(def_params, use_container_width=True, hide_index=True)

    horizonte = max(1, int(params_df.iloc[0]["Horizonte (años)"]))
    tasa_imp = float(params_df.iloc[0]["Impuesto (%)"]) / 100.0
    tasa_ke = float(params_df.iloc[0]["Ke (%)"]) / 100.0
    pct_deuda = float(params_df.iloc[0]["% Deuda"]) / 100.0
    tasa_kd = float(params_df.iloc[0]["Kd (%)"]) / 100.0
    plazo_deuda = max(1, int(params_df.iloc[0]["Plazo Créd."]))
    gasto_credito_pct = float(params_df.iloc[0]["Gasto Créd. (%)"]) / 100.0
    wacc_calc = tasa_ke * (1 - pct_deuda) + tasa_kd * (1 - tasa_imp) * pct_deuda

    st.caption("2. Inversiones Año 0, Valor de Rescate y Capital de Trabajo")
    inv_df = st.data_editor(
        def_inv,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Moneda": st.column_config.SelectboxColumn("Moneda", options=["Neto", "Bruto (/1.19)"]),
        }
    )

    st.caption("3. Operación Año 1 y Tasas de Crecimiento Anual (%)")
    op_df = st.data_editor(
        def_op_base,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Moneda Ing.": st.column_config.SelectboxColumn("Moneda Ing.", options=["Bruto (/1.19)", "Neto"]),
            "Moneda CV": st.column_config.SelectboxColumn("Moneda CV", options=["Neto", "Bruto (/1.19)"]),
            "Moneda CF": st.column_config.SelectboxColumn("Moneda CF", options=["Bruto (/1.19)", "Neto"]),
        }
    )

    if len(def_tasas) != horizonte:
        tasas_init = pd.DataFrame({
            "Año": list(range(1, horizonte + 1)),
            "Crec. Ingresos (%)": [0.0] + [4.5]*(horizonte - 1),
            "Crec. Costos Var. (%)": [0.0]*horizonte,
            "Crec. Costos Fijos (%)": [0.0] + [6.0]*min(2, horizonte-1) + [3.0]*max(0, horizonte-3)
        })
    else:
        tasas_init = def_tasas

    tasas_df = st.data_editor(tasas_init, use_container_width=True, hide_index=True)

r_op = op_df.iloc[0]
ing_bruto_o_neto = (float(r_op["Unidades A1 (opcional)"]) * float(r_op["Precio A1 ($)"])) if usar_q_por_p else float(r_op["Ingresos A1 ($)"])
ing_base_neto = ing_bruto_o_neto / 1.19 if "Bruto" in str(r_op["Moneda Ing."]) else ing_bruto_o_neto
cv_base_neto = float(r_op["CV Total A1 ($)"]) / 1.19 if "Bruto" in str(r_op["Moneda CV"]) else float(r_op["CV Total A1 ($)"])
cf_base_neto = float(r_op["CF Total A1 ($)"]) / 1.19 if "Bruto" in str(r_op["Moneda CF"]) else float(r_op["CF Total A1 ($)"])

inv_processed = []
for _, row in inv_df.iterrows():
    monto_raw = float(row["Monto"] or 0.0)
    es_bruto = "Bruto" in str(row["Moneda"])
    val_neto = monto_raw / (1.0 + IVA_RATE) if es_bruto else monto_raw
    afecto_iva = bool(row["IVA Créd"])
    val_iva = val_neto * IVA_RATE if afecto_iva else 0.0
    val_bruto = val_neto + val_iva
    vida = max(1, int(row["Vida Útil"] or horizonte))
    rescate = float(row["Valor Rescate ($)"] or 0.0)
    base_dep = (val_neto - rescate) if restar_rescate_dep else val_neto
    dep_anual = base_dep / vida if bool(row["Depreciable"]) else 0.0
    amort_anual = val_neto / vida if bool(row["Intangible"]) else 0.0
    inv_processed.append({
        "Ítem": str(row["Ítem"]),
        "Valor Neto": float(val_neto),
        "IVA (19%)": float(val_iva),
        "Valor Bruto": float(val_bruto),
        "Valor Rescate": float(rescate),
        "Depreciación Anual": float(dep_anual),
        "Amortización Anual": float(amort_anual),
        "Vida Útil": int(vida),
        "Es CT": bool(row["Cap. Trabajo"])
    })
df_inv_calc = pd.DataFrame(inv_processed)

total_inv_neta = float(df_inv_calc["Valor Neto"].sum())
total_inv_bruta = float(df_inv_calc["Valor Bruto"].sum())
total_rescate_af = float(df_inv_calc[~df_inv_calc["Es CT"]]["Valor Rescate"].sum())
total_recup_ct = float(df_inv_calc[df_inv_calc["Es CT"]]["Valor Neto"].sum())
dep_anual_total = float(df_inv_calc["Depreciación Anual"].sum())
amort_anual_total = float(df_inv_calc["Amortización Anual"].sum())

base_prestamo_val = total_inv_bruta if "Bruta" in base_credito else total_inv_neta
monto_credito = base_prestamo_val * pct_deuda * (1.0 + gasto_credito_pct)
cuota_anual = monto_credito * (tasa_kd * (1 + tasa_kd)**plazo_deuda) / ((1 + tasa_kd)**plazo_deuda - 1) if monto_credito > 0 else 0.0

ing_series, cv_series, cf_series = [], [], []
curr_i, curr_v, curr_f = ing_base_neto, cv_base_neto, cf_base_neto
for idx, r_row in tasas_df.iterrows():
    if idx > 0:
        curr_i *= (1.0 + float(r_row["Crec. Ingresos (%)"]) / 100.0)
        curr_v *= (1.0 + float(r_row["Crec. Costos Var. (%)"]) / 100.0)
        curr_f *= (1.0 + float(r_row["Crec. Costos Fijos (%)"]) / 100.0)
    ing_series.append(float(curr_i))
    cv_series.append(float(curr_v))
    cf_series.append(float(curr_f))

amort_rows = []
saldo = monto_credito
amort_rows.append({"Período": 0, "Cuota": 0.0, "Interés": 0.0, "Amortización": 0.0, "Saldo": saldo})
intereses_anuales, amort_deuda_anuales = [], []

for t in range(1, horizonte + 1):
    if t <= plazo_deuda and monto_credito > 0:
        interes = saldo * tasa_kd
        amort_cap = cuota_anual - interes
        saldo = max(0.0, saldo - amort_cap)
        c_val = cuota_anual
    else:
        interes, amort_cap, c_val = 0.0, 0.0, 0.0
    intereses_anuales.append(float(interes))
    amort_deuda_anuales.append(float(amort_cap))
    amort_rows.append({"Período": t, "Cuota": round(c_val), "Interés": round(interes), "Amortización": round(amort_cap), "Saldo": round(saldo)})

df_amort = pd.DataFrame(amort_rows)

def calc_metrics(flujos, tasa_desc):
    f_act = [flujos[t] / ((1.0 + tasa_desc)**t) for t in range(len(flujos))]
    van = sum(f_act)
    roots = np.roots(flujos)
    real_roots = roots[np.isreal(roots)].real
    pos_roots = [r for r in real_roots if r > 0]
    tir = (min(pos_roots, key=lambda r: abs((r - 1) - 0.15)) - 1.0) if pos_roots else 0.0
    acum = np.cumsum(flujos)
    pb_str = "No recupera"
    for t in range(1, len(flujos)):
        if acum[t] >= 0:
            prev_acum = abs(acum[t-1])
            frac_ano = prev_acum / flujos[t] if flujos[t] > 0 else 0
            anos = (t - 1) + frac_ano
            meses = int(round((anos - int(anos)) * 12))
            pb_str = f"{anos:.2f} años ({int(anos)}a {meses}m)"
            break
    return van, tir, pb_str, f_act, acum

tasa_fcl_usada = wacc_calc if "WACC" in descuento_fcl else tasa_ke

def construir_tabla_vista(es_inv=False, mult_ing=1.0, mult_cost=1.0):
    cols = ["Año 0"] + [f"Año {t}" for t in range(1, horizonte + 1)]
    data = {c: [] for c in cols}

    flujos_netos = []
    inv0 = -total_inv_neta
    prest0 = monto_credito if (es_inv and incluir_prestamo_fci0) else 0.0
    f0 = inv0 + prest0
    flujos_netos.append(f0)

    rows_a0 = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, inv0, prest0, 0, 0, 0, f0]
    for v in rows_a0:
        data["Año 0"].append(round(v))

    for t in range(horizonte):
        col_name = f"Año {t+1}"
        ing_t = ing_series[t] * mult_ing
        cv_t = -cv_series[t] * mult_cost
        cf_t = -cf_series[t] * mult_cost
        dep_t = -dep_anual_total
        amo_t = -amort_anual_total
        int_t = -intereses_anuales[t] if es_inv else 0.0
        bai_t = ing_t + cv_t + cf_t + dep_t + amo_t + int_t
        imp_t = -bai_t * tasa_imp
        bdi_t = bai_t + imp_t
        dep_pos = dep_anual_total
        amo_pos = amort_anual_total
        am_deuda = -amort_deuda_anuales[t] if es_inv else 0.0
        resc_t = total_rescate_af if t == horizonte - 1 else 0.0
        rec_ct = total_recup_ct if t == horizonte - 1 else 0.0
        fn_t = bdi_t + dep_pos + amo_pos + am_deuda + resc_t + rec_ct
        flujos_netos.append(fn_t)

        vals_t = [ing_t, cv_t, cf_t, dep_t, amo_t, int_t, bai_t, imp_t, bdi_t, dep_pos, amo_pos, 0, 0, am_deuda, resc_t, rec_ct, fn_t]
        for v in vals_t:
            data[col_name].append(round(v))

    t_desc = tasa_ke if es_inv else tasa_fcl_usada
    van, tir, pb_str, f_act, acum = calc_metrics(flujos_netos, t_desc)
    for idx_c, c_name in enumerate(cols):
        data[c_name].append(round(f_act[idx_c]))
        data[c_name].append(round(acum[idx_c]))

    index_names = [
        "(+) Ingresos", "(-) Costos Variables", "(-) Costos Fijos", "(-) Depreciación",
        "(-) Amort. Intangibles", "(-) Intereses Deuda", "(=) BAI", "(-) Impuesto (29.5%)",
        "(=) BDI", "(+) Depreciación", "(+) Amort. Intangibles", "(-) Inversión Inicial",
        "(+) Préstamo", "(-) Amortización Deuda", "(+) Valor Rescate", "(+) Recup. Cap. Trabajo",
        "(=) FLUJO CAJA NETO", "Flujo Actualizado (VP)", "Flujo Acumulado"
    ]
    df_vista = pd.DataFrame(data, index=index_names)
    return df_vista, van, tir, pb_str

df_fcl_vista, van_fcl, tir_fcl, pb_fcl = construir_tabla_vista(es_inv=False)
df_fci_vista, van_fci, tir_fci, pb_fci = construir_tabla_vista(es_inv=True)

decision_fcl = "ACEPTAR" if van_fcl >= 0 and tir_fcl > tasa_fcl_usada else "RECHAZAR"
decision_fci = "ACEPTAR" if van_fci >= 0 and tir_fci > tasa_ke else "RECHAZAR"

texto_interp_fcl = (
    f"Conclusión FCL: VAN = ${van_fcl:,.0f} (Tasa {tasa_fcl_usada*100:.2f}%) | TIR = {tir_fcl*100:.2f}% | Payback = {pb_fcl}. "
    f"El proyecto recupera la inversión neta (${total_inv_neta:,.0f}), cubre el costo de oportunidad del capital y genera riqueza neta por ${abs(van_fcl):,.0f}. "
    f"Decisión: {decision_fcl} (VAN >= 0 y TIR > Tasa exigida)."
)

texto_interp_fci = (
    f"Conclusión Inversionista: Préstamo = ${monto_credito:,.0f} ({pct_deuda*100:.0f}% Inv., Kd {tasa_kd*100:.2f}%, {plazo_deuda} años, Cuota PAGO = ${cuota_anual:,.0f}) | "
    f"VAN FCI = ${van_fci:,.0f} (Ke {tasa_ke*100:.2f}%) | TIR FCI = {tir_fci*100:.2f}% | Payback = {pb_fci}. "
    f"El apalancamiento financiero y el escudo fiscal de los intereses ({tasa_imp*100:.1f}%) incrementan la rentabilidad sobre el capital propio. Decisión: {decision_fci}."
)

def generar_excel_monocromatico():
    wb = openpyxl.Workbook()
    wb.calculation.fullCalcOnLoad = True

    font_bold = Font(name="Calibri", size=11, bold=True, color="000000")
    font_reg = Font(name="Calibri", size=11, bold=False, color="000000")
    border_bottom = Border(bottom=Side(style="thin", color="000000"))
    border_total = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))
    num_fmt_curr = "#,##0;(#,##0);\"-\""
    num_fmt_pct = "0.00%"

    ws1 = wb.active
    ws1.title = "Hoja1"
    ws1["B2"] = "Parámetros y Supuestos"
    ws1["B2"].font = font_bold
    ws1["B2"].border = border_bottom

    params = [
        ("Horizonte (años)", int(horizonte), "0"),
        ("Impuesto", float(tasa_imp), num_fmt_pct),
        ("Costo Capital (Ke)", float(tasa_ke), num_fmt_pct),
        ("Tasa Deuda (Kd)", float(tasa_kd), num_fmt_pct),
        ("% Deuda", float(pct_deuda), num_fmt_pct),
        ("Plazo Crédito (años)", int(plazo_deuda), "0"),
        ("IVA", float(IVA_RATE), num_fmt_pct),
    ]
    for i, (label, val, fmt) in enumerate(params, start=3):
        ws1[f"B{i}"] = label
        ws1[f"B{i}"].font = font_reg
        ws1[f"C{i}"] = val
        ws1[f"C{i}"].font = font_reg
        ws1[f"C{i}"].number_format = fmt

    ws1["B10"] = "WACC"
    ws1["B10"].font = font_bold
    ws1["C10"] = "=C5*(1-C7)+C6*(1-C4)*C7"
    ws1["C10"].number_format = num_fmt_pct
    ws1["C10"].font = font_bold

    ws1["B12"] = "Inversiones y Depreciación"
    ws1["B12"].font = font_bold

    headers_inv = ["Ítem", "Valor Neto", "IVA (19%)", "Valor Bruto", "Vida Útil", "Valor Rescate", "Depreciación Anual", "Amortización Anual"]
    for col_idx, h in enumerate(headers_inv, start=2):
        cell = ws1.cell(13, col_idx, h)
        cell.font = font_bold
        cell.border = border_bottom

    start_inv_row = 14
    for idx, r in df_inv_calc.iterrows():
        curr_r = start_inv_row + idx
        ws1.cell(curr_r, 2, str(r["Ítem"])).font = font_reg
        c_net = ws1.cell(curr_r, 3, float(r["Valor Neto"]))
        c_net.number_format = num_fmt_curr
        c_net.font = font_reg

        c_iva = ws1.cell(curr_r, 4, f"=C{curr_r}*$C$9" if r["IVA (19%)"] > 0 else 0)
        c_iva.number_format = num_fmt_curr
        c_iva.font = font_reg

        c_bru = ws1.cell(curr_r, 5, f"=C{curr_r}+D{curr_r}")
        c_bru.number_format = num_fmt_curr
        c_bru.font = font_reg

        c_vid = ws1.cell(curr_r, 6, int(r["Vida Útil"]))
        c_vid.font = font_reg

        c_res = ws1.cell(curr_r, 7, float(r["Valor Rescate"]))
        c_res.number_format = num_fmt_curr
        c_res.font = font_reg

        form_dep = f"=(C{curr_r}-G{curr_r})/F{curr_r}" if restar_rescate_dep else f"=C{curr_r}/F{curr_r}"
        c_dep = ws1.cell(curr_r, 8, form_dep if r["Depreciación Anual"] > 0 else 0)
        c_dep.number_format = num_fmt_curr
        c_dep.font = font_reg

        c_amo = ws1.cell(curr_r, 9, f"=C{curr_r}/F{curr_r}" if r["Amortización Anual"] > 0 else 0)
        c_amo.number_format = num_fmt_curr
        c_amo.font = font_reg

    tot_inv_row = start_inv_row + len(df_inv_calc)
    c_tot_lbl = ws1.cell(tot_inv_row, 2, "Total")
    c_tot_lbl.font = font_bold
    c_tot_lbl.border = border_total

    for col_i in [3, 4, 5, 7, 8, 9]:
        col_let = get_column_letter(col_i)
        c_tot = ws1.cell(tot_inv_row, col_i, f"=SUM({col_let}{start_inv_row}:{col_let}{tot_inv_row-1})")
        c_tot.font = font_bold
        c_tot.border = border_total
        c_tot.number_format = num_fmt_curr

    loan_row = tot_inv_row + 2
    ws1.cell(loan_row, 2, "Préstamo").font = font_bold
    base_col_let = "E" if "Bruta" in base_credito else "C"
    c_loan = ws1.cell(loan_row, 3, f"={base_col_let}{tot_inv_row}*C7*(1+{gasto_credito_pct})")
    c_loan.number_format = num_fmt_curr
    c_loan.font = font_bold

    ws3 = wb.create_sheet("Hoja3")
    ws3["B2"] = "Tabla de Amortización"
    ws3["B2"].font = font_bold

    ws3["B4"] = "Monto Crédito"
    ws3["C4"] = f"=Hoja1!C{loan_row}"
    ws3["C4"].number_format = num_fmt_curr
    ws3["B5"] = "Tasa Anual (Kd)"
    ws3["C5"] = "=Hoja1!C6"
    ws3["C5"].number_format = num_fmt_pct
    ws3["B6"] = "Plazo (Años)"
    ws3["C6"] = "=Hoja1!C8"
    ws3["B7"] = "Cuota Fija"
    ws3["B7"].font = font_bold
    ws3["C7"] = "=PMT(C5,C6,-C4)"
    ws3["C7"].font = font_bold
    ws3["C7"].number_format = num_fmt_curr

    headers_amo = ["Período", "Cuota", "Interés", "Amortización", "Saldo"]
    for idx, h in enumerate(headers_amo, start=2):
        c = ws3.cell(9, idx, h)
        c.font = font_bold
        c.border = border_bottom

    ws3.cell(10, 2, 0)
    ws3.cell(10, 6, "=C4").number_format = num_fmt_curr

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

    def poblar_hoja_flujo(ws, es_inversionista=False):
        titulo = "Flujo de Caja Inversionista" if es_inversionista else "Flujo de Caja Libre"
        ws["B2"] = titulo
        ws["B2"].font = font_bold

        c_hdr = ws.cell(4, 2, "Ítem / Año")
        c_hdr.font = font_bold
        c_hdr.border = border_bottom
        for t in range(horizonte + 1):
            c = ws.cell(4, 3 + t, t)
            c.font = font_bold
            c_hdr.border = border_bottom
            c.alignment = Alignment(horizontal="right")

        filas_labels = [
            (5, "(+) Ingresos"),
            (6, "(-) Costos Variables"),
            (7, "(-) Costos Fijos"),
            (8, "(-) Depreciación"),
            (9, "(-) Amortización Intangibles"),
            (10, "(-) Intereses Deuda"),
            (11, "(=) BAI"),
            (12, "(-) Impuesto (29,5%)"),
            (13, "(=) BDI"),
            (14, "(+) Depreciación"),
            (15, "(+) Amortización Intangibles"),
            (16, "(-) Inversión Inicial"),
            (17, "(+) Préstamo"),
            (18, "(-) Amortización Deuda"),
            (19, "(+) Valor Rescate"),
            (20, "(+) Recup. Capital Trabajo"),
            (21, "(=) FLUJO CAJA NETO"),
            (22, "Factor Actualización"),
            (23, "Flujo Actualizado"),
            (24, "Flujo Acumulado")
        ]
        for r_num, lbl in filas_labels:
            cell_lbl = ws.cell(r_num, 2, lbl)
            cell_lbl.font = font_bold if r_num in [11, 13, 21] else font_reg
            if r_num == 21:
                cell_lbl.border = border_total

        ws.cell(16, 3, f"=-Hoja1!C{tot_inv_row}").number_format = num_fmt_curr
        if es_inversionista and incluir_prestamo_fci0:
            ws.cell(17, 3, "=Hoja3!C4").number_format = num_fmt_curr
        else:
            ws.cell(17, 3, 0).number_format = num_fmt_curr

        c_fn0 = ws.cell(21, 3, "=SUM(C13:C20)")
        c_fn0.font = font_bold
        c_fn0.border = border_total
        c_fn0.number_format = num_fmt_curr

        tasa_ref = "Hoja1!$C$5" if (es_inversionista or "Ke" in descuento_fcl) else "Hoja1!$C$10"
        ws.cell(22, 3, f"=(1/(1+{tasa_ref}))^C4").number_format = "0.0000"
        ws.cell(23, 3, "=C21*C22").number_format = num_fmt_curr
        ws.cell(24, 3, "=C21").number_format = num_fmt_curr

        for t in range(1, horizonte + 1):
            col_i = 3 + t
            col_let = get_column_letter(col_i)
            prev_let = get_column_letter(col_i - 1)
            g_ing = float(tasas_df.iloc[t-1]["Crec. Ingresos (%)"]) / 100.0
            g_cv = float(tasas_df.iloc[t-1]["Crec. Costos Var. (%)"]) / 100.0
            g_cf = float(tasas_df.iloc[t-1]["Crec. Costos Fijos (%)"]) / 100.0

            if t == 1:
                ws.cell(5, col_i, float(ing_series[0]))
                ws.cell(6, col_i, float(-cv_series[0]))
                ws.cell(7, col_i, float(-cf_series[0]))
            else:
                ws.cell(5, col_i, f"={prev_let}5*(1+{g_ing})")
                ws.cell(6, col_i, f"={prev_let}6*(1+{g_cv})")
                ws.cell(7, col_i, f"={prev_let}7*(1+{g_cf})")

            ws.cell(8, col_i, f"=-Hoja1!$H${tot_inv_row}")
            ws.cell(9, col_i, f"=-Hoja1!$I${tot_inv_row}")
            if es_inversionista:
                ws.cell(10, col_i, f"=-Hoja3!D{10+t}")
            else:
                ws.cell(10, col_i, 0)

            ws.cell(11, col_i, f"=SUM({col_let}5:{col_let}10)").font = font_bold
            ws.cell(12, col_i, f"=-{col_let}11*Hoja1!$C$4")
            ws.cell(13, col_i, f"=SUM({col_let}11:{col_let}12)").font = font_bold
            ws.cell(14, col_i, f"=-{col_let}8")
            ws.cell(15, col_i, f"=-{col_let}9")
            ws.cell(16, col_i, 0)
            ws.cell(17, col_i, 0)
            if es_inversionista:
                ws.cell(18, col_i, f"=-Hoja3!E{10+t}")
            else:
                ws.cell(18, col_i, 0)

            if t == horizonte:
                ws.cell(19, col_i, float(total_rescate_af))
                ws.cell(20, col_i, float(total_recup_ct))
            else:
                ws.cell(19, col_i, 0)
                ws.cell(20, col_i, 0)

            c_fnt = ws.cell(21, col_i, f"=SUM({col_let}13:{col_let}20)")
            c_fnt.font = font_bold
            c_fnt.border = border_total

            ws.cell(22, col_i, f"=(1/(1+{tasa_ref}))^{col_let}4").number_format = "0.0000"
            ws.cell(23, col_i, f"={col_let}21*{col_let}22")
            ws.cell(24, col_i, f"={prev_let}24+{col_let}21")

            for r_curr in range(5, 22):
                ws.cell(r_curr, col_i).number_format = num_fmt_curr
            ws.cell(23, col_i).number_format = num_fmt_curr
            ws.cell(24, col_i).number_format = num_fmt_curr

        last_let = get_column_letter(3 + horizonte)
        ws.cell(26, 2, "VAN (Suma VP)").font = font_bold
        ws.cell(26, 3, f"=SUM(C23:{last_let}23)").font = font_bold
        ws.cell(26, 3).number_format = num_fmt_curr

        ws.cell(27, 2, "VAN (VNA)").font = font_bold
        ws.cell(27, 3, f"=C21+NPV({tasa_ref},D21:{last_let}21)").font = font_bold
        ws.cell(27, 3).number_format = num_fmt_curr

        ws.cell(28, 2, "TIR").font = font_bold
        ws.cell(28, 3, f"=IRR(C21:{last_let}21)").font = font_bold
        ws.cell(28, 3).number_format = num_fmt_pct

        ws.cell(29, 2, "Payback").font = font_bold
        ws.cell(29, 3, pb_fci if es_inversionista else pb_fcl).font = font_bold

        ws.cell(31, 2, "Respuesta:").font = font_bold
        ws.merge_cells(start_row=32, start_column=2, end_row=34, end_column=3 + horizonte)
        c_interp = ws.cell(32, 2, texto_interp_fci if es_inversionista else texto_interp_fcl)
        c_interp.alignment = Alignment(wrap_text=True, vertical="top")

    ws2 = wb.create_sheet("Hoja2")
    poblar_hoja_flujo(ws2, es_inversionista=False)

    ws4 = wb.create_sheet("Hoja4")
    poblar_hoja_flujo(ws4, es_inversionista=True)

    wb._sheets = [ws1, ws2, ws3, ws4]

    for sheet in wb.worksheets:
        sheet.column_dimensions["B"].width = 34
        for col_i in range(3, 15):
            sheet.column_dimensions[get_column_letter(col_i)].width = 16

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

excel_bytes = generar_excel_monocromatico()

top_left, top_right = st.columns([5.5, 1])
with top_left:
    st.caption(
        f"Inv. Neta: ${total_inv_neta:,.0f}  |  Inv. Bruta: ${total_inv_bruta:,.0f}  |  "
        f"Crédito: ${monto_credito:,.0f}  |  Cuota PAGO: ${cuota_anual:,.0f}  |  WACC: {wacc_calc*100:.2f}%"
    )
with top_right:
    st.download_button(
        label="Guardar (.xlsx)",
        data=excel_bytes,
        file_name="Libro1.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "Hoja1 (FCL)",
    "Hoja2 (Amortización)",
    "Hoja3 (Inversionista)",
    "Hoja4 (Inversiones/IVA)",
    "Hoja5 (Sensibilidad)",
    "Hoja6 (Monte Carlo)",
    "Hoja7 (Guía Repaso)"
])

with tab1:
    st.dataframe(df_fcl_vista.style.format("{:,.0f}"), use_container_width=True, height=450)
    st.markdown(f'<div class="nota-celda">{texto_interp_fcl}</div>', unsafe_allow_html=True)

with tab2:
    st.dataframe(df_amort.style.format("{:,.0f}"), use_container_width=True, hide_index=True)

with tab3:
    st.dataframe(df_fci_vista.style.format("{:,.0f}"), use_container_width=True, height=450)
    st.markdown(f'<div class="nota-celda">{texto_interp_fci}</div>', unsafe_allow_html=True)

with tab4:
    st.dataframe(
        df_inv_calc[["Ítem", "Valor Neto", "IVA (19%)", "Valor Bruto", "Vida Útil", "Valor Rescate", "Depreciación Anual", "Amortización Anual"]].style.format({
            "Valor Neto": "{:,.0f}", "IVA (19%)": "{:,.0f}", "Valor Bruto": "{:,.0f}",
            "Valor Rescate": "{:,.0f}", "Depreciación Anual": "{:,.0f}", "Amortización Anual": "{:,.0f}"
        }),
        use_container_width=True,
        hide_index=True
    )

with tab5:
    _, van_pes_fcl, tir_pes_fcl, pb_pes_fcl = construir_tabla_vista(es_inv=False, mult_ing=0.90)
    _, van_opt_fcl, tir_opt_fcl, pb_opt_fcl = construir_tabla_vista(es_inv=False, mult_ing=1.10)
    _, van_pes_fci, tir_pes_fci, pb_pes_fci = construir_tabla_vista(es_inv=True, mult_ing=0.90)
    _, van_opt_fci, tir_opt_fci, pb_opt_fci = construir_tabla_vista(es_inv=True, mult_ing=1.10)

    df_sens = pd.DataFrame([
        {"Escenario": "Pesimista (-10% Ingresos)", "VAN FCL ($)": round(van_pes_fcl), "TIR FCL (%)": round(tir_pes_fcl*100, 2), "Payback FCL": pb_pes_fcl, "VAN Inversionista ($)": round(van_pes_fci), "TIR Inversionista (%)": round(tir_pes_fci*100, 2)},
        {"Escenario": "Caso Base (100%)", "VAN FCL ($)": round(van_fcl), "TIR FCL (%)": round(tir_fcl*100, 2), "Payback FCL": pb_fcl, "VAN Inversionista ($)": round(van_fci), "TIR Inversionista (%)": round(tir_fci*100, 2)},
        {"Escenario": "Optimista (+10% Ingresos)", "VAN FCL ($)": round(van_opt_fcl), "TIR FCL (%)": round(tir_opt_fcl*100, 2), "Payback FCL": pb_opt_fcl, "VAN Inversionista ($)": round(van_opt_fci), "TIR Inversionista (%)": round(tir_opt_fci*100, 2)},
    ])
    st.dataframe(df_sens, use_container_width=True, hide_index=True)

    with st.expander("Calculadora de VAN Esperado por Probabilidades (Ej. 5 Guía)", expanded=False):
        c_p1, c_p2, c_p3 = st.columns(3)
        with c_p1:
            prob_opt = st.number_input("Prob. Optimista (%)", value=30.0, step=5.0, help="Probabilidad asignada al escenario optimista.")
            van_opt_in = st.number_input("VAN Optimista ($)", value=250000000.0, step=5000000.0, format="%.0f")
        with c_p2:
            prob_nor = st.number_input("Prob. Normal (%)", value=50.0, step=5.0, help="Probabilidad asignada al escenario base o normal.")
            van_nor_in = st.number_input("VAN Normal ($)", value=100000000.0, step=5000000.0, format="%.0f")
        with c_p3:
            prob_pes = st.number_input("Prob. Pesimista (%)", value=20.0, step=5.0, help="Probabilidad asignada al escenario pesimista.")
            van_pes_in = st.number_input("VAN Pesimista ($)", value=-80000000.0, step=5000000.0, format="%.0f")

        van_esperado_disc = (prob_opt/100.0)*van_opt_in + (prob_nor/100.0)*van_nor_in + (prob_pes/100.0)*van_pes_in
        desv_est_van = np.sqrt(
            (prob_opt/100.0)*((van_opt_in - van_esperado_disc)**2) +
            (prob_nor/100.0)*((van_nor_in - van_esperado_disc)**2) +
            (prob_pes/100.0)*((van_pes_in - van_esperado_disc)**2)
        )
        st.markdown(
            f'<div class="nota-celda"><b>Resultado VAN Esperado E(VAN):</b> ${van_esperado_disc:,.0f} &nbsp;|&nbsp; '
            f'<b>Desviación Estándar (Riesgo):</b> ${desv_est_van:,.0f} &nbsp;|&nbsp; '
            f'<b>Coeficiente de Variación:</b> {(desv_est_van/abs(van_esperado_disc) if van_esperado_disc != 0 else 0):.2f} '
            f'{"(ACEPTAR: E(VAN) > 0, pero monitorear probabilidad de pérdida de " + str(prob_pes) + "%)" if van_esperado_disc > 0 else "(RECHAZAR)"}</div>',
            unsafe_allow_html=True
        )

with tab6:
    st.caption("Simulación Estocástica de Monte Carlo (10.000 Iteraciones Triangulares sobre el Caso Activo)")
    mc_c1, mc_c2, mc_c3, mc_c4 = st.columns(4)
    with mc_c1:
        ing_min_pct = st.number_input(
            "Ingresos Mín. (% del Base)", value=80.0, step=5.0,
            help="Porcentaje de los ingresos proyectados en el peor escenario razonable (ej. 80% equivale a una caída del 20%)."
        )
        ing_max_pct = st.number_input(
            "Ingresos Máx. (% del Base)", value=125.0, step=5.0,
            help="Porcentaje de los ingresos proyectados en un escenario comercial favorable (ej. 125% equivale a un alza del 25%)."
        )
    with mc_c2:
        cv_min_pct = st.number_input(
            "Costo Var. Mín. (% del Base)", value=85.0, step=5.0,
            help="Porcentaje del costo variable en caso de eficiencias o descuentos por volumen."
        )
        cv_max_pct = st.number_input(
            "Costo Var. Máx. (% del Base)", value=120.0, step=5.0,
            help="Porcentaje del costo variable si suben las materias primas o insumos."
        )
    with mc_c3:
        cf_min_pct = st.number_input(
            "Costo Fijo Mín. (% del Base)", value=95.0, step=5.0,
            help="Variación mínima de los costos fijos anuales."
        )
        cf_max_pct = st.number_input(
            "Costo Fijo Máx. (% del Base)", value=110.0, step=5.0,
            help="Variación máxima de los costos fijos anuales por inflación o imprevistos."
        )
    with mc_c4:
        tipo_flujo_mc = st.selectbox(
            "Flujo a Simular",
            ["Flujo de Caja Libre (FCL)", "Flujo Inversionista (FCI)"],
            help="Elige si deseas evaluar el riesgo del proyecto puro (FCL) o el riesgo para el accionista con deuda (FCI)."
        )
        semilla_mc = st.number_input(
            "Iteraciones", value=10000, min_value=1000, max_value=50000, step=1000,
            help="Número de escenarios aleatorios evaluados mediante distribución triangular año a año."
        )

    es_fci_mc = "Inversionista" in tipo_flujo_mc
    n_iter = int(semilla_mc)
    tasa_mc = tasa_ke if es_fci_mc else tasa_fcl_usada

    # Año 0
    inv0_mc = -total_inv_neta + (monto_credito if (es_fci_mc and incluir_prestamo_fci0) else 0.0)
    van_mc = np.full(n_iter, inv0_mc, dtype=float)

    # Simulación año a año rompiendo la linealidad
    np.random.seed(42)
    f_ing_min, f_ing_max = min(ing_min_pct, 100.0)/100.0, max(ing_max_pct, 100.01)/100.0
    f_cv_min, f_cv_max = min(cv_min_pct, 100.0)/100.0, max(cv_max_pct, 100.01)/100.0
    f_cf_min, f_cf_max = min(cf_min_pct, 100.0)/100.0, max(cf_max_pct, 100.01)/100.0

    for t in range(horizonte):
        mult_i = np.random.triangular(f_ing_min, 1.0, f_ing_max, n_iter)
        mult_v = np.random.triangular(f_cv_min, 1.0, f_cv_max, n_iter)
        mult_f = np.random.triangular(f_cf_min, 1.0, f_cf_max, n_iter)

        ing_sim = ing_series[t] * mult_i
        cv_sim = cv_series[t] * mult_v
        cf_sim = cf_series[t] * mult_f
        int_sim = intereses_anuales[t] if es_fci_mc else 0.0

        bai_sim = ing_sim - cv_sim - cf_sim - dep_anual_total - amort_anual_total - int_sim
        imp_sim = bai_sim * tasa_imp
        bdi_sim = bai_sim - imp_sim

        am_deuda_sim = amort_deuda_anuales[t] if es_fci_mc else 0.0
        resc_sim = (total_rescate_af + total_recup_ct) if t == horizonte - 1 else 0.0

        fn_sim = bdi_sim + dep_anual_total + amort_anual_total - am_deuda_sim + resc_sim
        van_mc += fn_sim / ((1.0 + tasa_mc) ** (t + 1))

    van_prom_mc = float(van_mc.mean())
    prob_exito_mc = float((van_mc >= 0).mean() * 100.0)
    riesgo_dest_mc = 100.0 - prob_exito_mc
    var_95_mc = float(np.percentile(van_mc, 5))

    k_mc1, k_mc2, k_mc3, k_mc4 = st.columns(4)
    k_mc1.metric(
        "VAN Promedio Simulado", f"${van_prom_mc:,.0f}",
        help="Promedio de riqueza neta creada tras simular los 10.000 escenarios."
    )
    k_mc2.metric(
        "Probabilidad de Éxito (VAN >= 0)", f"{prob_exito_mc:.2f}%",
        help="Porcentaje de universos simulados en los que el proyecto supera la tasa exigida."
    )
    k_mc3.metric(
        "Riesgo Destrucción Valor", f"{riesgo_dest_mc:.2f}%",
        help="Porcentaje de escenarios donde el VAN resulta negativo."
    )
    k_mc4.metric(
        "VaR 95% (Percentil 5%)", f"${var_95_mc:,.0f}",
        help="Valor en Riesgo: Con un 95% de confianza, el VAN no caerá por debajo de este monto."
    )

    # Histograma de frecuencias compatible sin requerir librerías externas
    conteos, bordes = np.histogram(van_mc, bins=35)
    etiquetas_bins = [f"${(bordes[i]+bordes[i+1])/2000000:,.1f}M" for i in range(len(conteos))]
    df_hist = pd.DataFrame({
        "Rango VAN ($ Millones)": etiquetas_bins,
        "Frecuencia (Escenarios)": conteos
    }).set_index("Rango VAN ($ Millones)")
    st.bar_chart(df_hist, height=260, color="#107c41")

    df_resumen_mc = pd.DataFrame([
        {
            "Métrica Monte Carlo": "VAN Promedio Esperado",
            "Valor ($ / %)": f"${van_prom_mc:,.0f}",
            "Peor Escenario (Mín)": f"${van_mc.min():,.0f}",
            "Mejor Escenario (Máx)": f"${van_mc.max():,.0f}",
            "Percentil 5% (VaR 95%)": f"${var_95_mc:,.0f}",
            "Mediana (P50)": f"${np.median(van_mc):,.0f}",
            "Prob. Éxito (%)": f"{prob_exito_mc:.2f}%"
        }
    ])
    st.dataframe(df_resumen_mc, use_container_width=True, hide_index=True)

    st.markdown(
        f'<div class="nota-celda"><b>Conclusión Estocástica ({tipo_flujo_mc}):</b> Tras evaluar {n_iter:,} escenarios con distribución triangular '
        f'y tasa de descuento de {tasa_mc*100:.2f}%, el proyecto presenta una probabilidad de éxito comercial de <b>{prob_exito_mc:.2f}%</b> '
        f'y un VAN medio de <b>${van_prom_mc:,.0f}</b>. En el 5% de los peores escenarios (VaR 95%), el VAN se sitúa en <b>${var_95_mc:,.0f}</b>.</div>',
        unsafe_allow_html=True
    )

with tab7:
    st.caption("Pauta de Auditoría y Resolución: Guía de Análisis - Inversión en Activos Reales")
    guia_selec = st.selectbox(
        "Seleccionar Ejercicio de la Guía",
        [
            "Ejercicio 7: Auditoría de los 3 Errores del Flujo de Caja (Conservas)",
            "Ejercicio 8: Proyecto a 10 años (Ampliación, Reemplazo y Costo Hundido)",
            "Ejercicios 1 al 6: Respuestas Conceptuales y Criterios VAN vs. TIR vs. Payback"
        ]
    )

    if "Ejercicio 7" in guia_selec:
        df_err7 = pd.DataFrame([
            {
                "N° Error": "1. Inconsistencia de signos en Egresos, Depreciación e Inversiones",
                "Evidencia en la Tabla del Asesor": "Materia prima Año 1 (+13.000.000), MOI Años 3-5 (+5.040.000 y +6.500.000), Insumos Años 1, 4 y 5 (+), Gastos Generales (+1.000.000), Depreciación (+4.845.000) e Inversiones Año 0 en Act. Nominales (+1.000.000) y Cap. Trabajo (+1.100.000).",
                "Justificación Conceptual": "Los costos operativos, la depreciación pre-impuestos y los desembolsos de inversión en el Año 0 son salidas o deducciones y deben llevar signo negativo consistente. Mezclar signos distorsiona la Utilidad Antes de Impuestos y confunde la lectura contable."
            },
            {
                "N° Error": "2. Omisión del reverso de la Depreciación después de impuestos",
                "Evidencia en la Tabla del Asesor": "Después de la Utilidad Después de Impuestos solo se vuelve a sumar la Amortización (+200.000), pero se omite sumar de vuelta la Depreciación (+4.845.000).",
                "Justificación Conceptual": "La depreciación no constituye un flujo de salida de efectivo real; solo se resta antes de impuestos para capturar su escudo fiscal. Al no reversarla después de impuestos, se subestima el Flujo de Caja Libre anual en $4.845.000."
            },
            {
                "N° Error": "3. Multiplicación en vez de división por el Factor de Actualización",
                "Evidencia en la Tabla del Asesor": "En el Año 1 el Flujo de Caja es $6.327.737 y el Flujo Actualizado sube a $7.273.260 (multiplicó 6.327.737 x 1,149425 en vez de dividir).",
                "Justificación Conceptual": "Por el valor del dinero en el tiempo, un peso futuro vale menos hoy. Para traer flujos futuros a Valor Presente se debe dividir por (1 + r)^t o multiplicar por (1 + r)^(-t). El asesor capitalizó los flujos hacia el futuro en vez de descontarlos."
            }
        ])
        st.dataframe(df_err7, use_container_width=True, hide_index=True)
        st.markdown(
            '<div class="nota-celda"><b>Nota práctica:</b> Para ver este flujo corregido y cuadrado al peso, '
            'selecciona <b>"Caso 3 - Conservas (13M)"</b> en el menú lateral izquierdo y revisa la <b>Hoja1 (FCL)</b>.</div>',
            unsafe_allow_html=True
        )

    elif "Ejercicio 8" in guia_selec:
        # Modelo exacto Ejercicio 8 (10 años, valores netos, tasa 12%, impuesto 15%)
        cols_e8 = [f"Año {t}" for t in range(11)]
        ing_e8 = [0] + [25000000]*2 + [30000000]*3 + [36000000]*5
        cv_fab_e8 = [0] + [-3000000]*5 + [-3420000]*5
        comis_e8 = [0] + [-500000]*2 + [-600000]*3 + [-720000]*5
        cf_fab_e8 = [0] + [-2000000]*5 + [-2200000]*5
        gav_e8 = [0] + [-800000]*5 + [-820000]*5
        # Depreciación: Obras físicas 60M/20 = 3M (A1-5), +12M/20 = 0.6M -> 3.6M (A6-10)
        # Máquinas: 48M/10 = 4.8M (A1-5), +8M/10 = 0.8M -> 5.6M (A6-10, reemplazo 10M en A8 mantiene dep 1M/año)
        dep_e8 = [0] + [-7800000]*5 + [-9200000]*5
        # Amortización intangibles: Puesta en marcha $1.200.000 (excluye $800.000 de estudio de viabilidad por ser costo hundido) / 5 años = 240.000
        amo_e8 = [0] + [-240000]*5 + [0]*5
        vta_maq_e8 = [0]*8 + [2500000] + [0]*2
        vl_maq_e8 = [0]*8 + [-2000000] + [0]*2

        uai_e8 = [ing_e8[t]+cv_fab_e8[t]+comis_e8[t]+cf_fab_e8[t]+gav_e8[t]+dep_e8[t]+amo_e8[t]+vta_maq_e8[t]+vl_maq_e8[t] for t in range(11)]
        imp_e8 = [-max(0, uai_e8[t])*0.15 for t in range(11)]
        udi_e8 = [uai_e8[t]+imp_e8[t] for t in range(11)]
        rev_dep_e8 = [0] + [-dep_e8[t] for t in range(1, 11)]
        rev_amo_e8 = [0] + [-amo_e8[t] for t in range(1, 11)]
        rev_vl_e8 = [0]*8 + [2000000] + [0]*2

        # Inversiones: A0 = -(12M Terreno + 60M Obras + 48M Maq + 1.2M Puesta en Marcha + 2.9M Cap.Trab) = -124.100.000
        # A5 (fin año 5 para operar año 6) = -(12M Obras + 8M Maq) = -20.000.000 | A8 = -10.000.000 reemplazo maq
        inv_e8 = [-124100000, 0, 0, 0, 0, -20000000, 0, 0, -10000000, 0, 0]
        desecho_e8 = [0]*10 + [139117000]
        rec_ct_e8 = [0]*10 + [2900000]

        fcn_e8 = [udi_e8[t]+rev_dep_e8[t]+rev_amo_e8[t]+rev_vl_e8[t]+inv_e8[t]+desecho_e8[t]+rec_ct_e8[t] for t in range(11)]
        van_e8, tir_e8, pb_e8, vp_e8, _ = calc_metrics(fcn_e8, 0.12)

        df_e8 = pd.DataFrame(
            [ing_e8, cv_fab_e8, comis_e8, cf_fab_e8, gav_e8, dep_e8, amo_e8, vta_maq_e8, vl_maq_e8, uai_e8, imp_e8, udi_e8, rev_dep_e8, rev_amo_e8, rev_vl_e8, inv_e8, desecho_e8, rec_ct_e8, fcn_e8, [round(x) for x in vp_e8]],
            index=[
                "(+) Ingresos", "(-) Costos Var. Fabricación", "(-) Comisiones (2%)", "(-) Costos Fijos Fab.",
                "(-) Gastos Adm. y Ventas", "(-) Depreciación", "(-) Amort. Intangibles", "(+) Venta Máquina Usada",
                "(-) Valor Libro Máquina", "(=) Utilidad Antes Impuestos", "(-) Impuesto (15%)", "(=) Utilidad Neta",
                "(+) Depreciación", "(+) Amort. Intangibles", "(+) Valor Libro Máquina", "(-) Inversiones (Capex/CT)",
                "(+) Valor de Desecho", "(+) Recup. Cap. Trabajo", "(=) FLUJO DE CAJA NETO", "Flujo Actualizado (12%)"
            ],
            columns=cols_e8
        )
        st.dataframe(df_e8.style.format("{:,.0f}"), use_container_width=True, height=420)
        st.markdown(
            f'<div class="nota-celda"><b>Resultado Ejercicio 8 (10 años, Tasa 12%, Impuesto 15%):</b> '
            f'VAN = <b>${van_e8:,.0f}</b> | TIR = <b>{tir_e8*100:.2f}%</b> | Payback = <b>{pb_e8}</b>. '
            f'<i>Supuesto clave de cátedra:</i> El estudio de viabilidad por $800.000 es un <b>costo hundido</b> (irrelevante para la decisión), '
            f'por lo que solo se activan y amortizan $1.200.000 de puesta en marcha ($240.000/año durante 5 años).</div>',
            unsafe_allow_html=True
        )

    else:
        df_c1_6 = pd.DataFrame([
            {"Ejercicio": "1. Presupuesto $500M (Sucursales vs. Automatizar vs. Postergar)", "Respuesta Ejecutiva": "No se puede decidir por intuición. Se debe solicitar FCL proyectado, WACC, VAN, TIR y análisis de sensibilidad de cada opción. Abrir sucursales tiene mayor riesgo de mercado (demanda incierta); automatizar tiene menor riesgo (ahorro de costos controlable). Postergar tiene costo de oportunidad."},
            {"Ejercicio": "2. Clasificación de Proyectos", "Respuesta Ejecutiva": "Sustituir máquina: Reemplazo/Reducción de costos | Filtros ambientales: Obligatorio/Normativo (no rentable por sí mismo) | Operaciones en Perú: Expansión geográfica/Nuevos mercados | Casino trabajadores: Bienestar/Indirecto | Línea vegana: Expansión de nuevos productos."},
            {"Ejercicio": "3. Proyecto A (VAN $35M, TIR 18%) vs. Proyecto B (VAN $22M, TIR 24%)", "Respuesta Ejecutiva": "Ambos son aceptables (VAN > 0 y TIR > 14% exigido). Si son mutuamente excluyentes, el accionista debe elegir el Proyecto A porque maximiza la riqueza neta en pesos hoy ($35M > $22M). VAN y TIR difieren por el perfil temporal de los flujos y la tasa de reinversión implícita."},
            {"Ejercicio": "4. Máquina A (VAN $120M, TIR 17%, PB 3a) vs. Máquina B (VAN $90M, TIR 24%, PB 2a)", "Respuesta Ejecutiva": "Escoger Máquina A: aporta $30.000.000 adicionales de valor neto para la empresa. La Máquina B recupera antes (2 años) pero genera menos riqueza total durante su vida útil."},
            {"Ejercicio": "5. Nueva Planta bajo Incertidumbre (30% $250M, 50% $100M, 20% -$80M)", "Respuesta Ejecutiva": "E(VAN) = 0,3(250M) + 0,5(100M) + 0,2(-80M) = $109.000.000. Se recomienda invertir porque el valor esperado es ampliamente positivo, pero existe un 20% de probabilidad de destruir valor (-$80M). Mitigación: contratos de preventa, inversión modular por etapas o flexibilidad operativa."},
            {"Ejercicio": "6. Proyecto VAN $150M, TIR 15%, WACC 10%, Payback 7 años (Política máx. 5 años)", "Respuesta Ejecutiva": "No debe rechazarse automáticamente. El proyecto crea $150M de riqueza neta y rinde 5 puntos sobre el WACC (15% > 10%). Usar solo el Payback ignora todos los flujos posteriores al año 5 y castiga proyectos estratégicos de maduración larga."}
        ])
        st.dataframe(df_c1_6, use_container_width=True, hide_index=True)

# Script para ocultar elementos externos de Streamlit Cloud (corona roja, foto y barra inferior)
components.html(
    """
    <script>
        const hideStreamlitBranding = () => {
            try {
                const docs = [window.document, window.parent.document, window.top.document];
                docs.forEach(doc => {
                    if (doc && !doc.getElementById('hide-st-cloud-branding')) {
                        const style = doc.createElement('style');
                        style.id = 'hide-st-cloud-branding';
                        style.innerHTML = `
                            div[class*="viewerBadge"],
                            a[class*="viewerBadge"],
                            div[class*="_profileContainer"],
                            div[class*="_profilePreview"],
                            div[class*="embeddedAppMetaInfoBar"],
                            a[href*="streamlit.io"],
                            a[href*="share.streamlit.io/user"] {
                                display: none !important;
                                visibility: hidden !important;
                                opacity: 0 !important;
                                pointer-events: none !important;
                            }
                        `;
                        doc.head.appendChild(style);
                    }
                });
            } catch (e) {}
        };
        hideStreamlitBranding();
        setInterval(hideStreamlitBranding, 500);
    </script>
    """,
    height=0,
)
