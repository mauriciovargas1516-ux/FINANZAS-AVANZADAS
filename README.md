# Simulador de Finanzas Avanzadas y Evaluación de Activos Reales 2026

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://finanzas-avanzadas-2026-nxvqzsfrrtgu9e92javcli.streamlit.app/)
[![Lanzar Simulador en Streamlit](https://img.shields.io/badge/ABRIR_SIMULADOR_FINANCIERO_EN_STREAMLIT-107C41?style=for-the-badge&logo=streamlit&logoColor=white)](https://finanzas-avanzadas-2026-nxvqzsfrrtgu9e92javcli.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.11+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://github.com/mauriciovargas1516-ux/FINANZAS-AVANZADAS-2026)

Motor corporativo desarrollado en **Python + Streamlit** con interfaz estilo **Excel Online (`Libro1.xlsx`)** para la evaluación financiera de inversiones en activos reales, cálculo de estructura de capital (**CAPM / WACC**), construcción automatizada de **Flujo de Caja Libre (FCL)** y **Flujo del Inversionista (FCI)**, análisis de sensibilidad y **Simulación Estocástica de Monte Carlo** (hasta 50.000 iteraciones) con compilación de libros Excel con **fórmulas nativas vivas**.

**Acceso directo al simulador financiero en vivo:**  
https://finanzas-avanzadas-2026-nxvqzsfrrtgu9e92javcli.streamlit.app/

---

## Impacto Cuantitativo y Multiplicadores de Velocidad (Motor Python vs. Excel Manual)

| Proceso de Modelamiento Financiero Corporativo | Complejidad / Volumen de Cálculo | Tiempo Manual en Excel | Tiempo Motor Cloud | Horas Ahorradas | % Mejora (Multiplicador) |
| :--- | :--- | :---: | :---: | :---: | :---: |
| **1. Separación Tributaria CAPEX (Neto/IVA 19%), Depreciación y Valor de Rescate** | Múltiples activos fijos, intangibles y capital de trabajo | 2,0 horas | **0,01 hrs** (30 seg) | 1,99 hrs/caso | **99,5% (200,0x más rápido)** |
| **2. Estructura de Capital (CAPM / WACC) y Tabla de Amortización Francesa (`PAGO`)** | Cuota fija, intereses, escudo fiscal y saldo insoluto | 1,5 horas | **0,01 hrs** (20 seg) | 1,49 hrs/caso | **99,3% (150,0x más rápido)** |
| **3. Proyección Simultánea de Flujo de Caja Libre (FCL) y del Inversionista (FCI)** | Estados de resultados, reversos, VAN, TIR y Payback exacto (años/meses) | 4,5 horas | **0,02 hrs** (1 min) | 4,48 hrs/caso | **99,6% (225,0x más rápido)** |
| **4. Sensibilidad Multiescenario ($\pm 10\%$) y Calculadora de $E(\text{VAN}), \sigma, CV$** | 6 flujos simultáneos (Pesimista, Base, Optimista en FCL y FCI) | 3,0 horas | **0,01 hrs** (30 seg) | 2,99 hrs/caso | **99,7% (300,0x más rápido)** |
| **5. Simulación Estocástica de Monte Carlo (10.000 a 50.000 Iteraciones Triangulares)** | 50.000 flujos completos + cálculo de VaR 95%, P50 y Probabilidad de Éxito | **8,5 horas** (VBA / Macros) | **0,01 hrs** (1,2 seg) | 8,49 hrs/simulación | **99,9% (25.500x más rápido)** |
| **6. Exportación de `Libro1.xlsx` (6 Hojas con Fórmulas Nativas Vinculadas)** | Inyección de `=PMT`, `=NPV`, `=IRR`, `=SUMPRODUCT` sin errores `#¡REF!` | 5,0 horas | **0,01 hrs** (1 seg) | 4,99 hrs/entrega | **99,8% (18.000x más rápido)** |
| **TOTAL EVALUACIÓN INTEGRAL DE UN PROYECTO CORPORATIVO** | **Ciclo Completo FCL + FCI + Monte Carlo + Excel Vivo** | **24,5 horas** | **0,15 hrs (9 min)** | **24,35 horas (3 jornadas)** | **99,4% (163,3x más rápido)** |

---

## Estructura de Hojas y Módulos de la Aplicación

La plataforma se organiza en **7 pestañas interactivas** en pantalla y un panel lateral de configuración:

| Pestaña en Pantalla | Contenido y Cálculos Principales |
| :--- | :--- |
| **Hoja1 (FCL)** | Flujo de Caja Libre (proyecto puro sin deuda). Calcula VAN, TIR, Payback exacto en años y meses, Flujo Actualizado (VP), Flujo Acumulado y dictamen ejecutivo automático (`ACEPTAR` / `RECHAZAR`). |
| **Hoja2 (Amortización)** | Tabla de desarrollo del préstamo bancario período a período: Saldo Inicial, Cuota Fija (`PAGO`), Gasto Financiero (Intereses), Amortización de Capital y Saldo Insoluto. |
| **Hoja3 (Inversionista)** | Flujo de Caja del Inversionista (FCI) apalancado. Incorpora el desembolso neto de deuda en el Año 0, rebaja de intereses antes de impuestos y pago de amortizaciones, descontado a la tasa exigida por el accionista ($K_e$). |
| **Hoja4 (Inversiones/IVA)** | Cuadro detallado de activos fijos, intangibles y capital de trabajo. Desglosa Valor Neto, IVA (19%), Valor Bruto, Vida Útil, Valor de Rescate y cuotas anuales de depreciación y amortización. |
| **Hoja5 (Sensibilidad)** | Comparativa simultánea de escenarios Pesimista (-10% ingresos), Base (100%) y Optimista (+10% ingresos) para FCL y FCI. Incluye calculadora de VAN Esperado $E(\text{VAN})$, Desviación Estándar y Coeficiente de Variación ($CV$). |
| **Hoja6 (Monte Carlo)** | Simulación estocástica (10.000 a 50.000 iteraciones) mediante distribuciones triangulares para ingresos, costos variables y costos fijos. Entrega VAN Promedio, Mediana (P50), Probabilidad de Éxito ($\text{VAN} \ge 0$), Riesgo de Destrucción de Valor, VaR al 95% (Percentil 5%) e histograma de frecuencias. |
| **Hoja7 (Guía Repaso)** | Resolución completa de casos de estudio: Auditoría de los 3 errores conceptuales del flujo de caja (Ejercicio 7), proyecto a 10 años con ampliación, reemplazo de maquinaria, venta con valor libro y exclusión de costos hundidos (Ejercicio 8), y criterios teóricos VAN vs. TIR vs. Payback (Ejercicios 1 al 6). |

---

## Casos Preconfigurados en el Menú Lateral

* **Caso 1 — Estructura J ($120M):** Horizonte de 6 años, impuesto de 29,5%, $K_e$ de 13,0%, financiamiento con 30% de deuda a una tasa $K_d$ de 6,5% a 5 años, inversión en maquinaria por $120.000.000 netos, gastos de puesta en marcha, pruebas de producción, capacitación y capital de trabajo por $36.000.000.
* **Caso 2 — Guía P2 ($130M):** Horizonte de 6 años, impuesto de 29,5%, $K_e$ de 13,0%, deuda del 30% a $K_d$ de 7,5% por 5 años, activo principal de $130.000.000 con valor de rescate de $55.000.000 y capital de trabajo de $35.000.000.
* **Caso 3 — Conservas ($13M):** Proyecto agroindustrial a 5 años con impuesto de 29,5%, $K_e$ de 13,0%, deuda del 30% a $K_d$ de 8,5% en 4 años (con 2,0% de gasto de otorgamiento de crédito), múltiples partidas de activo fijo (maquinaria, implementos, terreno, vehículo, construcción), intangibles y tasas de crecimiento variables por año.
* **Caso 4 — En blanco:** Plantilla limpia para ingresar cualquier proyecto nuevo desde cero, permitiendo agregar o eliminar filas de inversión dinámicamente.

---

## Exportación a Excel con Fórmulas Vivas (`Libro1.xlsx`)

El botón superior **Guardar (.xlsx)** compila en memoria un archivo Excel de **6 hojas (`Hoja1` a `Hoja6`)** utilizando `openpyxl`. A diferencia de una exportación plana de datos, el archivo descargado contiene **fórmulas nativas de Excel vinculadas entre hojas**:

* Cálculo automático de WACC, IVA y depreciaciones (`=SUM`, `=(C14-G14)/F14`).
* Tabla de amortización vinculada con la función financiera `=PMT` (`PAGO`).
* Estados de resultados y flujos de caja vinculados entre hojas (`=Hoja1!C...`, `=Hoja3!D...`).
* Indicadores dinámicos al pie de cada flujo con `=NPV` (`VNA`), `=IRR` (`TIR`) y `=SUMPRODUCT` (`SUMAPRODUCTO`).

---

## Instalación y Uso Local

```bash
git clone [https://github.com/mauriciovargas1516-ux/FINANZAS-AVANZADAS-2026.git](https://github.com/mauriciovargas1516-ux/FINANZAS-AVANZADAS-2026.git)
cd FINANZAS-AVANZADAS-2026
pip install -r requirements.txt
streamlit run app.py
```
