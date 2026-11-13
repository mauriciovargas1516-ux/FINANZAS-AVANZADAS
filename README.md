# Simulador de Finanzas Avanzadas y Evaluacion de Activos Reales

Aplicacion web interactiva desarrollada en Python y Streamlit para la modelacion de presupuestos de capital, evaluacion de proyectos corporativos, analisis de estructura de financiamiento y simulacion estocastica de riesgo.

El sistema cuenta con una interfaz visual inspirada en una hoja de calculo financiera y permite exportar todos los modelos calculados a un archivo Excel (`Libro1.xlsx`) de 6 hojas con formulas dinamicas nativas.

---

## 1.Que Hace Este Sistema

Esta herramienta resuelve de forma automatizada el ciclo completo de evaluacion financiera de inversiones en activos reales:

1. **Separacion Tributaria y CAPEX (Ano 0):** Convierte automaticamente montos brutos a valores netos descontando el IVA (19%), calcula el IVA Credito Fiscal, determina la depreciacion lineal de activos fijos (con opcion de descontar el valor de rescate) y la amortizacion de activos intangibles.
2. **Estructura de Capital y Costo de Financiamiento:** Calcula el Costo Promedio Ponderado de Capital (WACC) y el Costo del Capital Propio (Ke) mediante una calculadora integrada del modelo CAPM. Genera la tabla de amortizacion del credito bancario bajo el metodo frances de cuota fija.
3. **Flujo de Caja Libre (FCL) y Flujo del Inversionista (FCI):** Construye ambos flujos proyectados, aplicando el escudo tributario de los intereses de la deuda, el impuesto de Primera Categoria, el reverso de gastos no desembolsables, el valor de rescate de activos y la recuperacion del capital de trabajo.
4. **Analisis de Riesgo e Incertidumbre:** Ejecuta analisis de sensibilidad por escenarios, calculo de VAN Esperado por probabilidades y una Simulacion de Monte Carlo de hasta 50.000 iteraciones con distribuciones triangulares.

---

## 2. Estructura de Hojas y Modulos de la Aplicacion

La plataforma se organiza en 7 pestanas interactivas en pantalla y un panel lateral de configuracion:

| Pestana en Pantalla | Contenido y Cálculos Principales |
| :--- | :--- |
| **Hoja1 (FCL)** | Flujo de Caja Libre (proyecto puro sin deuda). Calcula VAN, TIR, Payback exacto en anos y meses, Flujo Actualizado (VP), Flujo Acumulado y dictamen ejecutivo automatico (Aceptar / Rechazar). |
| **Hoja2 (Amortizacion)** | Tabla de desarrollo del prestamo bancario periodo a periodo: Saldo Inicial, Cuota Fija (`PAGO`), Gasto Financiero (Intereses), Amortizacion de Capital y Saldo Insoluto. |
| **Hoja3 (Inversionista)** | Flujo de Caja del Inversionista (FCI) apalancado. Incorpora el desembolso neto de deuda en el Ano 0, rebaja de intereses antes de impuestos y pago de amortizaciones, descontado a la tasa exigida por el accionista (Ke). |
| **Hoja4 (Inversiones/IVA)** | Cuadro detallado de activos fijos, intangibles y capital de trabajo. Desglosa Valor Neto, IVA (19%), Valor Bruto, Vida Util, Valor de Rescate y cuotas anuales de depreciacion y amortizacion. |
| **Hoja5 (Sensibilidad)** | Comparativa simultanea de escenarios Pesimista (-10% ingresos), Base (100%) y Optimista (+10% ingresos) para FCL y FCI. Incluye calculadora de VAN Esperado E(VAN), Desviacion Estandar y Coeficiente de Variacion. |
| **Hoja6 (Monte Carlo)** | Simulacion estocastica (10.000 iteraciones por defecto) mediante distribuciones triangulares para ingresos, costos variables y costos fijos. Entrega VAN Promedio, Mediana (P50), Probabilidad de Exito (VAN >= 0), Riesgo de Destruccion de Valor, VaR al 95% (Percentil 5%) e histograma de frecuencias. |
| **Hoja7 (Guia Repaso)** | Resolucion completa de casos de estudio: Auditoria de los 3 errores conceptuales del flujo de caja (Ejercicio 7), proyecto a 10 anos con ampliacion, reemplazo de maquinaria, venta con valor libro y exclusion de costos hundidos (Ejercicio 8), y criterios teoricos VAN vs. TIR vs. Payback (Ejercicios 1 al 6). |

---

## 3. Casos Preconfigurados en el Menu Lateral

El selector lateral permite cargar de forma inmediata 4 configuraciones base editables en pantalla:

* **Caso 1 — Estructura J (120M):** Horizonte de 6 anos, impuesto de 29,5%, Ke de 13,0%, financiamiento con 30% de deuda a una tasa Kd de 6,5% a 5 anos, inversion en maquinaria por $120.000.000 netos, gastos de puesta en marcha, pruebas de produccion, capacitacion y capital de trabajo por $36.000.000.
* **Caso 2 — Guia P2 (130M):** Horizonte de 6 anos, impuesto de 29,5%, Ke de 13,0%, deuda del 30% a Kd de 7,5% por 5 anos, activo principal de $130.000.000 con valor de rescate de $55.000.000 y capital de trabajo de $35.000.000.
* **Caso 3 — Conservas (13M):** Proyecto agroindustrial a 5 anos con impuesto de 29,5%, Ke de 13,0%, deuda del 30% a Kd de 8,5% en 4 anos (con 2,0% de gasto de otorgamiento de credito), multiples partidas de activo fijo (maquinaria, implementos, terreno, vehiculo, construccion), intangibles y tasas de crecimiento variables por ano.
* **Caso 4 — En blanco:** Plantilla limpia para ingresar cualquier proyecto nuevo desde cero, permitiendo agregar o eliminar filas de inversion dinamicamente.

---

## 4. Exportacion a Excel con Formulas Vivas (`Libro1.xlsx`)

El boton superior **Guardar (.xlsx)** compila en memoria un archivo Excel de 6 hojas (`Hoja1` a `Hoja6`) utilizando `openpyxl`. A diferencia de una exportacion plana de datos, el archivo descargado contiene formulas nativas de Excel vinculadas entre hojas:

* Calculo automatico de WACC, IVA y depreciaciones (`=SUM`, `=(C14-G14)/F14`).
* Tabla de amortizacion vinculada con la funcion financiera `=PMT` (`PAGO`).
* Estados de resultados y flujos de caja vinculados entre hojas (`=Hoja1!C...`, `=Hoja3!D...`).
* Indicadores dinamicos al pie de cada flujo con `=NPV` (`VNA`), `=IRR` (`TIR`) y `=SUMPRODUCT` (`SUMAPRODUCTO`).

---

## 5. Archivos del Repositorio

| Archivo | Funcion en el Proyecto |
| :--- | :--- |
| `app.py` | Codigo fuente principal en Python y Streamlit. Contiene la interfaz grafica, el motor de flujos de caja, el simulador de Monte Carlo y el generador de Excel con `openpyxl`. |
| `.streamlit/config.toml` | Configuracion del tema visual claro con paleta corporativa verde (`#107c41`) y barra de herramientas minimalista. |
| `requirements.txt` | Lista de librerias requeridas para la ejecucion en Streamlit Cloud o en entorno local (`streamlit`, `pandas`, `numpy`, `openpyxl`). |

---

## 6. Instalacion y Uso Local

1. Clonar este repositorio:
```bash
git clone [https://github.com/mauriciovargas1516-ux/FINANZAS-AVANZADAS.git](https://github.com/mauriciovargas1516-ux/FINANZAS-AVANZADAS.git)
cd FINANZAS-AVANZADAS
