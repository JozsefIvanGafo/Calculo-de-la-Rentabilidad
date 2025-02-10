import numpy as np
import pandas as pd
from scipy.stats import norm
import matplotlib.pyplot as plt

# Paso 1: Cargar datos
def cargar_datos(archivo_excel):
    # Lee el archivo Excel
    df = pd.read_excel(archivo_excel, sheet_name="Hoja1",engine='openpyxl', skiprows=1)
    return df

# Paso 2: Definir parámetros de simulación
def parametros_simulacion():
    # Rangos de variación para ingresos y costes
    ingresos_min, ingresos_max = 50000, 90000  # Ejemplo: rango de ingresos
    costes_min, costes_max = 15000, 30000     # Ejemplo: rango de costes
    inversion_inicial = 140000                # Inversión inicial
    tasa_descuento = 0.1                      # Tasa de descuento para VAN
    num_simulaciones = 1000                   # Número de simulaciones
    return ingresos_min, ingresos_max, costes_min, costes_max, inversion_inicial, tasa_descuento, num_simulaciones

# Paso 3: Simulación de Monte Carlo
def simulacion_monte_carlo(ingresos_min, ingresos_max, costes_min, costes_max, inversion_inicial, tasa_descuento, num_simulaciones):
    resultados_flujo_caja = []
    resultados_tir = []
    
    for _ in range(num_simulaciones):
        # Generar valores aleatorios para ingresos y costes
        ingresos = np.random.uniform(ingresos_min, ingresos_max)
        costes = np.random.uniform(costes_min, costes_max)
        
        # Calcular flujo de caja neto
        flujo_caja_bruto = ingresos - costes
        amortizacion = 14000  # Ejemplo: amortización anual
        impuestos = 0.3 * (flujo_caja_bruto - amortizacion)  # Impuestos al 30%
        flujo_caja_neto = flujo_caja_bruto - impuestos
        
        # Calcular TIR (simplificado para un solo año)
        tir = (flujo_caja_neto / inversion_inicial) - 1
        
        # Guardar resultados
        resultados_flujo_caja.append(flujo_caja_neto)
        resultados_tir.append(tir)
    
    return resultados_flujo_caja, resultados_tir

# Paso 4: Análisis de resultados
def analisis_resultados(resultados_flujo_caja, resultados_tir):
    # Estadísticas descriptivas
    media_flujo_caja = np.mean(resultados_flujo_caja)
    desviacion_flujo_caja = np.std(resultados_flujo_caja)
    media_tir = np.mean(resultados_tir)
    desviacion_tir = np.std(resultados_tir)
    
    print(f"Media del flujo de caja neto: {media_flujo_caja:.2f}")
    print(f"Desviación estándar del flujo de caja neto: {desviacion_flujo_caja:.2f}")
    print(f"Media de la TIR: {media_tir:.2%}")
    print(f"Desviación estándar de la TIR: {desviacion_tir:.2%}")
    
    # Histograma del flujo de caja neto
    plt.hist(resultados_flujo_caja, bins=30, color='blue', alpha=0.7)
    plt.title("Distribución del Flujo de Caja Neto")
    plt.xlabel("Flujo de Caja Neto")
    plt.ylabel("Frecuencia")
    plt.show()

# Paso 5: Ejecutar el programa
if __name__ == "__main__":
    # Cargar datos
    data_path="data\HOJA BASE PARA CALCULO1.xlsx"
    df = cargar_datos(data_path)
    
    # Obtener parámetros de simulación
    ingresos_min, ingresos_max, costes_min, costes_max, inversion_inicial, tasa_descuento, num_simulaciones = parametros_simulacion()
    
    # Realizar simulación
    resultados_flujo_caja, resultados_tir = simulacion_monte_carlo(
        ingresos_min, ingresos_max, costes_min, costes_max, inversion_inicial, tasa_descuento, num_simulaciones
    )
    
    # Analizar resultados
    analisis_resultados(resultados_flujo_caja, resultados_tir)