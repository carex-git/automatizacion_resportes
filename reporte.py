import subprocess
import sys
import os
base_dir = os.path.dirname(os.path.abspath(__file__))

# Lista de librerías necesarias
required_libraries = ['pandas','openpyxl','matplotlib','reportlab']

# Instalar automáticamente las librerías que falten
for lib in required_libraries:
    try:
        __import__(lib)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", lib])

print(sys.executable)
print("Todas las librerías están instaladas. Continuando con el reporte...")

import pandas as pd
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt

# Función para convertir formato colombiano a número (si es texto)
def convertir_formato_colombiano(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)  # Ya es numérico, no necesita conversión
    if isinstance(valor, str):
        # Remover puntos (separadores de miles) y reemplazar coma por punto (decimal)
        valor_limpio = valor.replace('.', '').replace(',', '.')
        try:
            return float(valor_limpio)
        except:
            return 0.0
    return float(valor)

# Función para formatear números como en Excel colombiano
def formatear_numero_colombiano(numero):
    if pd.isna(numero) or numero == 0:
        return "0"
    return f"{numero:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

# Cargar archivo de ejecutado
df = pd.read_excel(base_dir+"/data/Carex COL Reporte Vendedor.xlsx", sheet_name='BD')

# Limpiar nombres de columnas (remover espacios)
df.columns = df.columns.str.strip()

print("Columnas disponibles:", df.columns.tolist())
print("Tipos de datos antes:")
print(df.dtypes)

# Los datos ya están en formato numérico correcto, solo verificamos
print("\nValores numéricos (ya están correctos):")
print("Primeros 10 valores en notación normal:")
for i in range(min(10, len(df))):
    valor_cientifico = df['Valor Total USD'].iloc[i]
    valor_formateado = formatear_numero_colombiano(valor_cientifico)
    print(f"{i}: {valor_cientifico} = {valor_formateado}")

# Los datos ya están correctos numéricamente, no necesitamos convertir
# df['Valor Total USD'] ya contiene los valores correctos para cálculos

# Verificar valores únicos en columnas críticas para filtros
print("\nValores únicos en 'Concepto':")
print(df['Concepto'].unique())

print("\nValores únicos en 'Moneda':")
print(df['Moneda'].unique())

print("\nValores únicos en 'Mes':")
print(sorted(df['Mes'].unique()))

# Filtrar vendedores válidos - MEJORADO
vendedores_filtrados = df['Vendedor'].dropna()
vendedores_filtrados = vendedores_filtrados[
    vendedores_filtrados.str.upper() != 'COMERCIALIZADORA INTERNACIONAL CARIBBEAN EXOTICS S A'
]
# Remover espacios adicionales y valores vacíos
vendedores_filtrados = vendedores_filtrados.str.strip()
vendedores_filtrados = vendedores_filtrados[vendedores_filtrados != '']
vendedores_unicos = vendedores_filtrados.unique()

print(f"\nVendedores únicos encontrados: {len(vendedores_unicos)}")
for v in vendedores_unicos:
    print(f"- '{v}'")

# Ítems a excluir
excluir_items = [
    'AIR FREIGHT', 'INV PLANTAS', 'INV PRIMA - FLO', 'INV RECICLAJE', 'OTHER EXPORT COSTS',
    'SEA FREIGHT COST', 'HUMAGRO CALCIUM DRENCH', 'SULPHUR GULUPA X 20 LITROS',
    'HUMAGRO CALCIUM FOLIAR UCH X 20 LITROS', 'HUMAGRO KAGUACATE X 20',
    'INV RECUPERACIONES', 'UCHUVA X 400GR DESGRANAD NACIONAL EURO',
    'CONTENEDOR PET 19,0X12,0X7,5 500 GRS', 'HIGO X 1KG NACIONAL EXITO',
]

# Cargar archivo de Budget
df_BV = pd.read_excel(base_dir+"/data/Carex COL Reporte Vendedor.xlsx", sheet_name='Budget x Vendedor')

# Limpiar nombres de columnas del budget
df_BV.columns = df_BV.columns.str.strip()

# Convertir 'Valor Total USD' del budget también
if 'Valor Total USD' in df_BV.columns:
    df_BV['Valor Total USD'] = df_BV['Valor Total USD'].apply(convertir_formato_colombiano)

print("\nColumnas en Budget:", df_BV.columns.tolist())

# Mes actual
mes_actual = datetime.now().month
dia = datetime.now().day
print(f"\nDía actual: {dia}")
print(f"Mes actual: {mes_actual}")

# Crear lista de resultados
resultados = []

for vendedora in vendedores_unicos:
    print(f"\n--- Procesando vendedora: '{vendedora}' ---")
    
    # FILTRO MEJORADO - Ejecutado
    filtro = (
        (df['Vendedor'].str.strip() == vendedora.strip()) &  # Comparación exacta sin espacios
        (df['Mes'] == mes_actual) &
        (df['Concepto'].str.upper().isin(["FACTURA", "ANULACIÓN FE"])) &  # Usar upper() por si hay diferencias
        (df['Moneda'].str.upper().isin(["USD", "EUR"])) &  # Usar upper() por si hay diferencias
        (~df['Nombre Item'].isin(excluir_items)) &
        (df['Valor Total USD'].notna()) &  # Excluir valores nulos
        (df['Valor Total USD'] != 0)  # Excluir valores cero si es necesario
    )

    df_filtrado = df.loc[filtro]
    total_ejecutado = df_filtrado['Valor Total USD'].sum()

    print(f"Registros encontrados para {vendedora}: {len(df_filtrado)}")
    if len(df_filtrado) > 0:
        print("Primeros registros:")
        print(df_filtrado[['Mes', 'Concepto', 'Nombre Item', 'Valor Total USD']].head())
    
    # FILTRO MEJORADO - Budget
    filtro_bv = (
        (df_BV['Vendedor'].str.strip() == vendedora.strip()) &
        (df_BV['Mes'] == mes_actual) &
        (df_BV['Valor Total USD'].notna()) &  # Excluir valores nulos
        (df_BV['Valor Total USD'] != 0)  # Excluir valores cero si es necesario
    )
    
    df_bv_filtrado = df_BV.loc[filtro_bv]
    total_budget = df_bv_filtrado['Valor Total USD'].sum()
    
    print(f"Total ejecutado: {formatear_numero_colombiano(total_ejecutado)}")
    print(f"Total budget: {formatear_numero_colombiano(total_budget)}")

    # Cálculos
    porcentaje = (total_ejecutado / total_budget * 100) if total_budget > 0 else 0
    faltante = max(0, 100 - porcentaje)  # Asegurar que no sea negativo

    resultados.append({
        "Vendedor": vendedora,
        "Budget": round(total_budget, 2),
        "Ejecutado": round(total_ejecutado, 2),
        "% Ejecución": round(porcentaje, 2),
        "% Faltante": round(faltante, 2),
        "Meta": 100.0,
        "Mes": mes_actual,
    })

# Exportar a Excel
df_resultado = pd.DataFrame(resultados)

# Agregar TOTAL COMPAÑÍA - CORREGIDO
if len(df_resultado) > 0:
    total_budget = df_resultado["Budget"].sum()
    total_ejecutado = df_resultado["Ejecutado"].sum()
    total_porcentaje_real = (total_ejecutado / total_budget * 100) if total_budget > 0 else 0
    total_faltante = max(0, 100 - total_porcentaje_real)

    df_total = pd.DataFrame([{
        "Vendedor": "TOTAL COMPAÑÍA",
        "Budget": round(total_budget, 2),
        "Ejecutado": round(total_ejecutado, 2),
        "% Ejecución": round(total_porcentaje_real, 2),
        "% Faltante": round(total_faltante, 2),
        "Meta": 100.0,
        "Mes": mes_actual
    }])

    df_resultado = pd.concat([df_resultado, df_total], ignore_index=True)


excel_path = os.path.join(base_dir, "reporte_vendedoras.xlsx")
df_resultado.to_excel(excel_path, index=False)

print(f"\n✅ Excel generado como 'reporte_vendedoras.xlsx'")

def dividir_nombre(nombre, max_len=20):
    partes = nombre.split()
    linea = ""
    lineas = []
    for parte in partes:
        if len(linea + parte) < max_len:
            linea += parte + " "
        else:
            lineas.append(linea.strip())
            linea = parte + " "
    lineas.append(linea.strip())
    return "\n".join(lineas)

# Crear gráfico solo si hay datos
if len(df_resultado) > 0:
    image_path = os.path.join(base_dir, "grafico_ejecucion_vendedor.png")

    # Crear gráfico de barras apiladas
    fig, ax = plt.subplots(figsize=(12, 6))

    vendedoras = df_resultado["Vendedor"]
    ejecucion = df_resultado["% Ejecución"]
    faltante = df_resultado["% Faltante"]
    x = np.arange(len(vendedoras))
    bar_width = 0.4

    # Colores personalizados
    color_ejecucion = "#89CFF0"  # azul claro
    color_faltante = "#1f4e79"   # azul oscuro

    # Barras
    bars1 = ax.bar(x, ejecucion, width=bar_width, label="% Ejecución", color=color_ejecucion)
    bars2 = ax.bar(x, faltante, width=bar_width, bottom=ejecucion, label="% Faltante", color=color_faltante)

    # Etiquetas encima de cada segmento
    for i in range(len(x)):
        if ejecucion[i] > 5:  # Solo mostrar etiqueta si hay espacio
            ax.text(x[i], ejecucion[i]/2, f'{ejecucion[i]:.2f}%', ha='center', va='center', color='black', fontsize=9, fontweight='bold')
        if faltante[i] > 5:  # Solo mostrar etiqueta si hay espacio
            ax.text(x[i], ejecucion[i] + faltante[i]/2, f'{faltante[i]:.2f}%', ha='center', va='center', color='white', fontsize=9, fontweight='bold')

    # Personalización
    ax.set_title("Ejecución vs Faltante por Vendedor (Meta = 100%)", fontsize=14, fontweight='bold')
    ax.set_ylabel("Porcentaje", fontsize=12)
    ax.set_xlabel("Vendedor", fontsize=12, labelpad=20)
    ax.set_ylim(0, 110)
    ax.set_xticks(x)
    vendedoras_wrap = [dividir_nombre(v) for v in vendedoras]
    ax.set_xticklabels(vendedoras_wrap, rotation=0, ha="center", fontsize=10)

    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.25), ncol=2)
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.5)

    plt.tight_layout()
    plt.savefig(image_path, dpi=300)
    plt.show()
else:
    print("⚠️ No se encontraron datos para generar el gráfico")