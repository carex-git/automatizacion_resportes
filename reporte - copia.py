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
# Resto de tu código aquí...
print("Todas las librerías están instaladas. Continuando con el reporte...")


import pandas as pd
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt


# Cargar archivo de ejecutado
df = pd.read_csv(
    base_dir+"/data/Carex COL Reporte Vendedor(BD).csv",
    encoding="latin1",
    sep=";",
    on_bad_lines="skip"
)

# Conversión de valores a número
def convert_currency(value):
    if pd.isna(value):
        return 0.0
    value = str(value).strip().replace('.', '').replace(',', '.')
    try:
        return float(value)
    except ValueError:
        return 0.0

df['Valor Total USD'] = df['Valor Total USD'].apply(convert_currency)

df['Moneda'] = df['Moneda'].dropna()


# Filtrar vendedores válidos
vendedores_filtrados = df['Vendedor'].dropna()
vendedores_filtrados = vendedores_filtrados[
    vendedores_filtrados.str.upper() != 'COMERCIALIZADORA INTERNACIONAL CARIBBEAN EXOTICS S A'
]
vendedores_unicos = vendedores_filtrados.unique()

# Ítems a excluir
excluir_items = [
    'AIR FREIGHT', 'INV PLANTAS', 'INV PRIMA - FLO', 'INV RECICLAJE', 'OTHER EXPORT COST',
    'SEA FREIGHT COST', 'HUMAGRO CALCIUM DRENCH', 'SULPHUR GULUPA X 20 LITROS',
    'HUMAGRO CALCIUM FOLIAR UCH X 20 LITROS', 'HUMAGRO KAGUACATE X 20',
    'INV RECUPERACIONES', 'UCHUVA X 400GR DESGRANAD NACIONAL EURO',
    'CONTENEDOR PET 19,0X12,0X7,5 500 GRS', 'HIGO X 1KG NACIONAL EXITO',
]

# Cargar archivo de Budget
df_BV = pd.read_csv(base_dir+"/data/Carex COL Reporte Vendedor(Budget x Vendedor).csv", sep=";")
df_BV['Valor Total USD'] = df_BV['Valor Total USD'].apply(convert_currency)

# Asegurar que el mes también esté en formato texto con 2 dígitos
df_BV['Periodo'].astype(str).str[-3:]

# Mes actual
mes_actual = datetime.now().month
mes_actual_float = float(f"{mes_actual}")

# Crear lista de resultados
resultados = []

for vendedora in vendedores_unicos:
    # Ejecutado
    filtro = (
        (df['Vendedor'] == vendedora) &
        (df['Mes'] == mes_actual_float) &
        (df['Concepto'] == "FACTURA") &
        (df['Moneda'].isin(["USD", "EUR"])) &
        (~df['Nombre Item'].isin(excluir_items))
    )

    df_filtrado = df.loc[filtro]
    total_ejecutado = df_filtrado['Valor Total USD'].sum()

    print(df_filtrado[['Mes', 'Concepto', 'Nombre Item', 'Valor Total USD']].head(10))  # Ve
    # Budget
    filtro_bv = (
        (df_BV['Vendedor'] == vendedora) &
        (df_BV['Mes'] == mes_actual_float)
    )
    df_bv_filtrado = df_BV.loc[filtro_bv]
    
    total_budget = df_bv_filtrado['Valor Total USD'].sum()
    
    print(total_ejecutado)

    # Cálculos
    porcentaje = (total_ejecutado / total_budget * 100) if total_budget else 0
    faltante = 100 - porcentaje

    resultados.append({
        "Vendedor": vendedora,
        "Budget": round(total_budget, 2),
        "% Ejecución": round(porcentaje, 2),
        "% Faltante": round(faltante, 2),
        "Meta": porcentaje +faltante,
        "Mes": mes_actual_float,
        
    })

# Exportar a Excel
df_resultado = pd.DataFrame(resultados)

# Agregar TOTAL COMPAÑÍA
total_budget = df_resultado["Budget"].sum()
total_porcentaje = df_resultado["% Ejecución"].mean()  # O usa sum ejecutado / sum budget
total_ejecutado = (df_resultado["Budget"] * df_resultado["% Ejecución"] / 100).sum()
total_porcentaje_real = (total_ejecutado / total_budget) * 100 if total_budget else 0
total_faltante = 100 - total_porcentaje_real

df_total = pd.DataFrame([{
    "Vendedor": "TOTAL COMPAÑÍA",
    "Budget": round(total_budget, 2),
    "% Ejecución": round(total_porcentaje_real, 2),
    "% Faltante": round(total_faltante, 2),
    "Meta": 100.0,
    "Mes": mes_actual_float
}])

df_resultado = pd.concat([df_resultado, df_total], ignore_index=True)


excel_path = os.path.join(base_dir, "reporte_vendedoras.xlsx")
df_resultado.to_excel(excel_path, index=False)

print("✅ Excel generado como 'reporte_vendedoras.xlsx'")

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



# Leer el archivo generado
# Leer el archivo generado
df_resultado = pd.read_excel(base_dir+"reporte_vendedoras.xlsx")

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
    ax.text(x[i], ejecucion[i]/2, f'{ejecucion[i]:.2f}%', ha='center', va='center', color='black', fontsize=9, fontweight='bold')
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