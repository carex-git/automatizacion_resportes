import subprocess
import sys
import os
base_dir = os.path.dirname(os.path.abspath(__file__))

required_libraries = ['pandas', 'openpyxl', 'matplotlib']
for lib in required_libraries:
    try:
        __import__(lib)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", lib])

import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

def convertir_formato_colombiano(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, str):
        valor_limpio = valor.replace('.', '').replace(',', '.')
        try:
            return float(valor_limpio)
        except:
            return 0.0
    return float(valor)

archivo = os.path.join(base_dir, "data/Carex COL Reporte Vendedor.xlsx")
df = pd.read_excel(archivo, sheet_name='BD')
df.columns = df.columns.str.strip()

df_BV = pd.read_excel(archivo, sheet_name='Budget x Vendedor')
df_BV.columns = df_BV.columns.str.strip()
df_BV['Valor Total USD'] = df_BV['Valor Total USD'].apply(convertir_formato_colombiano)

nombre_mes = 'Total'

excluir_items = [
    'AIR FREIGHT', 'INV PLANTAS', 'INV PRIMA - FLO', 'INV RECICLAJE', 'OTHER EXPORT COSTS',
    'SEA FREIGHT COST', 'HUMAGRO CALCIUM DRENCH', 'SULPHUR GULUPA X 20 LITROS',
    'HUMAGRO CALCIUM FOLIAR UCH X 20 LITROS', 'HUMAGRO KAGUACATE X 20',
    'INV RECUPERACIONES', 'UCHUVA X 400GR DESGRANAD NACIONAL EURO',
    'CONTENEDOR PET 19,0X12,0X7,5 500 GRS', 'HIGO X 1KG NACIONAL EXITO',
]

df = df[df['Concepto'].str.upper() == 'FACTURA']
df = df[df['Valor Total USD'] > 0]
df = df[~df['Nombre Item'].isin(excluir_items)]
df = df[df['Moneda'].str.upper().isin(['USD', 'EUR'])]

ventas_total = df.groupby('Vendedor')['Valor Total USD'].sum().reset_index()
ventas_total.columns = ['Vendedor', 'Ejecutado']

budget_total = df_BV.groupby('Vendedor')['Valor Total USD'].sum().reset_index()
budget_total.columns = ['Vendedor', 'Budget']

df_merge = pd.merge(budget_total, ventas_total, on='Vendedor', how='outer').fillna(0)

df_merge['% Ejecución'] = df_merge.apply(
    lambda row: (row['Ejecutado'] / row['Budget']) * 100 if row['Budget'] > 0 else 0, axis=1)
df_merge['% Faltante'] = 100 - df_merge['% Ejecución']
df_merge['Meta'] = df_merge['Budget']
df_merge['Mes'] = nombre_mes

columnas_orden = ['Vendedor', 'Budget', 'Ejecutado', '% Ejecución', '% Faltante', 'Meta', 'Mes']
df_merge = df_merge[columnas_orden]

excel_path = os.path.join(base_dir, f"reporte_total_ejecucion_{datetime.now().strftime('%Y-%m-%d')}.xlsx")
df_merge.to_excel(excel_path, index=False)
print(f"✅ Archivo generado: {excel_path}")

# Gráfico
df_merge_plot = df_merge.sort_values('% Ejecución', ascending=False)
plt.figure(figsize=(12, 6))
plt.bar(df_merge_plot['Vendedor'], df_merge_plot['% Ejecución'], color='skyblue')
plt.axhline(y=100, color='red', linestyle='--', label='100%')
plt.xticks(rotation=45, ha='right')
plt.ylabel('% Ejecución')
plt.title('Porcentaje de Ejecución por Vendedor - TOTAL')
plt.tight_layout()
plt.legend()
plt.show()
