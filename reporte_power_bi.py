import subprocess
import sys
import os
from datetime import datetime
import pandas as pd

# Asegurar instalación de librerías necesarias
required_libraries = ['pandas', 'openpyxl']
for lib in required_libraries:
    try:
        __import__(lib)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", lib])

# Rutas
base_dir = os.path.dirname(os.path.abspath(__file__))
excel_input_path = os.path.join(base_dir, "data/Carex COL Reporte Vendedor.xlsx")
fecha_actual = datetime.now().strftime("%Y-%m-%d")
excel_resumen_path = os.path.join(base_dir, f"reporte_total_anual_{fecha_actual}.xlsx")

# Leer datos
df = pd.read_excel(excel_input_path, sheet_name='BD')
df.columns = df.columns.str.strip()

# Filtros
df['Vendedor'] = df['Vendedor'].str.strip()
df['Nombre Item'] = df['Nombre Item'].str.strip()
df['Concepto'] = df['Concepto'].str.strip()
df['Moneda'] = df['Moneda'].str.strip()

excluir_items = [
    'AIR FREIGHT', 'INV PLANTAS', 'INV PRIMA - FLO', 'INV RECICLAJE', 'OTHER EXPORT COSTS',
    'SEA FREIGHT COST', 'HUMAGRO CALCIUM DRENCH', 'SULPHUR GULUPA X 20 LITROS',
    'HUMAGRO CALCIUM FOLIAR UCH X 20 LITROS', 'HUMAGRO KAGUACATE X 20',
    'INV RECUPERACIONES', 'UCHUVA X 400GR DESGRANAD NACIONAL EURO',
    'CONTENEDOR PET 19,0X12,0X7,5 500 GRS', 'HIGO X 1KG NACIONAL EXITO',
]

# Aplicar filtros
df_filtrado = df[
    (df['Concepto'].isin(['FACTURA', 'ANULACIÓN FE'])) &
    (df['Moneda'].isin(["USD", "EUR"])) &
    (~df['Nombre Item'].isin(excluir_items)) &
    (df['Vendedor'].str.upper() != 'COMERCIALIZADORA INTERNACIONAL CARIBBEAN EXOTICS S A')
]

# Agrupar y resumir
df_resumen = df_filtrado.groupby('Vendedor', as_index=False)['Valor Total USD'].sum()
df_resumen.rename(columns={'Valor Total USD': 'Total Anual USD'}, inplace=True)

# Agregar total general
total_general = df_resumen['Total Anual USD'].sum()
df_resumen.loc[len(df_resumen)] = ['TOTAL COMPAÑÍA', total_general]

# Guardar resumen final a Excel
df_resumen.to_excel(excel_resumen_path, index=False)

print(f"✅ Resumen anual guardado en: '{excel_resumen_path}'")
