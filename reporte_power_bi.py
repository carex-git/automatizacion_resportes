import subprocess
import sys
import os
import math
from datetime import datetime
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont

# =============================================================================
# 1. CONFIGURACIÓN Y PREPARACIÓN
# =============================================================================

def install_required_libraries():
    """Verifica e instala las librerías necesarias si no están presentes."""
    required_libraries = ['pandas', 'openpyxl', 'plotly', 'kaleido', 'Pillow']
    for lib in required_libraries:
        try:
            __import__(lib)
        except ImportError:
            print(f"Instalando {lib}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", lib])

install_required_libraries()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

INPUT_FILENAME = "Carex COL Reporte Vendedor.xlsx"
INPUT_PATH = os.path.join(DATA_DIR, INPUT_FILENAME)
FECHA_ACTUAL = datetime.now().strftime("%Y-%m-%d")
ANIO_ACTUAL = datetime.now().year
MES_ACTUAL = datetime.now().month

# Diccionario para convertir el número del mes a su nombre en español
MESES_ESPANOL = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
    5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
    9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
}
MES_ACTUAL_NOMBRE = MESES_ESPANOL.get(MES_ACTUAL, str(MES_ACTUAL))

EXCLUIR_ITEMS = {
    'AIR FREIGHT', 'INV PLANTAS', 'INV PRIMA - FLO', 'INV RECICLAJE', 'OTHER EXPORT COSTS',
    'SEA FREIGHT COST', 'HUMAGRO CALCIUM DRENCH', 'SULPHUR GULUPA X 20 LITROS',
    'HUMAGRO CALCIUM FOLIAR UCH X 20 LITROS', 'HUMAGRO KAGUACATE X 20',
    'INV RECUPERACIONES', 'UCHUVA X 400GR DESGRANAD NACIONAL EURO',
    'CONTENEDOR PET 19,0X12,0X7,5 500 GRS', 'HIGO X 1KG NACIONAL EXITO',
}

# =============================================================================
# 2. FUNCIONES PRINCIPALES
# =============================================================================

def load_and_clean_data(file_path):
    """Carga y limpia los dataframes desde el archivo Excel."""
    print("📊 Cargando y limpiando datos...")
    required_columns = [
        'Año', 'Mes', 'Nombre Cliente_factura', 'Nombre Centro de Operacion',
        'Valor Total USD', 'Concepto', 'Moneda', 'Nombre Item', 'Vendedor', 'Desc Pais Cliente_factura'
    ]
    try:
        df = pd.read_excel(file_path, sheet_name='BD', usecols=required_columns)
        df_bv = pd.read_excel(file_path, sheet_name='Budget x Vendedor')
    except FileNotFoundError:
        print(f"❌ ERROR: Archivo no encontrado en la ruta: {file_path}")
        sys.exit()
    except ValueError as e:
        print(f"❌ ERROR: Fallo al leer el archivo Excel. Mensaje: {e}")
        sys.exit()
        
    df.columns = df.columns.str.strip()
    df_bv.columns = df_bv.columns.str.strip()

    df_filtered = df[
        (df['Concepto'].str.upper().isin(['FACTURA', 'ANULACIÓN FE'])) &
        (df['Moneda'].str.upper().isin(['USD', 'EUR'])) &
        (~df['Nombre Item'].isin(EXCLUIR_ITEMS)) &
        (df['Valor Total USD'].notna()) &
        (df['Valor Total USD'] != 0) &
        (df['Vendedor'].str.strip().str.upper() != 'COMERCIALIZADORA INTERNACIONAL CARIBBEAN EXOTICS S A')
    ].copy()
    return df_filtered, df_bv

def perform_analysis(df_filtered, df_bv):
    """Realiza todos los cálculos y agregaciones de los datos solicitados."""
    print("🔍 Realizando análisis...")
    df_anual = df_filtered[df_filtered['Año'] == ANIO_ACTUAL]
    df_mensual = df_anual[df_anual['Mes'] == MES_ACTUAL]

    # Cálculos anuales
    ventas_sede_anual = df_anual.groupby('Nombre Centro de Operacion')['Valor Total USD'].sum().sort_values(ascending=False)
    top_clientes_anual = df_anual.groupby('Nombre Cliente_factura')['Valor Total USD'].sum().sort_values(ascending=False).head(5)
    top_paises_anual = df_anual.groupby('Desc Pais Cliente_factura')['Valor Total USD'].sum().sort_values(ascending=False).head(4)
    ejecutado_anual = df_anual['Valor Total USD'].sum()
    budget_anual = df_bv['Valor Total USD'].sum()

    # Cálculos mensuales
    ventas_sede_mensual = df_mensual.groupby('Nombre Centro de Operacion')['Valor Total USD'].sum().sort_values(ascending=False)
    top_clientes_mensual = df_mensual.groupby('Nombre Cliente_factura')['Valor Total USD'].sum().sort_values(ascending=False).head(5)
    top_paises_mensual = df_mensual.groupby('Desc Pais Cliente_factura')['Valor Total USD'].sum().sort_values(ascending=False).head(4)
    ejecutado_mensual = df_mensual['Valor Total USD'].sum()
    budget_mensual = df_bv[df_bv['Mes'] == MES_ACTUAL]['Valor Total USD'].sum()
    
    return (ventas_sede_anual, ventas_sede_mensual, top_clientes_anual, top_clientes_mensual,
            top_paises_anual, top_paises_mensual, budget_anual, ejecutado_anual,
            budget_mensual, ejecutado_mensual)

def create_plots_in_memory(analysis_data):
    """Crea cada gráfico y tabla como un objeto BytesIO en memoria."""
    print("🎨 Creando gráficos en memoria...")
    (ventas_sede_anual, ventas_sede_mensual, top_clientes_anual, top_clientes_mensual,
     top_paises_anual, top_paises_mensual, budget_anual, ejecutado_anual,
     budget_mensual, ejecutado_mensual) = analysis_data

    colores_carex = ['#003366', '#ff7f0e', "#3d2ca0", '#d62728', '#9467bd']

    def create_plot_bytes(fig, title, width, height):
        fig.update_layout(
            title_text=f"<b>{title}</b>",
            title_x=0.5,
            title_font_size=24,
            height=height,
            width=width,
            plot_bgcolor='rgba(240,240,240,0.8)',
            paper_bgcolor='white',
            margin=dict(l=50, r=50, b=50, t=80)
        )
        try:
            img_bytes = fig.to_image(format="jpeg", scale=2)
            return BytesIO(img_bytes)
        except Exception as e:
            print(f"❌ ERROR al crear la imagen '{title}': {e}.")
            return None

    # Gráfico de Medidor Anual
    fig_gauge_anual = go.Figure(go.Indicator(
        mode="gauge+number",
        value=ejecutado_anual,
        number={'valueformat': '$,.2f', 'font': {'size': 50}},
        title={'text': f"<b>Venta Acumulada USD Anual {ANIO_ACTUAL}</b>", 'font': {'size': 20}},
        gauge={
            'axis': {'range': [0, budget_anual]},
            'bar': {'color': "#003366"},
            'steps': [{'range': [0, budget_anual], 'color': 'lightgray'}],
            'threshold': {'value': budget_anual}
        }
    ))
    gauge_anual_bytes = create_plot_bytes(fig_gauge_anual, "", 600, 400)
    
    # Gráfico de Medidor Mensual
    fig_gauge_mensual = go.Figure(go.Indicator(
        mode="gauge+number",
        value=ejecutado_mensual,
        number={'valueformat': '$,.2f', 'font': {'size': 50}},
        title={'text': f"<b>Venta Acumulada USD Mensual ({MES_ACTUAL_NOMBRE})</b>", 'font': {'size': 20}},
        gauge={
            'axis': {'range': [0, budget_mensual]},
            'bar': {'color': "#003366"},
            'steps': [{'range': [0, budget_mensual], 'color': 'lightgray'}],
            'threshold': {'value': budget_mensual}
        }
    ))
    gauge_mensual_bytes = create_plot_bytes(fig_gauge_mensual, "", 600, 400)
    
    # Gráfico de Pie de Sede Anual
    fig_anual = go.Figure(data=go.Pie(labels=ventas_sede_anual.index, values=ventas_sede_anual.values, textinfo='value+percent', texttemplate='$%{value:,.0f}<br>(%{percent})', insidetextfont={'size': 16, 'color': 'white'}, hoverinfo='label+percent+value', marker_colors=colores_carex, hole=0.5))
    total_anual = ventas_sede_anual.sum()
    if total_anual > 0: fig_anual.add_annotation(text=f"Total<br><b>${total_anual:,.0f}</b>", x=0.5, y=0.5, font=dict(size=20, color='#003366'), showarrow=False)
    pie_anual_bytes = create_plot_bytes(fig_anual, f"Ventas por Sede Anual ({ANIO_ACTUAL})", 1200, 800)

    # Gráfico de Pie de Sede Mensual
    fig_mensual = go.Figure(data=go.Pie(labels=ventas_sede_mensual.index, values=ventas_sede_mensual.values, textinfo='value+percent', texttemplate='$%{value:,.0f}<br>(%{percent})', insidetextfont={'size': 16, 'color': 'white'}, hoverinfo='label+percent+value', marker_colors=colores_carex, hole=0.5))
    total_mensual = ventas_sede_mensual.sum()
    if total_mensual > 0: fig_mensual.add_annotation(text=f"Total<br><b>${total_mensual:,.0f}</b>", x=0.5, y=0.5, font=dict(size=20, color='#003366'), showarrow=False)
    pie_mensual_bytes = create_plot_bytes(fig_mensual, f"Ventas por Sede ({MES_ACTUAL_NOMBRE})", 1200, 800)

    # Gráfico de Barras de Países Anual
    fig_paises_anual = go.Figure(data=go.Bar(x=top_paises_anual.index, y=top_paises_anual.values, marker_color=colores_carex, text=top_paises_anual.values, texttemplate='$%{text:,.0f}', textposition='outside'))
    fig_paises_anual.update_layout(xaxis_title_text='País', yaxis_title_text='Ventas USD')
    bar_paises_anual_bytes = create_plot_bytes(fig_paises_anual, f"Top 4 Ventas por País Anual ({ANIO_ACTUAL})", 1200, 800)

    # Gráfico de Barras de Países Mensual
    fig_paises_mensual = go.Figure(data=go.Bar(x=top_paises_mensual.index, y=top_paises_mensual.values, marker_color=colores_carex, text=top_paises_mensual.values, texttemplate='$%{text:,.0f}', textposition='outside'))
    fig_paises_mensual.update_layout(xaxis_title_text='País', yaxis_title_text='Ventas USD')
    bar_paises_mensual_bytes = create_plot_bytes(fig_paises_mensual, f"Top 4 Ventas por País ({MES_ACTUAL_NOMBRE})", 1200, 800)

    # Tabla de Clientes Anual
    fig_top_clientes_anual = go.Figure(data=[go.Table(
        header=dict(values=['<b>Cliente</b>', f'<b>Ventas {ANIO_ACTUAL} ($)</b>'], align=['left', 'right'], font=dict(color='white', size=16), fill_color='#003366', height=40),
        cells=dict(values=[top_clientes_anual.index, [f'${x:,.0f}' for x in top_clientes_anual.values]], align=['left', 'right'], fill_color=[['white', '#f0f0f0'] * (len(top_clientes_anual) // 2 + 1)], font=dict(color='black', size=14), height=30)
    )])
    tabla_clientes_anual_bytes = create_plot_bytes(fig_top_clientes_anual, f"Top 5 Clientes Anual ({ANIO_ACTUAL})", 800, 300)

    # Tabla de Clientes Mensual
    fig_top_clientes_mensual = go.Figure(data=[go.Table(
        header=dict(values=['<b>Cliente</b>', f'<b>Ventas Mes ({MES_ACTUAL_NOMBRE}) ($)</b>'], align=['left', 'right'], font=dict(color='white', size=16), fill_color='#003366', height=40),
        cells=dict(values=[top_clientes_mensual.index, [f'${x:,.0f}' for x in top_clientes_mensual.values]], align=['left', 'right'], fill_color=[['white', '#f0f0f0'] * (len(top_clientes_mensual) // 2 + 1)], font=dict(color='black', size=14), height=30)
    )])
    tabla_clientes_mensual_bytes = create_plot_bytes(fig_top_clientes_mensual, f"Top 5 Clientes ({MES_ACTUAL_NOMBRE})", 800, 300)

    return [
        gauge_anual_bytes, gauge_mensual_bytes, pie_anual_bytes, pie_mensual_bytes,
        bar_paises_anual_bytes, bar_paises_mensual_bytes, tabla_clientes_anual_bytes,
        tabla_clientes_mensual_bytes
    ]


def combine_images_into_single_report(image_bytes_list):
    """
    Combina las imágenes individuales en memoria en una sola imagen de reporte.
    """
    print("🖼️ Combinando gráficos en un reporte consolidado...")
    if any(img_bytes is None for img_bytes in image_bytes_list):
        print("❌ No se pudieron crear todas las imágenes en memoria para combinar. Saliendo...")
        return

    images = [Image.open(img_bytes) for img_bytes in image_bytes_list]

    gauge_width, gauge_height = images[0].size
    pie_width, pie_height = images[2].size
    bar_width, bar_height = images[4].size
    tabla_width, tabla_height = images[6].size

    combined_width = max(gauge_width * 2 + 100, pie_width * 2 + 100, bar_width * 2 + 100, tabla_width * 2 + 100)
    combined_height = sum([gauge_height, pie_height, bar_height, tabla_height]) + 250
    
    final_report = Image.new('RGB', (combined_width, combined_height), 'white')
    draw = ImageDraw.Draw(final_report)

    x_offset = (combined_width - gauge_width * 2 - 100) / 2
    y_offset = 100
    final_report.paste(images[0], (int(x_offset), y_offset))
    final_report.paste(images[1], (int(x_offset + gauge_width + 100), y_offset))

    y_offset += gauge_height + 50
    x_offset = (combined_width - pie_width * 2 - 100) / 2
    final_report.paste(images[2], (int(x_offset), y_offset))
    final_report.paste(images[3], (int(x_offset + pie_width + 100), y_offset))

    y_offset += pie_height + 50
    x_offset = (combined_width - bar_width * 2 - 100) / 2
    final_report.paste(images[4], (int(x_offset), y_offset))
    final_report.paste(images[5], (int(x_offset + bar_width + 100), y_offset))

    y_offset += bar_height + 50
    x_offset = (combined_width - tabla_width * 2 - 100) / 2
    final_report.paste(images[6], (int(x_offset), y_offset))
    final_report.paste(images[7], (int(x_offset + tabla_width + 100), y_offset))

    try:
        font_path = "arial.ttf"
        font = ImageFont.truetype(font_path, 80)
    except IOError:
        font = ImageFont.load_default()

    title_text = f"Reporte de Ventas ({FECHA_ACTUAL})"
    bbox = draw.textbbox((0, 0), title_text, font=font)
    text_width = bbox[2] - bbox[0]

    text_x = (combined_width - text_width) / 2
    draw.text((text_x, 15), title_text, fill='black', font=font)

    final_path = os.path.join(OUTPUT_DIR, f"dashboard_consolidado_{FECHA_ACTUAL}.jpg")
    final_report.save(final_path, 'JPEG', quality=95)
    print(f"✅ Dashboard consolidado guardado en: {final_path}")

def generate_excel_report(df, df_bv):
    """Genera el reporte de Excel anual por vendedor usando operaciones vectorizadas."""
    print("📋 Generando reporte de Excel anual...")
    output_path = os.path.join(OUTPUT_DIR, f"reporte_vendedoras_anual_{FECHA_ACTUAL}.xlsx")

    df_ejecutado = df.groupby('Vendedor')['Valor Total USD'].sum().rename('Ejecutado Total USD')
    df_budget = df_bv.groupby('Vendedor')['Valor Total USD'].sum().rename('Budget Total USD')

    df_resultado = pd.concat([df_ejecutado, df_budget], axis=1).fillna(0)
    df_resultado['% Ejecución Anual'] = (df_resultado['Ejecutado Total USD'] / df_resultado['Budget Total USD'] * 100).fillna(0)
    df_resultado['% Faltante'] = 100 - df_resultado['% Ejecución Anual']
    df_resultado['Meta'] = 100.0
    
    # Excluir el vendedor de la compañía si existe
    df_resultado = df_resultado[df_resultado.index.str.upper() != 'COMERCIALIZADORA INTERNACIONAL CARIBBEAN EXOTICS S A']

    # Crear la fila de totales
    total_row = pd.DataFrame([{
        'Vendedor': 'TOTAL COMPAÑÍA',
        'Budget Total USD': df_resultado['Budget Total USD'].sum(),
        'Ejecutado Total USD': df_resultado['Ejecutado Total USD'].sum(),
        '% Ejecución Anual': (df_resultado['Ejecutado Total USD'].sum() / df_resultado['Budget Total USD'].sum() * 100) if df_resultado['Budget Total USD'].sum() > 0 else 0,
        '% Faltante': 100 - ((df_resultado['Ejecutado Total USD'].sum() / df_resultado['Budget Total USD'].sum() * 100) if df_resultado['Budget Total USD'].sum() > 0 else 0),
        'Meta': 100.0
    }])
    
    df_final = df_resultado.reset_index().rename(columns={'index': 'Vendedor'})
    df_final = pd.concat([df_final, total_row], ignore_index=True)

    df_final.to_excel(output_path, index=False)
    print(f"✅ Reporte anual de vendedores generado en: {output_path}")

# =============================================================================
# 3. BLOQUE DE EJECUCIÓN PRINCIPAL
# =============================================================================

def main():
    """Función principal que orquesta el flujo de trabajo."""
    df_filtered, df_bv = load_and_clean_data(INPUT_PATH)

    if df_filtered.empty:
        print("Advertencia: No se encontraron datos que cumplan con los filtros.")
        sys.exit()

    analysis_data = perform_analysis(df_filtered, df_bv)
    
    individual_image_bytes = create_plots_in_memory(analysis_data)
    
    combine_images_into_single_report(individual_image_bytes)
    
    generate_excel_report(df_filtered, df_bv)

if __name__ == "__main__":
    main()