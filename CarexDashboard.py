import subprocess
import sys
import os
import math
from datetime import datetime
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont
import matplotlib.pyplot as plt

class CarexDashboard:
    def __init__(self):
        self.BASE_DIR = r"C:\Users\aprsistemas\Desktop\trabajo\automatizacion_resportes"
        self.DATA_DIR = os.path.join(self.BASE_DIR, "data")
        self.OUTPUT_DIR = os.path.join(self.BASE_DIR, "output")
        os.makedirs(self.OUTPUT_DIR, exist_ok=True)

        self.INPUT_FILENAME = "Carex COL Reporte Vendedor.xlsx"
        self.INPUT_PATH = os.path.join(self.DATA_DIR, self.INPUT_FILENAME)
        self.FECHA_ACTUAL = datetime.now().strftime("%Y-%m-%d")
        self.ANIO_ACTUAL = datetime.now().year
        self.MES_ACTUAL = datetime.now().month

        self.MESES_ESPANOL = {
            1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
            5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
            9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
        }
        self.MES_ACTUAL_NOMBRE = self.MESES_ESPANOL.get(self.MES_ACTUAL, str(self.MES_ACTUAL))

        self.EXCLUIR_ITEMS = {
            'AIR FREIGHT', 'INV PLANTAS', 'INV PRIMA - FLO', 'INV RECICLAJE', 'OTHER EXPORT COSTS',
            'SEA FREIGHT COST', 'HUMAGRO CALCIUM DRENCH', 'SULPHUR GULUPA X 20 LITROS',
            'HUMAGRO CALCIUM FOLIAR UCH X 20 LITROS', 'HUMAGRO KAGUACATE X 20',
            'INV RECUPERACIONES', 'UCHUVA X 400GR DESGRANAD NACIONAL EURO',
            'CONTENEDOR PET 19,0X12,0X7,5 500 GRS', 'HIGO X 1KG NACIONAL EXITO',
        }

    def install_required_libraries(self):
        required_libraries = ['pandas', 'openpyxl', 'plotly', 'kaleido', 'Pillow', 'matplotlib', 'numpy']
        for lib in required_libraries:
            try:
                __import__(lib)
            except ImportError:
                print(f"Instalando {lib}...")
                subprocess.check_call([sys.executable, "-m", "pip", "install", lib])

    # -------------------------
    # Datos principales (igual que antes)
    # -------------------------
    def load_and_clean_data(self):
        print("📊 Cargando y limpiando datos...")
        required_columns = [
            'Año', 'Mes', 'Nombre Cliente_factura', 'Nombre Centro de Operacion',
            'Valor Total USD', 'Concepto', 'Moneda', 'Nombre Item', 'Vendedor', 'Desc Pais Cliente_factura'
        ]
        try:
            df = pd.read_excel(self.INPUT_PATH, sheet_name='BD', usecols=required_columns)
            df_bv = pd.read_excel(self.INPUT_PATH, sheet_name='Budget x Vendedor')
        except Exception as e:
            print(f"❌ ERROR al cargar archivos: {e}")
            sys.exit()

        df.columns = df.columns.str.strip()
        df_bv.columns = df_bv.columns.str.strip()

        df_filtered = df[
            (df['Concepto'].str.upper().isin(['FACTURA', 'ANULACIÓN FE'])) &
            (df['Moneda'].str.upper().isin(['USD', 'EUR'])) &
            (~df['Nombre Item'].isin(self.EXCLUIR_ITEMS)) &
            (df['Valor Total USD'].notna()) &
            (df['Valor Total USD'] != 0) &
            (df['Vendedor'].str.strip().str.upper() != 'COMERCIALIZADORA INTERNACIONAL CARIBBEAN EXOTICS S A')
        ].copy()
        return df_filtered, df_bv

    def perform_analysis(self, df_filtered, df_bv):
        print("🔍 Realizando análisis...")
        df_anual = df_filtered[df_filtered['Año'] == self.ANIO_ACTUAL]
        df_mensual = df_anual[df_anual['Mes'] == self.MES_ACTUAL]

        ventas_sede_anual = df_anual.groupby('Nombre Centro de Operacion')['Valor Total USD'].sum().sort_values(ascending=False)
        top_clientes_anual = df_anual.groupby('Nombre Cliente_factura')['Valor Total USD'].sum().sort_values(ascending=False).head(5)
        top_paises_anual = df_anual.groupby('Desc Pais Cliente_factura')['Valor Total USD'].sum().sort_values(ascending=False).head(4)
        ejecutado_anual = df_anual['Valor Total USD'].sum()
        budget_anual = df_bv['Valor Total USD'].sum()

        ventas_sede_mensual = df_mensual.groupby('Nombre Centro de Operacion')['Valor Total USD'].sum().sort_values(ascending=False)
        top_clientes_mensual = df_mensual.groupby('Nombre Cliente_factura')['Valor Total USD'].sum().sort_values(ascending=False).head(5)
        top_paises_mensual = df_mensual.groupby('Desc Pais Cliente_factura')['Valor Total USD'].sum().sort_values(ascending=False).head(4)
        ejecutado_mensual = df_mensual['Valor Total USD'].sum()
        budget_mensual = df_bv[df_bv['Mes'] == self.MES_ACTUAL]['Valor Total USD'].sum()

        return (ventas_sede_anual, ventas_sede_mensual, top_clientes_anual, top_clientes_mensual,
                top_paises_anual, top_paises_mensual, budget_anual, ejecutado_anual,
                budget_mensual, ejecutado_mensual)

    # -------------------------
    # Métodos integrados de "ReporteVendedor" reutilizados dentro de Carex
    # -------------------------
    @staticmethod
    def _convertir_formato_colombiano(valor):
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

    @staticmethod
    def _dividir_nombre(nombre, max_len=20):
        partes = str(nombre).split()
        linea = ""
        lineas = []
        for parte in partes:
            if len(linea + parte) < max_len:
                linea += parte + " "
            else:
                lineas.append(linea.strip())
                linea = parte + " "
        if linea:
            lineas.append(linea.strip())
        return "\n".join(lineas)

    def _cargar_datos_vendedores(self):
        # carga las mismas hojas usadas antes, aplicando conversión si es necesario
        try:
            df = pd.read_excel(self.INPUT_PATH, sheet_name='BD')
            df.columns = df.columns.str.strip()
            df_bv = pd.read_excel(self.INPUT_PATH, sheet_name='Budget x Vendedor')
            df_bv.columns = df_bv.columns.str.strip()
            if 'Valor Total USD' in df_bv.columns:
                df_bv['Valor Total USD'] = df_bv['Valor Total USD'].apply(self._convertir_formato_colombiano)
            return df, df_bv
        except Exception as e:
            print(f"❌ ERROR al cargar datos de vendedoras: {e}")
            return None, None

    def _procesar_vendedores(self, df, df_bv, anual=False):
        # reproduce la lógica de ReporteVendedor para anual o mensual (mes actual)
        if df is None or df_bv is None:
            return pd.DataFrame()

        vendedores_filtrados = df['Vendedor'].dropna()
        vendedores_filtrados = vendedores_filtrados[
            vendedores_filtrados.str.upper() != 'COMERCIALIZADORA INTERNACIONAL CARIBBEAN EXOTICS S A'
        ].str.strip()
        vendedores_unicos = vendedores_filtrados[vendedores_filtrados != ''].unique()

        mes_actual = self.MES_ACTUAL
        resultados = []

        for vendedora in vendedores_unicos:
            if anual:
                filtro = (
                    (df['Vendedor'].str.strip() == vendedora.strip()) &
                    (df['Concepto'].str.upper().isin(["FACTURA", "ANULACIÓN FE"])) &
                    (df['Moneda'].str.upper().isin(["USD", "EUR"])) &
                    (~df['Nombre Item'].isin(self.EXCLUIR_ITEMS)) &
                    (df['Valor Total USD'].notna()) &
                    (df['Valor Total USD'] != 0)
                )
                filtro_bv = (
                    (df_bv['Vendedor'].str.strip() == vendedora.strip()) &
                    (df_bv['Valor Total USD'].notna()) &
                    (df_bv['Valor Total USD'] != 0)
                )
            else:
                filtro = (
                    (df['Vendedor'].str.strip() == vendedora.strip()) &
                    (df['Mes'] == mes_actual) &
                    (df['Concepto'].str.upper().isin(["FACTURA", "ANULACIÓN FE"])) &
                    (df['Moneda'].str.upper().isin(["USD", "EUR"])) &
                    (~df['Nombre Item'].isin(self.EXCLUIR_ITEMS)) &
                    (df['Valor Total USD'].notna()) &
                    (df['Valor Total USD'] != 0)
                )
                filtro_bv = (
                    (df_bv['Vendedor'].str.strip() == vendedora.strip()) &
                    (df_bv['Mes'] == mes_actual) &
                    (df_bv['Valor Total USD'].notna()) &
                    (df_bv['Valor Total USD'] != 0)
                )

            total_ejecutado = df.loc[filtro, 'Valor Total USD'].sum()
            total_budget = df_bv.loc[filtro_bv, 'Valor Total USD'].sum()
            porcentaje = (total_ejecutado / total_budget * 100) if total_budget > 0 else 0
            faltante = max(0, 100 - porcentaje)

            resultados.append({
                "Vendedor": vendedora,
                "Budget": round(total_budget, 2),
                "Ejecutado": round(total_ejecutado, 2),
                "% Ejecución": round(porcentaje, 2),
                "% Faltante": round(faltante, 2),
                "Meta": 100.0,
                "Mes": mes_actual if not anual else None
            })

        # Total compañía
        if resultados:
            df_res = pd.DataFrame(resultados)
            total_budget = df_res["Budget"].sum()
            total_ejecutado = df_res["Ejecutado"].sum()
            total_porcentaje = (total_ejecutado / total_budget * 100) if total_budget > 0 else 0
            total_faltante = max(0, 100 - total_porcentaje)
            resultados.append({
                "Vendedor": "TOTAL COMPAÑÍA",
                "Budget": round(total_budget, 2),
                "Ejecutado": round(total_ejecutado, 2),
                "% Ejecución": round(total_porcentaje, 2),
                "% Faltante": round(total_faltante, 2),
                "Meta": 100.0,
                "Mes": mes_actual if not anual else None
            })

        df_resultado = pd.DataFrame(resultados)
        return df_resultado

    def _generar_grafico_vendedores_memoria(self, df_resultado, anual=False):
        # genera el plot (matplotlib) y devuelve BytesIO
        if df_resultado is None or df_resultado.empty:
            return None

        fig, ax = plt.subplots(figsize=(12, 5))
        x = np.arange(len(df_resultado["Vendedor"]))
        ejecucion = df_resultado["% Ejecución"]
        faltante = df_resultado["% Faltante"]

        bar1 = ax.bar(x, ejecucion, width=0.4, label="% Ejecución", color="#89CFF0")
        bar2 = ax.bar(x, faltante, width=0.4, bottom=ejecucion, label="% Faltante", color="#1f4e79")

        titulo = "Ejecución vs Faltante por Vendedor - " + ("Anual" if anual else f"Mes {self.MES_ACTUAL_NOMBRE}")
        ax.set_title(titulo, fontsize=14, fontweight='bold')
        ax.set_ylabel("Porcentaje")
        ax.set_xticks(x)
        ax.set_xticklabels([self._dividir_nombre(v) for v in df_resultado["Vendedor"]])

        ax.legend(loc='lower center', bbox_to_anchor=(0.5, -0.25), ncol=2, fontsize=10)

        # Mostrar porcentaje sobre cada barra
        for rects, valores in zip([bar1, bar2], [ejecucion, faltante]):
            for rect, val in zip(rects, valores):
                height = rect.get_height()
                if height > 0:
                    ax.text(
                        rect.get_x() + rect.get_width() / 2,
                        rect.get_y() + height / 2,
                        f"{val:.1f}%",
                        ha='center',
                        va='center',
                        color='white' if rects is bar2 else 'black',
                        fontsize=9,
                        fontweight='bold'
                    )

        plt.tight_layout()
        buf = BytesIO()
        plt.savefig(buf, format='png', dpi=150)
        plt.close(fig)
        buf.seek(0)
        return buf

    # -------------------------
    # Gráficos del dashboard (se añaden aquí también los gráficos de vendedor)
    # -------------------------
    def create_plots_in_memory(self, analysis_data):
        print("🎨 Creando gráficos en memoria...")
        (ventas_sede_anual, ventas_sede_mensual, top_clientes_anual, top_clientes_mensual,
         top_paises_anual, top_paises_mensual, budget_anual, ejecutado_anual,
         budget_mensual, ejecutado_mensual) = analysis_data

        colores_carex = ['#003366', '#ff7f0e', "#3d2ca0", '#d62728', '#9467bd']

        def create_plot_bytes(fig, title, width, height):
            fig.update_layout(
                title_text=f"<b>{title}</b>",
                title_x=0.5,
                title_font_size=18,
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

        fig_gauge_anual = go.Figure(go.Indicator(
            mode="gauge+number",
            value=ejecutado_anual,
            number={'valueformat': '$,.2f', 'font': {'size': 40}},
            title={'text': f"Venta Acumulada USD Anual {self.ANIO_ACTUAL}", 'font': {'size': 14}},
            gauge={'axis': {'range': [0, budget_anual]}, 'bar': {'color': "#003366"}}
        ))
        gauge_anual_bytes = create_plot_bytes(fig_gauge_anual, "Gauge Anual", 600, 350)

        fig_gauge_mensual = go.Figure(go.Indicator(
            mode="gauge+number",
            value=ejecutado_mensual,
            number={'valueformat': '$,.2f', 'font': {'size': 40}},
            title={'text': f"Venta Acumulada USD Mensual ({self.MES_ACTUAL_NOMBRE})", 'font': {'size': 14}},
            gauge={'axis': {'range': [0, budget_mensual]}, 'bar': {'color': "#003366"}}
        ))
        gauge_mensual_bytes = create_plot_bytes(fig_gauge_mensual, "Gauge Mensual", 600, 350)

        # Pie anual y mensual
        fig_anual = go.Figure(data=go.Pie(
            labels=ventas_sede_anual.index,
            values=ventas_sede_anual.values,
            textinfo='value+percent',
            texttemplate='$%{value:,.0f}<br>(%{percent})',
            insidetextfont={'size': 12, 'color': 'white'},
            hoverinfo='label+percent+value',
            marker_colors=colores_carex,
            hole=0.45
        ))
        total_anual = ventas_sede_anual.sum()
        if total_anual > 0:
            fig_anual.add_annotation(text=f"Total<br><b>${total_anual:,.0f}</b>", x=0.5, y=0.5,
                                    font=dict(size=16, color='#003366'), showarrow=False)
        pie_anual_bytes = create_plot_bytes(fig_anual, f"Ventas por Sede Anual ({self.ANIO_ACTUAL})", 800, 500)

        fig_mensual = go.Figure(data=go.Pie(
            labels=ventas_sede_mensual.index,
            values=ventas_sede_mensual.values,
            textinfo='value+percent',
            texttemplate='$%{value:,.0f}<br>(%{percent})',
            insidetextfont={'size': 12, 'color': 'white'},
            hoverinfo='label+percent+value',
            marker_colors=colores_carex,
            hole=0.45
        ))
        total_mensual = ventas_sede_mensual.sum()
        if total_mensual > 0:
            fig_mensual.add_annotation(text=f"Total<br><b>${total_mensual:,.0f}</b>", x=0.5, y=0.5,
                                      font=dict(size=16, color='#003366'), showarrow=False)
        pie_mensual_bytes = create_plot_bytes(fig_mensual, f"Ventas por Sede ({self.MES_ACTUAL_NOMBRE})", 800, 500)

        # Top paises (bar) anual/mensual
        fig_paises_anual = go.Figure(data=go.Bar(
            x=top_paises_anual.index,
            y=top_paises_anual.values,
            marker_color=colores_carex[::],
            text=[f'${v:,.0f}' for v in top_paises_anual.values],
            textposition='outside'
        ))
        fig_paises_anual.update_layout(xaxis_title_text='País', yaxis_title_text='Ventas USD')
        bar_paises_anual_bytes = create_plot_bytes(fig_paises_anual, f"Top 4 Ventas por País Anual ({self.ANIO_ACTUAL})", 800, 450)

        fig_paises_mensual = go.Figure(data=go.Bar(
            x=top_paises_mensual.index,
            y=top_paises_mensual.values,
            marker_color=colores_carex[::],
            text=[f'${v:,.0f}' for v in top_paises_mensual.values],
            textposition='outside'
        ))
        fig_paises_mensual.update_layout(xaxis_title_text='País', yaxis_title_text='Ventas USD')
        bar_paises_mensual_bytes = create_plot_bytes(fig_paises_mensual, f"Top 4 Ventas por País ({self.MES_ACTUAL_NOMBRE})", 800, 450)

        # Tablas top clientes (anual / mensual)
        fig_top_clientes_anual = go.Figure(data=[go.Table(
            header=dict(values=['<b>Cliente</b>', f'<b>Ventas {self.ANIO_ACTUAL} ($)</b>'],
                        align=['left', 'right'], font=dict(color='white', size=12), fill_color='#003366', height=30),
            cells=dict(values=[top_clientes_anual.index, [f'${x:,.0f}' for x in top_clientes_anual.values]],
                       align=['left', 'right'], fill_color=[['white', '#f0f0f0'] * (len(top_clientes_anual) // 2 + 1)],
                       font=dict(color='black', size=11), height=25)
        )])
        tabla_clientes_anual_bytes = create_plot_bytes(fig_top_clientes_anual, f"Top 5 Clientes Anual ({self.ANIO_ACTUAL})", 800, 300)

        fig_top_clientes_mensual = go.Figure(data=[go.Table(
            header=dict(values=['<b>Cliente</b>', f'<b>Ventas Mes ({self.MES_ACTUAL_NOMBRE}) ($)</b>'],
                        align=['left', 'right'], font=dict(color='white', size=12), fill_color='#003366', height=30),
            cells=dict(values=[top_clientes_mensual.index, [f'${x:,.0f}' for x in top_clientes_mensual.values]],
                       align=['left', 'right'], fill_color=[['white', '#f0f0f0'] * (len(top_clientes_mensual) // 2 + 1)],
                       font=dict(color='black', size=11), height=25)
        )])
        tabla_clientes_mensual_bytes = create_plot_bytes(fig_top_clientes_mensual, f"Top 5 Clientes ({self.MES_ACTUAL_NOMBRE})", 800, 300)

        # --- ahora integramos los gráficos de vendedoras (anual y mensual) reutilizando la lógica ---
        df_full, df_bv = self._cargar_datos_vendedores()
        df_vendedores_anual = self._procesar_vendedores(df_full, df_bv, anual=True)
        df_vendedores_mensual = self._procesar_vendedores(df_full, df_bv, anual=False)

        graf_vendedores_anual_bytes = self._generar_grafico_vendedores_memoria(df_vendedores_anual, anual=True)
        graf_vendedores_mensual_bytes = self._generar_grafico_vendedores_memoria(df_vendedores_mensual, anual=False)

        # Devolvemos la lista completa de imágenes (algunas pueden ser None si falló)
        images = [
            gauge_anual_bytes, gauge_mensual_bytes,
            pie_anual_bytes, pie_mensual_bytes,
            bar_paises_anual_bytes, bar_paises_mensual_bytes,
            graf_vendedores_anual_bytes, graf_vendedores_mensual_bytes,
            tabla_clientes_anual_bytes, tabla_clientes_mensual_bytes,
        ]
        return images

    # -------------------------
    # Combinar imágenes en rejilla 2 columnas (dinámico)
    # -------------------------
    def combine_images_into_single_report(self, image_bytes_list, cols=2, padding=40):
        print("🖼️ Combinando gráficos en un reporte consolidado (rejilla)...")
        # Filtrar None
        image_bytes_list = [b for b in image_bytes_list if b is not None]
        if not image_bytes_list:
            print("❌ No hay imágenes para combinar.")
            return

        pil_images = [Image.open(b).convert("RGB") for b in image_bytes_list]
        widths, heights = zip(*(im.size for im in pil_images))
        max_w = max(widths)
        max_h = max(heights)
        
        scale_factor = 5.3  # 30% más alto
        max_h = int(max_h * scale_factor)
        max_w = int(max_w * scale_factor)  # opcional si quieres aumentar también el ancho


        rows = math.ceil(len(pil_images) / cols)
        final_w = cols * max_w + (cols + 1) * padding
        final_h = rows * max_h + (rows + 1) * padding + 120  # espacio para título

        final = Image.new("RGB", (final_w, final_h), "white")
        draw = ImageDraw.Draw(final)

        # título
        try:
            font_path = "arial.ttf"
            title_font = ImageFont.truetype(font_path, 48)
        except Exception:
            title_font = ImageFont.load_default()
        title = f"Reporte Consolidado - {self.FECHA_ACTUAL}"
        bbox = draw.textbbox((0, 0), title, font=title_font)
        w_t = bbox[2] - bbox[0]
        draw.text(((final_w - w_t) / 2, 20), title, fill="black", font=title_font)

        # pegar imágenes en rejilla
        for idx, im in enumerate(pil_images):
            row = idx // cols
            col = idx % cols
            x = padding + col * (max_w + padding)
            y = padding + row * (max_h + padding) + 80
            final.paste(im.resize((max_w, max_h)), (int(x), int(y)))

        out_path = os.path.join(self.OUTPUT_DIR, f"dashboard_consolidado_{self.FECHA_ACTUAL}.jpg")
        final.save(out_path, "JPEG", quality=95)
        print(f"✅ Dashboard consolidado guardado en: {out_path}")
        return out_path

    # -------------------------
    # Excel report (igual que antes)
    # -------------------------
    def generate_excel_report(self, df, df_bv):
        print("📋 Generando reporte de Excel anual...")
        output_path = os.path.join(self.OUTPUT_DIR, f"reporte_vendedoras_anual_{self.FECHA_ACTUAL}.xlsx")

        df_ejecutado = df.groupby('Vendedor')['Valor Total USD'].sum().rename('Ejecutado Total USD')
        df_budget = df_bv.groupby('Vendedor')['Valor Total USD'].sum().rename('Budget Total USD')

        df_resultado = pd.concat([df_ejecutado, df_budget], axis=1).fillna(0)
        df_resultado['% Ejecución Anual'] = (df_resultado['Ejecutado Total USD'] / df_resultado['Budget Total USD'] * 100).fillna(0)
        df_resultado['% Faltante'] = 100 - df_resultado['% Ejecución Anual']
        df_resultado['Meta'] = 100.0

        df_resultado = df_resultado[df_resultado.index.str.upper() != 'COMERCIALIZADORA INTERNACIONAL CARIBBEAN EXOTICS S A']

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

    # -------------------------
    # Flujo principal
    # -------------------------
    def generate_all_reports(self):
        df_filtered, df_bv = self.load_and_clean_data()
        if df_filtered.empty:
            print("⚠ No se encontraron datos válidos.")
            return

        analysis_data = self.perform_analysis(df_filtered, df_bv)
        images = self.create_plots_in_memory(analysis_data)
        # combinar (rejilla 2 columnas)
        self.combine_images_into_single_report(images, cols=2)
        # excel anual
        self.generate_excel_report(df_filtered, df_bv)
