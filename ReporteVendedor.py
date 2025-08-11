import os
import pandas as pd
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt

class ReporteVendedor:
    def __init__(self, archivo_excel, base_dir):
        self.archivo_excel = archivo_excel
        self.base_dir = base_dir
        self.carpeta_salida = os.path.join(base_dir, "output")
        os.makedirs(self.carpeta_salida, exist_ok=True)
        self.df = None
        self.df_BV = None
        self.resultados = []

    @staticmethod
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

    @staticmethod
    def formatear_numero_colombiano(numero):
        if pd.isna(numero) or numero == 0:
            return "0"
        return f"{numero:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

    @staticmethod
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

    def cargar_datos(self):
        self.df = pd.read_excel(self.archivo_excel, sheet_name='BD')
        self.df.columns = self.df.columns.str.strip()

        self.df_BV = pd.read_excel(self.archivo_excel, sheet_name='Budget x Vendedor')
        self.df_BV.columns = self.df_BV.columns.str.strip()
        if 'Valor Total USD' in self.df_BV.columns:
            self.df_BV['Valor Total USD'] = self.df_BV['Valor Total USD'].apply(self.convertir_formato_colombiano)

    def procesar(self):
        vendedores_filtrados = self.df['Vendedor'].dropna()
        vendedores_filtrados = vendedores_filtrados[
            vendedores_filtrados.str.upper() != 'COMERCIALIZADORA INTERNACIONAL CARIBBEAN EXOTICS S A'
        ].str.strip()
        vendedores_unicos = vendedores_filtrados[vendedores_filtrados != ''].unique()

        excluir_items = [
            'AIR FREIGHT', 'INV PLANTAS', 'INV PRIMA - FLO', 'INV RECICLAJE', 'OTHER EXPORT COSTS',
            'SEA FREIGHT COST', 'HUMAGRO CALCIUM DRENCH', 'SULPHUR GULUPA X 20 LITROS',
            'HUMAGRO CALCIUM FOLIAR UCH X 20 LITROS', 'HUMAGRO KAGUACATE X 20',
            'INV RECUPERACIONES', 'UCHUVA X 400GR DESGRANAD NACIONAL EURO',
            'CONTENEDOR PET 19,0X12,0X7,5 500 GRS', 'HIGO X 1KG NACIONAL EXITO',
        ]

        mes_actual = datetime.now().month

        self.resultados.clear()

        for vendedora in vendedores_unicos:
            filtro = (
                (self.df['Vendedor'].str.strip() == vendedora.strip()) &
                (self.df['Mes'] == mes_actual) &
                (self.df['Concepto'].str.upper().isin(["FACTURA", "ANULACIÓN FE"])) &
                (self.df['Moneda'].str.upper().isin(["USD", "EUR"])) &
                (~self.df['Nombre Item'].isin(excluir_items)) &
                (self.df['Valor Total USD'].notna()) &
                (self.df['Valor Total USD'] != 0)
            )

            total_ejecutado = self.df.loc[filtro, 'Valor Total USD'].sum()

            filtro_bv = (
                (self.df_BV['Vendedor'].str.strip() == vendedora.strip()) &
                (self.df_BV['Mes'] == mes_actual) &
                (self.df_BV['Valor Total USD'].notna()) &
                (self.df_BV['Valor Total USD'] != 0)
            )
            total_budget = self.df_BV.loc[filtro_bv, 'Valor Total USD'].sum()

            porcentaje = (total_ejecutado / total_budget * 100) if total_budget > 0 else 0
            faltante = max(0, 100 - porcentaje)

            self.resultados.append({
                "Vendedor": vendedora,
                "Budget": round(total_budget, 2),
                "Ejecutado": round(total_ejecutado, 2),
                "% Ejecución": round(porcentaje, 2),
                "% Faltante": round(faltante, 2),
                "Meta": 100.0,
                "Mes": mes_actual,
            })

        # Total Compañía
        if self.resultados:
            df_resultado = pd.DataFrame(self.resultados)
            total_budget = df_resultado["Budget"].sum()
            total_ejecutado = df_resultado["Ejecutado"].sum()
            total_porcentaje = (total_ejecutado / total_budget * 100) if total_budget > 0 else 0
            total_faltante = max(0, 100 - total_porcentaje)
            self.resultados.append({
                "Vendedor": "TOTAL COMPAÑÍA",
                "Budget": round(total_budget, 2),
                "Ejecutado": round(total_ejecutado, 2),
                "% Ejecución": round(total_porcentaje, 2),
                "% Faltante": round(total_faltante, 2),
                "Meta": 100.0,
                "Mes": mes_actual
            })

    def exportar_excel(self):
        df_res = pd.DataFrame(self.resultados)
        excel_path = os.path.join(self.carpeta_salida, "reporte_vendedoras.xlsx")
        df_res.to_excel(excel_path, index=False)
        print(f"✅ Excel generado en: {excel_path}")

    def generar_grafico(self):
        df_res = pd.DataFrame(self.resultados)
        if df_res.empty:
            print("⚠️ No hay datos para generar gráfico.")
            return

        image_path = os.path.join(self.carpeta_salida, "grafico_ejecucion_vendedor.png")
        fig, ax = plt.subplots(figsize=(12, 6))
        x = np.arange(len(df_res["Vendedor"]))
        ejecucion = df_res["% Ejecución"]
        faltante = df_res["% Faltante"]

        bar1 = ax.bar(x, ejecucion, width=0.4, label="% Ejecución", color="#89CFF0")
        bar2 = ax.bar(x, faltante, width=0.4, bottom=ejecucion, label="% Faltante", color="#1f4e79")

        ax.set_title("Ejecución vs Faltante por Vendedor (Meta = 100%)", fontsize=14, fontweight='bold')
        ax.set_ylabel("Porcentaje")
        ax.set_xticks(x)
        ax.set_xticklabels([self.dividir_nombre(v) for v in df_res["Vendedor"]])

        # Leyenda centrada abajo y fuera del gráfico
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
                        color='white' if rects == bar2 else 'black',
                        fontsize=9,
                        fontweight='bold'
                    )

        plt.tight_layout()
        plt.savefig(image_path, dpi=300)
        plt.close(fig)
        print(f"📊 Gráfico generado en: {image_path}")


    def generar_reporte(self):
        self.cargar_datos()
        self.procesar()
        self.exportar_excel()
        self.generar_grafico()
