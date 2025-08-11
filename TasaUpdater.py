import requests
from datetime import datetime
import os
import xml.etree.ElementTree as ET
import shutil
import xlwings as xw


class TasaUpdater:
    def __init__(self,base_dir):
        self.BASE_DIR = base_dir
        self.DATA_DIR = os.path.join(self.BASE_DIR, "data")
        self.INPUT_FILENAME = "Carex COL Reporte Vendedor.xlsx"
        self.INPUT_PATH = os.path.join(self.DATA_DIR, self.INPUT_FILENAME)

        self.BACKUP_DIR = os.path.join(self.DATA_DIR, "backups")
        os.makedirs(self.BACKUP_DIR, exist_ok=True)

    def hacer_backup(self, path):
        fecha_str = datetime.today().strftime("%Y%m%d_%H%M%S")
        backup_path = os.path.join(self.BACKUP_DIR, f"backup_{fecha_str}.xlsx")
        shutil.copy2(path, backup_path)
        print(f"📂 Copia de seguridad creada: {backup_path}")
        return backup_path

    def obtener_tasa_eur_usd(self):
        url = "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml"
        try:
            resp = requests.get(url)
            resp.raise_for_status()
            tree = ET.fromstring(resp.content)
            for cube in tree.findall(".//{*}Cube[@currency='USD']"):
                return float(cube.attrib['rate'])
        except Exception as e:
            print(f"❌ Error obteniendo EUR/USD: {e}")
        return None

    def obtener_tasa_cop_usd(self):
        url = "https://www.datos.gov.co/resource/ceyp-9c7c.json?$order=vigenciadesde DESC&$limit=1"
        try:
            resp = requests.get(url)
            resp.raise_for_status()
            data = resp.json()
            if data:
                return float(data[0]['valor'])
        except Exception as e:
            print(f"⚠️ Error obteniendo TRM COP/USD: {e}")
            try:
                return float(input("🔁 Ingresa manualmente la TRM COP/USD: "))
            except:
                print("❌ Valor inválido.")
        return None

    def actualizar_excel_sin_corromper(self, path, fecha, cop_usd, eur_usd):
        eur_cop = eur_usd * cop_usd

        app = xw.App(visible=False)
        libro = app.books.open(path)
        hoja = libro.sheets['TC']

        fecha_str = str(fecha)
        encontrada = False

        ultima_fila = hoja.range("A" + str(hoja.cells.last_cell.row)).end('up').row

        for fila in range(2, ultima_fila + 1):
            if str(hoja.range(f"A{fila}").value) == fecha_str:
                hoja.range(f"B{fila}").value = cop_usd
                hoja.range(f"C{fila}").value = eur_cop
                hoja.range(f"D{fila}").value = eur_usd
                encontrada = True
                print(f"🔁 Actualizada fila {fila} para fecha {fecha}")
                break

        if not encontrada:
            nueva_fila = ultima_fila + 1
            hoja.range(f"A{nueva_fila}").value = fecha
            hoja.range(f"B{nueva_fila}").value = cop_usd
            hoja.range(f"C{nueva_fila}").value = eur_cop
            hoja.range(f"D{nueva_fila}").value = eur_usd
            print(f"➕ Añadida fila {nueva_fila} para fecha {fecha}")

        libro.save()
        libro.close()
        app.quit()
        print("✅ Archivo actualizado sin corromper.")

    def main(self):
        if not os.path.exists(self.INPUT_PATH):
            print(f"❌ Archivo no encontrado: {self.INPUT_PATH}")
            return

        self.hacer_backup(self.INPUT_PATH)

        fecha_actual = int(datetime.today().strftime('%Y%m%d'))
        eur_usd = self.obtener_tasa_eur_usd()
        cop_usd = self.obtener_tasa_cop_usd()

        if eur_usd and cop_usd:
            self.actualizar_excel_sin_corromper(self.INPUT_PATH, fecha_actual, cop_usd, eur_usd)
        else:
            print("❌ No se pudieron obtener todas las tasas.")
