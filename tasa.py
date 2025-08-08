import requests
import pandas as pd
from datetime import datetime
import os
import xml.etree.ElementTree as ET

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
INPUT_FILENAME = "Carex COL Reporte Vendedor.xlsx"
INPUT_PATH = os.path.join(DATA_DIR, INPUT_FILENAME)

def obtener_tasa_eur_usd():
    url = "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml"
    try:
        response = requests.get(url)
        response.raise_for_status()
        tree = ET.fromstring(response.content)
        for cube in tree.findall(".//{*}Cube[@currency='USD']"):
            return float(cube.attrib['rate'])
    except Exception as e:
        print(f"❌ Error al obtener EUR/USD desde BCE: {e}")
    return None

def obtener_tasa_cop_usd():
    try:
        url = "https://www.datos.gov.co/resource/ceyp-9c7c.json?$order=vigenciadesde DESC&$limit=1"
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        if data:
            return float(data[0]['valor'])
    except Exception as e:
        print(f"⚠️ Error al obtener la TRM de Colombia: {e}")
        try:
            return float(input("🔁 Ingresa manualmente la TRM COP/USD: "))
        except:
            print("❌ Valor ingresado inválido.")
    return None

def actualizar_excel(path, fecha, cop_usd, eur_usd):
    eur_cop = eur_usd * cop_usd
    usd_eur = 1 / eur_usd

    df = pd.read_excel(path, sheet_name='TC')

    if fecha in df['Fecha'].values:
        idx = df[df['Fecha'] == fecha].index[0]
        df.loc[idx, 'COP/USD'] = cop_usd
        df.loc[idx, 'EUR/COP'] = eur_cop
        df.loc[idx, 'USD/EUR'] = usd_eur
        print(f"🔁 Actualizada la fila existente para la fecha {fecha}")
    else:
        nueva_fila = {
            'Fecha': fecha,
            'COP/USD': cop_usd,
            'EUR/COP': eur_cop,
            'USD/EUR': eur_usd
        }
        df = df.append(nueva_fila, ignore_index=True)
        print(f"➕ Añadida nueva fila para la fecha {fecha}")

    with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='TC', index=False)

    print("\n✅ Tasas insertadas en Excel:")
    print(f"Fecha:     {fecha}")
    print(f"COP/USD:   {cop_usd}")
    print(f"EUR/USD:   {eur_usd}")
    print(f"EUR/COP:   {eur_cop}")
    print(f"USD/EUR:   {usd_eur}")

def main():
    if not os.path.exists(INPUT_PATH):
        print(f"❌ Archivo no encontrado: {INPUT_PATH}")
        return

    fecha_actual = int(datetime.today().strftime('%Y%m%d'))
    eur_usd = obtener_tasa_eur_usd()
    cop_usd = obtener_tasa_cop_usd()

    if eur_usd and cop_usd:
        actualizar_excel(INPUT_PATH, fecha_actual, cop_usd, eur_usd)
    else:
        print("❌ No se pudieron obtener todas las tasas necesarias.")

if __name__ == "__main__":
    main()
