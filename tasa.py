import requests
import xml.etree.ElementTree as ET

#fuente del Banco Central Europeo

def obtener_tasa_bce():
    url = "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml"
    response = requests.get(url)
    
    if response.status_code != 200:
        raise Exception(f"Error al acceder al BCE: {response.status_code}")
    
    tree = ET.fromstring(response.content)

    # Busca la tasa USD
    namespace = {'ns': 'http://www.ecb.int/vocabulary/2002-08-01/eurofxref'}
    for cube in tree.findall(".//{*}Cube[@currency='USD']"):
        return float(cube.attrib['rate'])
    
    raise Exception("No se encontró la tasa USD en el XML del BCE.")

try:
    tasa = obtener_tasa_bce()
    print(f"Tasa EUR a USD actual (BCE): {tasa}")
except Exception as e:
    print(f"Error al obtener la tasa: {e}")