import os
import shutil
import json
from CarexDashboard import CarexDashboard
from ReportEmailSender import ReportEmailSender
from TasaUpdater import TasaUpdater
from UnoBiableUpdater import UnoBiableUpdater

# Cargar configuración
with open("config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

def eliminar_carpeta(path):
    if os.path.exists(path) and os.path.isdir(path):
        shutil.rmtree(path)

if __name__ == "__main__":
    
    base_dir = config['base_dir']
    
    eliminar_carpeta(os.path.join(base_dir, 'output'))
    eliminar_carpeta(os.path.join(base_dir, 'data/backups'))
    
    if config.get('tasa_updater', True):
        print("== Ejecutando TasaUpdater ==")
        TasaUpdater(base_dir=base_dir).main()
        
    if config.get('uno_biable_updater', True):
        print("== Ejecutando Updater UnoBiable ==")
        UnoBiableUpdater(base_dir=base_dir).main()
        
    CarexDashboard(base_dir=base_dir).generate_all_reports()
    ReportEmailSender(
        base_dir=base_dir,
        remitente=config["remitente"],
        password=config["password"],
        destinatarios=config["destinatarios"],
        asunto=config["asunto"],
        cuerpo=config["cuerpo"]
    ).send_mail()
