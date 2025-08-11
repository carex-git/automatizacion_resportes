import os
import json
from CarexDashboard import CarexDashboard
from ReportEmailSender import ReportEmailSender
from TasaUpdater import TasaUpdater

# Cargar configuración
with open("config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

if __name__ == "__main__":
    TasaUpdater.main()
    CarexDashboard().generate_all_reports()
    ReportEmailSender(
        remitente=config["remitente"],
        password=config["password"],
        destinatarios=config["destinatarios"],
        asunto=config["asunto"],
        cuerpo=config["cuerpo"]
    ).send_mail()