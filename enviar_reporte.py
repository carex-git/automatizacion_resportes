import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
from email.mime.text import MIMEText
import os

base_dir = os.path.dirname(os.path.abspath(__file__))

# Configuración del correo
remitente = "cm090457@gmail.com"
destinatario = "cristian.marin5062@soyuco.edu.co"
asunto = "Reporte automático"
cuerpo = "Hola,\n\nAdjunto encontrarás el reporte en Excel y la imagen generada.\n\nSaludos."
password = "xrjw xnuz jtox qrmv"

# Archivos a enviar
archivo_excel = base_dir + "/reporte_vendedoras.xlsx"
archivo_imagen = base_dir + "/grafico_ejecucion_vendedor.png"

# Crear mensaje
mensaje = MIMEMultipart()
mensaje["From"] = remitente
mensaje["To"] = destinatario
mensaje["Subject"] = asunto
mensaje.attach(MIMEText(cuerpo, "plain"))

# Adjuntar archivo Excel
with open(archivo_excel, "rb") as adj:
    parte = MIMEBase("application", "octet-stream")
    parte.set_payload(adj.read())
    encoders.encode_base64(parte)
    parte.add_header("Content-Disposition", f"attachment; filename={os.path.basename(archivo_excel)}")
    mensaje.attach(parte)

# Adjuntar imagen
with open(archivo_imagen, "rb") as img:
    imagen = MIMEImage(img.read(), name=os.path.basename(archivo_imagen))
    mensaje.attach(imagen)

# Enviar el correo
try:
    servidor = smtplib.SMTP("smtp.gmail.com", 587)
    servidor.starttls()
    servidor.login(remitente, password)
    servidor.send_message(mensaje)
    servidor.quit()
    print("Correo enviado correctamente.")
except Exception as e:
    print("Error al enviar el correo:", e)
