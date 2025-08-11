import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
import os

class ReportEmailSender:
    def __init__(self, base_dir, remitente, password, destinatarios, asunto, cuerpo):
        self.remitente = remitente
        self.password = password
        self.destinatarios = destinatarios
        self.asunto = asunto
        self.cuerpo = cuerpo

        base_dir = base_dir
        self.output_dir = os.path.join(base_dir, "output")

    def send_mail(self):
        mensaje = MIMEMultipart()
        mensaje["From"] = self.remitente
        mensaje["To"] = ", ".join(self.destinatarios)
        mensaje["Subject"] = self.asunto
        mensaje.attach(MIMEText(self.cuerpo, "plain"))

        """# Adjuntar archivos extra (ejemplo: Excel)
        for ruta in self.archivos_extra:
            if not os.path.exists(ruta):
                print(f"⚠ Archivo no encontrado: {ruta}")
                continue
            with open(ruta, "rb") as adj:
                parte = MIMEBase("application", "octet-stream")
                parte.set_payload(adj.read())
                encoders.encode_base64(parte)
                parte.add_header("Content-Disposition", f"attachment; filename={os.path.basename(ruta)}")
                mensaje.attach(parte)
                print(f"📎 Archivo adjuntado: {ruta}")
"""
        # Adjuntar todas las imágenes de la carpeta output (png, jpg, jpeg, gif)
        extensiones_img = (".png", ".jpg", ".jpeg", ".gif")
        for archivo in os.listdir(self.output_dir):
            if archivo.lower().endswith(extensiones_img):
                ruta_img = os.path.join(self.output_dir, archivo)
                with open(ruta_img, "rb") as img:
                    imagen = MIMEImage(img.read(), name=archivo)
                    mensaje.attach(imagen)
                    print(f"🖼 Imagen adjuntada: {ruta_img}")

        # Enviar correo
        try:
            servidor = smtplib.SMTP("smtp.gmail.com", 587)
            servidor.starttls()
            servidor.login(self.remitente, self.password)
            servidor.send_message(mensaje)
            servidor.quit()
            print("✅ Correo enviado correctamente.")
        except Exception as e:
            print("❌ Error al enviar el correo:", e)
