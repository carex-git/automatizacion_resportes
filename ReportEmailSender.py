import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
import os

class ReportEmailSender:
    def __init__(self, base_dir, remitente, password, destinatarios, asunto, cuerpo):
        self.remitente = remitente
        self.password = password
        self.destinatarios = destinatarios
        self.asunto = asunto
        self.cuerpo = cuerpo  # Texto o HTML
        self.output_dir = os.path.join(base_dir, "output")
        self.imagen_extra = os.path.join(base_dir, "image.png")  # ✅ tu imagen debajo de Saludes

    def send_mail(self):
        # Crear mensaje raíz tipo "related"
        mensaje = MIMEMultipart("related")
        mensaje["From"] = self.remitente
        mensaje["To"] = ", ".join(self.destinatarios)
        mensaje["Subject"] = self.asunto

        # Contenedor "alternative" (texto plano y HTML)
        parte_alternativa = MIMEMultipart("alternative")
        mensaje.attach(parte_alternativa)

        # Texto plano (fallback)
        parte_alternativa.attach(
            MIMEText("Este correo contiene un reporte con imágenes embebidas.", "plain")
        )

        # HTML base
        html_cuerpo = f"<html><body><p>{self.cuerpo}</p>"

        # Insertar la primera imagen del reporte
        extensiones_img = (".png", ".jpg", ".jpeg", ".gif")
        img_index = 1
        for archivo in os.listdir(self.output_dir):
            if archivo.lower().endswith(extensiones_img):
                ruta_img = os.path.join(self.output_dir, archivo)
                with open(ruta_img, "rb") as img:
                    mime_img = MIMEImage(img.read())
                    cid = f"imagen{img_index}"
                    mime_img.add_header("Content-ID", f"<{cid}>")
                    mime_img.add_header("Content-Disposition", "inline")
                    mensaje.attach(mime_img)

                    # HTML referencia inline
                    html_cuerpo += f'<br><img src="cid:{cid}" style="max-width:100%;"><br>'
                    img_index += 1
                    print(f"🖼 Imagen de reporte embebida en el correo: {ruta_img}")

                break  # ✅ solo insertamos la primera

        # ✅ Agregar saludo después del reporte
        html_cuerpo += "<p style='font-size:18px; color:#333;'>Saludes,</p>"

        # ✅ Insertar la imagen `imagen.png` debajo del saludo
        if os.path.exists(self.imagen_extra):
            with open(self.imagen_extra, "rb") as fimg:
                mime_img = MIMEImage(fimg.read())
                cid_extra = "imagen_extra"
                mime_img.add_header("Content-ID", f"<{cid_extra}>")
                mime_img.add_header("Content-Disposition", "inline")
                mensaje.attach(mime_img)

                html_cuerpo += f'<br><img src="cid:{cid_extra}" style="max-width:100%;"><br>'
                print(f"🖋 Imagen adicional añadida: {self.imagen_extra}")

        html_cuerpo += "</body></html>"

        # Agregar HTML al bloque "alternative"
        parte_alternativa.attach(MIMEText(html_cuerpo, "html"))

        # Enviar correo
        try:
            servidor = smtplib.SMTP("smtp.gmail.com", 587)
            servidor.starttls()
            servidor.login(self.remitente, self.password)
            servidor.send_message(mensaje)
            servidor.quit()
            print("✅ Correo enviado correctamente (imágenes incrustadas en el cuerpo).")
        except Exception as e:
            print("❌ Error al enviar el correo:", e)
