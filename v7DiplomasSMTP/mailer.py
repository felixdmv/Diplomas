import os
import pandas as pd
import smtplib  # <-- La librería clave para SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# --- MANTENEMOS LA LÓGICA DE OUTLOOK ---
import win32com.client as win32
import pythoncom
import shutil
import tempfile
from utils import email_a_id, limpiar_nombre_curso

# --- NUEVO: BASE DE DATOS DE PROVEEDORES SMTP ---
SMTP_PROVIDERS = {
    "gmail": {
        "server": "smtp.gmail.com",
        "port": 587,
        "tls": True
    },
    "office365": {
        "server": "smtp.office365.com",
        "port": 587,
        "tls": True
    },
    "hotmail": {
        "server": "smtp.live.com",
        "port": 587,
        "tls": True
    }
}

# --- NUEVA FUNCIÓN DE ENVÍO POR SMTP ---
def enviar_masivo_smtp(provider_config, user, password, excel_path, pdf_folder, callback_log):
    """
    Motor de envío masivo usando el protocolo SMTP.
    provider_config: Un diccionario con 'server', 'port', 'tls'.
    """
    try:
        # 1. Conectar al servidor
        callback_log(f"🔌 Conectando a {provider_config['server']}...")
        server = smtplib.SMTP(provider_config['server'], provider_config['port'])
        if provider_config.get('tls', False):
            server.starttls() # Iniciar conexión segura
        
        # 2. Iniciar sesión
        server.login(user, password)
        callback_log("✅ Login SMTP correcto.")
        
    except Exception as e:
        callback_log(f"❌ ERROR DE CONEXIÓN/LOGIN SMTP: {e}")
        callback_log("   Asegúrate de que la contraseña es correcta.")
        callback_log("   Si usas 2FA, necesitas una 'Contraseña de Aplicación'.")
        return

    df = pd.read_excel(excel_path)
    df.columns = [c.lower().strip() for c in df.columns]
    
    enviados = 0
    errores = 0

    for idx, row in df.iterrows():
        try:
            email_dest = str(row["email"]).strip()
            nombre = str(row.get("nombre", "")).strip()
            curso_raw = str(row.get("curso_nombre", "")).strip()

            if "@" not in email_dest:
                continue

            email_id = email_a_id(email_dest)
            curso_id = limpiar_nombre_curso(curso_raw).lower()
            
            pdf_path = buscar_pdf_especifico(pdf_folder, email_id, curso_id)

            if not pdf_path:
                callback_log(f"⚠️  No encuentro PDF para: {email_dest} del curso '{curso_raw}'")
                errores += 1
                continue

            # Crear el mensaje
            msg = MIMEMultipart()
            msg['From'] = user
            msg['To'] = email_dest
            msg['Subject'] = f"Tu diploma del curso: {curso_raw}"

            # Cuerpo del email
            body = f"""<p>Hola {nombre},</p>
                       <p>Adjunto te enviamos tu diploma firmado correspondiente al curso <strong>{curso_raw}</strong>.</p>
                       <p>Un saludo.</p>"""
            msg.attach(MIMEText(body, 'html'))

            # Adjuntar el PDF
            with open(pdf_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            
            # Nombre visible del adjunto
            curso_limpio = limpiar_nombre_curso(curso_raw)
            nombre_visible_pdf = f"Diploma_{curso_limpio}.pdf"
            part.add_header("Content-Disposition", f"attachment; filename= {nombre_visible_pdf}")
            msg.attach(part)
            
            # Enviar
            server.send_message(msg)
            callback_log(f"🚀 [ENVIADO SMTP] a {email_dest}")
            enviados += 1

        except Exception as e:
            callback_log(f"[ERROR SMTP] Fila {idx+2}: {e}")
            errores += 1
            
    server.quit()
    callback_log(f"FIN SMTP. Enviados: {enviados} | Errores: {errores}")


def enviar_masivo_outlook(excel_path, pdf_folder, dry_run, callback_log):
    """
    Lógica original usando Outlook de escritorio (Classic).
    dry_run = True -> mail.Display() (abre la ventana)
    dry_run = False -> mail.Send() (envía directo)
    """
    
    # IMPORTANTE: Inicializar COM en este hilo
    pythoncom.CoInitialize()

    try:
        if not os.path.exists(excel_path):
            callback_log("❌ Error: No se encuentra el Excel.")
            return

        df = pd.read_excel(excel_path)
        df.columns = [c.lower().strip() for c in df.columns]

        # Verificar columnas
        if "email" not in df.columns:
            callback_log("❌ Error: El Excel no tiene columna 'email'.")
            return

        try:
            outlook = win32.Dispatch("Outlook.Application")
            callback_log("✅ Conectado a Outlook local.")
        except Exception as e:
            callback_log(f"❌ Error al conectar con Outlook: {e}")
            callback_log("Asegúrate de tener el Outlook 'Clásico' abierto.")
            return

        enviados = 0
        errores = 0
        
        callback_log(f"Iniciando proceso... (Modo Prueba: {dry_run})")

        for idx, row in df.iterrows():
            try:
                email = str(row["email"]).strip()
                nombre = str(row.get("nombre", "")).strip()
                curso_raw = str(row.get("curso_nombre", "")).strip()
                if "@" not in email: continue

                email_id = email_a_id(email)
                curso_id = limpiar_nombre_curso(curso_raw).lower()
                
                # 1. Encontrar el PDF original
                pdf_original_path = buscar_pdf_especifico(pdf_folder, email_id, curso_id)

                if not pdf_original_path:
                    callback_log(f"⚠️ [ERROR] No encuentro PDF para: {email} del curso '{curso_raw}'")
                    errores += 1
                    continue

                mail = outlook.CreateItem(0)
                mail.To = email
                mail.Subject = f"Tu diploma del curso: {curso_raw}"

                mail.HTMLBody = f"""
                <p>Hola {nombre},</p>
                <p>Adjunto te enviamos tu diploma firmado correspondiente al curso <strong>{curso_raw}</strong>.</p>
                <p>Un saludo.</p>
                """

                # --- LÓGICA DE RENOMBRADO UNIVERSAL ---
                with tempfile.TemporaryDirectory() as temp_dir:
                    # 2. Crear nombre de archivo bonito
                    curso_limpio = limpiar_nombre_curso(curso_raw)
                    nombre_visible_pdf = f"Diploma_{curso_limpio}.pdf"
                    
                    # 3. Ruta del archivo temporal
                    ruta_temp_pdf = os.path.join(temp_dir, nombre_visible_pdf)
                    
                    # 4. Copiar el original a la nueva ruta con el nuevo nombre
                    shutil.copy2(pdf_original_path, ruta_temp_pdf)
                    
                    # 5. Adjuntar la copia con el nombre correcto
                    mail.Attachments.Add(ruta_temp_pdf)

                    # Ya no necesitamos DisplayName, el nombre del archivo es el correcto.

                    # 6. Enviar o mostrar borrador
                    if dry_run:
                        callback_log(f"🔭 [PRUEBA] Abriendo borrador para {email} con adjunto '{nombre_visible_pdf}'...")
                        mail.Display()
                    else:
                        mail.Send()
                        callback_log(f"🚀 [ENVIADO] {email} (Curso: {curso_raw})")
                        enviados += 1

            except Exception as e:
                callback_log(f"[ERROR] Fila {idx+2}: {e}")
                errores += 1

        callback_log("-" * 30)
        callback_log(f"FIN PROCESO. Enviados: {enviados} | Errores: {errores}")

    finally:
        # Liberar recursos COM
        pythoncom.CoUninitialize()
        
def buscar_pdf_especifico(carpeta, email_id, curso_id):
    """
    Busca un archivo que contenga TANTO el email_id COMO el curso_id.
    Ejemplo busca: juan_perez + Curso_Python
    Archivo real: juan_perez__Curso_Python.pdf
    """
    if not os.path.exists(carpeta):
        return None

    candidatos = []
    
    for f in os.listdir(carpeta):
        if not f.lower().endswith(".pdf"):
            continue
            
        # El archivo debe contener AMBAS partes
        # Usamos lower() para evitar problemas de mayúsculas/minúsculas
        nombre_archivo = f.lower()
        
        if email_id in nombre_archivo and curso_id in nombre_archivo:
            ruta = os.path.join(carpeta, f)
            candidatos.append(ruta)

    if not candidatos:
        return None

    # Si por algún motivo hay duplicados, cogemos el más reciente
    candidatos.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return candidatos[0]