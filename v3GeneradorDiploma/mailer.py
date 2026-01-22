import os
import pandas as pd
import win32com.client as win32
import pythoncom  # Necesario para multihilo con COM
from utils import email_a_id, buscar_pdf_correcto

NOMBRE_VISIBLE_PDF = "DiplomaBiblioteca.pdf"

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

                if "@" not in email:
                    callback_log(f"[SKIP] Fila {idx+2}: Email inválido ({email})")
                    errores += 1
                    continue

                email_id = email_a_id(email)
                pdf_path = buscar_pdf_correcto(pdf_folder, email_id)

                if not pdf_path:
                    callback_log(f"[ERROR] No hay PDF para: {email}")
                    errores += 1
                    continue

                # Crear correo
                mail = outlook.CreateItem(0)
                mail.To = email
                mail.Subject = "Tu diploma del curso"
                
                mail.HTMLBody = f"""
                <p>Hola {nombre},</p>
                <p>Adjunto te enviamos tu diploma firmado en formato PDF.</p>
                <p>Un saludo.</p>
                """

                # Adjuntar y renombrar visualmente
                adjunto = mail.Attachments.Add(pdf_path)
                try:
                    adjunto.DisplayName = NOMBRE_VISIBLE_PDF
                except:
                    pass # A veces falla renombrar, no es crítico

                if dry_run:
                    callback_log(f"🔭 [PRUEBA] Abriendo borrador para {email}...")
                    mail.Display() # Abre la ventana de Outlook
                else:
                    mail.Send()
                    callback_log(f"🚀 [ENVIADO] {email}")
                    enviados += 1

            except Exception as e:
                callback_log(f"[ERROR] Fila {idx+2}: {e}")
                errores += 1

        callback_log("-" * 30)
        callback_log(f"FIN PROCESO. Enviados: {enviados} | Errores: {errores}")

    finally:
        # Liberar recursos COM
        pythoncom.CoUninitialize()