
import os
import pandas as pd
import win32com.client as win32

# ---------------- CONFIGURACIÓN ---------------- #

EXCEL_PATH = "alumnos_diplomas_ejemplo.xlsx"  # nombre del Excel
SHEET_NAME = "Hoja1"                          # cámbialo si tu hoja se llama distinto
PDF_FOLDER = "pdfs"                           # carpeta donde guardas los PDFs

DRY_RUN = True   # True = mostrar borrador en Outlook / False = enviar

# ------------------------------------------------ #


def main():
    # Leer Excel
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, engine="openpyxl")

    # Normalizar nombres de columnas
    df.columns = [c.lower().strip() for c in df.columns]

    if "email" not in df.columns:
        raise ValueError("ERROR: el Excel debe tener una columna llamada 'email'.")

    # Preparar Outlook
    outlook = win32.Dispatch("Outlook.Application")

    enviados = 0
    errores = 0

    for idx, row in df.iterrows():
        try:
            email = str(row["email"]).strip()

            if "@" not in email:
                print(f"[ERROR] Email inválido en fila {idx+2}: {email}")
                errores += 1
                continue

            # Nombre del PDF basado en la parte local del email
            pdf_filename = email.split("@")[0] + ".pdf"
            pdf_path = os.path.join(PDF_FOLDER, pdf_filename)

            # Comprobar que existe el PDF
            if not os.path.exists(pdf_path):
                print(f"[ERROR] No existe el PDF esperado: {pdf_filename}")
                errores += 1
                continue

            # Crear correo
            mail = outlook.CreateItem(0)
            mail.To = email
            mail.Subject = "Tu diploma en PDF"

            mail.HTMLBody = """
            <p>Hola,</p>
            <p>Adjunto te enviamos tu diploma en formato PDF.</p>
            <p>Un saludo.</p>
            """

            # Adjuntar PDF
            mail.Attachments.Add(pdf_path)

            if DRY_RUN:
                print(f"[PRUEBA] Abriendo borrador para: {email}")
                mail.Display()
            else:
                mail.Send()
                print(f"[OK] Enviado a: {email}")
                enviados += 1

        except Exception as e:
            print(f"[ERROR] En fila {idx+2} enviando a {email}: {e}")
            errores += 1

    print("\n--------- RESUMEN ---------")
    print("Enviados: ", enviados)
    print("Errores:  ", errores)
    print("Modo prueba: ", DRY_RUN)
    print("----------------------------")


if __name__ == "__main__":
    main()
