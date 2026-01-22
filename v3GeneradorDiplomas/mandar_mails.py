import os
import re
import unicodedata
import pandas as pd
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog, messagebox


# =========================
# CONFIGURACIÓN
# =========================

DRY_RUN = True   # True = abrir borrador / False = enviar
NOMBRE_VISIBLE_PDF = "DiplomaBiblioteca.pdf"


# =========================
# UTILIDADES
# =========================

def email_a_id(email):
    return email.lower().strip().replace("@", "_").replace(".", "_")


def seleccionar_excel():
    return filedialog.askopenfilename(
        title="Seleccionar Excel de alumnos",
        filetypes=[("Excel files", "*.xlsx")]
    )


def seleccionar_carpeta():
    return filedialog.askdirectory(
        title="Seleccionar carpeta con diplomas firmados"
    )


def buscar_pdf_correcto(carpeta, email_id):
    """
    Busca PDFs que empiecen por email_id.
    Si hay varios (ej. _signed), devuelve el más reciente.
    """
    candidatos = []

    for f in os.listdir(carpeta):
        if not f.lower().endswith(".pdf"):
            continue
        if f.startswith(email_id):
            ruta = os.path.join(carpeta, f)
            candidatos.append(ruta)

    if not candidatos:
        return None

    # Elegimos el más reciente
    candidatos.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return candidatos[0]


# =========================
# MAIN
# =========================

def main():
    root = tk.Tk()
    root.withdraw()

    excel_path = seleccionar_excel()
    if not excel_path:
        messagebox.showerror("Error", "No se seleccionó ningún Excel")
        return

    pdf_folder = seleccionar_carpeta()
    if not pdf_folder:
        messagebox.showerror("Error", "No se seleccionó carpeta de PDFs")
        return

    df = pd.read_excel(excel_path)
    df.columns = [c.lower().strip() for c in df.columns]

    campos_necesarios = {"email", "nombre"}
    if not campos_necesarios.issubset(df.columns):
        messagebox.showerror(
            "Error",
            "El Excel debe contener al menos las columnas: email, nombre"
        )
        return

    outlook = win32.Dispatch("Outlook.Application")

    enviados = 0
    errores = 0

    for idx, row in df.iterrows():
        try:
            email = str(row["email"]).strip()

            if "@" not in email:
                print(f"[ERROR] Email inválido fila {idx+2}: {email}")
                errores += 1
                continue

            email_id = email_a_id(email)

            pdf_path = buscar_pdf_correcto(pdf_folder, email_id)

            if not pdf_path:
                print(f"[ERROR] No se encontró PDF para {email}")
                errores += 1
                continue

            nombre = str(row.get("nombre", "")).strip()

            mail = outlook.CreateItem(0)
            mail.To = email
            mail.Subject = "Tu diploma del curso"

            mail.HTMLBody = f"""
            <p>Hola {nombre},</p>
            <p>Adjunto te enviamos tu diploma firmado en formato PDF.</p>
            <p>Un saludo.</p>
            """

            adjunto = mail.Attachments.Add(pdf_path)
            adjunto.DisplayName = NOMBRE_VISIBLE_PDF

            if DRY_RUN:
                print(f"[PRUEBA] Borrador para {email} → {os.path.basename(pdf_path)}")
                mail.Display()
            else:
                mail.Send()
                print(f"[OK] Enviado a {email}")
                enviados += 1

        except Exception as e:
            print(f"[ERROR] Fila {idx+2}: {e}")
            errores += 1

    print("\n--------- RESUMEN ---------")
    print("Enviados:", enviados)
    print("Errores: ", errores)
    print("Modo prueba:", DRY_RUN)
    print("----------------------------")


if __name__ == "__main__":
    main()
