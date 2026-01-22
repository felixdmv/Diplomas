import os
import sys
import re
import pandas as pd

def resource_path(relative_path):
    """Ruta absoluta para recursos, compatible con PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def texto_seguro(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()

def email_a_id(email):
    return email.lower().strip().replace("@", "_").replace(".", "_")

def parse_calificacion(valor):
    if pd.isna(valor):
        return None
    s = str(valor).replace(",", ".")
    try:
        f = float(s)
        if 0 <= f <= 10:
            return f
    except:
        return None
    return None

def buscar_pdf_correcto(carpeta, email_id):
    """Busca el PDF más reciente que empiece por el email_id"""
    if not os.path.exists(carpeta):
        return None

    candidatos = []
    for f in os.listdir(carpeta):
        if not f.lower().endswith(".pdf"):
            continue
        if f.startswith(email_id):
            ruta = os.path.join(carpeta, f)
            candidatos.append(ruta)

    if not candidatos:
        return None

    # Ordenar por fecha de modificación (el más nuevo primero)
    candidatos.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return candidatos[0]

def abrir_archivo(ruta):
    """Abre el archivo con el programa predeterminado del sistema (Windows/Mac/Linux)"""
    import sys
    import subprocess
    
    try:
        if sys.platform == 'win32':
            os.startfile(ruta)
        elif sys.platform == 'darwin':
            subprocess.call(('open', ruta))
        else:
            subprocess.call(('xdg-open', ruta))
    except Exception as e:
        print(f"No se pudo abrir el archivo: {e}")