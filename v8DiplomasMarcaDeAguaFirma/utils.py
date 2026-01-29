import os
import sys
import re
import pandas as pd
import json
import unicodedata
from cryptography.hazmat.primitives.serialization import pkcs12
from cryptography.hazmat.primitives import serialization
from cryptography.x509.oid import NameOID

def obtener_ruta_plantillas():
    """Devuelve la ruta de la carpeta 'plantillas' junto al ejecutable"""
    if getattr(sys, 'frozen', False):
        # Estamos en modo EXE
        base_dir = os.path.dirname(sys.executable)
    else:
        # Estamos en modo script .py
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    ruta = os.path.join(base_dir, "plantillas")
    if not os.path.exists(ruta):
        os.makedirs(ruta) # Crear si no existe para evitar errores
    return ruta

def cargar_configuracion_plantilla(nombre_plantilla):
    """
    Busca la plantilla en este orden:
    1. Carpeta 'plantillas' junto al EXE (Personalizadas por el usuario).
    2. Carpeta 'plantillas' dentro del EXE (Integradas de fábrica).
    """
    if not nombre_plantilla or pd.isna(nombre_plantilla):
        nombre_plantilla = "default"
    
    nombre_plantilla = str(nombre_plantilla).strip()
    
    # Rutas posibles
    # A) Externa (Junto al .exe)
    if getattr(sys, 'frozen', False):
        base_externa = os.path.dirname(sys.executable)
    else:
        base_externa = os.path.dirname(os.path.abspath(__file__))
        
    ruta_externa = os.path.join(base_externa, "plantillas", nombre_plantilla)
    
    # B) Interna (Dentro del .exe / _MEIPASS)
    ruta_interna = resource_path(os.path.join("plantillas", nombre_plantilla))
    
    # Decisión: ¿Cual usamos?
    carpeta_tema = None
    
    if os.path.exists(ruta_externa):
        # Prioridad 1: El usuario ha creado una carpeta fuera
        carpeta_tema = ruta_externa
    elif os.path.exists(ruta_interna):
        # Prioridad 2: Usamos la que viene dentro del EXE
        carpeta_tema = ruta_interna
    else:
        # Fallback: Si no existe la pedida, intentamos buscar 'default' interna
        print(f"⚠️ Plantilla '{nombre_plantilla}' no encontrada. Usando default.")
        carpeta_tema = resource_path(os.path.join("plantillas", "default"))

    # Cargar Config
    ruta_json = os.path.join(carpeta_tema, "config.json")
    
    # Buscar imagen fondo
    ruta_fondo = None
    for ext in [".png", ".jpg", ".jpeg"]:
        temp = os.path.join(carpeta_tema, "fondo" + ext)
        if os.path.exists(temp):
            ruta_fondo = temp
            break
            
    config = {"orientacion": "landscape", "elementos": {}}
    
    if os.path.exists(ruta_json):
        try:
            with open(ruta_json, 'r', encoding='utf-8') as f:
                config = json.load(f)
        except Exception as e:
            print(f"Error JSON: {e}")

    return config, ruta_fondo

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
        

def limpiar_nombre_curso(texto):
    """
    Convierte 'Curso de Introducción a Python 3.0!' en 'Curso_de_Introduccion_a_Python_3_0'
    Para usarlo en nombres de archivo.
    """
    if not texto:
        return "diploma"
    
    # 1. Convertir a string y quitar espacios extremos
    s = str(texto).strip()
    
    # 2. Normalizar unicode (quitar tildes: á -> a, ñ -> n)
    # (NFD separa caracteres, ascii ignore descarta los acentos separados)
    s = unicodedata.normalize('NFD', s).encode('ascii', 'ignore').decode('utf-8')
    
    # 3. Reemplazar espacios y símbolos raros por guiones bajos
    # Solo dejamos letras, números y guiones bajos
    s = re.sub(r'[^a-zA-Z0-9]', '_', s)
    
    # 4. Quitar guiones bajos duplicados (___ -> _)
    s = re.sub(r'_+', '_', s)
    
    # 5. Si queda muy largo, cortar (Windows tiene límite de 260 chars en rutas)
    return s[:50]

def obtener_nombre_del_pfx(ruta_pfx, password):
    """
    Abre un archivo .pfx y extrae el 'Common Name' (Nombre y Apellidos) del propietario.
    Devuelve un string (ej: 'FELIX DE MIGUEL') o None si falla.
    """
    if not os.path.exists(ruta_pfx):
        return None

    try:
        with open(ruta_pfx, "rb") as f:
            pfx_data = f.read()

        # Cargar el PFX
        private_key, certificate, additional_certificates = pkcs12.load_key_and_certificates(
            pfx_data, 
            password.encode("utf-8") if password else None
        )

        # Extraer el sujeto (Subject) del certificado
        subject = certificate.subject
        
        # Buscar el campo "Common Name" (CN)
        # Esto suele devolver una lista, cogemos el primero
        common_names = subject.get_attributes_for_oid(NameOID.COMMON_NAME)
        
        if common_names:
            return common_names[0].value
        else:
            return "Firma Digital Verificada" # Fallback si no tiene CN
            
    except Exception as e:
        print(f"Error leyendo PFX: {e}")
        return None