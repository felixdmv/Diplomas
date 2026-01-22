import os
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
import pandas as pd
from utils import texto_seguro, parse_calificacion, email_a_id, abrir_archivo

def crear_pdf_individual(row, opciones, ruta_salida, logo_path):
    mostrar_horas, mostrar_calif, mostrar_extra = opciones
    
    c = canvas.Canvas(ruta_salida, pagesize=landscape(A4))
    width, height = landscape(A4)

    # Marco
    c.setLineWidth(3)
    c.rect(2*cm, 2*cm, width-4*cm, height-4*cm)

    # Logo
    if logo_path and os.path.exists(logo_path):
        try:
            img = ImageReader(logo_path)
            logo_width = 4*cm
            iw, ih = img.getSize()
            factor = logo_width / iw
            logo_height = ih * factor
            c.drawImage(img, 2.5*cm, height - 2.5*cm - logo_height, width=logo_width, height=logo_height, mask='auto')
        except Exception:
            pass

    y = height - 8*cm

    # Contenido
    c.setFont("Times-Bold", 32)
    c.drawCentredString(width/2, y, "DIPLOMA")

    y -= 2*cm
    c.setFont("Times-Roman", 16)
    c.drawCentredString(width/2, y, "Se certifica que")

    y -= 2*cm
    c.setFont("Times-Bold", 26)
    nombre_completo = " ".join([
        texto_seguro(row.get("nombre")),
        texto_seguro(row.get("apellido1")),
        texto_seguro(row.get("apellido2"))
    ]).strip()
    c.drawCentredString(width/2, y, nombre_completo)

    y -= 2*cm
    c.setFont("Times-Roman", 16)
    c.drawCentredString(width/2, y, "ha superado el curso")

    y -= 1.2*cm
    c.setFont("Times-Bold", 18)
    c.drawCentredString(width/2, y, texto_seguro(row.get("curso_nombre")))

    y -= 2*cm
    c.setFont("Times-Roman", 14)

    if mostrar_horas:
        horas = texto_seguro(row.get("horas"))
        if horas:
            c.drawCentredString(width/2, y, f"Duración: {horas} horas")
            y -= 1*cm

    if mostrar_calif:
        calif = parse_calificacion(row.get("calificacion_num"))
        nota_txt = texto_seguro(row.get("calificacion_texto"))
        if calif is not None:
            c.drawCentredString(width/2, y, f"Calificación: {calif}")
            y -= 1*cm
        elif nota_txt:
            c.drawCentredString(width/2, y, f"Resultado: {nota_txt}")
            y -= 1*cm

    if mostrar_extra:
        extra = texto_seguro(row.get("extra_1"))
        if extra:
            c.drawCentredString(width/2, y, extra)
            y -= 1*cm

    # Fecha y Firma
    c.setFont("Times-Roman", 12)
    c.drawRightString(width - 3*cm, 3*cm, f"Fecha: {texto_seguro(row.get('fecha'))}")
    c.line(width - 10*cm, 5*cm, width - 3*cm, 5*cm)
    c.drawRightString(width - 3*cm, 4.3*cm, "Firma")

    c.showPage()
    c.save()

def procesar_excel_y_generar(excel_path, output_folder, opciones, logo_path, callback_log):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    df = pd.read_excel(excel_path)
    df.columns = [c.lower().strip() for c in df.columns]

    generados = 0
    errores = 0

    callback_log(f"Iniciando generación de {len(df)} diplomas...")

    for i, row in df.iterrows():
        try:
            email = texto_seguro(row.get("email"))
            if not email or "@" not in email:
                callback_log(f"[SKIP] Fila {i+2}: Email inválido")
                errores += 1
                continue

            identificador = email_a_id(email)
            filename = f"{identificador}.pdf"
            ruta_completa = os.path.join(output_folder, filename)

            crear_pdf_individual(row, opciones, ruta_completa, logo_path)
            generados += 1
            
        except Exception as e:
            callback_log(f"[ERROR] Fila {i+2}: {e}")
            errores += 1

    callback_log("-" * 30)
    callback_log(f"FIN GENERACIÓN. OK: {generados} | Errores: {errores}")
    

def generar_preview(excel_path, opciones, logo_path, callback_log):
    """Genera un único diploma basado en la primera fila del Excel y lo abre"""
    try:
        df = pd.read_excel(excel_path)
        df.columns = [c.lower().strip() for c in df.columns]
        
        if df.empty:
            callback_log("❌ El Excel está vacío.")
            return

        # Buscamos la primera fila que tenga datos mínimos
        row_demo = None
        for i, row in df.iterrows():
            if texto_seguro(row.get("email")) and texto_seguro(row.get("nombre")):
                row_demo = row
                break
        
        if row_demo is None:
            callback_log("❌ No se encontró ninguna fila válida (con email y nombre) en el Excel.")
            return

        ruta_temp = "preview_diploma.pdf"
        
        # Generamos el PDF
        crear_pdf_individual(row_demo, opciones, ruta_temp, logo_path)
        
        callback_log(f"👁️ Vista previa generada: {ruta_temp}")
        
        # Abrimos el archivo automáticamente
        abrir_archivo(ruta_temp)

    except Exception as e:
        callback_log(f"❌ Error en vista previa: {e}")