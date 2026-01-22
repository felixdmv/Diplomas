import os
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm, mm
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import HexColor, Color
import pandas as pd
from utils import texto_seguro, parse_calificacion, email_a_id, resource_path

# --- CONFIGURACIÓN DE ESTILO ---
COLOR_TEXTO_FIJO = HexColor("#555555")  # Gris elegante
COLOR_DATOS = HexColor("#000000")       # Negro puro
FUENTE_TITULO = "Helvetica-Bold"
FUENTE_TEXTO = "Helvetica"

def dibujar_fondo(c, width, height, logo_path):
    """
    Intenta cargar 'fondo.png' o 'fondo.jpg' desde la carpeta imgs.
    Si no existe, dibuja un marco clásico elegante.
    """
    # Rutas posibles para el fondo
    ruta_fondo_png = resource_path(os.path.join("imgs", "fondo.png"))
    ruta_fondo_jpg = resource_path(os.path.join("imgs", "fondo.jpg"))
    
    fondo_usado = None
    if os.path.exists(ruta_fondo_png):
        fondo_usado = ruta_fondo_png
    elif os.path.exists(ruta_fondo_jpg):
        fondo_usado = ruta_fondo_jpg

    if fondo_usado:
        # OPCIÓN A: DIBUJAR IMAGEN DE FONDO (DISEÑO PRO)
        try:
            c.drawImage(fondo_usado, 0, 0, width=width, height=height)
            # Si usas fondo, asumimos que el logo ya está en el diseño.
            # Si quieres forzar el logo encima, descomenta las líneas de abajo.
        except Exception as e:
            print(f"Error cargando fondo: {e}")
    else:
        # OPCIÓN B: MARCO CLÁSICO GENERADO POR CÓDIGO (SI NO HAY IMAGEN)
        # Marco doble elegante
        c.setStrokeColor(HexColor("#333333"))
        c.setLineWidth(3)
        c.rect(1.5*cm, 1.5*cm, width-3*cm, height-3*cm) # Exterior
        
        c.setLineWidth(1)
        c.rect(1.7*cm, 1.7*cm, width-3.4*cm, height-3.4*cm) # Interior

        # Logo automático si no hay fondo
        if logo_path and os.path.exists(logo_path):
            try:
                img = ImageReader(logo_path)
                logo_width = 4.5*cm
                iw, ih = img.getSize()
                factor = logo_width / iw
                logo_height = ih * factor
                # Centrado arriba
                c.drawImage(img, (width - logo_width)/2, height - 2.5*cm - logo_height, 
                           width=logo_width, height=logo_height, mask='auto')
            except:
                pass

def crear_pdf_individual(row, opciones, ruta_salida, logo_path):
    mostrar_horas, mostrar_calif, mostrar_extra = opciones
    
    c = canvas.Canvas(ruta_salida, pagesize=landscape(A4))
    width, height = landscape(A4)
    centro_x = width / 2

    # 1. DIBUJAR EL "LIENZO" (Fondo o Marco)
    dibujar_fondo(c, width, height, logo_path)

    # 2. DEFINIR DATOS
    nombre_completo = " ".join([
        texto_seguro(row.get("nombre")),
        texto_seguro(row.get("apellido1")),
        texto_seguro(row.get("apellido2"))
    ]).strip().upper() # Nombre en mayúsculas queda mejor
    
    curso = texto_seguro(row.get("curso_nombre"))
    fecha = texto_seguro(row.get("fecha"))

    # --- COORDENADAS (AJUSTAR AQUÍ SI CAMBIAS EL DISEÑO) ---
    # Empezamos más arriba para tener espacio
    y_actual = height - 7.5*cm 

    # TÍTULO (Si tu fondo ya tiene la palabra DIPLOMA, comenta estas 2 líneas)
    c.setFillColor(COLOR_TEXTO_FIJO)
    c.setFont(FUENTE_TITULO, 36)
    c.drawCentredString(centro_x, y_actual, "DIPLOMA DE CERTIFICACIÓN")
    y_actual -= 2.0*cm

    # "Se otorga a"
    c.setFillColor(COLOR_TEXTO_FIJO)
    c.setFont(FUENTE_TEXTO, 14)
    c.drawCentredString(centro_x, y_actual, "Se otorga el presente reconocimiento a:")
    y_actual -= 1.5*cm

    # NOMBRE DEL ALUMNO (Destacado)
    c.setFillColor(COLOR_DATOS)
    c.setFont(FUENTE_TITULO, 26)
    c.drawCentredString(centro_x, y_actual, nombre_completo)
    y_actual -= 0.5*cm
    
    # Línea decorativa bajo el nombre
    c.setStrokeColor(COLOR_TEXTO_FIJO)
    c.setLineWidth(0.5)
    c.line(centro_x - 6*cm, y_actual, centro_x + 6*cm, y_actual)
    y_actual -= 1.5*cm

    # "Por haber completado..."
    c.setFillColor(COLOR_TEXTO_FIJO)
    c.setFont(FUENTE_TEXTO, 14)
    c.drawCentredString(centro_x, y_actual, "Por haber completado satisfactoriamente el curso:")
    y_actual -= 1.2*cm

    # NOMBRE DEL CURSO
    c.setFillColor(COLOR_DATOS)
    c.setFont(FUENTE_TITULO, 20)
    c.drawCentredString(centro_x, y_actual, curso)
    y_actual -= 1.5*cm

    # --- BLOQUE DE DETALLES (Horas, Nota, Extra) ---
    # Usaremos un tamaño de fuente menor para que no ocupe tanto
    c.setFont(FUENTE_TEXTO, 12)
    c.setFillColor(HexColor("#444444"))

    detalles = []
    if mostrar_horas:
        h = texto_seguro(row.get("horas"))
        if h: detalles.append(f"Duración: {h} horas")
    
    if mostrar_calif:
        calif = parse_calificacion(row.get("calificacion_num"))
        nota_txt = texto_seguro(row.get("calificacion_texto"))
        if calif is not None: detalles.append(f"Calificación: {calif}")
        elif nota_txt: detalles.append(f"Resultado: {nota_txt}")

    if mostrar_extra:
        ext = texto_seguro(row.get("extra_1"))
        if ext: detalles.append(ext)

    # Imprimir detalles centrados con separación pequeña
    for detalle in detalles:
        c.drawCentredString(centro_x, y_actual, detalle)
        y_actual -= 0.6*cm  # Salto pequeño entre detalles

    # --- PIE DE PÁGINA (FECHA Y FIRMA) ---
    # Lo posicionamos fijo abajo, independiente de lo anterior para evitar solapamiento
    pos_pie_y = 4*cm
    
    # Fecha (Izquierda o Centro, según prefieras. Aquí la pongo a la izquierda alineada)
    c.setFont(FUENTE_TEXTO, 12)
    c.drawString(3*cm, pos_pie_y, f"Fecha: {fecha}")

    # Firma (Derecha)
    c.line(width - 9*cm, pos_pie_y + 0.5*cm, width - 3*cm, pos_pie_y + 0.5*cm) # Línea firma
    c.drawCentredString(width - 6*cm, pos_pie_y, "Firma del Responsable")

    c.showPage()
    c.save()

def generar_preview(excel_path, opciones, logo_path, callback_log):
    """Genera un único diploma basado en la primera fila del Excel y lo abre"""
    # IMPORTACIÓN LOCAL PARA EVITAR CICLOS Y ASEGURAR UTILS
    from utils import texto_seguro, abrir_archivo 
    
    try:
        df = pd.read_excel(excel_path)
        df.columns = [c.lower().strip() for c in df.columns]
        
        if df.empty:
            callback_log("❌ El Excel está vacío.")
            return

        row_demo = None
        for i, row in df.iterrows():
            if texto_seguro(row.get("email")) and texto_seguro(row.get("nombre")):
                row_demo = row
                break
        
        if row_demo is None:
            callback_log("❌ No se encontró ninguna fila válida para la preview.")
            return

        ruta_temp = "preview_diploma.pdf"
        crear_pdf_individual(row_demo, opciones, ruta_temp, logo_path)
        callback_log(f"👁️ Abriendo vista previa...")
        abrir_archivo(ruta_temp)

    except Exception as e:
        callback_log(f"❌ Error preview: {e}")

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

    callback_log(f"FIN. Generados: {generados} | Errores: {errores}")