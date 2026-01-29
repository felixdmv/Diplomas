import os
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import HexColor
import pandas as pd
from utils import texto_seguro, parse_calificacion, cargar_configuracion_plantilla, email_a_id, limpiar_nombre_curso
import firmador

# Valores por defecto por si el JSON está incompleto
DEFAULT_FONT = "Helvetica"
DEFAULT_COLOR = "#000000"

def dibujar_texto_config(c, config_elem, texto, width):
    """Dibuja un elemento de texto basado en su configuración JSON"""
    if not config_elem or not texto:
        return

    # Extraer valores del JSON (con valores seguros por defecto)
    x = config_elem.get("x_mm", 0) * mm
    y = config_elem.get("y_mm", 0) * mm
    tamano = config_elem.get("tamano", 12)
    align = config_elem.get("alineacion", "center")
    fuente = config_elem.get("fuente", DEFAULT_FONT)
    color = config_elem.get("color", DEFAULT_COLOR)

    # Configurar canvas
    try:
        c.setFont(fuente, tamano)
        c.setFillColor(HexColor(color))
    except:
        c.setFont("Helvetica", tamano) # Fallback si la fuente no existe
        c.setFillColor(HexColor("#000000"))

    # Dibujar según alineación
    if align == "center":
        # Si x es 0 en el JSON, asumimos centro de página
        pos_x = x if x > 0 else (width / 2)
        c.drawCentredString(pos_x, y, str(texto))
    elif align == "right":
        c.drawRightString(x, y, str(texto))
    else: # Left
        c.drawString(x, y, str(texto))

def crear_pdf_individual(row, opciones, ruta_salida, logo_global_path, nombre_firmante_personalizado=None):
    mostrar_horas, mostrar_calif, mostrar_extra = opciones
    
    # 1. IDENTIFICAR QUÉ PLANTILLA USAR
    # Buscamos la columna 'id_plantilla' en el excel. Si no existe, None.
    id_plantilla = row.get("id_plantilla")
    
    # 2. CARGAR CONFIGURACIÓN
    config, ruta_fondo = cargar_configuracion_plantilla(id_plantilla)
    
    # 3. PREPARAR PÁGINA
    orientacion = config.get("orientacion", "landscape").lower()
    if orientacion == "portrait":
        pagesize = portrait(A4)
        width, height = portrait(A4)
    else:
        pagesize = landscape(A4)
        width, height = landscape(A4)
        
    c = canvas.Canvas(ruta_salida, pagesize=pagesize)

    # 4. DIBUJAR FONDO
    if ruta_fondo:
        try:
            c.drawImage(ruta_fondo, 0, 0, width=width, height=height)
        except Exception:
            pass # Si falla, sale blanco
    else:
        # Si no hay imagen, un marco básico de seguridad
        c.setLineWidth(1)
        c.rect(10*mm, 10*mm, width-20*mm, height-20*mm)

    # 5. DIBUJAR ELEMENTOS DEFINIDOS EN JSON
    elems = config.get("elementos", {})
    
    # -- Datos básicos --
    
    # Nombre Alumno
    nombre_completo = " ".join([
        texto_seguro(row.get("nombre")),
        texto_seguro(row.get("apellido1")),
        texto_seguro(row.get("apellido2"))
    ]).strip().upper()
    dibujar_texto_config(c, elems.get("nombre_alumno"), nombre_completo, width)

    # Nombre Curso
    curso = texto_seguro(row.get("curso_nombre"))
    dibujar_texto_config(c, elems.get("nombre_curso"), curso, width)

    # Fecha
    fecha = texto_seguro(row.get("fecha"))
    if fecha:
        txt_fecha = f"Fecha: {fecha}" # Puedes ajustar el prefijo aquí o en el Excel
        dibujar_texto_config(c, elems.get("fecha"), txt_fecha, width)
        
    # 5. FIRMA / RESPONSABLE
    elem_firma = elems.get("firma", {})
    
    # Coordenadas base
    x = elem_firma.get("x_mm", 200) * mm
    y = elem_firma.get("y_mm", 35) * mm
    tamano_base = elem_firma.get("tamano", 12)
    alineacion = elem_firma.get("alineacion", "center")
    
    texto_a_pintar = "Firma Responsable"
    es_digital = False

    if nombre_firmante_personalizado:
        texto_a_pintar = nombre_firmante_personalizado.upper()
        es_digital = True

    # --- LÓGICA ANTI-DESBORDAMIENTO (AUTO-AJUSTE) ---
    from reportlab.pdfbase.pdfmetrics import stringWidth
    
    # 1. Definimos el ancho máximo permitido para la firma (aprox 9 cm)
    max_ancho_permitido = 90 * mm 
    
    # 2. Empezamos con el tamaño de fuente deseado
    tamano_actual = tamano_base
    font_name = "Helvetica-Bold"
    
    # 3. Calculamos cuánto ocupa el texto
    ancho_texto = stringWidth(texto_a_pintar, font_name, tamano_actual)
    
    # 4. Si se pasa, reducimos el tamaño en un bucle hasta que quepa
    while ancho_texto > max_ancho_permitido and tamano_actual > 6:
        tamano_actual -= 0.5
        ancho_texto = stringWidth(texto_a_pintar, font_name, tamano_actual)
        
    # --- FIN LÓGICA AUTO-AJUSTE ---

    # Configuramos fuente con el tamaño calculado (puede ser menor que el base)
    c.setFont(font_name, tamano_actual)
    c.setFillColor(HexColor("#000000"))

    # DIBUJAMOS
    if alineacion == "center":
        c.drawCentredString(x, y, texto_a_pintar)
        if es_digital:
            c.setFont("Helvetica", tamano_actual - 3) # Subtítulo proporcional
            c.setFillColor(HexColor("#555555"))
            c.drawCentredString(x, y - (4*mm), "(Firmado Digitalmente)")
            
    elif alineacion == "right":
        c.drawRightString(x, y, texto_a_pintar)
        if es_digital:
            c.setFont("Helvetica", tamano_actual - 3)
            c.setFillColor(HexColor("#555555"))
            c.drawRightString(x, y - (4*mm), "(Firmado Digitalmente)")
            
    else: # Left
        c.drawString(x, y, texto_a_pintar)
        if es_digital:
            c.setFont("Helvetica", tamano_actual - 3)
            c.setFillColor(HexColor("#555555"))
            c.drawString(x, y - (4*mm), "(Firmado Digitalmente)")

    c.showPage()
    c.save()

# --- Funciones de control (igual que antes) ---

def generar_preview(excel_path, opciones, logo_path, callback_log):
    # (Igual que antes pero usa crear_pdf_individual actualizado)
    from utils import texto_seguro, abrir_archivo 
    try:
        df = pd.read_excel(excel_path)
        df.columns = [c.lower().strip() for c in df.columns]
        if df.empty: return

        row_demo = None
        for i, row in df.iterrows():
            if texto_seguro(row.get("email")) and texto_seguro(row.get("nombre")):
                row_demo = row
                break
        
        if row_demo is None:
            callback_log("❌ No hay datos válidos para preview.")
            return

        ruta_temp = "preview_diploma.pdf"
        crear_pdf_individual(row_demo, opciones, ruta_temp, logo_path)
        abrir_archivo(ruta_temp)

    except Exception as e:
        callback_log(f"❌ Error preview: {e}")

def procesar_excel_y_generar(excel_path, output_folder, opciones, logo_path, callback_log, datos_firma=None, nombre_firmante=None):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    df = pd.read_excel(excel_path)
    df.columns = [c.lower().strip() for c in df.columns]

    # Validar si existe columna id_plantilla, si no, avisar
    if "id_plantilla" not in df.columns:
        callback_log("ℹ️ Columna 'id_plantilla' no encontrada. Usando plantilla 'default'.")

    generados = 0
    errores = 0

    for i, row in df.iterrows():
        try:
            email = texto_seguro(row.get("email"))
            curso = texto_seguro(row.get("curso_nombre"))

            if not email or "@" not in email:
                errores += 1
                continue

            # 1. Identificador del Email
            id_email = email_a_id(email)
            
            # 2. Identificador del Curso (Nuevo)
            id_curso = limpiar_nombre_curso(curso)
            
            # 3. Nombre compuesto único
            filename = f"{id_email}__{id_curso}.pdf"  # Uso doble guion bajo para separar visualmente
            
            ruta_completa = os.path.join(output_folder, filename)

            crear_pdf_individual(row, opciones, ruta_completa, logo_path, nombre_firmante)
            msg_extra = ""
            
            # 2. FIRMAR EL PDF (Si toca)
            if datos_firma:
                ruta_pfx, pass_pfx = datos_firma
                ok_firma = firmador.firmar_pdf(ruta_completa, ruta_pfx, pass_pfx)
                if ok_firma:
                    msg_extra = " [🔏 FIRMADO]"
                else:
                    msg_extra = " [❌ ERROR FIRMA]"
                    # Opcional: Contar como error o dejarlo pasar sin firmar
            
            generados += 1
            # Añadimos feedback visual al log
            callback_log(f"Generado: {filename}{msg_extra}") # Si quieres mucho detalle
            
        except Exception as e:
            callback_log(f"[ERROR] Fila {i+2}: {e}")
            errores += 1

    callback_log(f"FIN. Generados: {generados} | Errores: {errores}")