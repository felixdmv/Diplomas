import os
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import HexColor
import pandas as pd
from utils import texto_seguro, parse_calificacion, cargar_configuracion_plantilla, email_a_id, limpiar_nombre_curso

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

def crear_pdf_individual(row, opciones, ruta_salida, logo_global_path):
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
        
    # Firma (Etiqueta de texto)
    dibujar_texto_config(c, elems.get("firma"), "Firma Responsable", width)

    # -- Detalles (Horas, Nota, Extra) --
    # Este es especial porque son líneas dinámicas.
    # Usaremos la config "detalles" para saber dónde empieza (Y) y cuánto baja por línea.
    
    conf_det = elems.get("detalles")
    if conf_det:
        lineas = []
        if mostrar_horas:
            h = texto_seguro(row.get("horas"))
            if h: lineas.append(f"Duración: {h} horas")
        
        if mostrar_calif:
            cal = parse_calificacion(row.get("calificacion_num"))
            txt = texto_seguro(row.get("calificacion_texto"))
            if cal: lineas.append(f"Calificación: {cal}")
            elif txt: lineas.append(f"Resultado: {txt}")
            
        if mostrar_extra:
            ext = texto_seguro(row.get("extra_1"))
            if ext: lineas.append(ext)
            
        # Dibujar líneas en bucle
        y_inicial = conf_det.get("y_mm", 50) * mm
        salto = conf_det.get("interlineado", 5) * mm
        
        for i, linea in enumerate(lineas):
            # Clonamos la config para alterar solo la Y
            conf_temp = conf_det.copy()
            conf_temp["y_mm"] = (y_inicial - (i * salto)) / mm
            dibujar_texto_config(c, conf_temp, linea, width)

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

def procesar_excel_y_generar(excel_path, output_folder, opciones, logo_path, callback_log):
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

            crear_pdf_individual(row, opciones, ruta_completa, logo_path)
            generados += 1
            
        except Exception as e:
            callback_log(f"[ERROR] Fila {i+2}: {e}")
            errores += 1

    callback_log(f"FIN. Generados: {generados} | Errores: {errores}")