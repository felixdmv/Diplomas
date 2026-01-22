import os
import sys
import unicodedata
import re
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader


# =========================
# UTILIDADES
# =========================

def texto_seguro(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def email_a_id(email):
    email = email.lower().strip()
    return email.replace("@", "_").replace(".", "_")


def abrir_pdf(ruta):
    if sys.platform == "win32":
        os.startfile(ruta)
    elif sys.platform == "darwin":
        os.system(f"open {ruta}")
    else:
        os.system(f"xdg-open {ruta}")


def validar_texto_letras(texto):
    texto = texto if texto is not None else ""
    texto = str(texto)
    return bool(re.match(r"^[A-Za-zÁÉÍÓÚáéíóúÑñüÜ\s]+$", texto))


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


# =========================
# GENERACIÓN PDF
# =========================

def generar_diploma_pdf(row, opciones, salida, logo_path=None):
    mostrar_horas, mostrar_calif, mostrar_extra = opciones

    c = canvas.Canvas(salida, pagesize=landscape(A4))
    width, height = landscape(A4)

    # Marco
    c.setLineWidth(3)
    c.rect(2*cm, 2*cm, width-4*cm, height-4*cm)

    # Logo (esquina superior izquierda)
    if logo_path and os.path.exists(logo_path):
        try:
            img = ImageReader(logo_path)
            logo_width = 4*cm
            iw, ih = img.getSize()
            factor = logo_width / iw
            logo_height = ih * factor
            c.drawImage(
                img,
                2.5*cm,
                height - 2.5*cm - logo_height,
                width=logo_width,
                height=logo_height,
                mask='auto'
            )
        except Exception as e:
            print(f"No se pudo cargar el logo: {e}")

    y = height - 8*cm

    # Título
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

    horas = texto_seguro(row.get("horas"))
    if mostrar_horas and horas:
        c.drawCentredString(width/2, y, f"Duración: {horas} horas")
        y -= 1*cm

    calif = parse_calificacion(row.get("calificacion_num"))
    nota_txt = texto_seguro(row.get("calificacion_texto"))

    if mostrar_calif:
        if calif is not None:
            c.drawCentredString(width/2, y, f"Calificación: {calif}")
            y -= 1*cm
        elif nota_txt:
            c.drawCentredString(width/2, y, f"Resultado: {nota_txt}")
            y -= 1*cm

    extra = texto_seguro(row.get("extra_1"))
    if mostrar_extra and extra:
        c.drawCentredString(width/2, y, extra)
        y -= 1*cm

    # Fecha
    c.setFont("Times-Roman", 12)
    c.drawRightString(width - 3*cm, 3*cm, f"Fecha: {texto_seguro(row.get('fecha'))}")

    # Firma (espacio)
    c.line(width - 10*cm, 5*cm, width - 3*cm, 5*cm)
    c.drawRightString(width - 3*cm, 4.3*cm, "Firma")

    c.showPage()
    c.save()


# =========================
# APLICACIÓN
# =========================

class GeneradorDiplomasApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Diplomas")

        self.excel_path = None
        self.logo_path = os.path.join("imgs", "ubu_logo.png")

        self.var_horas = tk.BooleanVar()
        self.var_calif = tk.BooleanVar()
        self.var_extra = tk.BooleanVar()

        tk.Button(root, text="Seleccionar Excel", command=self.seleccionar_excel).pack(pady=5)

        tk.Checkbutton(root, text="Mostrar horas", variable=self.var_horas).pack(anchor="w")
        tk.Checkbutton(root, text="Mostrar calificación", variable=self.var_calif).pack(anchor="w")
        tk.Checkbutton(root, text="Mostrar campo extra", variable=self.var_extra).pack(anchor="w")

        tk.Button(root, text="Vista previa", command=self.vista_previa).pack(pady=10)
        tk.Button(root, text="Generar diplomas", command=self.generar_todos).pack(pady=5)

    def seleccionar_excel(self):
        path = filedialog.askopenfilename(
            title="Seleccionar Excel",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            self.excel_path = path
            messagebox.showinfo("Excel cargado", os.path.basename(path))

    def obtener_opciones(self):
        return (
            self.var_horas.get(),
            self.var_calif.get(),
            self.var_extra.get()
        )

    def comprobar_excel(self, df):
        filas_invalidas = []
        for i, row in df.iterrows():
            errores = []

            if not texto_seguro(row.get("nombre")):
                errores.append("Nombre vacío")
            if not texto_seguro(row.get("apellido1")):
                errores.append("Apellido1 vacío")
            if not texto_seguro(row.get("curso_nombre")):
                errores.append("Curso vacío")
            if not texto_seguro(row.get("email")):
                errores.append("Email vacío")

            if not validar_texto_letras(row.get("nombre")):
                errores.append("Nombre inválido")
            if not validar_texto_letras(row.get("apellido1")):
                errores.append("Apellido1 inválido")
            if row.get("apellido2") and not validar_texto_letras(row.get("apellido2")):
                errores.append("Apellido2 inválido")

            if errores:
                filas_invalidas.append((i, errores))

        return filas_invalidas

    def vista_previa(self):
        if not self.excel_path:
            messagebox.showerror("Error", "Selecciona un Excel primero")
            return

        df = pd.read_excel(self.excel_path)

        invalidas = self.comprobar_excel(df)
        if invalidas:
            msg = "Filas con errores (no se generarán):\n"
            for i, errores in invalidas:
                msg += f"Fila {i+2}: {', '.join(errores)}\n"
            messagebox.showwarning("Advertencia", msg)

        opciones = self.obtener_opciones()
        ruta = "preview_diploma.pdf"
        generar_diploma_pdf(df.iloc[0], opciones, ruta, self.logo_path)
        abrir_pdf(ruta)

    def generar_todos(self):
        if not self.excel_path:
            messagebox.showerror("Error", "Selecciona un Excel primero")
            return

        df = pd.read_excel(self.excel_path)
        opciones = self.obtener_opciones()

        os.makedirs("pdfs", exist_ok=True)

        invalidas = self.comprobar_excel(df)
        invalidas_idx = [i for i, _ in invalidas]

        for i, row in df.iterrows():
            if i in invalidas_idx:
                continue

            email_id = email_a_id(row["email"])
            nombre_pdf = f"{email_id}.pdf"
            ruta = os.path.join("pdfs", nombre_pdf)

            generar_diploma_pdf(row, opciones, ruta, self.logo_path)

        messagebox.showinfo("OK", "Diplomas generados correctamente")
        self.root.quit()


# =========================
# MAIN
# =========================

if __name__ == "__main__":
    root = tk.Tk()
    app = GeneradorDiplomasApp(root)
    root.mainloop()
