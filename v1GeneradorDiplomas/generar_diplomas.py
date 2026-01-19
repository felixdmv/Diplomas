import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog

from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm


# =========================
# UTILIDADES
# =========================

def texto_seguro(valor):
    if pd.isna(valor):
        return ""
    return str(valor)


def abrir_pdf(ruta):
    if sys.platform == "win32":
        os.startfile(ruta)
    elif sys.platform == "darwin":
        os.system(f"open {ruta}")
    else:
        os.system(f"xdg-open {ruta}")


# =========================
# GENERACIÓN PDF
# =========================

def generar_diploma_pdf(row, opciones, salida):
    mostrar_horas, mostrar_calif, mostrar_extra = opciones

    c = canvas.Canvas(salida, pagesize=landscape(A4))
    width, height = landscape(A4)

    # Marco
    c.setLineWidth(3)
    c.rect(2*cm, 2*cm, width-4*cm, height-4*cm)

    # Título
    c.setFont("Times-Bold", 32)
    c.drawCentredString(width/2, height-4*cm, "DIPLOMA")

    y = height - 6*cm

    c.setFont("Times-Roman", 16)
    c.drawCentredString(width/2, y, "Se certifica que")

    y -= 2*cm
    c.setFont("Times-Bold", 26)
    nombre = f"{texto_seguro(row['nombre'])} {texto_seguro(row['apellidos'])}"
    c.drawCentredString(width/2, y, nombre)

    y -= 2*cm
    c.setFont("Times-Roman", 16)
    c.drawCentredString(width/2, y, "ha superado el curso")

    y -= 1.2*cm
    c.setFont("Times-Bold", 18)
    c.drawCentredString(width/2, y, texto_seguro(row["curso_nombre"]))

    y -= 2*cm
    c.setFont("Times-Roman", 14)

    horas = texto_seguro(row["horas"])
    if mostrar_horas and horas:
        c.drawCentredString(width/2, y, f"Duración: {horas} horas")
        y -= 1*cm

    nota_num = texto_seguro(row["calificacion_num"])
    nota_txt = texto_seguro(row["calificacion_texto"])

    if mostrar_calif:
        if nota_num:
            c.drawCentredString(width/2, y, f"Calificación: {nota_num}")
            y -= 1*cm
        elif nota_txt:
            c.drawCentredString(width/2, y, f"Resultado: {nota_txt}")
            y -= 1*cm

    extra = texto_seguro(row["extra_1"])
    if mostrar_extra and extra:
        c.drawCentredString(width/2, y, extra)
        y -= 1*cm

    # Fecha
    c.setFont("Times-Roman", 12)
    c.drawRightString(width - 3*cm, 3*cm, f"Fecha: {texto_seguro(row['fecha'])}")

    # Firma
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

    def vista_previa(self):
        if not self.excel_path:
            messagebox.showerror("Error", "Selecciona un Excel primero")
            return

        df = pd.read_excel(self.excel_path)
        if df.empty:
            messagebox.showerror("Error", "El Excel está vacío")
            return

        opciones = self.obtener_opciones()
        ruta = "preview_diploma.pdf"
        generar_diploma_pdf(df.iloc[0], opciones, ruta)
        abrir_pdf(ruta)

        decision = messagebox.askyesnocancel(
            "Confirmación",
            "¿Generar diplomas para todos los alumnos?\n\n"
            "Sí = Generar\n"
            "No = Volver al menú\n"
            "Cancelar = Salir"
        )

        if decision is True:
            self.generar_todos()
        elif decision is False:
            return
        else:
            self.root.quit()

    def generar_todos(self):
        if not self.excel_path:
            messagebox.showerror("Error", "Selecciona un Excel primero")
            return

        df = pd.read_excel(self.excel_path)
        opciones = self.obtener_opciones()

        os.makedirs("pdfs", exist_ok=True)

        for _, row in df.iterrows():
            nombre = f"diploma_{row['alumno_id']}.pdf"
            ruta = os.path.join("pdfs", nombre)
            generar_diploma_pdf(row, opciones, ruta)

        messagebox.showinfo("OK", "Diplomas generados correctamente")
        self.root.quit()


# =========================
# MAIN
# =========================

if __name__ == "__main__":
    root = tk.Tk()
    app = GeneradorDiplomasApp(root)
    root.mainloop()
