import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading

# Módulos propios
from utils import resource_path
import generador
import mailer

class AppUnificada:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestor de Diplomas (Outlook Clásico)")
        self.root.geometry("700x600")
        
        self.logo_path = resource_path(os.path.join("imgs", "ubu_logo.png"))
        
        # --- UI LAYOUT ---
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Gestor de Diplomas", font=("Arial", 16, "bold")).pack(pady=10)

        # Pestañas
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=5)

        # Tab 1: Generar
        self.tab_gen = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_gen, text="1. Generar Diplomas")
        self.construir_tab_generador()

        # Tab 2: Enviar
        self.tab_env = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_env, text="2. Enviar por Outlook")
        self.construir_tab_enviar()

        # Consola
        ttk.Label(main_frame, text="Registro:").pack(anchor="w", pady=(10,0))
        self.txt_log = scrolledtext.ScrolledText(main_frame, height=10, state='disabled', font=("Consolas", 9))
        self.txt_log.pack(fill=tk.BOTH, expand=True, pady=5)

    def log(self, mensaje):
        self.txt_log.config(state='normal')
        self.txt_log.insert(tk.END, mensaje + "\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state='disabled')

    # --- PESTAÑA 1: GENERADOR ---
    def construir_tab_generador(self):
        frame = ttk.Frame(self.tab_gen, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        f_excel = ttk.Frame(frame)
        f_excel.pack(fill=tk.X, pady=5)
        ttk.Label(f_excel, text="Excel Alumnos:").pack(side=tk.LEFT)
        self.ent_excel_gen = ttk.Entry(f_excel)
        self.ent_excel_gen.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(f_excel, text="...", width=3, command=lambda: self.sel_archivo(self.ent_excel_gen)).pack(side=tk.LEFT)

        f_opts = ttk.LabelFrame(frame, text="Opciones", padding=10)
        f_opts.pack(fill=tk.X, pady=10)
        
        self.var_horas = tk.BooleanVar(value=True)
        self.var_calif = tk.BooleanVar(value=False)
        self.var_extra = tk.BooleanVar(value=False)

        ttk.Checkbutton(f_opts, text="Mostrar horas", variable=self.var_horas).pack(anchor="w")
        ttk.Checkbutton(f_opts, text="Mostrar calificación", variable=self.var_calif).pack(anchor="w")
        ttk.Checkbutton(f_opts, text="Mostrar extra", variable=self.var_extra).pack(anchor="w")

        # --- AQUÍ ESTÁN LOS BOTONES ---
        f_btns = ttk.Frame(frame)
        f_btns.pack(fill=tk.X, pady=20)

        # Botón Vista Previa (Izquierda)
        ttk.Button(f_btns, text="👁️ VISTA PREVIA", command=self.ejecutar_preview).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        # Botón Generar Todo (Derecha)
        ttk.Button(f_btns, text="🚀 GENERAR TODOS", command=self.ejecutar_generacion).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

    # --- PESTAÑA 2: ENVIAR ---
    def construir_tab_enviar(self):
        frame = ttk.Frame(self.tab_env, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Usa tu Outlook de escritorio (Clásico) para enviar.", foreground="blue").pack(pady=5)

        f_ex = ttk.Frame(frame)
        f_ex.pack(fill=tk.X, pady=5)
        ttk.Label(f_ex, text="Excel Alumnos:").pack(side=tk.LEFT)
        self.ent_excel_env = ttk.Entry(f_ex)
        self.ent_excel_env.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(f_ex, text="...", width=3, command=lambda: self.sel_archivo(self.ent_excel_env)).pack(side=tk.LEFT)

        f_dir = ttk.Frame(frame)
        f_dir.pack(fill=tk.X, pady=5)
        ttk.Label(f_dir, text="Carpeta PDFs Firmados:").pack(side=tk.LEFT)
        self.ent_dir_pdfs = ttk.Entry(f_dir)
        self.ent_dir_pdfs.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(f_dir, text="...", width=3, command=lambda: self.sel_carpeta(self.ent_dir_pdfs)).pack(side=tk.LEFT)

        self.var_dryrun = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="MODO PRUEBA (Solo abrir borradores)", variable=self.var_dryrun).pack(pady=10)

        ttk.Button(frame, text="ENVIAR CORREOS", command=self.ejecutar_envio).pack(pady=10, fill=tk.X)

    # --- UTILIDADES ---
    def sel_archivo(self, entry):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if p:
            entry.delete(0, tk.END)
            entry.insert(0, p)

    def sel_carpeta(self, entry):
        p = filedialog.askdirectory()
        if p:
            entry.delete(0, tk.END)
            entry.insert(0, p)

    # --- LOGICA ---
    def ejecutar_generacion(self):
        excel = self.ent_excel_gen.get()
        if not excel: return
        
        opts = (self.var_horas.get(), self.var_calif.get(), self.var_extra.get())
        out = "pdfs_generados"
        threading.Thread(target=self._hilo_gen, args=(excel, out, opts)).start()

    def _hilo_gen(self, excel, out, opts):
        generador.procesar_excel_y_generar(excel, out, opts, self.logo_path, self.log)
        messagebox.showinfo("Fin", "Generación completada")

    def ejecutar_envio(self):
        excel = self.ent_excel_env.get()
        folder = self.ent_dir_pdfs.get()
        if not excel or not folder:
            messagebox.showerror("Error", "Faltan rutas")
            return
        
        dry = self.var_dryrun.get()
        threading.Thread(target=self._hilo_envio, args=(excel, folder, dry)).start()

    def _hilo_envio(self, excel, folder, dry):
        mailer.enviar_masivo_outlook(excel, folder, dry, self.log)
        if not dry:
            messagebox.showinfo("Fin", "Envío completado")
        else:
            self.log("Fin de la prueba.")
    
    def ejecutar_preview(self):
        excel = self.ent_excel_gen.get()
        if not excel:
            messagebox.showwarning("Atención", "Selecciona primero el Excel")
            return
        
        opts = (self.var_horas.get(), self.var_calif.get(), self.var_extra.get())
        
        # Ejecutamos la preview (no hace falta hilo porque es solo 1 pdf y es rápido)
        self.log("--- Generando Vista Previa ---")
        generador.generar_preview(excel, opts, self.logo_path, self.log)

if __name__ == "__main__":
    root = tk.Tk()
    app = AppUnificada(root)
    root.mainloop()