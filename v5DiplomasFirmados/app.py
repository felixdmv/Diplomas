import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
from PIL import Image, ImageTk  # Necesario: pip install Pillow

# Módulos propios
from utils import resource_path
import generador
import mailer

# --- CONFIGURACIÓN DE COLORES Y ESTILO ---
COLOR_FONDO = "#F4F6F9"       # Gris muy clarito (moderno)
COLOR_BLANCO = "#FFFFFF"      # Blanco puro para contenedores
COLOR_PRIMARIO = "#005A9C"    # Azul corporativo (tipo UBU o estándar profesional)
COLOR_SECUNDARIO = "#6C757D"  # Gris para botones secundarios
COLOR_TEXTO = "#333333"       # Texto casi negro

class AppUnificada:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestor de Diplomas Oficiales")
        self.root.geometry("850x700")
        self.root.configure(bg=COLOR_FONDO)

        # 1. ICONO DE LA VENTANA (.ico)
        ruta_ico = resource_path(os.path.join("imgs", "icono_final.ico"))
        if os.path.exists(ruta_ico):
            try:
                self.root.iconbitmap(ruta_ico)
            except:
                pass # Si falla, usa el de Python por defecto

        # 2. RUTAS DE IMÁGENES
        self.logo_path = resource_path(os.path.join("imgs", "ubu_logo.png"))
        
        # 3. CONFIGURAR ESTILOS (LOOK & FEEL)
        self.configurar_estilos()

        # --- LAYOUT PRINCIPAL ---
        
        # CABECERA (Header)
        self.crear_cabecera()

        # CONTENEDOR PRINCIPAL
        main_frame = ttk.Frame(root, style="Main.TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # PESTAÑAS (Notebook)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Tab 1: Generar
        self.tab_gen = ttk.Frame(self.notebook, style="Tab.TFrame")
        self.notebook.add(self.tab_gen, text="  1. GENERAR DIPLOMAS  ")
        self.construir_tab_generador()

        # Tab 2: Enviar
        self.tab_env = ttk.Frame(self.notebook, style="Tab.TFrame")
        self.notebook.add(self.tab_env, text="  2. ENVIAR POR OUTLOOK  ")
        self.construir_tab_enviar()

        # CONSOLA DE REGISTRO
        lbl_log = ttk.Label(main_frame, text="Registro de actividad:", background=COLOR_FONDO, font=("Segoe UI", 9, "bold"))
        lbl_log.pack(anchor="w", pady=(15, 5))
        
        self.txt_log = scrolledtext.ScrolledText(main_frame, height=8, state='disabled', 
                                                 font=("Consolas", 9), bg="#FFFFFF", fg="#333333", borderwidth=1, relief="solid")
        self.txt_log.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

    def configurar_estilos(self):
        style = ttk.Style()
        style.theme_use("clam")  # 'clam' permite cambiar colores de fondo y botones mejor que 'vista'

        # Estilo general
        style.configure("TLabel", background=COLOR_BLANCO, foreground=COLOR_TEXTO, font=("Segoe UI", 10))
        style.configure("TEntry", fieldbackground=COLOR_BLANCO, font=("Segoe UI", 10))
        style.configure("Main.TFrame", background=COLOR_FONDO)
        style.configure("Tab.TFrame", background=COLOR_BLANCO)
        style.configure("White.TFrame", background=COLOR_BLANCO)
        
        # Pestañas
        style.configure("TNotebook", background=COLOR_FONDO, tabposition='n')
        style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=[15, 8], background="#E0E0E0")
        style.map("TNotebook.Tab", background=[("selected", COLOR_PRIMARIO)], foreground=[("selected", "white")])

        # Botones Primarios (Azules)
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), background=COLOR_PRIMARIO, foreground="white", borderwidth=0)
        style.map("Accent.TButton", background=[("active", "#004475")]) # Azul más oscuro al pasar ratón

        # Botones Secundarios (Grises)
        style.configure("TButton", font=("Segoe UI", 10), background="#E0E0E0", borderwidth=0)
        style.map("TButton", background=[("active", "#D0D0D0")])

        # Checkbuttons
        style.configure("TCheckbutton", background=COLOR_BLANCO, font=("Segoe UI", 10))

        # LabelFrame
        style.configure("TLabelframe", background=COLOR_BLANCO, bordercolor="#DDDDDD")
        style.configure("TLabelframe.Label", background=COLOR_BLANCO, font=("Segoe UI", 10, "bold"), foreground=COLOR_PRIMARIO)

    def crear_cabecera(self):
        # Frame superior blanco para el logo y título
        header_frame = tk.Frame(self.root, bg=COLOR_BLANCO, height=80)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        
        # Sombra simulada (línea gris abajo)
        tk.Frame(self.root, bg="#DDDDDD", height=1).pack(fill=tk.X)

        # Contenedor interno con padding
        inner = tk.Frame(header_frame, bg=COLOR_BLANCO)
        inner.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # LOGO (Redimensionado con Pillow)
        if os.path.exists(self.logo_path):
            try:
                pil_img = Image.open(self.logo_path)
                # Redimensionar manteniendo proporción (altura 50px)
                aspect = pil_img.width / pil_img.height
                new_h = 50
                new_w = int(new_h * aspect)
                pil_img = pil_img.resize((new_w, new_h), Image.LANCZOS)
                
                self.tk_logo = ImageTk.PhotoImage(pil_img)
                lbl_img = tk.Label(inner, image=self.tk_logo, bg=COLOR_BLANCO)
                lbl_img.pack(side=tk.LEFT, padx=(0, 15))
            except Exception as e:
                print(f"Error cargando logo UI: {e}")

        # TEXTOS CABECERA
        text_frame = tk.Frame(inner, bg=COLOR_BLANCO)
        text_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        tk.Label(text_frame, text="GENERADOR DE DIPLOMAS", font=("Segoe UI", 16, "bold"), 
                 bg=COLOR_BLANCO, fg=COLOR_TEXTO).pack(anchor="w")
        tk.Label(text_frame, text="Automatización de certificados y envío por correo", font=("Segoe UI", 10), 
                 bg=COLOR_BLANCO, fg="#777777").pack(anchor="w")

    # --- PESTAÑA 1: GENERADOR ---
    def construir_tab_generador(self):
        # Usamos un frame interno blanco con padding
        frame = ttk.Frame(self.tab_gen, style="White.TFrame", padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # Sección 1: Archivo
        lbl_info = ttk.Label(frame, text="Paso 1: Selecciona los datos", font=("Segoe UI", 11, "bold"), foreground=COLOR_PRIMARIO)
        lbl_info.pack(anchor="w", pady=(0, 10))

        f_excel = ttk.Frame(frame, style="White.TFrame")
        f_excel.pack(fill=tk.X, pady=5)
        
        ttk.Label(f_excel, text="Archivo Excel:").pack(side=tk.LEFT)
        self.ent_excel_gen = ttk.Entry(f_excel)
        self.ent_excel_gen.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        ttk.Button(f_excel, text="Examinar...", command=lambda: self.sel_archivo(self.ent_excel_gen)).pack(side=tk.LEFT)

        ttk.Separator(frame, orient="horizontal").pack(fill=tk.X, pady=20)

        # Sección 2: Configuración
        f_opts = ttk.LabelFrame(frame, text=" Configuración del Diploma ", padding=15)
        f_opts.pack(fill=tk.X, pady=5)
        
        # Grid para opciones
        self.var_horas = tk.BooleanVar(value=True)
        self.var_calif = tk.BooleanVar(value=False)
        self.var_extra = tk.BooleanVar(value=False)

        ttk.Checkbutton(f_opts, text="Incluir duración (horas)", variable=self.var_horas).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ttk.Checkbutton(f_opts, text="Incluir calificación / nota", variable=self.var_calif).grid(row=0, column=1, sticky="w", padx=10, pady=5)
        ttk.Checkbutton(f_opts, text="Incluir campo extra personalizado", variable=self.var_extra).grid(row=0, column=2, sticky="w", padx=10, pady=5)

        ttk.Separator(frame, orient="horizontal").pack(fill=tk.X, pady=15)

        # SECCIÓN FIRMA DIGITAL
        f_firma = ttk.LabelFrame(frame, text=" 🔐 Firma Digital (Opcional) ", padding=10)
        f_firma.pack(fill=tk.X, pady=5)

        # Checkbox activar
        self.var_firmar = tk.BooleanVar(value=False)
        chk_firmar = ttk.Checkbutton(f_firma, text="Firmar digitalmente los PDFs (Requiere certificado .pfx)", 
                                     variable=self.var_firmar, command=self.toggle_firma)
        chk_firmar.pack(anchor="w")

        # Contenedor para archivo y pass (se oculta/muestra según checkbox)
        self.f_datos_firma = ttk.Frame(f_firma)
        self.f_datos_firma.pack(fill=tk.X, pady=5)
        
        # Archivo PFX
        ttk.Label(self.f_datos_firma, text="Certificado (.pfx / .p12):").pack(side=tk.LEFT)
        self.ent_pfx = ttk.Entry(self.f_datos_firma, width=30)
        self.ent_pfx.pack(side=tk.LEFT, padx=5)
        ttk.Button(self.f_datos_firma, text="...", width=3, 
                   command=lambda: self.sel_archivo(self.ent_pfx, [("Certificados", "*.pfx *.p12")])).pack(side=tk.LEFT)

        # Contraseña
        ttk.Label(self.f_datos_firma, text="Contraseña:").pack(side=tk.LEFT, padx=(15, 0))
        self.ent_pass = ttk.Entry(self.f_datos_firma, show="*", width=15) # show="*" oculta caracteres
        self.ent_pass.pack(side=tk.LEFT, padx=5)

        # Estado inicial (desactivado)
        self.toggle_firma()

        # Sección 3: Botones de Acción
        f_btns = ttk.Frame(frame, style="White.TFrame")
        f_btns.pack(fill=tk.X, pady=10)

        # Botón Vista Previa (Gris)
        btn_prev = ttk.Button(f_btns, text="👁️  VISTA PREVIA", command=self.ejecutar_preview)
        btn_prev.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10), ipady=5)
        
        # Botón Generar (Azul Grande)
        btn_gen = ttk.Button(f_btns, text="🚀  GENERAR TODOS LOS DIPLOMAS", style="Accent.TButton", command=self.ejecutar_generacion)
        btn_gen.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0), ipady=5)

    # --- PESTAÑA 2: ENVIAR ---
    def construir_tab_enviar(self):
        frame = ttk.Frame(self.tab_env, style="White.TFrame", padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # Aviso Outlook
        f_info = tk.Frame(frame, bg="#E3F2FD", padx=10, pady=10) # Fondo azul muy claro
        f_info.pack(fill=tk.X, pady=(0, 20))
        tk.Label(f_info, text="ℹ️  NOTA: Se utilizará tu cuenta de Outlook (Clásico) actualmente abierta.", 
                 bg="#E3F2FD", fg="#0c5460", font=("Segoe UI", 9)).pack(anchor="w")

        # Configuración
        f_inputs = ttk.Frame(frame, style="White.TFrame")
        f_inputs.pack(fill=tk.X)

        # Excel
        ttk.Label(f_inputs, text="Excel de Alumnos (con emails):").pack(anchor="w", pady=(5, 2))
        f_ex = ttk.Frame(f_inputs, style="White.TFrame")
        f_ex.pack(fill=tk.X, pady=(0, 10))
        self.ent_excel_env = ttk.Entry(f_ex)
        self.ent_excel_env.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(f_ex, text="Examinar...", command=lambda: self.sel_archivo(self.ent_excel_env)).pack(side=tk.LEFT)

        # Carpeta PDFs
        ttk.Label(f_inputs, text="Carpeta donde están los PDFs Firmados:").pack(anchor="w", pady=(5, 2))
        f_dir = ttk.Frame(f_inputs, style="White.TFrame")
        f_dir.pack(fill=tk.X, pady=(0, 10))
        self.ent_dir_pdfs = ttk.Entry(f_dir)
        self.ent_dir_pdfs.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(f_dir, text="Seleccionar Carpeta", command=lambda: self.sel_carpeta(self.ent_dir_pdfs)).pack(side=tk.LEFT)

        ttk.Separator(frame, orient="horizontal").pack(fill=tk.X, pady=20)

        # Modo Prueba
        self.var_dryrun = tk.BooleanVar(value=True)
        f_check = tk.Frame(frame, bg="#FFF3CD", padx=10, pady=10) # Fondo amarillo suave
        f_check.pack(fill=tk.X, pady=10)
        
        chk_dry = tk.Checkbutton(f_check, text="MODO PRUEBA (Recomendado)", variable=self.var_dryrun, 
                                 bg="#FFF3CD", font=("Segoe UI", 10, "bold"), activebackground="#FFF3CD")
        chk_dry.pack(anchor="w")
        tk.Label(f_check, text="Si está marcado, solo se crearán borradores en Outlook. No se enviará nada real.", 
                 bg="#FFF3CD", fg="#856404", font=("Segoe UI", 9)).pack(anchor="w", padx=25)

        # Botón Enviar
        btn_send = ttk.Button(frame, text="📤  ENVIAR CORREOS MASIVOS", style="Accent.TButton", command=self.ejecutar_envio)
        btn_send.pack(fill=tk.X, pady=20, ipady=8)

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

    def log(self, mensaje):
        self.txt_log.config(state='normal')
        self.txt_log.insert(tk.END, " >> " + mensaje + "\n") # Añadimos flechita visual
        self.txt_log.see(tk.END)
        self.txt_log.config(state='disabled')

    # --- LOGICA (Conectada igual que antes) ---
    def ejecutar_preview(self):
        excel = self.ent_excel_gen.get()
        if not excel:
            messagebox.showwarning("Atención", "Selecciona primero el Excel")
            return
        
        opts = (self.var_horas.get(), self.var_calif.get(), self.var_extra.get())
        self.log("--- Generando Vista Previa ---")
        generador.generar_preview(excel, opts, self.logo_path, self.log)

    def ejecutar_generacion(self):
        excel_path = self.ent_excel_gen.get()
        if not excel_path:
            messagebox.showwarning("Atención", "Selecciona primero el archivo Excel.")
            return
        
        carpeta_excel = os.path.dirname(excel_path)
        out_folder = os.path.join(carpeta_excel, "Diplomas_Generados")
        
        opts = (self.var_horas.get(), self.var_calif.get(), self.var_extra.get())
        
        # DATOS DE FIRMA
        datos_firma = None
        if self.var_firmar.get():
            pfx = self.ent_pfx.get()
            pwd = self.ent_pass.get()
            if not pfx or not os.path.exists(pfx):
                messagebox.showerror("Error Firma", "Selecciona un archivo .pfx válido.")
                return
            if not pwd:
                messagebox.showerror("Error Firma", "Introduce la contraseña del certificado.")
                return
            datos_firma = (pfx, pwd) # Tupla con ruta y pass
        
        # Pasamos datos_firma al hilo
        threading.Thread(target=self._hilo_gen, args=(excel_path, out_folder, opts, datos_firma)).start()

    def _hilo_gen(self, excel, out, opts, datos_firma=None):
        # Llamamos al generador (él creará la carpeta si no existe)
        generador.procesar_excel_y_generar(excel, out, opts, self.logo_path, self.log, datos_firma)
        
        # AVISO AL USUARIO
        # Le mostramos la ruta y, para ser más amables, abrimos la carpeta automáticamente.
        msg = f"¡Proceso finalizado!\n\nLos diplomas se han guardado en:\n{out}"
        self.root.after(0, lambda: messagebox.showinfo("Generación Completada", msg))
        
        # Truco UX: Abrir la carpeta automáticamente en el Explorador de archivos
        try:
            os.startfile(out)
        except Exception:
            pass # Si falla abrir la carpeta (raro en Windows), no pasa nada

    def ejecutar_envio(self):
        excel = self.ent_excel_env.get()
        folder = self.ent_dir_pdfs.get()
        if not excel or not folder:
            messagebox.showerror("Error", "Faltan rutas (Excel o Carpeta)")
            return
        dry = self.var_dryrun.get()
        threading.Thread(target=self._hilo_envio, args=(excel, folder, dry)).start()

    def _hilo_envio(self, excel, folder, dry):
        mailer.enviar_masivo_outlook(excel, folder, dry, self.log)
        if not dry:
            messagebox.showinfo("Fin", "Envío completado")
        else:
            self.log("Fin de la prueba. Revisa la carpeta Borradores de Outlook.")
            
    # AÑADIR ESTA FUNCIÓN A LA CLASE
    def toggle_firma(self):
        if self.var_firmar.get():
            for child in self.f_datos_firma.winfo_children():
                child.configure(state='normal')
        else:
            for child in self.f_datos_firma.winfo_children():
                child.configure(state='disabled')

    # AÑADIR MODIFICACIÓN EN sel_archivo PARA ACEPTAR FILTROS
    def sel_archivo(self, entry, tipos=[("Excel", "*.xlsx")]):
        p = filedialog.askopenfilename(filetypes=tipos)
        if p:
            entry.delete(0, tk.END)
            entry.insert(0, p)

if __name__ == "__main__":
    # --- TRUCO PARA EL ICONO EN LA BARRA DE TAREAS DE WINDOWS ---
    try:
        from ctypes import windll
        # Este string arbitrario le dice a Windows que es una app única
        myappid = 'ubu.generador.diplomas.v3.0' 
        windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except ImportError:
        pass
    # ------------------------------------------------------------

    root = tk.Tk()
    app = AppUnificada(root)
    root.mainloop()