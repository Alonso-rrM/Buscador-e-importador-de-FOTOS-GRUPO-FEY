import os
import sys
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from unificador import generate_pdf, export_images, index_images_from_db
# pyrefly: ignore [missing-import]
from PIL import Image, ImageTk

# Intenta habilitar la nitidez de alta resolución (DPI) en Windows
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Unificador de Catálogos a PDF")
        self.geometry("700x700")
        self.minsize(650, 650)
        
        # Paleta de colores (Modo Oscuro Elegante)
        self.bg_color = "#121212"       # Fondo general
        self.card_color = "#1e1e1e"     # Tarjetas/paneles
        self.primary_color = "#bb86fc"  # Acento (Púrpura Material)
        self.primary_active = "#9965f4" # Acento activo
        self.border_color = "#333333"   # Bordes
        self.text_color = "#e0e0e0"     # Texto principal
        self.text_muted = "#a0a0a0"     # Texto secundario/apagado
        
        self.configure(bg=self.bg_color)
        
        # Configuración de estilos TTK
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        
        # Configurar colores de fondo generales de ttk
        self.style.configure(".", background=self.bg_color, foreground=self.text_color)
        self.style.configure("TLabel", background=self.bg_color, foreground=self.text_color, font=("Segoe UI", 10))
        self.style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"), foreground=self.primary_color)
        self.style.configure("Sub.TLabel", font=("Segoe UI", 9), foreground=self.text_muted)
        
        # Frame de tarjeta (panel oscuro)
        self.style.configure("Card.TFrame", background=self.card_color, relief="flat")
        self.style.configure("CardLabel.TLabel", background=self.card_color, foreground=self.text_color, font=("Segoe UI", 10, "bold"))
        self.style.configure("CardSubLabel.TLabel", background=self.card_color, foreground=self.text_muted, font=("Segoe UI", 9))
        
        # Entradas y comboboxes (ajustes para modo oscuro)
        self.style.configure("TEntry", fieldbackground="#2c2c2c", foreground="#ffffff", bordercolor=self.border_color, lightcolor=self.border_color, darkcolor=self.border_color)
        self.style.configure("TCombobox", fieldbackground="#2c2c2c", foreground="#ffffff", bordercolor=self.border_color, lightcolor=self.border_color, darkcolor=self.border_color)
        
        # Botones secundarios
        self.style.configure("TButton", font=("Segoe UI", 10), bordercolor=self.border_color, lightcolor=self.border_color, darkcolor=self.border_color, background="#2c2c2c", foreground=self.text_color)
        self.style.map("TButton",
            background=[('active', '#383838'), ('disabled', '#1a1a1a')],
            foreground=[('disabled', '#555555')]
        )
        
        # Botones primarios
        self.style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), background=self.primary_color, foreground="#121212", bordercolor=self.primary_color)
        self.style.map("Primary.TButton",
            background=[('active', self.primary_active), ('disabled', '#1a1a1a')],
            foreground=[('disabled', '#555555')]
        )

        # Estilo del Notebook (pestañas)
        self.style.configure("TNotebook", background=self.bg_color, borderwidth=0, tabmargins=[0, 5, 0, 0])
        self.style.configure("TNotebook.Tab", background="#2c2c2c", foreground=self.text_color, font=("Segoe UI", 10, "bold"), padding=[15, 8])
        self.style.map("TNotebook.Tab",
            background=[('selected', self.card_color), ('active', '#383838')],
            foreground=[('selected', self.primary_color), ('active', self.text_color)]
        )

        # Variables de estado (Pestaña 1 — Catálogo Excel)
        self.excel_path_var = tk.StringVar()
        self.db_path_var = tk.StringVar(value=r"\\feyserver\GRUPOS\Publico\VENTAS\CONTENEDORES 2025\BASE_DE_DATOS_FOTOS_MASTER.db")
        self.pdf_output_var = tk.StringVar()
        self.missing_codes_list = []

        # Variables de estado (Pestaña 2 — Búsqueda Rápida)
        self.quick_search_var = tk.StringVar()
        self._quick_found_path = None
        self._cached_images_dict = None
        self._cached_db_path = None
        self._preview_image = None  # Referencia para evitar recolección de basura
        
        self.create_widgets()
        
    def create_widgets(self):
        # Contenedor principal con márgenes
        main_container = tk.Frame(self, bg=self.bg_color)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 1. Cabecera (siempre visible, fuera del Notebook)
        header_frame = tk.Frame(main_container, bg=self.bg_color)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        lbl_title = ttk.Label(header_frame, text="Unificador de Catálogos a PDF", style="Header.TLabel")
        lbl_title.pack(anchor="w")
        
        lbl_subtitle = ttk.Label(header_frame, text="Busca y exporta imágenes de productos desde la Base de Datos Maestra.", style="Sub.TLabel")
        lbl_subtitle.pack(anchor="w", pady=(2, 0))
        
        # 2. Notebook con pestañas
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Pestaña 1: Catálogo desde Excel
        tab1 = tk.Frame(self.notebook, bg=self.bg_color)
        self.notebook.add(tab1, text="  📋 Catálogo (Excel)  ")
        self._create_catalog_tab(tab1)
        
        # Pestaña 2: Búsqueda Rápida de un solo código
        tab2 = tk.Frame(self.notebook, bg=self.bg_color)
        self.notebook.add(tab2, text="  🔍 Búsqueda Rápida  ")
        self._create_quick_search_tab(tab2)

    # =========================================================================
    #  PESTAÑA 1 — Catálogo desde Excel (contenido original, sin cambios)
    # =========================================================================

    def _create_catalog_tab(self, parent):
        """Construye todos los widgets de la pestaña de catálogo Excel.
        Este es el contenido original de create_widgets() movido aquí."""

        # 2. Panel de Configuración (Card Blanca)
        self.card_frame = ttk.Frame(parent, style="Card.TFrame", padding=15)
        self.card_frame.pack(fill=tk.X, pady=(10, 15))
        
        # Fila 1: Archivo Excel
        lbl_excel = ttk.Label(self.card_frame, text="Archivo Excel de Códigos:", style="CardLabel.TLabel")
        lbl_excel.grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        self.ent_excel = ttk.Entry(self.card_frame, textvariable=self.excel_path_var, width=55)
        self.ent_excel.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0, 10), pady=(0, 15))
        self.ent_excel.bind("<FocusOut>", lambda e: self.on_excel_selected())
        self.ent_excel.bind("<Return>", lambda e: self.on_excel_selected())
        
        self.btn_excel = ttk.Button(self.card_frame, text="Seleccionar Excel", command=self.browse_excel)
        self.btn_excel.grid(row=1, column=2, sticky="ew", pady=(0, 15))
        
        # Fila 2: Hoja y Columna (lado a lado)
        self.combo_frame = tk.Frame(self.card_frame, bg=self.card_color)
        self.combo_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 15))
        self.combo_frame.columnconfigure(0, weight=1)
        self.combo_frame.columnconfigure(1, weight=1)
        
        # Sub-panel izquierdo: Hoja
        sheet_panel = tk.Frame(self.combo_frame, bg=self.card_color)
        sheet_panel.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        lbl_sheet = ttk.Label(sheet_panel, text="Hoja del Excel:", style="CardLabel.TLabel")
        lbl_sheet.pack(anchor="w", pady=(0, 5))
        self.sheet_combo = ttk.Combobox(sheet_panel, state="readonly")
        self.sheet_combo.pack(fill=tk.X)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_selected)
        
        # Sub-panel derecho: Columna
        col_panel = tk.Frame(self.combo_frame, bg=self.card_color)
        col_panel.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))
        lbl_col = ttk.Label(col_panel, text="Columna de Códigos:", style="CardLabel.TLabel")
        lbl_col.pack(anchor="w", pady=(0, 5))
        self.column_combo = ttk.Combobox(col_panel, state="readonly")
        self.column_combo.pack(fill=tk.X)
        
        # Fila 3: Base de Datos de imágenes
        lbl_images = ttk.Label(self.card_frame, text="Archivo de Base de Datos Maestro (.db / .csv):", style="CardLabel.TLabel")
        lbl_images.grid(row=3, column=0, sticky="w", pady=(0, 5))
        
        self.ent_images = ttk.Entry(self.card_frame, textvariable=self.db_path_var, width=55)
        self.ent_images.grid(row=4, column=0, columnspan=2, sticky="ew", padx=(0, 10), pady=(0, 15))
        
        self.btn_images = ttk.Button(self.card_frame, text="Seleccionar Base de Datos", command=self.browse_database_file)
        self.btn_images.grid(row=4, column=2, sticky="ew", pady=(0, 15))
        
        # Fila 4: Guardar PDF
        lbl_pdf = ttk.Label(self.card_frame, text="Guardar PDF como (opcional si solo exporta imágenes):", style="CardLabel.TLabel")
        lbl_pdf.grid(row=5, column=0, sticky="w", pady=(0, 5))
        
        self.ent_pdf = ttk.Entry(self.card_frame, textvariable=self.pdf_output_var, width=55)
        self.ent_pdf.grid(row=6, column=0, columnspan=2, sticky="ew", padx=(0, 10), pady=(0, 5))
        
        self.btn_pdf = ttk.Button(self.card_frame, text="Definir Destino", command=self.browse_pdf_output)
        self.btn_pdf.grid(row=6, column=2, sticky="ew", pady=(0, 5))
        
        self.card_frame.columnconfigure(0, weight=1)
        self.card_frame.columnconfigure(1, weight=1)
        
        # 3. Panel de Progreso y Acción
        action_frame = tk.Frame(parent, bg=self.bg_color)
        action_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.progress = ttk.Progressbar(action_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill=tk.X, pady=(0, 8))
        
        self.lbl_status = ttk.Label(action_frame, text="Listo para comenzar.", style="Sub.TLabel")
        self.lbl_status.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Contenedor de botones de acción
        buttons_frame = tk.Frame(action_frame, bg=self.bg_color)
        buttons_frame.pack(side=tk.RIGHT)
        
        self.btn_clear = ttk.Button(buttons_frame, text="🧹 Limpiar", command=self.clear_all)
        self.btn_clear.pack(side=tk.LEFT, ipadx=10, ipady=3, padx=(0, 8))
        
        self.btn_export = ttk.Button(buttons_frame, text="📁 Exportar Imágenes", command=self.start_export)
        self.btn_export.pack(side=tk.LEFT, ipadx=10, ipady=3, padx=(0, 8))
        
        self.btn_generate = ttk.Button(buttons_frame, text="📄 Generar PDF", style="Primary.TButton", command=self.start_generation)
        self.btn_generate.pack(side=tk.LEFT, ipadx=15, ipady=3)
        
        # 4. Panel de Resultados (Log de Proceso)
        results_card = ttk.Frame(parent, style="Card.TFrame", padding=15)
        results_card.pack(fill=tk.BOTH, expand=True)
        
        lbl_results = ttk.Label(results_card, text="Reporte de Ejecución y Códigos Faltantes:", style="CardLabel.TLabel")
        lbl_results.pack(anchor="w", pady=(0, 5))
        
        # Log Text Box con Scrollbar
        log_container = tk.Frame(results_card, bg=self.card_color)
        log_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        scrollbar = tk.Scrollbar(log_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(log_container, wrap=tk.WORD, yscrollcommand=scrollbar.set, font=("Consolas", 9), bg="#1e1e1e", fg="#e0e0e0", relief="flat", highlightbackground=self.border_color, highlightthickness=1)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)
        
        # Botón para copiar códigos faltantes
        self.btn_copy_missing = ttk.Button(results_card, text="Copiar Códigos Faltantes al Portapapeles", command=self.copy_missing_to_clipboard, state=tk.DISABLED)
        self.btn_copy_missing.pack(anchor="e")

    # =========================================================================
    #  PESTAÑA 2 — Búsqueda Rápida de un solo código
    # =========================================================================

    def _create_quick_search_tab(self, parent):
        """Construye los widgets de la pestaña de búsqueda rápida."""

        container = tk.Frame(parent, bg=self.bg_color)
        container.pack(fill=tk.BOTH, expand=True, padx=15, pady=(10, 15))

        # --- Panel de búsqueda ---
        search_card = ttk.Frame(container, style="Card.TFrame", padding=15)
        search_card.pack(fill=tk.X, pady=(0, 15))

        # Fila 1: Campo de código
        lbl_code = ttk.Label(search_card, text="Código del producto:", style="CardLabel.TLabel")
        lbl_code.grid(row=0, column=0, sticky="w", pady=(0, 5))

        self.ent_quick_code = ttk.Entry(search_card, textvariable=self.quick_search_var, width=55, font=("Segoe UI", 12))
        self.ent_quick_code.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0, 10), pady=(0, 15))
        self.ent_quick_code.bind("<Return>", lambda e: self.quick_search())

        self.btn_quick_search = ttk.Button(search_card, text="🔍 Buscar", style="Primary.TButton", command=self.quick_search)
        self.btn_quick_search.grid(row=1, column=2, sticky="ew", pady=(0, 15), ipady=4)

        # Fila 2: Base de Datos (comparte variable con Pestaña 1)
        lbl_db = ttk.Label(search_card, text="Base de Datos Maestro (.db / .csv):", style="CardLabel.TLabel")
        lbl_db.grid(row=2, column=0, sticky="w", pady=(0, 5))

        self.ent_quick_db = ttk.Entry(search_card, textvariable=self.db_path_var, width=55)
        self.ent_quick_db.grid(row=3, column=0, columnspan=2, sticky="ew", padx=(0, 10), pady=(0, 5))

        self.btn_quick_db = ttk.Button(search_card, text="Seleccionar Base de Datos", command=self.browse_database_file)
        self.btn_quick_db.grid(row=3, column=2, sticky="ew", pady=(0, 5))

        search_card.columnconfigure(0, weight=1)
        search_card.columnconfigure(1, weight=1)

        # --- Panel de resultado ---
        result_card = ttk.Frame(container, style="Card.TFrame", padding=15)
        result_card.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        lbl_result_title = ttk.Label(result_card, text="Resultado:", style="CardLabel.TLabel")
        lbl_result_title.pack(anchor="w", pady=(0, 10))

        # Contenido: previsualización (izquierda) + info (derecha)
        result_content = tk.Frame(result_card, bg=self.card_color)
        result_content.pack(fill=tk.BOTH, expand=True)

        # Panel de previsualización (cuadro oscuro fijo)
        self.preview_frame = tk.Frame(result_content, bg="#2c2c2c", width=200, height=200,
                                      highlightbackground=self.border_color, highlightthickness=1)
        self.preview_frame.pack(side=tk.LEFT, padx=(0, 15), pady=5)
        self.preview_frame.pack_propagate(False)

        self.preview_label = tk.Label(self.preview_frame, bg="#2c2c2c", text="Sin imagen",
                                      fg=self.text_muted, font=("Segoe UI", 9))
        self.preview_label.pack(expand=True)

        # Panel de información (derecha)
        info_frame = tk.Frame(result_content, bg=self.card_color)
        info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=5)

        self.lbl_quick_status = ttk.Label(info_frame, text="Ingresa un código y presiona Buscar.",
                                          style="CardSubLabel.TLabel")
        self.lbl_quick_status.pack(anchor="w", pady=(0, 8))

        self.lbl_quick_path = ttk.Label(info_frame, text="", style="CardSubLabel.TLabel", wraplength=350)
        self.lbl_quick_path.pack(anchor="w", pady=(0, 5))

        # --- Botones de acción ---
        buttons_frame = tk.Frame(container, bg=self.bg_color)
        buttons_frame.pack(fill=tk.X)

        self.btn_quick_clear = ttk.Button(buttons_frame, text="🧹 Limpiar", command=self.quick_clear)
        self.btn_quick_clear.pack(side=tk.LEFT, ipadx=10, ipady=3, padx=(0, 8))

        self.btn_quick_export = ttk.Button(buttons_frame, text="📁 Exportar Copia",
                                           command=self.quick_export_image, state=tk.DISABLED)
        self.btn_quick_export.pack(side=tk.RIGHT, ipadx=10, ipady=3, padx=(8, 0))

        self.btn_quick_open = ttk.Button(buttons_frame, text="📂 Abrir Imagen",
                                         command=self.quick_open_image, state=tk.DISABLED)
        self.btn_quick_open.pack(side=tk.RIGHT, ipadx=10, ipady=3)

    # =========================================================================
    #  ACCIONES DE EXAMINAR (compartidas entre pestañas)
    # =========================================================================
    
    def browse_excel(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if path:
            self.excel_path_var.set(path)
            self.on_excel_selected()
            
    def browse_database_file(self):
        path = filedialog.askopenfilename(
            title="Seleccionar Base de Datos de Imágenes",
            filetypes=[("Archivos Base de Datos", "*.db *.csv"), ("Todos los archivos", "*.*")]
        )
        if path:
            self.db_path_var.set(path)
            # Intentar autocompletar la ruta de salida del PDF en la misma carpeta del excel si no está definida
            if not self.pdf_output_var.get() and self.excel_path_var.get():
                excel_dir = os.path.dirname(self.excel_path_var.get())
                self.pdf_output_var.set(os.path.join(excel_dir, "catalogo_unificado.pdf"))
            
    def browse_pdf_output(self):
        path = filedialog.asksaveasfilename(
            title="Guardar PDF como",
            defaultextension=".pdf",
            filetypes=[("Documento PDF", "*.pdf")]
        )
        if path:
            self.pdf_output_var.set(path)

    # =========================================================================
    #  CARGAR DATOS DEL EXCEL (Pestaña 1)
    # =========================================================================
    
    def on_excel_selected(self):
        path = self.excel_path_var.get()
        if not path or not os.path.exists(path):
            return
            
        def load_sheets():
            self.set_status("Cargando hojas de Excel...")
            try:
                xl = pd.ExcelFile(path)
                sheets = xl.sheet_names
                self.after(0, lambda: self.update_sheets(sheets))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error de lectura", f"No se pudo abrir el Excel:\n{str(e)}"))
                self.after(0, lambda: self.set_status("Error al abrir Excel"))
                
        threading.Thread(target=load_sheets, daemon=True).start()
        
    def update_sheets(self, sheets):
        self.sheet_combo['values'] = sheets
        if sheets:
            self.sheet_combo.set(sheets[0])
            self.on_sheet_selected()
        self.set_status("Excel cargado correctamente.")
        
    def on_sheet_selected(self, event=None):
        path = self.excel_path_var.get()
        sheet = self.sheet_combo.get()
        if not path or not sheet:
            return
            
        def load_columns():
            self.set_status("Cargando columnas...")
            try:
                # Leer solo los encabezados para no tardar si el archivo es gigante
                df_headers = pd.read_excel(path, sheet_name=sheet, nrows=0)
                cols = list(df_headers.columns)
                self.after(0, lambda: self.update_columns(cols))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", f"No se pudieron leer las columnas de la hoja '{sheet}':\n{str(e)}"))
                self.after(0, lambda: self.set_status("Error al cargar columnas"))
                
        threading.Thread(target=load_columns, daemon=True).start()
        
    def update_columns(self, cols):
        self.column_combo['values'] = cols
        
        # Intentar pre-seleccionar una columna que suene a código
        selected_col = ""
        for col in cols:
            col_lower = str(col).lower()
            if any(kw in col_lower for kw in ("codigo", "código", "id", "sku", "articulo", "artículo", "nombre", "code")):
                selected_col = col
                break
        if not selected_col and cols:
            selected_col = cols[0]
            
        if selected_col:
            self.column_combo.set(selected_col)
        self.set_status("Listo para comenzar.")

    # =========================================================================
    #  LIMPIAR INTERFAZ (Pestaña 1)
    # =========================================================================
    
    def clear_all(self):
        """Limpia todos los campos de la interfaz y reinicia el estado."""
        self.excel_path_var.set("")
        self.sheet_combo.set("")
        self.sheet_combo['values'] = []
        self.column_combo.set("")
        self.column_combo['values'] = []
        # Restablece la base de datos a su valor por defecto
        self.db_path_var.set(r"\\feyserver\GRUPOS\Publico\VENTAS\CONTENEDORES 2025\BASE_DE_DATOS_FOTOS_MASTER.db")
        self.pdf_output_var.set("")
        self.missing_codes_list = []
        self.btn_copy_missing.config(state=tk.DISABLED)
        self.progress['value'] = 0
        self.set_status("Listo para comenzar.")
        self.log_text.delete("1.0", tk.END)

    # =========================================================================
    #  GENERACIÓN DEL PDF (Pestaña 1)
    # =========================================================================
    
    def set_status(self, text):
        self.lbl_status.config(text=text)
        
    def set_controls_state(self, state):
        self.btn_excel.config(state=state)
        self.sheet_combo.config(state="readonly" if state == tk.NORMAL else state)
        self.column_combo.config(state="readonly" if state == tk.NORMAL else state)
        self.btn_images.config(state=state)
        self.btn_pdf.config(state=state)
        self.btn_generate.config(state=state)
        self.btn_export.config(state=state)
        self.btn_clear.config(state=state)
        self.ent_excel.config(state=state)
        self.ent_images.config(state=state)
        self.ent_pdf.config(state=state)
        if state == tk.NORMAL:
            self.btn_copy_missing.config(state=tk.NORMAL if self.missing_codes_list else tk.DISABLED)
        else:
            self.btn_copy_missing.config(state=state)
            
    def update_progress(self, percent, msg):
        self.progress['value'] = percent
        self.set_status(msg)
        self.log_text.insert(tk.END, f"{msg}\n")
        self.log_text.see(tk.END)
        
    def start_generation(self):
        excel = self.excel_path_var.get()
        sheet = self.sheet_combo.get()
        col = self.column_combo.get()
        images = self.db_path_var.get()
        pdf = self.pdf_output_var.get()
        
        if not excel or not os.path.exists(excel):
            messagebox.showerror("Campos incompletos", "Por favor, selecciona un archivo Excel válido.")
            return
        if not sheet:
            messagebox.showerror("Campos incompletos", "Por favor, selecciona una hoja de Excel.")
            return
        if not col:
            messagebox.showerror("Campos incompletos", "Por favor, selecciona la columna que contiene los códigos.")
            return
        if not images or not os.path.isfile(images):
            messagebox.showerror("Campos incompletos", "Por favor, selecciona el archivo de la Base de Datos (.db o .csv).")
            return
        if not pdf:
            messagebox.showerror("Campos incompletos", "Por favor, define la ruta de salida del PDF.")
            return
            
        # Iniciar generación en segundo plano
        self.set_controls_state(tk.DISABLED)
        self.progress['value'] = 0
        self.log_text.delete("1.0", tk.END)
        self.missing_codes_list = []
        self.btn_copy_missing.config(state=tk.DISABLED)
        
        def run():
            try:
                def progress_cb(current, total, msg):
                    percent = (current / total) * 100 if total > 0 else 0
                    self.after(0, lambda: self.update_progress(percent, msg))
                
                matched_count, missing = generate_pdf(
                    excel_path=excel,
                    sheet_name=sheet,
                    column_name=col,
                    db_path=images,
                    output_pdf_path=pdf,
                    progress_callback=progress_cb
                )
                self.after(0, lambda: self.on_success(matched_count, missing))
            except Exception as e:
                error_msg = str(e)
                self.after(0, lambda: self.on_failure(error_msg))
                
        threading.Thread(target=run, daemon=True).start()
        
    def on_success(self, matched_count, missing):
        self.set_controls_state(tk.NORMAL)
        self.progress['value'] = 100
        self.set_status("¡Proceso completado con éxito!")
        
        # Escribir reporte final
        self.log_text.insert(tk.END, "\n=========================================\n")
        self.log_text.insert(tk.END, "              REPORTE FINAL              \n")
        self.log_text.insert(tk.END, "=========================================\n")
        self.log_text.insert(tk.END, f"Imágenes añadidas al PDF: {matched_count}\n")
        self.log_text.insert(tk.END, f"Códigos no encontrados: {len(missing)}\n\n")
        
        if missing:
            self.missing_codes_list = missing
            self.btn_copy_missing.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, "CÓDIGOS FALTANTES (Sin imagen asociada):\n")
            for c in missing:
                self.log_text.insert(tk.END, f" - {c}\n")
        else:
            self.log_text.insert(tk.END, "¡Felicidades! Se encontraron imágenes para todos los códigos.\n")
            
        self.log_text.see(tk.END)
        messagebox.showinfo("Éxito", f"Se ha generado el PDF con éxito.\nTotal de imágenes unificadas: {matched_count}\nCódigos faltantes: {len(missing)}")
        
    def on_failure(self, error_msg):
        self.set_controls_state(tk.NORMAL)
        self.progress['value'] = 0
        self.set_status("Error en el proceso.")
        self.log_text.insert(tk.END, f"\n[ERROR] El proceso falló: {error_msg}\n")
        self.log_text.see(tk.END)
        messagebox.showerror("Error de Generación", f"No se pudo crear el PDF:\n{error_msg}")
        
    def copy_missing_to_clipboard(self):
        if not self.missing_codes_list:
            return
        text = "\n".join(self.missing_codes_list)
        self.clipboard_clear()
        self.clipboard_append(text)
        messagebox.showinfo("Portapapeles", "La lista de códigos faltantes se copió al portapapeles.")

    # =========================================================================
    #  EXPORTACIÓN DE IMÁGENES INDIVIDUALES (Pestaña 1)
    # =========================================================================
    
    def start_export(self):
        """Valida los campos comunes y pide una carpeta destino para exportar las imágenes."""
        excel = self.excel_path_var.get()
        sheet = self.sheet_combo.get()
        col = self.column_combo.get()
        db = self.db_path_var.get()
        
        if not excel or not os.path.exists(excel):
            messagebox.showerror("Campos incompletos", "Por favor, selecciona un archivo Excel válido.")
            return
        if not sheet:
            messagebox.showerror("Campos incompletos", "Por favor, selecciona una hoja de Excel.")
            return
        if not col:
            messagebox.showerror("Campos incompletos", "Por favor, selecciona la columna que contiene los códigos.")
            return
        if not db or not os.path.isfile(db):
            messagebox.showerror("Campos incompletos", "Por favor, selecciona el archivo de la Base de Datos (.db o .csv).")
            return
        
        # Pedir carpeta de destino
        output_folder = filedialog.askdirectory(
            title="Seleccionar carpeta para exportar las imágenes"
        )
        if not output_folder:
            return  # El usuario canceló
        
        # Iniciar exportación en segundo plano
        self.set_controls_state(tk.DISABLED)
        self.progress['value'] = 0
        self.log_text.delete("1.0", tk.END)
        self.missing_codes_list = []
        self.btn_copy_missing.config(state=tk.DISABLED)
        
        def run():
            try:
                def progress_cb(current, total, msg):
                    percent = (current / total) * 100 if total > 0 else 0
                    self.after(0, lambda: self.update_progress(percent, msg))
                
                exported_count, missing = export_images(
                    excel_path=excel,
                    sheet_name=sheet,
                    column_name=col,
                    db_path=db,
                    output_folder=output_folder,
                    progress_callback=progress_cb
                )
                self.after(0, lambda: self.on_export_success(exported_count, missing, output_folder))
            except Exception as e:
                error_msg = str(e)
                self.after(0, lambda: self.on_export_failure(error_msg))
                
        threading.Thread(target=run, daemon=True).start()
    
    def on_export_success(self, exported_count, missing, output_folder):
        self.set_controls_state(tk.NORMAL)
        self.progress['value'] = 100
        self.set_status("¡Exportación completada con éxito!")
        
        self.log_text.insert(tk.END, "\n=========================================\n")
        self.log_text.insert(tk.END, "         REPORTE DE EXPORTACIÓN          \n")
        self.log_text.insert(tk.END, "=========================================\n")
        self.log_text.insert(tk.END, f"Imágenes exportadas: {exported_count}\n")
        self.log_text.insert(tk.END, f"Carpeta destino: {output_folder}\n")
        self.log_text.insert(tk.END, f"Códigos no encontrados: {len(missing)}\n\n")
        
        if missing:
            self.missing_codes_list = missing
            self.btn_copy_missing.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, "CÓDIGOS FALTANTES (Sin imagen asociada):\n")
            for c in missing:
                self.log_text.insert(tk.END, f" - {c}\n")
        else:
            self.log_text.insert(tk.END, "¡Felicidades! Se encontraron imágenes para todos los códigos.\n")
            
        self.log_text.see(tk.END)
        messagebox.showinfo("Éxito", f"Se exportaron {exported_count} imágenes individuales a:\n{output_folder}\nCódigos faltantes: {len(missing)}")
    
    def on_export_failure(self, error_msg):
        self.set_controls_state(tk.NORMAL)
        self.progress['value'] = 0
        self.set_status("Error en la exportación.")
        self.log_text.insert(tk.END, f"\n[ERROR] La exportación falló: {error_msg}\n")
        self.log_text.see(tk.END)
        messagebox.showerror("Error de Exportación", f"No se pudieron exportar las imágenes:\n{error_msg}")

    # =========================================================================
    #  BÚSQUEDA RÁPIDA (Pestaña 2)
    # =========================================================================

    def _get_images_dict(self):
        """Retorna el diccionario de imágenes cacheado. Solo recarga si la ruta de la DB cambió."""
        db_path = self.db_path_var.get()
        if self._cached_images_dict is None or self._cached_db_path != db_path:
            self._cached_images_dict = index_images_from_db(db_path)
            self._cached_db_path = db_path
        return self._cached_images_dict

    def quick_search(self):
        """Busca un solo código en la Base de Datos Maestra y muestra el resultado."""
        code = self.quick_search_var.get().strip()
        if not code:
            messagebox.showwarning("Campo vacío", "Por favor, escribe un código para buscar.")
            return

        db_path = self.db_path_var.get()
        if not db_path or not os.path.isfile(db_path):
            messagebox.showerror("Base de Datos", "La ruta de la Base de Datos no es válida. Selecciona un archivo .db o .csv.")
            return

        # Feedback visual inmediato
        self.lbl_quick_status.config(text="Buscando...", foreground=self.text_muted)
        self.lbl_quick_path.config(text="")
        self._quick_found_path = None
        self.btn_quick_open.config(state=tk.DISABLED)
        self.btn_quick_export.config(state=tk.DISABLED)
        self._clear_preview()
        self.btn_quick_search.config(state=tk.DISABLED)

        def search():
            try:
                images_dict = self._get_images_dict()
                code_lower = code.lower().strip()

                if code_lower in images_dict:
                    found_path = images_dict[code_lower]
                    self.after(0, lambda: self._show_quick_result(code, found_path))
                else:
                    self.after(0, lambda: self._show_quick_not_found(code))
            except Exception as e:
                error_msg = str(e)
                self.after(0, lambda: self._show_quick_error(error_msg))
            finally:
                self.after(0, lambda: self.btn_quick_search.config(state=tk.NORMAL))

        threading.Thread(target=search, daemon=True).start()

    def _show_quick_result(self, code, path):
        """Muestra el resultado exitoso de una búsqueda rápida."""
        self._quick_found_path = path
        self.lbl_quick_status.config(text=f"✅ Código '{code}' encontrado:", foreground="#4caf50")
        self.lbl_quick_path.config(text=path)
        self.btn_quick_open.config(state=tk.NORMAL)
        self.btn_quick_export.config(state=tk.NORMAL)
        self._show_preview(path)

    def _show_quick_not_found(self, code):
        """Muestra el resultado cuando un código no se encuentra."""
        self._quick_found_path = None
        self.lbl_quick_status.config(text=f"❌ Código '{code}' no encontrado en la Base de Datos.", foreground="#f44336")
        self.lbl_quick_path.config(text="")
        self.btn_quick_open.config(state=tk.DISABLED)
        self.btn_quick_export.config(state=tk.DISABLED)
        self._clear_preview()

    def _show_quick_error(self, error_msg):
        """Muestra un error ocurrido durante la búsqueda."""
        self.lbl_quick_status.config(text=f"⚠ Error: {error_msg}", foreground="#ff9800")
        self.lbl_quick_path.config(text="")
        # Invalidar caché en caso de error (la DB puede haber cambiado)
        self._cached_images_dict = None
        self._cached_db_path = None

    def _show_preview(self, path):
        """Carga una miniatura de la imagen en el panel de previsualización."""
        try:
            img = Image.open(path)
            img.thumbnail((190, 190), Image.LANCZOS)
            self._preview_image = ImageTk.PhotoImage(img)
            self.preview_label.config(image=self._preview_image, text="")
            img.close()
        except Exception:
            self._clear_preview()
            self.preview_label.config(text="No se pudo\ncargar vista previa")

    def _clear_preview(self):
        """Limpia la previsualización de imagen."""
        self._preview_image = None
        self.preview_label.config(image="", text="Sin imagen")

    def quick_open_image(self):
        """Abre la imagen encontrada con el programa predeterminado del sistema."""
        if self._quick_found_path and os.path.isfile(self._quick_found_path):
            os.startfile(self._quick_found_path)
        else:
            messagebox.showerror("Error", "La imagen no existe o no se puede acceder. Realiza una nueva búsqueda.")

    def quick_export_image(self):
        """Copia la imagen encontrada a una carpeta elegida por el usuario."""
        if not self._quick_found_path or not os.path.isfile(self._quick_found_path):
            messagebox.showerror("Error", "No hay imagen para exportar. Realiza una búsqueda primero.")
            return

        dest_folder = filedialog.askdirectory(title="Seleccionar carpeta de destino")
        if not dest_folder:
            return

        try:
            code = self.quick_search_var.get().strip()
            ext = os.path.splitext(self._quick_found_path)[1]
            dest_path = os.path.join(dest_folder, f"{code}{ext}")
            shutil.copy2(self._quick_found_path, dest_path)
            messagebox.showinfo("Éxito", f"Imagen exportada correctamente a:\n{dest_path}")
        except Exception as e:
            messagebox.showerror("Error de Exportación", f"No se pudo exportar la imagen:\n{str(e)}")

    def quick_clear(self):
        """Limpia todos los campos de la pestaña de búsqueda rápida e invalida la caché."""
        self.quick_search_var.set("")
        self._quick_found_path = None
        self.btn_quick_open.config(state=tk.DISABLED)
        self.btn_quick_export.config(state=tk.DISABLED)
        self.lbl_quick_status.config(text="Ingresa un código y presiona Buscar.", foreground=self.text_muted)
        self.lbl_quick_path.config(text="")
        self._clear_preview()
        # Invalidar caché para forzar recarga si la DB cambió
        self._cached_images_dict = None
        self._cached_db_path = None

if __name__ == "__main__":
    app = App()
    app.mainloop()
