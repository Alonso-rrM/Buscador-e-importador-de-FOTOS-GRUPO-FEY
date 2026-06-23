import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from unificador import generate_pdf, export_images

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
        self.geometry("700x650")
        self.minsize(650, 600)
        
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

        # Variables de estado
        self.excel_path_var = tk.StringVar()
        self.db_path_var = tk.StringVar(value=r"\\feyserver\GRUPOS\Publico\VENTAS\CONTENEDORES 2025\BASE_DE_DATOS_FOTOS_MASTER.db")
        self.pdf_output_var = tk.StringVar()
        self.missing_codes_list = []
        
        self.create_widgets()
        
    def create_widgets(self):
        # Contenedor principal con márgenes
        main_container = tk.Frame(self, bg=self.bg_color)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 1. Cabecera
        header_frame = tk.Frame(main_container, bg=self.bg_color)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        lbl_title = ttk.Label(header_frame, text="Unificador de Catálogos a PDF", style="Header.TLabel")
        lbl_title.pack(anchor="w")
        
        lbl_subtitle = ttk.Label(header_frame, text="Busca imágenes recursivamente por sus códigos de Excel y genera un PDF ordenado.", style="Sub.TLabel")
        lbl_subtitle.pack(anchor="w", pady=(2, 0))
        
        # 2. Panel de Configuración (Card Blanca)
        self.card_frame = ttk.Frame(main_container, style="Card.TFrame", padding=15)
        self.card_frame.pack(fill=tk.X, pady=(0, 15))
        
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
        action_frame = tk.Frame(main_container, bg=self.bg_color)
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
        results_card = ttk.Frame(main_container, style="Card.TFrame", padding=15)
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

    # --- Acciones de Examinar ---
    
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

    # --- Cargar datos del Excel dinámicamente ---
    
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

    # --- Limpiar Interfaz ---
    
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

    # --- Generación del PDF ---
    
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

    # --- Exportación de Imágenes Individuales ---
    
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

if __name__ == "__main__":
    app = App()
    app.mainloop()
