"""
Interfaz gráfica (Tkinter) del Unificador de Catálogos.

Toda la lógica de negocio vive en `unificador.py`; aquí solo se recogen datos
del usuario, se lanzan las tareas en segundo plano y se muestran resultados.
"""
import os
import queue
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# pyrefly: ignore [missing-import]
from PIL import Image, ImageTk

from unificador import (
    export_images, generate_pdf, index_images_from_db, list_columns, list_sheets,
    normalize_code, sanitize_filename,
)

DEFAULT_DB_PATH = r"\\feyserver\GRUPOS\Publico\VENTAS\CONTENEDORES 2025\BASE_DE_DATOS_FOTOS_MASTER.db"
DEFAULT_PDF_NAME = "catalogo_unificado.pdf"
CODE_COLUMN_KEYWORDS = ("codigo", "código", "id", "sku", "articulo", "artículo", "nombre", "code")
PREVIEW_SIZE = 200
READY_TEXT = "Listo para comenzar."


class Theme:
    """Paleta Material Dark y tipografías de la aplicación."""
    BG = "#121212"
    CARD = "#1e1e1e"
    FIELD = "#2c2c2c"
    HOVER = "#383838"
    BORDER = "#333333"
    PRIMARY = "#bb86fc"
    PRIMARY_ACTIVE = "#9965f4"
    TEXT = "#e0e0e0"
    TEXT_MUTED = "#a0a0a0"
    TEXT_ON_FIELD = "#ffffff"
    DISABLED_BG = "#1a1a1a"
    DISABLED_FG = "#555555"
    SUCCESS = "#4caf50"
    ERROR = "#f44336"
    WARNING = "#ff9800"
    FONT = "Segoe UI"
    MONO = "Consolas"


def apply_dark_theme(root):
    """Configura los estilos ttk usados por la aplicación."""
    t = Theme
    style = ttk.Style(root)
    style.theme_use("clam")
    border = dict(bordercolor=t.BORDER, lightcolor=t.BORDER, darkcolor=t.BORDER)

    style.configure(".", background=t.BG, foreground=t.TEXT)
    style.configure("TLabel", background=t.BG, foreground=t.TEXT, font=(t.FONT, 10))
    style.configure("Header.TLabel", font=(t.FONT, 14, "bold"), foreground=t.PRIMARY)
    style.configure("Sub.TLabel", font=(t.FONT, 9), foreground=t.TEXT_MUTED)

    style.configure("Card.TFrame", background=t.CARD, relief="flat")
    style.configure("CardLabel.TLabel", background=t.CARD, foreground=t.TEXT, font=(t.FONT, 10, "bold"))
    style.configure("CardSubLabel.TLabel", background=t.CARD, foreground=t.TEXT_MUTED, font=(t.FONT, 9))

    for widget in ("TEntry", "TCombobox"):
        style.configure(widget, fieldbackground=t.FIELD, foreground=t.TEXT_ON_FIELD, **border)

    disabled_fg = [("disabled", t.DISABLED_FG)]
    style.configure("TButton", font=(t.FONT, 10), background=t.FIELD, foreground=t.TEXT, **border)
    style.map("TButton", background=[("active", t.HOVER), ("disabled", t.DISABLED_BG)], foreground=disabled_fg)
    style.configure("Primary.TButton", font=(t.FONT, 10, "bold"), background=t.PRIMARY,
                    foreground=t.BG, bordercolor=t.PRIMARY)
    style.map("Primary.TButton", background=[("active", t.PRIMARY_ACTIVE), ("disabled", t.DISABLED_BG)],
              foreground=disabled_fg)

    style.configure("TNotebook", background=t.BG, borderwidth=0, tabmargins=[0, 5, 0, 0])
    style.configure("TNotebook.Tab", background=t.FIELD, foreground=t.TEXT,
                    font=(t.FONT, 10, "bold"), padding=[15, 8])
    style.map("TNotebook.Tab", background=[("selected", t.CARD), ("active", t.HOVER)],
              foreground=[("selected", t.PRIMARY), ("active", t.TEXT)])


# =============================================================================
#  Utilidades compartidas
# =============================================================================

class MainThreadDispatcher:
    """
    Puente entre hilos: Tkinter no es seguro entre hilos, así que los hilos de
    trabajo solo encolan funciones (cola FIFO thread-safe) y el hilo de la
    interfaz las ejecuta en orden al revisar la cola periódicamente.
    """

    def __init__(self, root, interval_ms=50):
        self._root = root
        self._queue = queue.Queue()
        self._interval_ms = interval_ms
        self._poll()

    def call(self, func, *args):
        """Programa `func(*args)` en el hilo de Tkinter. Se puede llamar desde cualquier hilo."""
        self._queue.put((func, args))

    def _poll(self):
        try:
            while True:
                try:
                    func, args = self._queue.get_nowait()
                except queue.Empty:
                    break
                func(*args)
        finally:
            self._root.after(self._interval_ms, self._poll)


def run_in_background(dispatcher, task, on_success, on_error):
    """
    Ejecuta `task()` en un hilo para no congelar la ventana y entrega el
    resultado (o el mensaje de error) en el hilo de Tkinter.
    `task` NO debe tocar widgets.
    """
    def worker():
        try:
            result = task()
        except Exception as e:  # Frontera de la GUI: todo fallo se muestra al usuario
            dispatcher.call(on_error, str(e))
        else:
            dispatcher.call(on_success, result)

    threading.Thread(target=worker, daemon=True).start()


def guess_code_column(columns):
    """Elige la columna cuyo nombre suena a 'código'; si no hay, la primera."""
    for col in columns:
        if any(keyword in str(col).lower() for keyword in CODE_COLUMN_KEYWORDS):
            return col
    return columns[0] if columns else ""


def ask_database_file():
    return filedialog.askopenfilename(
        title="Seleccionar Base de Datos de Imágenes",
        filetypes=[("Archivos Base de Datos", "*.db *.csv"), ("Todos los archivos", "*.*")],
    )


def file_row(parent, row, label, variable, button_text, command, bottom_pad=15):
    """Fila 'etiqueta / campo de texto + botón' dentro de una tarjeta con grid."""
    ttk.Label(parent, text=label, style="CardLabel.TLabel").grid(row=row, column=0, sticky="w", pady=(0, 5))
    entry = ttk.Entry(parent, textvariable=variable, width=55)
    entry.grid(row=row + 1, column=0, columnspan=2, sticky="ew", padx=(0, 10), pady=(0, bottom_pad))
    button = ttk.Button(parent, text=button_text, command=command)
    button.grid(row=row + 1, column=2, sticky="ew", pady=(0, bottom_pad))
    return entry, button


def card(parent, **pack_options):
    frame = ttk.Frame(parent, style="Card.TFrame", padding=15)
    frame.pack(**pack_options)
    return frame


# =============================================================================
#  Pestaña 1 — Catálogo desde Excel
# =============================================================================

class CatalogTab(tk.Frame):
    """Genera un PDF o exporta imágenes a partir de una lista de códigos en Excel."""

    def __init__(self, parent, db_path_var, dispatcher):
        super().__init__(parent, bg=Theme.BG)
        self.db_path_var = db_path_var
        self.dispatcher = dispatcher
        self.excel_path_var = tk.StringVar()
        self.pdf_output_var = tk.StringVar()
        self.missing_codes = []
        self._build()

    # ---------------------------------------------------------------- Layout
    def _build(self):
        form = card(self, fill=tk.X, pady=(10, 15))

        self.ent_excel, self.btn_excel = file_row(
            form, 0, "Archivo Excel de Códigos:", self.excel_path_var, "Seleccionar Excel", self.browse_excel)
        self.ent_excel.bind("<FocusOut>", lambda e: self.on_excel_selected())
        self.ent_excel.bind("<Return>", lambda e: self.on_excel_selected())

        combos = tk.Frame(form, bg=Theme.CARD)
        combos.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 15))
        self.sheet_combo = self._combo(combos, "Hoja del Excel:", padx=(0, 10))
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda e: self.on_sheet_selected())
        self.column_combo = self._combo(combos, "Columna de Códigos:", padx=(10, 0))

        self.ent_db, self.btn_db = file_row(
            form, 3, "Archivo de Base de Datos Maestro (.db / .csv):", self.db_path_var,
            "Seleccionar Base de Datos", self.browse_database_file)
        self.ent_pdf, self.btn_pdf = file_row(
            form, 5, "Guardar PDF como (opcional si solo exporta imágenes):", self.pdf_output_var,
            "Definir Destino", self.browse_pdf_output, bottom_pad=5)
        form.columnconfigure(0, weight=1)
        form.columnconfigure(1, weight=1)

        actions = tk.Frame(self, bg=Theme.BG)
        actions.pack(fill=tk.X, pady=(0, 15))
        self.progress = ttk.Progressbar(actions, orient="horizontal", mode="determinate")
        self.progress.pack(fill=tk.X, pady=(0, 8))
        self.lbl_status = ttk.Label(actions, text=READY_TEXT, style="Sub.TLabel")
        self.lbl_status.pack(side=tk.LEFT, fill=tk.X, expand=True)

        buttons = tk.Frame(actions, bg=Theme.BG)
        buttons.pack(side=tk.RIGHT)
        self.btn_clear = ttk.Button(buttons, text="🧹 Limpiar", command=self.clear_all)
        self.btn_clear.pack(side=tk.LEFT, ipadx=10, ipady=3, padx=(0, 8))
        self.btn_export = ttk.Button(buttons, text="📁 Exportar Imágenes", command=self.start_export)
        self.btn_export.pack(side=tk.LEFT, ipadx=10, ipady=3, padx=(0, 8))
        self.btn_generate = ttk.Button(buttons, text="📄 Generar PDF", style="Primary.TButton",
                                       command=self.start_generation)
        self.btn_generate.pack(side=tk.LEFT, ipadx=15, ipady=3)

        results = card(self, fill=tk.BOTH, expand=True)
        ttk.Label(results, text="Reporte de Ejecución y Códigos Faltantes:",
                  style="CardLabel.TLabel").pack(anchor="w", pady=(0, 5))
        log_container = tk.Frame(results, bg=Theme.CARD)
        log_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        scrollbar = tk.Scrollbar(log_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text = tk.Text(log_container, wrap=tk.WORD, yscrollcommand=scrollbar.set,
                                font=(Theme.MONO, 9), bg=Theme.CARD, fg=Theme.TEXT, relief="flat",
                                highlightbackground=Theme.BORDER, highlightthickness=1)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)

        self.btn_copy_missing = ttk.Button(results, text="Copiar Códigos Faltantes al Portapapeles",
                                           command=self.copy_missing_to_clipboard, state=tk.DISABLED)
        self.btn_copy_missing.pack(anchor="e")

    @staticmethod
    def _combo(parent, label, padx):
        panel = tk.Frame(parent, bg=Theme.CARD)
        panel.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=padx)
        ttk.Label(panel, text=label, style="CardLabel.TLabel").pack(anchor="w", pady=(0, 5))
        combo = ttk.Combobox(panel, state="readonly")
        combo.pack(fill=tk.X)
        return combo

    # --------------------------------------------------------------- Examinar
    def browse_excel(self):
        path = filedialog.askopenfilename(title="Seleccionar archivo Excel",
                                          filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if path:
            self.excel_path_var.set(path)
            self.on_excel_selected()

    def browse_database_file(self):
        path = ask_database_file()
        if not path:
            return
        self.db_path_var.set(path)
        # Proponer el PDF de salida junto al Excel si aún no hay destino
        if not self.pdf_output_var.get() and self.excel_path_var.get():
            excel_dir = os.path.dirname(self.excel_path_var.get())
            self.pdf_output_var.set(os.path.join(excel_dir, DEFAULT_PDF_NAME))

    def browse_pdf_output(self):
        path = filedialog.asksaveasfilename(title="Guardar PDF como", defaultextension=".pdf",
                                            filetypes=[("Documento PDF", "*.pdf")])
        if path:
            self.pdf_output_var.set(path)

    # ------------------------------------------------- Hojas y columnas Excel
    def on_excel_selected(self):
        path = self.excel_path_var.get()
        if not path or not os.path.exists(path):
            return
        self.set_status("Cargando hojas de Excel...")
        run_in_background(self.dispatcher, lambda: list_sheets(path), self._update_sheets,
                          lambda msg: self._show_load_error("Error de lectura", f"No se pudo abrir el Excel:\n{msg}",
                                                            "Error al abrir Excel"))

    def _update_sheets(self, sheets):
        self.sheet_combo["values"] = sheets
        if sheets:
            self.sheet_combo.set(sheets[0])
            self.on_sheet_selected()
        self.set_status("Excel cargado correctamente.")

    def on_sheet_selected(self):
        path, sheet = self.excel_path_var.get(), self.sheet_combo.get()
        if not path or not sheet:
            return
        self.set_status("Cargando columnas...")
        run_in_background(self.dispatcher, lambda: list_columns(path, sheet), self._update_columns,
                          lambda msg: self._show_load_error(
                              "Error", f"No se pudieron leer las columnas de la hoja '{sheet}':\n{msg}",
                              "Error al cargar columnas"))

    def _update_columns(self, columns):
        self.column_combo["values"] = columns
        selected = guess_code_column(columns)
        if selected:
            self.column_combo.set(selected)
        self.set_status(READY_TEXT)

    def _show_load_error(self, title, message, status):
        messagebox.showerror(title, message)
        self.set_status(status)

    # ------------------------------------------------------- Estado y reporte
    def set_status(self, text):
        self.lbl_status.config(text=text)

    def set_controls_state(self, state):
        combo_state = "readonly" if state == tk.NORMAL else state
        for widget in (self.btn_excel, self.btn_db, self.btn_pdf, self.btn_generate, self.btn_export,
                       self.btn_clear, self.ent_excel, self.ent_db, self.ent_pdf):
            widget.config(state=state)
        for combo in (self.sheet_combo, self.column_combo):
            combo.config(state=combo_state)
        copy_enabled = state == tk.NORMAL and self.missing_codes
        self.btn_copy_missing.config(state=tk.NORMAL if copy_enabled else tk.DISABLED)

    def _log(self, text):
        self.log_text.insert(tk.END, text)
        self.log_text.see(tk.END)

    def update_progress(self, percent, message):
        self.progress["value"] = percent
        self.set_status(message)
        self._log(f"{message}\n")

    def _report_progress_from_worker(self, current, total, message):
        """Callback de progreso llamado desde el hilo de trabajo."""
        percent = (current / total) * 100 if total > 0 else 0
        self.dispatcher.call(self.update_progress, percent, message)

    def _write_report(self, title, lines, missing):
        self._log("\n" + "=" * 41 + "\n" + title.center(41) + "\n" + "=" * 41 + "\n")
        self._log("".join(f"{line}\n" for line in lines))
        self._log(f"Códigos no encontrados: {len(missing)}\n\n")
        if missing:
            self._log("CÓDIGOS FALTANTES (Sin imagen asociada):\n")
            self._log("".join(f" - {code}\n" for code in missing))
        else:
            self._log("¡Felicidades! Se encontraron imágenes para todos los códigos.\n")

    def clear_all(self):
        """Limpia todos los campos de la pestaña y reinicia el estado."""
        self.excel_path_var.set("")
        for combo in (self.sheet_combo, self.column_combo):
            combo.set("")
            combo["values"] = []
        self.db_path_var.set(DEFAULT_DB_PATH)
        self.pdf_output_var.set("")
        self.missing_codes = []
        self.btn_copy_missing.config(state=tk.DISABLED)
        self.progress["value"] = 0
        self.set_status(READY_TEXT)
        self.log_text.delete("1.0", tk.END)

    def copy_missing_to_clipboard(self):
        if not self.missing_codes:
            return
        self.clipboard_clear()
        self.clipboard_append("\n".join(self.missing_codes))
        messagebox.showinfo("Portapapeles", "La lista de códigos faltantes se copió al portapapeles.")

    # ------------------------------------------------------- Tareas (casos de uso)
    def _validated_inputs(self, require_pdf):
        """Valida el formulario. Devuelve los argumentos comunes o None si falta algo."""
        inputs = dict(excel_path=self.excel_path_var.get(), sheet_name=self.sheet_combo.get(),
                      column_name=self.column_combo.get(), db_path=self.db_path_var.get())
        checks = [
            (os.path.exists(inputs["excel_path"]), "selecciona un archivo Excel válido."),
            (inputs["sheet_name"], "selecciona una hoja de Excel."),
            (inputs["column_name"], "selecciona la columna que contiene los códigos."),
            (os.path.isfile(inputs["db_path"]), "selecciona el archivo de la Base de Datos (.db o .csv)."),
        ]
        if require_pdf:
            checks.append((self.pdf_output_var.get(), "define la ruta de salida del PDF."))
        for ok, message in checks:
            if not ok:
                messagebox.showerror("Campos incompletos", f"Por favor, {message}")
                return None
        return inputs

    def _run_task(self, task, on_done, failure):
        """Prepara la interfaz y ejecuta `task(progress_callback)` en segundo plano."""
        self.set_controls_state(tk.DISABLED)
        self.progress["value"] = 0
        self.log_text.delete("1.0", tk.END)
        self.missing_codes = []
        self.btn_copy_missing.config(state=tk.DISABLED)
        run_in_background(self.dispatcher, lambda: task(self._report_progress_from_worker), on_done,
                          lambda msg: self._on_task_failed(msg, *failure))

    def _finish_task(self, status, missing):
        self.missing_codes = missing
        self.set_controls_state(tk.NORMAL)
        self.progress["value"] = 100
        self.set_status(status)

    def _on_task_failed(self, error_msg, status, title, detail):
        self.set_controls_state(tk.NORMAL)
        self.progress["value"] = 0
        self.set_status(status)
        self._log(f"\n[ERROR] {detail}: {error_msg}\n")
        messagebox.showerror(title, f"{detail}:\n{error_msg}")

    def start_generation(self):
        inputs = self._validated_inputs(require_pdf=True)
        if not inputs:
            return
        output_pdf = self.pdf_output_var.get()
        self._run_task(
            lambda progress: generate_pdf(**inputs, output_pdf_path=output_pdf, progress_callback=progress),
            on_done=lambda result: self._on_pdf_done(*result),
            failure=("Error en el proceso.", "Error de Generación", "No se pudo crear el PDF"),
        )

    def _on_pdf_done(self, pages, missing):
        self._finish_task("¡Proceso completado con éxito!", missing)
        self._write_report("REPORTE FINAL", [f"Imágenes añadidas al PDF: {pages}"], missing)
        messagebox.showinfo("Éxito", f"Se ha generado el PDF con éxito.\nTotal de imágenes unificadas: {pages}"
                                     f"\nCódigos faltantes: {len(missing)}")

    def start_export(self):
        inputs = self._validated_inputs(require_pdf=False)
        if not inputs:
            return
        output_folder = filedialog.askdirectory(title="Seleccionar carpeta para exportar las imágenes")
        if not output_folder:
            return
        self._run_task(
            lambda progress: export_images(**inputs, output_folder=output_folder, progress_callback=progress),
            on_done=lambda result: self._on_export_done(*result, output_folder),
            failure=("Error en la exportación.", "Error de Exportación", "No se pudieron exportar las imágenes"),
        )

    def _on_export_done(self, exported, missing, output_folder):
        self._finish_task("¡Exportación completada con éxito!", missing)
        self._write_report("REPORTE DE EXPORTACIÓN",
                           [f"Imágenes exportadas: {exported}", f"Carpeta destino: {output_folder}"], missing)
        messagebox.showinfo("Éxito", f"Se exportaron {exported} imágenes individuales a:\n{output_folder}"
                                     f"\nCódigos faltantes: {len(missing)}")


# =============================================================================
#  Pestaña 2 — Búsqueda rápida de un solo código
# =============================================================================

class QuickSearchTab(tk.Frame):
    """Busca un código, muestra su miniatura y permite abrirla o copiarla."""

    def __init__(self, parent, db_path_var, dispatcher):
        super().__init__(parent, bg=Theme.BG)
        self.db_path_var = db_path_var
        self.dispatcher = dispatcher
        self.code_var = tk.StringVar()
        self._found = None              # (código, ruta) del último resultado
        self._index = None              # Índice cacheado de la BD...
        self._index_db_path = None      # ...y la ruta con la que se construyó
        self._preview_image = None      # Referencia para evitar recolección de basura
        self._build()

    # ---------------------------------------------------------------- Layout
    def _build(self):
        container = tk.Frame(self, bg=Theme.BG)
        container.pack(fill=tk.BOTH, expand=True, padx=15, pady=(10, 15))

        search = card(container, fill=tk.X, pady=(0, 15))
        ttk.Label(search, text="Código del producto:", style="CardLabel.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 5))
        self.ent_code = ttk.Entry(search, textvariable=self.code_var, width=55, font=(Theme.FONT, 12))
        self.ent_code.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0, 10), pady=(0, 15))
        self.ent_code.bind("<Return>", lambda e: self.search())
        self.btn_search = ttk.Button(search, text="🔍 Buscar", style="Primary.TButton", command=self.search)
        self.btn_search.grid(row=1, column=2, sticky="ew", pady=(0, 15), ipady=4)
        # La base de datos comparte variable con la pestaña de catálogo
        file_row(search, 2, "Base de Datos Maestro (.db / .csv):", self.db_path_var,
                 "Seleccionar Base de Datos", self.browse_database_file, bottom_pad=5)
        search.columnconfigure(0, weight=1)
        search.columnconfigure(1, weight=1)

        result = card(container, fill=tk.BOTH, expand=True, pady=(0, 15))
        ttk.Label(result, text="Resultado:", style="CardLabel.TLabel").pack(anchor="w", pady=(0, 10))
        content = tk.Frame(result, bg=Theme.CARD)
        content.pack(fill=tk.BOTH, expand=True)

        preview = tk.Frame(content, bg=Theme.FIELD, width=PREVIEW_SIZE, height=PREVIEW_SIZE,
                           highlightbackground=Theme.BORDER, highlightthickness=1)
        preview.pack(side=tk.LEFT, padx=(0, 15), pady=5)
        preview.pack_propagate(False)
        self.preview_label = tk.Label(preview, bg=Theme.FIELD, text="Sin imagen",
                                      fg=Theme.TEXT_MUTED, font=(Theme.FONT, 9))
        self.preview_label.pack(expand=True)

        info = tk.Frame(content, bg=Theme.CARD)
        info.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=5)
        self.lbl_status = ttk.Label(info, style="CardSubLabel.TLabel")
        self.lbl_status.pack(anchor="w", pady=(0, 8))
        self.lbl_path = ttk.Label(info, text="", style="CardSubLabel.TLabel", wraplength=350)
        self.lbl_path.pack(anchor="w", pady=(0, 5))

        buttons = tk.Frame(container, bg=Theme.BG)
        buttons.pack(fill=tk.X)
        ttk.Button(buttons, text="🧹 Limpiar", command=self.clear).pack(side=tk.LEFT, ipadx=10, ipady=3, padx=(0, 8))
        self.btn_export = ttk.Button(buttons, text="📁 Exportar Copia", command=self.export_image)
        self.btn_export.pack(side=tk.RIGHT, ipadx=10, ipady=3, padx=(8, 0))
        self.btn_open = ttk.Button(buttons, text="📂 Abrir Imagen", command=self.open_image)
        self.btn_open.pack(side=tk.RIGHT, ipadx=10, ipady=3)

        self._show_result(None, "Ingresa un código y presiona Buscar.", Theme.TEXT_MUTED)

    # ---------------------------------------------------------------- Acciones
    def browse_database_file(self):
        path = ask_database_file()
        if path:
            self.db_path_var.set(path)

    def search(self):
        code = self.code_var.get().strip()
        if not code:
            messagebox.showwarning("Campo vacío", "Por favor, escribe un código para buscar.")
            return
        db_path = self.db_path_var.get()
        if not os.path.isfile(db_path):
            messagebox.showerror("Base de Datos",
                                 "La ruta de la Base de Datos no es válida. Selecciona un archivo .db o .csv.")
            return

        self._show_result(None, "Buscando...", Theme.TEXT_MUTED)
        self.btn_search.config(state=tk.DISABLED)
        # Los widgets y el caché solo se leen/escriben aquí, en el hilo de Tkinter
        cached = self._index if self._index_db_path == db_path else None

        def task():
            index = cached if cached is not None else index_images_from_db(db_path)
            return index, index.get(normalize_code(code))

        def on_done(result):
            self._index, self._index_db_path = result[0], db_path
            self.btn_search.config(state=tk.NORMAL)
            path = result[1]
            if path:
                self._show_result((code, path), f"✅ Código '{code}' encontrado:", Theme.SUCCESS)
            else:
                self._show_result(None, f"❌ Código '{code}' no encontrado en la Base de Datos.", Theme.ERROR)

        def on_error(message):
            self._invalidate_cache()  # La BD pudo haber cambiado
            self.btn_search.config(state=tk.NORMAL)
            self._show_result(None, f"⚠ Error: {message}", Theme.WARNING)

        run_in_background(self.dispatcher, task, on_done, on_error)

    def open_image(self):
        if self._found and os.path.isfile(self._found[1]):
            os.startfile(self._found[1])
        else:
            messagebox.showerror("Error", "La imagen no existe o no se puede acceder. Realiza una nueva búsqueda.")

    def export_image(self):
        if not self._found or not os.path.isfile(self._found[1]):
            messagebox.showerror("Error", "No hay imagen para exportar. Realiza una búsqueda primero.")
            return
        dest_folder = filedialog.askdirectory(title="Seleccionar carpeta de destino")
        if not dest_folder:
            return
        code, src_path = self._found
        dest_path = os.path.join(dest_folder, sanitize_filename(code) + os.path.splitext(src_path)[1])
        try:
            shutil.copy2(src_path, dest_path)
        except OSError as e:
            messagebox.showerror("Error de Exportación", f"No se pudo exportar la imagen:\n{e}")
            return
        messagebox.showinfo("Éxito", f"Imagen exportada correctamente a:\n{dest_path}")

    def clear(self):
        """Limpia la pestaña e invalida el caché para forzar recarga de la BD."""
        self.code_var.set("")
        self._invalidate_cache()
        self._show_result(None, "Ingresa un código y presiona Buscar.", Theme.TEXT_MUTED)

    # ------------------------------------------------------------- Presentación
    def _invalidate_cache(self):
        self._index = self._index_db_path = None

    def _show_result(self, found, status_text, color):
        """Único punto que actualiza el panel de resultado (evita estados incoherentes)."""
        self._found = found
        self.lbl_status.config(text=status_text, foreground=color)
        self.lbl_path.config(text=found[1] if found else "")
        button_state = tk.NORMAL if found else tk.DISABLED
        self.btn_open.config(state=button_state)
        self.btn_export.config(state=button_state)
        if found:
            self._show_preview(found[1])
        else:
            self._clear_preview()

    def _show_preview(self, path):
        try:
            with Image.open(path) as img:
                img.thumbnail((PREVIEW_SIZE - 10, PREVIEW_SIZE - 10), Image.LANCZOS)
                self._preview_image = ImageTk.PhotoImage(img)
            self.preview_label.config(image=self._preview_image, text="")
        except (OSError, ValueError, Image.DecompressionBombError):
            self._clear_preview("No se pudo\ncargar vista previa")

    def _clear_preview(self, text="Sin imagen"):
        self._preview_image = None
        self.preview_label.config(image="", text=text)


# =============================================================================
#  Ventana principal
# =============================================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Unificador de Catálogos a PDF")
        self.geometry("700x700")
        self.minsize(650, 650)
        self.configure(bg=Theme.BG)
        apply_dark_theme(self)

        # Estado compartido entre pestañas: ruta de la BD y puente entre hilos
        self.db_path_var = tk.StringVar(value=DEFAULT_DB_PATH)
        self.dispatcher = MainThreadDispatcher(self)

        main = tk.Frame(self, bg=Theme.BG)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        header = tk.Frame(main, bg=Theme.BG)
        header.pack(fill=tk.X, pady=(0, 15))
        ttk.Label(header, text="Unificador de Catálogos a PDF", style="Header.TLabel").pack(anchor="w")
        ttk.Label(header, text="Busca y exporta imágenes de productos desde la Base de Datos Maestra.",
                  style="Sub.TLabel").pack(anchor="w", pady=(2, 0))

        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.catalog_tab = CatalogTab(self.notebook, self.db_path_var, self.dispatcher)
        self.notebook.add(self.catalog_tab, text="  📋 Catálogo (Excel)  ")
        self.quick_tab = QuickSearchTab(self.notebook, self.db_path_var, self.dispatcher)
        self.notebook.add(self.quick_tab, text="  🔍 Búsqueda Rápida  ")


def enable_high_dpi():
    """Nitidez en pantallas de alta resolución (solo Windows; se ignora en otros SO)."""
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except (ImportError, AttributeError, OSError):
        pass


if __name__ == "__main__":
    enable_high_dpi()
    App().mainloop()
