"""
Lógica de negocio del Unificador de Catálogos.

No conoce la interfaz gráfica: recibe rutas y un callback de progreso
opcional `progress_callback(actual, total, mensaje)`, y devuelve resultados
o lanza excepciones con mensajes listos para mostrar al usuario.
"""
import os
import re
import shutil

import pandas as pd
# pyrefly: ignore [missing-import]
from PIL import Image

# --- Reglas de negocio -------------------------------------------------------
CATALOG_PREFIX = "xz"                       # Solo las fotos finales del catálogo
SUPPORTED_EXTENSIONS = (".jpg", ".jpeg", ".png")
DB_REQUIRED_COLUMNS = ("Nombre_Archivo", "Ruta_Completa", "Fecha_Modificacion")

# "xzGF-5657.jpg" -> code="gf-5657", ext="jpg" (sobre el nombre en minúsculas)
_CATALOG_NAME_PATTERN = rf"^{re.escape(CATALOG_PREFIX)}(?P<code>.+)\.(?P<ext>[^.]+)$"
_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# Errores esperables al abrir una imagen: corrupta, inaccesible o demasiado grande
_IMAGE_ERRORS = (OSError, ValueError, Image.DecompressionBombError)


def _ignore_progress(current, total, message):
    """Callback por defecto: no reporta nada."""


# =============================================================================
#  Funciones puras (sin E/S): fáciles de probar de forma aislada
# =============================================================================

def normalize_code(code):
    """Forma canónica de un código para comparar: sin espacios y en minúsculas."""
    return code.strip().lower()


def clean_codes(values):
    """Quita espacios y descarta valores vacíos o 'nan', conservando el orden."""
    stripped = (str(value).strip() for value in values)
    return [code for code in stripped if code and code.lower() != "nan"]


def build_image_index(df):
    """
    Construye el índice {código: ruta} a partir de las filas de la base maestra.

    REGLAS DE NEGOCIO:
    1. Solo se indexan imágenes cuyo nombre empieza con el prefijo "xz" y
       tienen una extensión soportada.
    2. Si un código aparece varias veces, gana la 'Fecha_Modificacion' más
       reciente. Las fechas inválidas pierden siempre y, ante un empate, gana
       la primera fila del archivo (ordenamiento estable).

    Llave: código del producto en minúsculas, sin prefijo ni extensión.
    Búsqueda posterior: O(1) por código (tabla hash).
    """
    parts = df["Nombre_Archivo"].str.lower().str.extract(_CATALOG_NAME_PATTERN)
    candidates = df.assign(
        code=parts["code"],
        ext="." + parts["ext"],
        fecha=pd.to_datetime(df["Fecha_Modificacion"], errors="coerce"),
    )
    valid = candidates[candidates["code"].notna() & candidates["ext"].isin(SUPPORTED_EXTENSIONS)]
    latest = (valid
              .sort_values("fecha", ascending=False, kind="stable", na_position="last")
              .drop_duplicates(subset="code", keep="first"))
    return dict(zip(latest["code"], latest["Ruta_Completa"]))


def match_codes(codes, images_dict):
    """
    Busca cada código en el índice respetando el orden de entrada.
    Devuelve ([(código, ruta), ...] encontrados, [códigos faltantes]).
    """
    matches, missing = [], []
    for code in codes:
        path = images_dict.get(normalize_code(code))
        if path is None:
            missing.append(code)
        else:
            matches.append((code, path))
    return matches, missing


def sanitize_filename(name):
    """
    Convierte un texto (ej. un código del Excel) en un nombre de archivo seguro:
    reemplaza separadores de ruta y caracteres inválidos en Windows por '_',
    y elimina puntos/espacios finales. Evita escribir fuera de la carpeta destino.
    """
    clean = _INVALID_FILENAME_CHARS.sub("_", name).strip(" .")
    return clean or "_"


def convert_to_rgb(img):
    """
    Convierte una imagen de cualquier formato a RGB.
    Si tiene transparencia, la coloca sobre un fondo blanco en lugar de negro.
    """
    if img.mode == "RGB":
        return img
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        rgba_img = img.convert("RGBA")
        background = Image.new("RGB", rgba_img.size, (255, 255, 255))
        background.paste(rgba_img, mask=rgba_img.split()[-1])
        return background
    return img.convert("RGB")


# =============================================================================
#  Lectura de archivos (Excel y Base de Datos Maestra)
# =============================================================================

def list_sheets(excel_path):
    """Nombres de las hojas de un Excel."""
    with pd.ExcelFile(excel_path) as workbook:
        return list(workbook.sheet_names)


def list_columns(excel_path, sheet_name):
    """Encabezados de una hoja (solo lee la primera fila)."""
    return list(pd.read_excel(excel_path, sheet_name=sheet_name, nrows=0).columns)


def read_codes(excel_path, sheet_name, column_name):
    """Lee y limpia los códigos de una columna del Excel."""
    try:
        # dtype=str evita que los enteros se lean como flotantes (ej. 1234.0)
        df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str)
    except Exception as e:  # pandas lanza tipos muy variados según el problema
        raise ValueError(f"Error al leer el archivo Excel: {e}") from e

    if column_name not in df.columns:
        raise ValueError(f"La columna '{column_name}' no existe en la hoja seleccionada del Excel.")

    codes = clean_codes(df[column_name].dropna())
    if not codes:
        raise ValueError("No se encontraron códigos válidos en la columna seleccionada.")
    return codes


def index_images_from_db(db_path):
    """Lee la base de datos maestra (.db / .csv separado por ';') y la indexa."""
    if not os.path.isfile(db_path):
        raise ValueError(f"El archivo de base de datos no existe: {db_path}")

    df = pd.read_csv(db_path, sep=";", dtype=str, encoding="utf-8-sig", na_filter=False)
    if not set(DB_REQUIRED_COLUMNS).issubset(df.columns):
        raise ValueError("La base de datos no contiene las columnas necesarias "
                         f"({', '.join(DB_REQUIRED_COLUMNS)}).")
    return build_image_index(df)


def _find_matches(excel_path, sheet_name, column_name, db_path, report):
    """Paso común a PDF y exportación: Excel -> índice de la BD -> coincidencias."""
    codes = read_codes(excel_path, sheet_name, column_name)
    report(0, len(codes), "Leyendo Base de Datos Maestra...")
    matches, missing = match_codes(codes, index_images_from_db(db_path))
    if not matches:
        raise ValueError("No se encontró ninguna imagen que coincida exactamente con los códigos de la lista.")
    return matches, missing


# =============================================================================
#  Casos de uso
# =============================================================================

def _iter_rgb_images(matches, report, opened, skipped):
    """
    Genera las imágenes en RGB una a una (ahorra memoria), omitiendo las ilegibles.
    Registra en `opened` todo lo abierto (Pillow lo necesita abierto hasta terminar
    de guardar el PDF) y en `skipped` los códigos omitidos.
    """
    total = len(matches)
    for idx, (code, path) in enumerate(matches, start=1):
        try:
            img = Image.open(path)
            opened.append(img)
            rgb_img = convert_to_rgb(img)
        except _IMAGE_ERRORS as e:
            skipped.append(code)
            report(idx, total, f"Omitiendo por error al abrir {os.path.basename(path)}: {e}")
            continue
        if rgb_img is not img:
            opened.append(rgb_img)
        report(idx, total, f"Agregando página ({idx}/{total}): {code}")
        yield rgb_img


def generate_pdf(excel_path, sheet_name, column_name, db_path, output_pdf_path, progress_callback=None):
    """
    Une en un PDF las imágenes de los códigos del Excel, en el mismo orden.
    Devuelve (páginas escritas, códigos sin imagen en la BD).
    """
    report = progress_callback or _ignore_progress
    matches, missing = _find_matches(excel_path, sheet_name, column_name, db_path, report)

    opened, skipped = [], []
    try:
        pages = _iter_rgb_images(matches, report, opened, skipped)
        first_page = next(pages, None)
        if first_page is None:
            raise RuntimeError("Todas las imágenes encontradas están corruptas y no se pudo generar el PDF.")
        try:
            first_page.save(output_pdf_path, save_all=True, append_images=pages)
        except PermissionError as e:
            raise RuntimeError("El archivo PDF de destino está abierto o bloqueado por otro programa. "
                               "Por favor, ciérrelo e intente nuevamente.") from e
        except Exception as e:  # Pillow puede fallar de muchas formas al escribir
            raise RuntimeError(f"Error al escribir el archivo PDF: {e}") from e
    finally:
        for img in opened:
            img.close()

    report(len(matches), len(matches), "¡PDF generado con éxito!")
    return len(matches) - len(skipped), missing


def export_images(excel_path, sheet_name, column_name, db_path, output_folder, progress_callback=None):
    """
    Copia a `output_folder` la imagen de cada código del Excel, nombrada con el
    código (saneado). Los códigos repetidos se copian una sola vez.
    Devuelve (imágenes copiadas, códigos sin imagen en la BD).
    """
    report = progress_callback or _ignore_progress
    matches, missing = _find_matches(excel_path, sheet_name, column_name, db_path, report)
    os.makedirs(output_folder, exist_ok=True)

    exported, seen = 0, set()
    total = len(matches)
    for idx, (code, src_path) in enumerate(matches, start=1):
        key = normalize_code(code)
        if key in seen:
            report(idx, total, f"Código duplicado omitido: {code}")
            continue
        seen.add(key)

        ext = os.path.splitext(src_path)[1]
        dest_path = os.path.join(output_folder, f"{sanitize_filename(code)}{ext}")
        try:
            shutil.copy2(src_path, dest_path)
        except OSError as e:
            report(idx, total, f"Error al copiar {code}: {e}")
            continue
        exported += 1
        report(idx, total, f"Exportando imagen ({idx}/{total}): {code}")

    if exported == 0:
        raise ValueError("No se pudo copiar ninguna de las imágenes encontradas.")

    report(total, total, f"¡Exportación completada! {exported} imágenes copiadas.")
    return exported, missing
