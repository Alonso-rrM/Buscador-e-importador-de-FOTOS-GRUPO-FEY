import os
import re
import shutil
import pandas as pd
# pyrefly: ignore [missing-import]
from PIL import Image

SUPPORTED_EXTENSIONS = ('.jpg', '.jpeg', '.png')

def index_images_from_db(db_path):
    """
    Lee la base de datos maestra (.db / .csv) y devuelve un Diccionario (Hash Table).
    
    REGLA DE NEGOCIO: 
    1. Solo se indexan imágenes cuyo nombre empiece con el prefijo "xz" (foto final).
    2. Si hay múltiples versiones (duplicados de código), se conserva 
       la que tenga la 'Fecha_Modificacion' más reciente.
    
    Llave: Código real del producto (nombre sin el prefijo "xz", en minúsculas).
    Valor: Ruta completa de la imagen.
    Complejidad de búsqueda: O(1) por código.
    """
    if not os.path.isfile(db_path):
        raise ValueError(f"El archivo de base de datos no existe: {db_path}")

    # Leer el archivo CSV con separador punto y coma
    df = pd.read_csv(db_path, sep=';', dtype=str, encoding='utf-8-sig', na_filter=False)
    
    required_cols = {'Nombre_Archivo', 'Ruta_Completa', 'Fecha_Modificacion'}
    if not required_cols.issubset(set(df.columns)):
        raise ValueError("La base de datos no contiene las columnas necesarias (Nombre_Archivo, Ruta_Completa, Fecha_Modificacion).")
    
    # 1. Convertir nombres a minúsculas
    df['Nombre_Lower'] = df['Nombre_Archivo'].str.lower()
    
    # 2. Filtrar solo los que empiezan por 'xz'
    df_xz = df[df['Nombre_Lower'].str.startswith('xz')]
    if df_xz.empty:
        return {}
    
    # 3. Extraer extensión y filtrar las soportadas
    df_xz = df_xz.copy()
    df_xz['Ext_Lower'] = df_xz['Nombre_Lower'].apply(lambda x: os.path.splitext(x)[1])
    df_valid = df_xz[df_xz['Ext_Lower'].isin(SUPPORTED_EXTENSIONS)].copy()
    if df_valid.empty:
        return {}
        
    # 4. Extraer el código limpio (quitar "xz" y la extensión)
    # Ejemplo: "xzgf-5657.jpg" -> "gf-5657"
    df_valid['Clean_Code'] = df_valid['Nombre_Lower'].apply(lambda x: os.path.splitext(x)[0][2:])
    
    # 5. Ordenar por Fecha_Modificacion (descendente)
    # pd.to_datetime convierte correctamente fechas como "2026-02-06 10:08:04"
    df_valid['Fecha_Date'] = pd.to_datetime(df_valid['Fecha_Modificacion'], errors='coerce')
    df_sorted = df_valid.sort_values(by='Fecha_Date', ascending=False)
    
    # 6. Eliminar duplicados quedándonos con el primero (el más reciente)
    df_final = df_sorted.drop_duplicates(subset=['Clean_Code'], keep='first')
    
    # 7. Construir y retornar el diccionario Hash
    images_dict = dict(zip(df_final['Clean_Code'], df_final['Ruta_Completa']))
    
    return images_dict

_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

def sanitize_filename(name):
    """
    Convierte un texto (ej. un código del Excel) en un nombre de archivo seguro:
    reemplaza separadores de ruta y caracteres inválidos en Windows por '_',
    y elimina puntos/espacios finales. Evita escribir fuera de la carpeta destino.
    """
    clean = _INVALID_FILENAME_CHARS.sub('_', name).strip(' .')
    return clean or '_'

def convert_to_rgb(img):
    """
    Convierte una imagen de cualquier formato a RGB.
    Si tiene transparencia (PNG/WebP), la coloca sobre un fondo blanco en lugar de negro.
    """
    if img.mode == 'RGB':
        return img
    # Si tiene transparencia o canal alfa, la procesamos con fondo blanco
    if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
        rgba_img = img.convert('RGBA')
        background = Image.new('RGB', rgba_img.size, (255, 255, 255))
        background.paste(rgba_img, mask=rgba_img.split()[-1])
        return background
    else:
        return img.convert('RGB')

def generate_pdf(excel_path, sheet_name, column_name, db_path, output_pdf_path, progress_callback=None):
    """
    Lee los códigos de Excel, busca imágenes coincidentes en la Base de Datos Maestra,
    y las une en un único archivo PDF en el orden del Excel.
    """
    # 1. Leer Excel
    try:
        # Se usa dtype=str para evitar que enteros se lean como flotantes (ej. 1234.0)
        df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str)
    except Exception as e:
        raise ValueError(f"Error al leer el archivo Excel: {str(e)}")
        
    if column_name not in df.columns:
        raise ValueError(f"La columna '{column_name}' no existe en la hoja seleccionada del Excel.")
    
    # Extraer códigos, limpiar espacios y omitir vacíos (incluyendo 'nan' literal)
    raw_codes = df[column_name].dropna().astype(str).tolist()
    codes = [c.strip() for c in raw_codes if c.strip() and c.strip().lower() != 'nan']
    
    if not codes:
        raise ValueError("No se encontraron códigos válidos en la columna seleccionada.")
    
    # 2. Leer Base de Datos Maestra e indexar imágenes (Hash Table O(1))
    if progress_callback:
        progress_callback(0, len(codes), "Leyendo Base de Datos Maestra...")
    
    # --- IMPLEMENTACIÓN DE HASH TABLE DESDE DB ---
    # Se crea el diccionario con las rutas de las imágenes pre-procesadas
    images_dict = index_images_from_db(db_path)
    
    # 3. Buscar coincidencias manteniendo el orden del Excel
    pdf_images_paths = []
    missing_codes = []
    
    total_codes = len(codes)
    for idx, code in enumerate(codes):
        if progress_callback:
            progress_callback(idx, total_codes, f"Buscando código ({idx+1}/{total_codes}): {code}")
            
        # Limpieza del código del Excel
        code_lower = code.lower().strip()
        
        # --- COINCIDENCIA ESTRICTA (Strict Matching O(1)) ---
        # Búsqueda directa en el diccionario, evitando falsos positivos de .endswith()
        if code_lower in images_dict:
            pdf_images_paths.append(images_dict[code_lower])
        else:
            missing_codes.append(code)
            
    if not pdf_images_paths:
        raise ValueError("No se encontró ninguna imagen que coincida exactamente con los códigos de la lista.")
        
    if progress_callback:
        progress_callback(total_codes - 1, total_codes, "Generando archivo PDF...")
        
    # 4. Cargar imágenes y unirlas en el PDF (uso de generador para ahorrar memoria)
    # NOTA: Pillow necesita que las imágenes permanezcan abiertas hasta que .save()
    # termine de escribir el PDF completo. Por eso las acumulamos y cerramos al final.
    opened_images = []
    pages_written = 1  # La primera imagen válida ya inicializa el PDF

    def image_generator(paths_iterator):
        nonlocal pages_written
        for img_path in paths_iterator:
            try:
                img = Image.open(img_path)
                rgb_img = convert_to_rgb(img)
                # Registrar para cerrar después del guardado
                opened_images.append(img)
                if rgb_img is not img:
                    opened_images.append(rgb_img)
                pages_written += 1
                yield rgb_img
            except Exception as e:
                # --- MANEJO DE ERRORES: Mostrar el error real ---
                if progress_callback:
                    progress_callback(total_codes - 1, total_codes, f"Omitiendo por error al abrir {os.path.basename(img_path)}: {str(e)}")
                continue
                
    # Convertimos la lista en un iterador para consumir la primera imagen válida y pasar el resto al generador
    paths_iter = iter(pdf_images_paths)
    first_img = None
    first_orig = None
    
    # Buscamos la primera imagen que no esté corrupta para inicializar el PDF
    for path in paths_iter:
        try:
            first_orig = Image.open(path)
            first_img = convert_to_rgb(first_orig)
            break
        except Exception as e:
            if progress_callback:
                progress_callback(total_codes - 1, total_codes, f"Omitiendo imagen inicial por error al abrir {os.path.basename(path)}: {str(e)}")
            continue
            
    if not first_img:
        raise RuntimeError("Todas las imágenes encontradas están corruptas y no se pudo generar el PDF.")
            
    try:
        # Guardado del PDF consumiendo el generador
        first_img.save(
            output_pdf_path,
            save_all=True,
            append_images=image_generator(paths_iter)
        )
    except PermissionError:
        # --- MANEJO DE ERRORES: Archivos Bloqueados ---
        # Captura específica si el usuario tiene el PDF abierto en otro programa
        raise RuntimeError("El archivo PDF de destino está abierto o bloqueado por otro programa. Por favor, ciérrelo e intente nuevamente.")
    except Exception as e:
        raise RuntimeError(f"Error al escribir el archivo PDF: {str(e)}")
    finally:
        # Cerrar TODAS las imágenes después de que el PDF se haya escrito
        if first_orig:
            first_orig.close()
        if first_img and first_img is not first_orig:
            first_img.close()
        for img in opened_images:
            try:
                img.close()
            except Exception:
                pass
            
    if progress_callback:
        progress_callback(total_codes, total_codes, "¡PDF generado con éxito!")
        
    # Páginas realmente escritas (excluye imágenes omitidas por estar corruptas)
    return pages_written, missing_codes


def export_images(excel_path, sheet_name, column_name, db_path, output_folder, progress_callback=None):
    """
    Lee los códigos de Excel, busca imágenes coincidentes en la Base de Datos Maestra,
    y copia cada imagen individual a la carpeta de destino.
    
    Cada archivo se nombra con el código del Excel para fácil identificación.
    Si ya existe un archivo con el mismo nombre, se sobreescribe.
    """
    # 1. Leer Excel (misma lógica que generate_pdf)
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str)
    except Exception as e:
        raise ValueError(f"Error al leer el archivo Excel: {str(e)}")
        
    if column_name not in df.columns:
        raise ValueError(f"La columna '{column_name}' no existe en la hoja seleccionada del Excel.")
    
    raw_codes = df[column_name].dropna().astype(str).tolist()
    codes = [c.strip() for c in raw_codes if c.strip() and c.strip().lower() != 'nan']
    
    if not codes:
        raise ValueError("No se encontraron códigos válidos en la columna seleccionada.")
    
    # 2. Leer Base de Datos Maestra e indexar imágenes
    if progress_callback:
        progress_callback(0, len(codes), "Leyendo Base de Datos Maestra...")
    
    images_dict = index_images_from_db(db_path)
    
    # 3. Crear la carpeta de destino si no existe
    os.makedirs(output_folder, exist_ok=True)
    
    # 4. Buscar coincidencias y copiar cada imagen individual
    exported_count = 0
    missing_codes = []
    seen_codes = set()

    total_codes = len(codes)
    for idx, code in enumerate(codes):
        if progress_callback:
            progress_callback(idx, total_codes, f"Exportando imagen ({idx+1}/{total_codes}): {code}")

        code_lower = code.lower().strip()

        # Código repetido en el Excel: ya se exportó, no se vuelve a copiar
        # (evita sobrescribir el mismo archivo y contarlo dos veces).
        if code_lower in seen_codes:
            if progress_callback:
                progress_callback(idx, total_codes, f"Código duplicado omitido: {code}")
            continue

        if code_lower in images_dict:
            seen_codes.add(code_lower)
            src_path = images_dict[code_lower]
            # Mantener la extensión original del archivo fuente
            ext = os.path.splitext(src_path)[1]
            # Nombrar el archivo con el código del Excel, saneado para que no
            # pueda contener separadores de ruta ni caracteres inválidos.
            dest_filename = f"{sanitize_filename(code)}{ext}"
            dest_path = os.path.join(output_folder, dest_filename)

            try:
                shutil.copy2(src_path, dest_path)
                exported_count += 1
            except Exception as e:
                if progress_callback:
                    progress_callback(idx, total_codes, f"Error al copiar {code}: {str(e)}")
        else:
            missing_codes.append(code)
    
    if exported_count == 0:
        raise ValueError("No se encontró ninguna imagen que coincida con los códigos de la lista.")
    
    if progress_callback:
        progress_callback(total_codes, total_codes, f"¡Exportación completada! {exported_count} imágenes copiadas.")
        
    return exported_count, missing_codes
