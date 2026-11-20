import os
import pandas as pd
# pyrefly: ignore [missing-import]
from PIL import Image, ImageDraw, ImageFont

def create_test_image(path, text, color, size=(800, 600)):
    """Crea una imagen de prueba con texto y color de fondo."""
    img = Image.new('RGB', size, color=color)
    d = ImageDraw.Draw(img)
    
    try:
        font = ImageFont.truetype("arial.ttf", 40)
    except IOError:
        font = ImageFont.load_default()
        
    d.text((100, 250), text, fill=(255, 255, 255), font=font)
    img.save(path)

def create_corrupt_image(path):
    """Crea un archivo con datos basura para simular una imagen corrupta."""
    with open(path, 'wb') as f:
        f.write(b'\x00\x01\x02CORRUPTO\xff\xfe')

def setup_test_environment(base_dir=None):
    """Crea imágenes, Base de Datos Maestra y Excel de prueba. Devuelve la carpeta usada."""
    if base_dir is None:
        base_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_env")
    images_dir = os.path.join(base_dir, "imagenes")
    os.makedirs(images_dir, exist_ok=True)
    
    # Crear subcarpetas para simular el servidor real
    sub1 = os.path.join(images_dir, "categoria_A")
    sub2 = os.path.join(images_dir, "categoria_B")
    sub3 = os.path.join(sub2, "sub_sub_B")
    
    os.makedirs(sub1, exist_ok=True)
    os.makedirs(sub2, exist_ok=True)
    os.makedirs(sub3, exist_ok=True)
    
    # ===================================================================
    # SIMULACIÓN DEL SERVIDOR REAL:
    # - Las fotos con prefijo "xz" son las FINALES del catálogo (se usan)
    # - Las fotos SIN prefijo "xz" son tomas alternativas (se ignoran)
    # ===================================================================
    
    # Lista para armar la Base de Datos Maestra (CSV)
    db_rows = []
    
    def add_to_db(path, is_corrupt=False, custom_date=None):
        import datetime
        fname = os.path.basename(path)
        ext = os.path.splitext(fname)[1]
        posible = f"'{os.path.splitext(fname)[0]}"
        folder = os.path.basename(os.path.dirname(path))
        
        # Fecha por defecto: hoy
        fecha = custom_date if custom_date else datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        db_rows.append({
            "Posible_Codigo": posible,
            "Ruta_Completa": path,
            "Nombre_Archivo": fname,
            "Extension": ext,
            "Fecha_Modificacion": fecha,
            "Carpeta_Padre": folder
        })
    
    # --- FOTOS FINALES DEL CATÁLOGO (prefijo "xz") → DEBEN incluirse ---
    p1 = os.path.join(sub1, "xzGF-5657.jpg")
    create_test_image(p1, "CATALOGO FINAL\nGF-5657", (26, 115, 232))
    add_to_db(p1)
    
    p2 = os.path.join(sub2, "xzGF-1111.png")
    create_test_image(p2, "CATALOGO FINAL\nGF-1111", (219, 68, 85))
    add_to_db(p2)
    
    # DUPLICADO PARA PROBAR FECHAS (xzGF-2222.jpg aparece en dos carpetas)
    p3_old = os.path.join(sub1, "xzGF-2222.jpg")
    create_test_image(p3_old, "CATALOGO FINAL VIEJO\nGF-2222", (50, 50, 50))
    add_to_db(p3_old, custom_date="2023-01-01 10:00:00") # Muy viejo
    
    p3_new = os.path.join(sub3, "xzGF-2222.jpg")
    create_test_image(p3_new, "CATALOGO FINAL NUEVO\nGF-2222", (15, 157, 88))
    add_to_db(p3_new, custom_date="2026-01-01 10:00:00") # Más reciente -> DEBE GANAR ESTE
    
    # --- TOMAS ALTERNATIVAS (sin prefijo "xz") → NO deben incluirse ---
    p4 = os.path.join(sub1, "GF-5657.jpg")
    create_test_image(p4, "ANGULO LATERAL\nGF-5657 (NO USAR)", (100, 100, 100))
    add_to_db(p4)
    
    p5 = os.path.join(sub1, "GF-5657_detalle.jpg")
    create_test_image(p5, "DETALLE\nGF-5657 (NO USAR)", (80, 80, 80))
    add_to_db(p5)
    
    p6 = os.path.join(sub2, "GF-1111.png")
    create_test_image(p6, "TOMA TRASERA\nGF-1111 (NO USAR)", (120, 120, 120))
    add_to_db(p6)
    
    p7 = os.path.join(sub1, "GF-3333.jpg")
    create_test_image(p7, "TOMA UNICA\nGF-3333 (NO HAY xz)", (60, 60, 60))
    add_to_db(p7)
    
    # --- FOTO CORRUPTA con prefijo "xz" (debe omitirse sin crashear) ---
    p8 = os.path.join(sub2, "xzGF-4444.jpg")
    create_corrupt_image(p8)
    add_to_db(p8)
    
    # Generar la Base de Datos Maestra (CSV)
    db_path = os.path.join(base_dir, "BASE_DE_DATOS_FOTOS_MASTER.db")
    df_db = pd.DataFrame(db_rows)
    # Se guarda con sep=';' y comillas
    df_db.to_csv(db_path, sep=';', index=False, encoding='utf-8-sig', quoting=2) # quoting=2 es csv.QUOTE_NONNUMERIC
    
    # Crear archivo Excel de prueba
    excel_path = os.path.join(base_dir, "lista_codigos.xlsx")
    data = {
        "Código Producto": [
            "GF-2222",    # Tiene 2 fotos "xz". Debe elegir la más reciente
            "GF-5657",    # Tiene foto "xz" (+ tomas alternativas sin xz)
            "GF-9999",    # NO existe
            "GF-1111",    # Tiene foto "xz" (+ toma alternativa)
            "GF-3333",    # Solo tiene foto SIN prefijo "xz"
            "GF-4444",    # Tiene foto "xz" corrupta
        ]
    }
    df = pd.DataFrame(data)
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Lista Principal", index=False)
        pd.DataFrame({"Otros": [1, 2, 3]}).to_excel(writer, sheet_name="Hoja Secundaria", index=False)

    return base_dir

if __name__ == "__main__":
    base_dir = setup_test_environment()
    print(f"Entorno de pruebas configurado exitosamente en: {base_dir}")
