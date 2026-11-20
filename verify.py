import os
import re
import sys
import shutil
import tempfile
import pandas as pd
# pyrefly: ignore [missing-import]
from PIL import Image

# Forzar UTF-8 en la consola de Windows para evitar errores con caracteres especiales
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from generate_test_data import setup_test_environment
from unificador import generate_pdf, export_images, index_images_from_db, sanitize_filename

# ============================================================
#  SUITE DE TESTS - UNIFICADOR DE CATÁLOGOS
# ============================================================

PASS = 0
FAIL = 0

def check(condition, test_name):
    """Helper para registrar resultados de tests."""
    global PASS, FAIL
    if condition:
        PASS += 1
        print(f"  ✅ PASS: {test_name}")
    else:
        FAIL += 1
        print(f"  ❌ FAIL: {test_name}")

def get_base_dir():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_env")

# ----------------------------------------------------------
# TEST 1: Filtro por prefijo "xz" (solo fotos del catálogo)
# ----------------------------------------------------------
def test_index_images():
    print("\n" + "=" * 60)
    print("  TEST 1: Filtro por prefijo 'xz' (solo fotos del catálogo)")
    print("=" * 60)
    
    base_dir = get_base_dir()
    db_path = os.path.join(base_dir, "BASE_DE_DATOS_FOTOS_MASTER.db")
    
    images_dict = index_images_from_db(db_path)
    
    # Las fotos con prefijo "xz" DEBEN aparecer (son las del catálogo final)
    check("gf-5657" in images_dict, "Foto catálogo 'xzGF-5657' indexada como 'gf-5657'")
    check("gf-1111" in images_dict, "Foto catálogo 'xzGF-1111' indexada como 'gf-1111'")
    check("gf-2222" in images_dict, "Foto catálogo 'xzGF-2222' indexada como 'gf-2222'")
    
    # Validar que para GF-2222 se eligió la versión más nueva de sub_sub_B y no la vieja
    if "gf-2222" in images_dict:
        path_chosen = images_dict["gf-2222"]
        check("sub_sub_B" in path_chosen, "GF-2222: Se eligió la versión más reciente (Fecha_Modificacion)")
    
    check("gf-4444" in images_dict, "Foto corrupta 'xzGF-4444' indexada (se filtrará al generar PDF)")
    
    # Las tomas alternativas SIN prefijo "xz" NO deben aparecer
    check("gf-3333" not in images_dict, "Toma alternativa 'GF-3333' (sin xz) NO indexada - CORRECTO")
    check("gf-5657_detalle" not in images_dict, "Toma detalle 'GF-5657_detalle' (sin xz) NO indexada - CORRECTO")
    
    # No debe haber claves con el prefijo "xz" aún presente
    xz_keys = [k for k in images_dict if k.startswith("xz")]
    check(len(xz_keys) == 0, f"Ninguna clave conserva el prefijo 'xz' (encontradas: {len(xz_keys)})")
    
    print(f"\n  Diccionario indexado ({len(images_dict)} fotos de catálogo):")
    for key, path in sorted(images_dict.items()):
        print(f"    '{key}' → {os.path.basename(path)}")

# ----------------------------------------------------------
# TEST 2: Generación del PDF (flujo principal)
# ----------------------------------------------------------
def test_generate_pdf():
    print("\n" + "=" * 60)
    print("  TEST 2: Generación del PDF (flujo principal)")
    print("=" * 60)
    
    base_dir = get_base_dir()
    excel_path = os.path.join(base_dir, "lista_codigos.xlsx")
    db_path = os.path.join(base_dir, "BASE_DE_DATOS_FOTOS_MASTER.db")
    output_pdf = os.path.join(base_dir, "resultado.pdf")
    
    if os.path.exists(output_pdf):
        os.remove(output_pdf)
    
    progress_messages = []
    
    def progress_cb(curr, tot, msg):
        progress_messages.append(msg)
        print(f"    [Progreso {curr}/{tot}] {msg}")
    
    try:
        matched_count, missing = generate_pdf(
            excel_path=excel_path,
            sheet_name="Lista Principal",
            column_name="Código Producto",
            db_path=db_path,
            output_pdf_path=output_pdf,
            progress_callback=progress_cb
        )
        generation_ok = True
    except Exception as e:
        print(f"    ⚠️  Excepción durante generación: {e}")
        matched_count, missing = 0, []
        generation_ok = False
    
    check(generation_ok, "generate_pdf() se ejecutó sin excepciones fatales")
    # 4 coincidencias en la BD, pero xzGF-4444 está corrupta: solo 3 páginas reales
    check(matched_count == 3, f"Páginas escritas en el PDF: {matched_count} (esperado: 3, la corrupta no cuenta)")
    
    expected_missing = ["GF-9999", "GF-3333"]
    check(missing == expected_missing, f"Códigos faltantes: {missing} (esperado: {expected_missing})")
    check("GF-3333" in missing, "GF-3333 reportado faltante (solo tiene toma alternativa sin 'xz')")
    
    check(os.path.exists(output_pdf), "Archivo PDF creado en disco")
    
    if os.path.exists(output_pdf):
        with open(output_pdf, 'rb') as f:
            content = f.read()
        header = content[:5]
        check(header == b'%PDF-', f"Encabezado PDF válido: {header}")
        real_pages = len(re.findall(rb"/Type\s*/Page\b(?!s)", content))
        check(real_pages == matched_count, f"Conteo reportado ({matched_count}) = páginas reales del PDF ({real_pages})")
        
        file_size = os.path.getsize(output_pdf)
        check(file_size > 1000, f"Tamaño del PDF: {file_size:,} bytes (> 1KB)")
    
    check(len(progress_messages) > 0, f"Callback de progreso invocado {len(progress_messages)} veces")
    check(any("Leyendo Base de Datos Maestra" in m for m in progress_messages), "Mensaje de indexación presente")
    check(any("PDF generado" in m for m in progress_messages), "Mensaje de éxito final presente")
    check(any("Omitiendo" in m and "4444" in m for m in progress_messages),
          "Imagen corrupta xzGF-4444 detectada y omitida")

# ----------------------------------------------------------
# TEST 3: Manejo de errores
# ----------------------------------------------------------
def test_error_handling():
    print("\n" + "=" * 60)
    print("  TEST 3: Manejo de errores")
    print("=" * 60)
    
    base_dir = get_base_dir()
    db_path = os.path.join(base_dir, "BASE_DE_DATOS_FOTOS_MASTER.db")
    excel_path = os.path.join(base_dir, "lista_codigos.xlsx")
    output_pdf = os.path.join(base_dir, "error_test.pdf")
    
    try:
        generate_pdf("C:\\ruta\\falsa\\no_existe.xlsx", "Hoja1", "Col", db_path, output_pdf)
        check(False, "Excel inexistente: debió lanzar excepción")
    except ValueError:
        check(True, "Excel inexistente: ValueError lanzado correctamente")
    except Exception as e:
        check(False, f"Excel inexistente: excepción inesperada ({type(e).__name__})")
    
    try:
        generate_pdf(excel_path, "Lista Principal", "COLUMNA_FALSA", db_path, output_pdf)
        check(False, "Columna inexistente: debió lanzar excepción")
    except ValueError as e:
        check("no existe" in str(e).lower() or "columna" in str(e).lower(),
              "Columna inexistente: ValueError con mensaje claro")
    except Exception as e:
        check(False, f"Columna inexistente: excepción inesperada ({type(e).__name__})")
    
    # BD Falsa
    fake_db = os.path.join(base_dir, "FAKE_DB.db")
    with open(fake_db, 'w') as f:
        f.write("Posible_Codigo;Ruta_Completa;Nombre_Archivo;Extension;Fecha_Modificacion;Carpeta_Padre\n") # Vacío
    try:
        generate_pdf(excel_path, "Lista Principal", "Código Producto", fake_db, output_pdf)
        check(False, "BD vacía o sin fotos 'xz': debió lanzar excepción")
    except (ValueError, RuntimeError):
        check(True, "BD vacía o sin fotos 'xz': excepción lanzada correctamente")
    except Exception as e:
        check(False, f"Carpeta sin fotos 'xz': excepción inesperada ({type(e).__name__})")
    finally:
        if os.path.exists(fake_db):
            os.remove(fake_db)

# ----------------------------------------------------------
# TEST 4: Imagen corrupta no crashea el proceso
# ----------------------------------------------------------
def test_corrupt_image_handling():
    print("\n" + "=" * 60)
    print("  TEST 4: Manejo de imagen corrupta con prefijo 'xz'")
    print("=" * 60)
    
    base_dir = get_base_dir()
    output_pdf = os.path.join(base_dir, "resultado.pdf")
    
    if os.path.exists(output_pdf):
        check(True, "PDF generado a pesar de imagen corrupta (xzGF-4444)")
    else:
        check(False, "PDF no se generó - posible crash por imagen corrupta")

# ----------------------------------------------------------
# TEST 5: Exportación de imágenes individuales
# ----------------------------------------------------------
def test_export_images():
    print("\n" + "=" * 60)
    print("  TEST 5: Exportación de imágenes (nombres seguros y duplicados)")
    print("=" * 60)

    base_dir = get_base_dir()
    db_path = os.path.join(base_dir, "BASE_DE_DATOS_FOTOS_MASTER.db")
    tmp = tempfile.mkdtemp()
    try:
        # Flujo normal
        out = os.path.join(tmp, "normal")
        count, missing = export_images(os.path.join(base_dir, "lista_codigos.xlsx"),
                                       "Lista Principal", "Código Producto", db_path, out)
        check(count == len(os.listdir(out)), f"Conteo exportado ({count}) = archivos en carpeta ({len(os.listdir(out))})")
        check(missing == ["GF-9999", "GF-3333"], f"Códigos faltantes en export: {missing}")

        # Códigos duplicados (misma clave con distinto formato)
        excel_dup = os.path.join(tmp, "dup.xlsx")
        pd.DataFrame({"c": ["GF-1111", "gf-1111", " GF-1111 "]}).to_excel(excel_dup, index=False)
        out = os.path.join(tmp, "dup")
        count, _ = export_images(excel_dup, 0, "c", db_path, out)
        check(count == 1 and len(os.listdir(out)) == 1, f"Duplicados: 1 archivo y conteo 1 (conteo={count})")

        # Código con separadores de ruta: no debe escapar de la carpeta destino
        evil = "../escape"
        rows = pd.read_csv(db_path, sep=";", dtype=str, encoding="utf-8-sig")
        rows.loc[len(rows)] = ["'xz" + evil, rows.Ruta_Completa[0], "xz" + evil + ".jpg",
                               ".jpg", "2026-01-01 00:00:00", "x"]
        db_evil = os.path.join(tmp, "evil.db")
        rows.to_csv(db_evil, sep=";", index=False, encoding="utf-8-sig")
        excel_evil = os.path.join(tmp, "evil.xlsx")
        pd.DataFrame({"c": [evil]}).to_excel(excel_evil, index=False)
        out = os.path.join(tmp, "dest", "sub")
        export_images(excel_evil, 0, "c", db_evil, out)
        check(not os.path.exists(os.path.join(tmp, "dest", "escape.jpg")),
              "Código '../escape' NO escribe fuera de la carpeta destino")
        check(len(os.listdir(out)) == 1, f"Archivo saneado creado dentro del destino: {os.listdir(out)}")

        check(sanitize_filename('a<b>c:d"e/f\\g|h?i*j') == "a_b_c_d_e_f_g_h_i_j", "sanitize_filename reemplaza caracteres inválidos")
        check(sanitize_filename("...") == "_", "sanitize_filename nunca devuelve nombre vacío")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

# ============================================================
#  EJECUCIÓN PRINCIPAL
# ============================================================
def main():
    print("╔════════════════════════════════════════════════════════════╗")
    print("║    SUITE DE VERIFICACIÓN - UNIFICADOR DE CATÁLOGOS       ║")
    print("╠════════════════════════════════════════════════════════════╣")
    print("║  Regla de negocio: Solo fotos con prefijo 'xz' van al   ║")
    print("║  PDF. Las demás tomas/ángulos se ignoran.                ║")
    print("╚════════════════════════════════════════════════════════════╝")
    
    setup_test_environment()
    
    test_index_images()
    test_generate_pdf()
    test_error_handling()
    test_corrupt_image_handling()
    test_export_images()

    total = PASS + FAIL
    print("\n" + "═" * 60)
    print(f"  RESUMEN FINAL: {PASS}/{total} tests pasaron")
    if FAIL > 0:
        print(f"  ⚠️  {FAIL} test(s) FALLARON")
    else:
        print(f"  🎉 ¡TODOS LOS TESTS PASARON EXITOSAMENTE!")
    print("═" * 60)
    
    if FAIL > 0:
        sys.exit(1)

if __name__ == "__main__":
    main()
