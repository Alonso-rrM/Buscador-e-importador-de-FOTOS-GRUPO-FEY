"""
Suite de pruebas del Unificador de Catálogos (unittest, sin dependencias extra).

    python verify.py

- Pruebas unitarias: funciones puras, sin archivos ni red (rápidas).
- Pruebas de integración: Excel, Base de Datos y PDF reales en una carpeta
  temporal que se crea y borra en cada ejecución.
Cada prueba sigue el patrón AAA (Arrange / Act / Assert) y es independiente
del orden de ejecución.
"""
import os
import re
import shutil
import sys
import tempfile
import unittest

import pandas as pd

from generate_test_data import setup_test_environment
from gui import guess_code_column
from unificador import (
    build_image_index, clean_codes, export_images, generate_pdf, index_images_from_db,
    list_columns, list_sheets, match_codes, normalize_code, read_codes, sanitize_filename,
)

SHEET = "Lista Principal"
COLUMN = "Código Producto"


def db_frame(rows):
    """Arma un DataFrame con las columnas de la BD maestra: [(nombre, ruta, fecha), ...]."""
    return pd.DataFrame(rows, columns=["Nombre_Archivo", "Ruta_Completa", "Fecha_Modificacion"])


def count_pdf_pages(path):
    with open(path, "rb") as f:
        return len(re.findall(rb"/Type\s*/Page\b(?!s)", f.read()))


# =============================================================================
#  Pruebas unitarias (funciones puras)
# =============================================================================

class TestCodes(unittest.TestCase):
    def test_clean_codes_strips_and_drops_empty_values(self):
        raw = [" GF-1 ", "", "   ", "nan", "NaN", "GF-2", 1234]

        codes = clean_codes(raw)

        self.assertEqual(codes, ["GF-1", "GF-2", "1234"])

    def test_normalize_code_ignores_case_and_spaces(self):
        self.assertEqual(normalize_code("  GF-5657 "), "gf-5657")

    def test_match_codes_keeps_excel_order_and_reports_missing(self):
        index = {"gf-1": "a.jpg", "gf-2": "b.jpg"}

        matches, missing = match_codes(["GF-2", "X-9", "gf-1"], index)

        self.assertEqual(matches, [("GF-2", "b.jpg"), ("gf-1", "a.jpg")])
        self.assertEqual(missing, ["X-9"])


class TestBuildImageIndex(unittest.TestCase):
    def test_only_catalog_prefix_and_supported_extensions_are_indexed(self):
        df = db_frame([
            ("xzGF-1.jpg", "p1", "2020-01-01"),
            ("GF-2.jpg", "p2", "2020-01-01"),       # sin prefijo xz
            ("xzGF-3.webp", "p3", "2020-01-01"),    # extensión no soportada
            ("xzGF-4.PNG", "p4", "2020-01-01"),     # extensión en mayúsculas
            ("xz.jpg", "p5", "2020-01-01"),         # sin código tras el prefijo
        ])

        index = build_image_index(df)

        self.assertEqual(index, {"gf-1": "p1", "gf-4": "p4"})

    def test_most_recent_date_wins(self):
        df = db_frame([("xzA.jpg", "old", "2023-01-01 10:00:00"), ("xzA.jpg", "new", "2026-01-01 10:00:00")])

        self.assertEqual(build_image_index(df), {"a": "new"})

    def test_invalid_date_loses_against_valid_one(self):
        df = db_frame([("xzA.jpg", "bad", "basura"), ("xzA.jpg", "good", "2020-01-01")])

        self.assertEqual(build_image_index(df), {"a": "good"})

    def test_date_tie_keeps_first_row_of_file(self):
        df = db_frame([("xzA.jpg", "first", "2020-01-01"), ("xzA.jpg", "second", "2020-01-01")])

        self.assertEqual(build_image_index(df), {"a": "first"})


class TestSanitizeFilename(unittest.TestCase):
    def test_replaces_path_separators_and_invalid_chars(self):
        self.assertEqual(sanitize_filename('a<b>c:d"e/f\\g|h?i*j'), "a_b_c_d_e_f_g_h_i_j")

    def test_never_returns_empty_name(self):
        self.assertEqual(sanitize_filename("..."), "_")


class TestGuessCodeColumn(unittest.TestCase):
    def test_prefers_column_that_sounds_like_code(self):
        self.assertEqual(guess_code_column(["Precio", "Código Producto"]), "Código Producto")

    def test_falls_back_to_first_column_or_empty(self):
        self.assertEqual(guess_code_column(["Precio", "Stock"]), "Precio")
        self.assertEqual(guess_code_column([]), "")


# =============================================================================
#  Pruebas de integración (archivos reales en carpeta temporal)
# =============================================================================

class IntegrationTestCase(unittest.TestCase):
    """Crea el entorno de prueba una vez por clase y una carpeta de salida por prueba."""

    @classmethod
    def setUpClass(cls):
        cls.env_dir = setup_test_environment(tempfile.mkdtemp(prefix="unificador_env_"))
        cls.excel = os.path.join(cls.env_dir, "lista_codigos.xlsx")
        cls.db = os.path.join(cls.env_dir, "BASE_DE_DATOS_FOTOS_MASTER.db")

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.env_dir, ignore_errors=True)

    def setUp(self):
        self.out_dir = tempfile.mkdtemp(prefix="unificador_out_")
        self.addCleanup(shutil.rmtree, self.out_dir, ignore_errors=True)

    def write_excel(self, codes, name="codigos.xlsx"):
        path = os.path.join(self.out_dir, name)
        pd.DataFrame({"c": codes}).to_excel(path, index=False)
        return path


class TestReadingFiles(IntegrationTestCase):
    def test_index_from_db_file_applies_business_rules(self):
        index = index_images_from_db(self.db)

        self.assertEqual(sorted(index), ["gf-1111", "gf-2222", "gf-4444", "gf-5657"])
        self.assertIn("sub_sub_B", index["gf-2222"])  # la versión más reciente

    def test_lists_sheets_and_columns(self):
        self.assertEqual(list_sheets(self.excel), [SHEET, "Hoja Secundaria"])
        self.assertEqual(list_columns(self.excel, SHEET), [COLUMN])

    def test_read_codes_returns_cleaned_codes(self):
        codes = read_codes(self.excel, SHEET, COLUMN)

        self.assertEqual(codes, ["GF-2222", "GF-5657", "GF-9999", "GF-1111", "GF-3333", "GF-4444"])


class TestGeneratePdf(IntegrationTestCase):
    def test_generates_pdf_skipping_corrupt_image(self):
        output = os.path.join(self.out_dir, "resultado.pdf")
        messages = []

        pages, missing = generate_pdf(self.excel, SHEET, COLUMN, self.db, output,
                                      progress_callback=lambda cur, tot, msg: messages.append(msg))

        self.assertEqual(pages, 3)  # 4 coincidencias, xzGF-4444 está corrupta
        self.assertEqual(count_pdf_pages(output), pages)
        self.assertEqual(missing, ["GF-9999", "GF-3333"])
        self.assertTrue(any("Omitiendo" in m and "4444" in m for m in messages))
        self.assertEqual(messages[-1], "¡PDF generado con éxito!")

    def test_progress_reaches_total(self):
        output = os.path.join(self.out_dir, "resultado.pdf")
        progress = []

        generate_pdf(self.excel, SHEET, COLUMN, self.db, output,
                     progress_callback=lambda cur, tot, msg: progress.append((cur, tot)))

        current, total = progress[-1]
        self.assertEqual(current, total)

    def test_all_images_corrupt_raises_runtime_error(self):
        excel = self.write_excel(["GF-4444"])

        with self.assertRaises(RuntimeError):
            generate_pdf(excel, 0, "c", self.db, os.path.join(self.out_dir, "x.pdf"))


class TestErrors(IntegrationTestCase):
    def test_missing_excel_raises_value_error(self):
        with self.assertRaises(ValueError):
            generate_pdf(os.path.join(self.out_dir, "no_existe.xlsx"), SHEET, COLUMN, self.db,
                         os.path.join(self.out_dir, "x.pdf"))

    def test_missing_column_raises_clear_message(self):
        with self.assertRaisesRegex(ValueError, "COLUMNA_FALSA"):
            generate_pdf(self.excel, SHEET, "COLUMNA_FALSA", self.db, os.path.join(self.out_dir, "x.pdf"))

    def test_db_without_catalog_photos_raises_value_error(self):
        empty_db = os.path.join(self.out_dir, "vacia.db")
        db_frame([]).to_csv(empty_db, sep=";", index=False)

        with self.assertRaises(ValueError):
            generate_pdf(self.excel, SHEET, COLUMN, empty_db, os.path.join(self.out_dir, "x.pdf"))

    def test_db_missing_required_columns_raises_value_error(self):
        bad_db = os.path.join(self.out_dir, "mala.db")
        pd.DataFrame({"Otra": ["x"]}).to_csv(bad_db, sep=";", index=False)

        with self.assertRaisesRegex(ValueError, "columnas"):
            index_images_from_db(bad_db)

    def test_nonexistent_db_raises_value_error(self):
        with self.assertRaises(ValueError):
            index_images_from_db(os.path.join(self.out_dir, "no_existe.db"))


class TestExportImages(IntegrationTestCase):
    def test_exports_one_file_per_found_code(self):
        folder = os.path.join(self.out_dir, "export")

        exported, missing = export_images(self.excel, SHEET, COLUMN, self.db, folder)

        self.assertEqual(exported, 4)
        self.assertEqual(sorted(os.listdir(folder)), ["GF-1111.png", "GF-2222.jpg", "GF-4444.jpg", "GF-5657.jpg"])
        self.assertEqual(missing, ["GF-9999", "GF-3333"])

    def test_duplicate_codes_are_copied_once(self):
        excel = self.write_excel(["GF-1111", "gf-1111", " GF-1111 "])
        folder = os.path.join(self.out_dir, "export")

        exported, _ = export_images(excel, 0, "c", self.db, folder)

        self.assertEqual(exported, 1)
        self.assertEqual(os.listdir(folder), ["GF-1111.png"])

    def test_code_with_path_separators_stays_inside_destination(self):
        evil = "../escape"
        rows = pd.read_csv(self.db, sep=";", dtype=str, encoding="utf-8-sig")
        rows.loc[len(rows)] = ["'xz" + evil, rows.Ruta_Completa[0], "xz" + evil + ".jpg",
                               ".jpg", "2026-01-01 00:00:00", "x"]
        db_evil = os.path.join(self.out_dir, "evil.db")
        rows.to_csv(db_evil, sep=";", index=False, encoding="utf-8-sig")
        excel = self.write_excel([evil])
        folder = os.path.join(self.out_dir, "dest", "sub")

        export_images(excel, 0, "c", db_evil, folder)

        self.assertFalse(os.path.exists(os.path.join(self.out_dir, "dest", "escape.jpg")))
        self.assertEqual(os.listdir(folder), ["_escape.jpg"])


if __name__ == "__main__":
    # Forzar UTF-8 en la consola de Windows para evitar errores con caracteres especiales
    if sys.stdout and hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    unittest.main(verbosity=2)
