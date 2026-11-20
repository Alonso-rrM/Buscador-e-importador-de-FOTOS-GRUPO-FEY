import csv

path = r"\\feyserver\GRUPOS\ANALISIS COMERCIAL\VENTAS\CONTENEDORES 2025\BASE_DE_DATOS_FOTOS_MASTER.db"

def count_columns(file_path):
    counts = {}
    total_lines = 0
    with open(file_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f, delimiter=';')
        for i, row in enumerate(reader):
            total_lines += 1
            num_cols = len(row)
            counts[num_cols] = counts.get(num_cols, 0) + 1
            if i < 5:
                print(f"Linea {i+1} tiene {num_cols} columnas: {row}")

    print("\nResumen de columnas por linea:")
    for cols, count in sorted(counts.items()):
        print(f"  {count} lineas tienen exactamente {cols} columnas.")
    print(f"Total de lineas leidas: {total_lines}")

if __name__ == "__main__":
    import sys
    # Uso: python check_csv.py [ruta_de_la_bd]  (por defecto, la ruta de red)
    target = sys.argv[1] if len(sys.argv) > 1 else path
    try:
        count_columns(target)
    except (OSError, csv.Error, UnicodeDecodeError) as e:
        print("Error:", e)
