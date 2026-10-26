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

try:
    count_columns(path)
except Exception as e:
    print("Error:", e)
