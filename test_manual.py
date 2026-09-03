"""Prueba manual del flujo completo RPA Excel."""
from pathlib import Path
from rpa_excel_ui_automation import ExcelManager

input_file = Path(".data/input/origen.xlsx")
output_file = Path(".data/output/destino.xlsx")

output_file.parent.mkdir(parents=True, exist_ok=True)

print("=== PRUEBA MANUAL RPA EXCEL ===\n")

with ExcelManager() as excel:
    print("1. Abriendo archivo...")
    if excel.open_file(input_file):
        print("   [OK] Archivo abierto correctamente\n")
        
        print("2. Guardando como...")
        if excel.save_as(output_file):
            print("   [OK] Archivo guardado correctamente\n")
        else:
            print("   [FAIL] Error al guardar\n")
    else:
        print("   [FAIL] Error al abrir\n")
    
    # Esperar antes de cerrar
    import time
    time.sleep(2)

    print("3. Cerrando Excel...")
    excel.close()
    print("   [OK] Excel cerrado\n")

print("=== PRUEBA COMPLETADA ===")