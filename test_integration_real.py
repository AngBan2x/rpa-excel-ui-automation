"""Pruebas de integracion TC01 y TC02 con Excel real.

Verifica:
1. TC01: Abrir archivo existente via dialogo (Backstage -> Examinar -> #32770)
2. TC02: Guardar como con reemplazo de archivo existente
"""
from __future__ import annotations

import subprocess
from pathlib import Path

import uiautomation as auto

from rpa_excel_ui_automation import ExcelManager

INPUT_FILE = Path(".data/input/origen.xlsx")
OUTPUT_FILE = Path(".data/output/destino.xlsx")
OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)


def cleanup_output_files() -> None:
    """Elimina archivos de salida de ejecuciones anteriores."""
    output_dir = OUTPUT_FILE.parent
    if output_dir.exists():
        for f in output_dir.glob("*.xlsx"):
            f.unlink(missing_ok=True)


def kill_excel() -> None:
    """Mata todos los procesos Excel y espera a que se cierre."""
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "EXCEL.EXE"],
            capture_output=True,
            check=False,
            timeout=10,
        )
    except (subprocess.TimeoutExpired, FileNotFoundError):
        pass
    # Esperar a que Excel se cierre completamente
    # NOTA: despues de taskkill, las referencias COM pueden quedar invalidas,
    # asi que envolvemos en try/except.
    import time
    time.sleep(1)
    try:
        excel_window = auto.WindowControl(searchDepth=1, ClassName="XLMAIN")
        for _ in range(10):
            if not excel_window.Exists(0, 0):
                break
            time.sleep(0.5)
    except Exception:
        pass


def test_tc01_open_existing_file() -> bool:
    """TC01: Abrir archivo existente via dialogo."""
    print("\n=== TC01: Abrir archivo existente ===")
    kill_excel()

    if not INPUT_FILE.exists():
        print(f"  [FAIL] Archivo de prueba no existe: {INPUT_FILE}")
        return False

    with ExcelManager() as excel:
        result = excel.open_file(INPUT_FILE, use_dialog=True)
        if result:
            print("  [PASS] Archivo abierto correctamente")
            print(f"  Ventana: {excel.app.Name if excel.app else 'N/A'}")
        else:
            print("  [FAIL] No se pudo abrir el archivo")
            return False

    # Verificar que Excel se cerro
    check = auto.WindowControl(searchDepth=1, ClassName="XLMAIN")
    if not check.Exists(2, 0.5):
        print("  [PASS] Excel se cerro correctamente")
        return True
    else:
        print("  [FAIL] Excel no se cerro")
        return False


def test_tc02_save_as_with_replace() -> bool:
    """TC02: Guardar como con reemplazo de archivo existente."""
    print("\n=== TC02: Guardar como con reemplazo ===")
    kill_excel()

    if not INPUT_FILE.exists():
        print(f"  [FAIL] Archivo de prueba no existe: {INPUT_FILE}")
        return False

    # Crear archivo destino previo para probar reemplazo
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_FILE.write_bytes(b"contenido previo para probar reemplazo")
    print(f"  [INFO] Archivo previo creado: {OUTPUT_FILE} ({OUTPUT_FILE.stat().st_size} bytes)")

    with ExcelManager() as excel:
        # Abrir archivo
        if not excel.open_file(INPUT_FILE):
            print("  [FAIL] No se pudo abrir el archivo de entrada")
            return False

        # Guardar como
        result = excel.save_as(OUTPUT_FILE)
        if result:
            print("  [PASS] Archivo guardado como exitosamente")
        else:
            print("  [FAIL] No se pudo guardar como")
            return False

    # Verificar que el archivo se creo (deberia ser mas grande que el previo)
    if OUTPUT_FILE.exists():
        size = OUTPUT_FILE.stat().st_size
        print(f"  [PASS] Archivo destino creado: {OUTPUT_FILE} ({size} bytes)")
        return True
    else:
        print(f"  [FAIL] Archivo destino no se creo: {OUTPUT_FILE}")
        return False


def main() -> None:
    """Ejecutar TC01 y TC02 de integracion."""
    print("=" * 60)
    print("PRUEBAS DE INTEGRACION: TC01 + TC02")
    print("=" * 60)

    if not INPUT_FILE.exists():
        print(f"\n[FATAL] Archivo de prueba no existe: {INPUT_FILE}")
        print("Ejecuta: pdm run python create_test_data.py")
        return

    print("\n[INFO] Limpiando archivos de salida anteriores...")
    cleanup_output_files()
    print("[INFO] Limpieza completada")

    results = []
    results.append(("TC01: Abrir archivo", test_tc01_open_existing_file()))
    results.append(("TC02: Guardar como", test_tc02_save_as_with_replace()))

    print("\n" + "=" * 60)
    print("RESUMEN DE PRUEBAS")
    print("=" * 60)

    passed = sum(1 for _, result in results if result)
    total = len(results)

    for name, result in results:
        status = "[PASS]" if result else "[FAIL]"
        print(f"  {status} {name}")

    print(f"\nTotal: {passed}/{total} pruebas pasaron")

    kill_excel()

    if passed == total:
        print("\nTODAS LAS PRUEBAS PASARON!")
    else:
        print(f"\n{total - passed} prueba(s) fallaron")


if __name__ == "__main__":
    main()
