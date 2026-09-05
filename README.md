# RPA Excel UI Automation

Automatización RPA de la interfaz de usuario de Microsoft Excel utilizando **UI Automation** (`uiautomation`). Abre archivos, guarda como y maneja modales de reemplazo sin intervención humana, siguiendo principios de Clean Code y anti-fragilidad.

## Requisitos

- **Python** 3.12+
- **Microsoft Excel** instalado (Office 365 / 2016+)
- **PDM** (gestor de paquetes)

## Instalación

```bash
git clone https://github.com/AngBan2x/rpa-excel-ui-automation.git
cd rpa-excel-ui-automation
pdm install
```

## Uso rápido

```bash
# TC01: Abrir archivo existente via Backstage
pdm run python tc01.py

# TC02: Guardar como con reemplazo
pdm run python tc02.py

# Ambos
pdm run python tc_all.py
```

## Arquitectura

| Clase | Responsabilidad |
|-------|----------------|
| `ExcelManager` | Gestiona la instancia de Excel (abrir, guardar como, cerrar) |
| `FileExplorer` | Interactúa con diálogos de Windows (Abrir, Guardar como, modal de reemplazo) |

### Flujo TC01 — Abrir archivo existente

```
ExcelManager.open_file(use_dialog=True)
  → subprocess.Popen([excel])           # Lanza Excel vacío
  → Backstage "Abrir" → "Examinar"      # Navega al Backstage de Office
  → #32770 "Abrir"                       # Diálogo nativo de Windows
  → FileExplorer.open_file_dialog()      # Inyecta ruta + Enter en Edit
```

### Flujo TC02 — Guardar como con reemplazo

```
ExcelManager.save_as()
  → F12                                  # Abre "Guardar como"
  → FileExplorer.save_file_dialog()      # Inyecta ruta + Click "Guardar"
  → handle_replace_modal()              # Detecta "Confirmar Guardar como" → Click "Sí"
```

## Casos de prueba

### TC01: Abrir archivo existente

1. Lanza Excel vacío (sin archivo)
2. Navega al Backstage → selecciona "Abrir" → clickea "Examinar"
3. Inyecta la ruta de `origen.xlsx` en el diálogo #32770 nativo de Windows
4. Verifica que el archivo se abrió correctamente (título de ventana)
5. Cierra Excel con `{Alt}{F4}`

### TC02: Guardar como con reemplazo

1. Crea `destino.xlsx` previo con contenido basura (38 bytes) para forzar reemplazo
2. Abre `origen.xlsx` en Excel
3. Presiona F12 → inyecta la ruta `destino.xlsx` en el diálogo "Guardar como"
4. Detecta y confirma el modal "Confirmar Guardar como" (reemplazo)
5. Verifica que `destino.xlsx` tiene contenido real (~8 KB)

## Ejecución de tests

```bash
# Unit tests (mocked, no requiere Excel)
pdm run pytest tests/ -v

# Tests de integracion extras (requiere Excel real)
pdm run pytest -m integration tests/test_integration_extras.py -v

# Coverage
pdm run pytest -v --cov=src --cov-report=term-missing
```

## Criterios técnicos

| Criterio | Implementación |
|----------|----------------|
| Sin `time.sleep()` | Sincronización vía `auto.WaitForExist()` y polling con `time.monotonic()` |
| Sin navegación con `Tab` | Interacción directa por selectores UI Automation (AutoId, Name, ClassName) |
| Uso exclusivo de `pathlib` | Todas las rutas son `Path`, sin `os.path` |
| Control de modales | `handle_replace_modal()` con búsqueda dual (Name + ClassName) y polling |
| Logging estructurado | Logs en `ExcelManager` y `FileExplorer` para cada acción |
| Context manager | `ExcelManager` implementa `__enter__`/`__exit__` para cierre automático |
| Anti-fragilidad | Selectores UIA robustos, tolerancia a COM errors, fallback con `taskkill` |

## Estructura del proyecto

```
excel-ui/
├── src/rpa_excel_ui_automation/
│   ├── __init__.py
│   ├── excel_manager.py          # ExcelManager
│   └── file_explorer.py          # FileExplorer
├── tests/
│   ├── conftest.py
│   ├── test_excel_manager.py     # Unit tests (mocked)
│   ├── test_file_explorer.py     # Unit tests (mocked)
│   └── test_integration_extras.py  # Integration extras (marker: integration)
├── .data/
│   ├── input/origen.xlsx
│   └── output/
├── tc01.py                       # Runner TC01
├── tc02.py                       # Runner TC02
├── tc_all.py                     # Runner TC01 + TC02
├── test_integration_real.py      # Implementación TC01 + TC02
├── pyproject.toml
└── README.md
```

## Tecnologías

- **uiautomation** — Selectores UIA nativos de Windows
- **openpyxl** — Verificación de contenido de archivos Excel
- **pytest** — Framework de testing
- **PDM** — Gestión de dependencias y scripts
- **pathlib** — Manejo de rutas

## License

MIT
