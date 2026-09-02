---
name: rpa-excel-automation
description: Patrones UI Automation para Excel: selectores robustos, manejo modales, sincronizacion sin sleep, pathlib
license: MIT
compatibility: opencode
metadata:
  audience: developers
  domain: rpa
---

## Qué cubre

- **UI Automation (uiautomation)**: `autoit`, `uia_client`, `Application`, `Window`, `Control`
- **Selectores robustos**: `child_window(control_type="Edit", title="Nombre de archivo:")`, `by_id`, `by_automation_id`
- **Sincronización nativa**: `.exists(timeout=10)`, `.wait(wait_for="ready", timeout=15)`, `.wait_for_input_idle()`
- **Manejo de modales**: Detección ventana "Confirmar guardado" -> `.exists()` -> click botón "Sí" directo
- **Anti-patrones PROHIBIDOS**: `time.sleep()`, `send_keys("{TAB}")`, clics por coordenadas X/Y
- **Rutas**: `pathlib.Path` obligatorio, `.resolve()`, `.absolute()`, `/` operator

## Cuándo usar

- Implementar `ExcelManager` y `FileExplorer`
- Cualquier interacción con ventanas nativas Windows
- Tests que validen robustez UI

## Patrones clave

```python
# BUENO: Selector directo + wait nativo
dlg = app.window(title="Abrir")
dlg.child_window(control_type="Edit", title="Nombre de archivo:").wait("visible", timeout=10)
dlg.child_window(control_type="Edit", title="Nombre de archivo:").set_edit_text(str(path))
dlg.child_window(control_type="Button", title="Abrir").click()

# BUENO: Modal reemplazo
if replace_dlg.exists(timeout=5):
    replace_dlg.child_window(control_type="Button", title="Sí").click()

# MALO: time.sleep(2) -> FRAGIL
# MALO: send_keys("{TAB 3}") -> FRAGIL
```