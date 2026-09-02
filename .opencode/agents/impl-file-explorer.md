---
description: Implementa FileExplorer: dialogos Windows, modal reemplazo
mode: subagent
model: opencode/nemotron-3.5-lightning-free
temperature: 0.2
permission:
  edit: allow
  bash: allow
---

# Implementador FileExplorer

Implementas `src/rpa_excel_ui_automation/file_explorer.py` basado en `spec.json` y `rpa-excel-automation` skill.

## ENTRADA
- `TaskSpec` con `context.spec`, `acceptance_criteria` (TC02: inyectar ruta, click Guardar, manejar modal reemplazo)

## SALIDA
- `TaskResult` con `files_changed`: ["src/rpa_excel_ui_automation/file_explorer.py"]

## REQUISITOS OBLIGATORIOS

```python
# src/rpa_excel_ui_automation/file_explorer.py
from pathlib import Path
import logging
import uiautomation as auto

logger = logging.getLogger(__name__)

class FileExplorer:
    """Interaccion exclusiva con dialogos nativos Windows (Explorador de archivos)."""
    
    def __init__(self) -> None:
        self.current_dialog: auto.WindowControl | None = None
        logger.info("FileExplorer inicializado")
    
    def _get_file_edit_control(self, dialog: auto.WindowControl) -> auto.EditControl | None:
        """Obtiene control Edit 'Nombre de archivo:' directamente."""
        try:
            edit = dialog.EditControl(searchDepth=10, Name="Nombre de archivo:")
            if edit.Exists(3, 0.2):
                return edit
            # Fallback: buscar por control_type
            edit = dialog.Control(searchDepth=10, ControlType=auto.ControlType.EditControl)
            if edit.Exists(3, 0.2):
                return edit
        except Exception as e:
            logger.debug("Error buscando control Edit: %s", e)
        return None
    
    def _get_button(self, dialog: auto.WindowControl, name: str) -> auto.ButtonControl | None:
        """Obtiene boton por nombre directamente."""
        try:
            btn = dialog.ButtonControl(searchDepth=10, Name=name)
            if btn.Exists(3, 0.2):
                return btn
        except Exception as e:
            logger.debug("Error buscando boton %s: %s", name, e)
        return None
    
    def open_file_dialog(self, file_path: Path) -> bool:
        """Inyecta ruta en dialogo 'Abrir' y click 'Abrir'."""
        logger.info("Iniciando open_file_dialog(%s)", file_path)
        try:
            # Detectar dialogo activo "Abrir"
            self.current_dialog = auto.WindowControl(searchDepth=1, Name="Abrir")
            if not self.current_dialog.WaitForExist(5, 0.5):
                logger.error("Dialogo 'Abrir' no encontrado")
                return False
            
            logger.debug("Dialogo 'Abrir' detectado: %s", self.current_dialog.Name)
            
            # Inyectar ruta directamente en control Edit
            edit = self._get_file_edit_control(self.current_dialog)
            if not edit:
                logger.error("Control 'Nombre de archivo:' no encontrado")
                return False
            
            absolute_path = file_path.resolve()
            logger.debug("Inyectando ruta: %s", absolute_path)
            edit.SetValue(str(absolute_path))
            
            # Click boton "Abrir" directamente
            open_btn = self._get_button(self.current_dialog, "Abrir")
            if not open_btn:
                logger.error("Boton 'Abrir' no encontrado")
                return False
            
            open_btn.Click()
            logger.info("Archivo abierto correctamente: %s", absolute_path)
            return True
            
        except Exception as e:
            logger.exception("Error en open_file_dialog: %s", e)
            return False
    
    def save_file_dialog(self, file_path: Path) -> bool:
        """Inyecta ruta en dialogo 'Guardar como', click 'Guardar', maneja reemplazo."""
        logger.info("Iniciando save_file_dialog(%s)", file_path)
        try:
            # Detectar dialogo activo "Guardar como"
            self.current_dialog = auto.WindowControl(searchDepth=1, Name="Guardar como")
            if not self.current_dialog.WaitForExist(5, 0.5):
                logger.error("Dialogo 'Guardar como' no encontrado")
                return False
            
            logger.debug("Dialogo 'Guardar como' detectado: %s", self.current_dialog.Name)
            
            # Inyectar ruta destino
            edit = self._get_file_edit_control(self.current_dialog)
            if not edit:
                logger.error("Control 'Nombre de archivo:' no encontrado en Guardar como")
                return False
            
            absolute_path = file_path.resolve()
            logger.debug("Inyectando ruta destino: %s", absolute_path)
            edit.SetValue(str(absolute_path))
            
            # Click boton "Guardar" directamente
            save_btn = self._get_button(self.current_dialog, "Guardar")
            if not save_btn:
                logger.error("Boton 'Guardar' no encontrado")
                return False
            
            save_btn.Click()
            
            # MANEJO MODAL REEMPLAZO: Detectar y click "Si"
            if self.handle_replace_modal():
                logger.info("Modal reemplazo manejado correctamente")
            else:
                logger.debug("No hubo modal de reemplazo (archivo nuevo)")
            
            logger.info("Archivo guardado correctamente: %s", absolute_path)
            return True
            
        except Exception as e:
            logger.exception("Error en save_file_dialog: %s", e)
            return False
    
    def handle_replace_modal(self) -> bool:
        """Detecta modal 'Confirmar guardado' y click 'Si' para reemplazar."""
        logger.debug("Verificando modal de reemplazo...")
        try:
            # Buscar modal con titulo tipico
            replace_modal = auto.WindowControl(searchDepth=1, Name="Confirmar guardado")
            if not replace_modal.Exists(3, 0.2):
                # Fallback: buscar por patron
                replace_modal = auto.WindowControl(searchDepth=1, NameRe=".*ya existe.*|.*reemplazar.*|.*Confirmar.*", searchFromControl=auto.GetRootControl())
                if not replace_modal.Exists(3, 0.2):
                    return False
            
            logger.info("Modal de reemplazo detectado: %s", replace_modal.Name)
            
            # Click boton "Si" directamente
            yes_btn = replace_modal.ButtonControl(searchDepth=10, Name="Si")
            if not yes_btn.Exists(3, 0.2):
                yes_btn = replace_modal.ButtonControl(searchDepth=10, Name="S&iacute;")  # Con acento HTML
            if not yes_btn.Exists(3, 0.2):
                yes_btn = replace_modal.ButtonControl(searchDepth=10, Name="&Yes")  # Ingles
            
            if not yes_btn:
                logger.error("Boton 'Si' no encontrado en modal reemplazo")
                return False
            
            yes_btn.Click()
            logger.info("Confirmacion de reemplazo ejecutada")
            return True
            
        except Exception as e:
            logger.debug("Error/No modal reemplazo: %s", e)
            return False
```

## REGLAS CRITICAS
- **NO** `time.sleep()` -> `.WaitForExist()`, `.Exists(timeout, interval)`
- **NO** `Tab` -> selectores directos `EditControl(Name="...")`, `ButtonControl(Name="...")`
- **SOLO** `pathlib.Path` -> `.resolve()` antes de inyectar
- **LOGGING** en cada metodo: inicio, exito, fallo, deteccion modal
- `SetValue()` para inyectar texto (no SendKeys)
- Manejar variaciones de idioma: "Si"/"Sí"/"Yes", "Guardar como"/"Save As"