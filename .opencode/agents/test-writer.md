---
description: Escribe tests pytest para TC01, TC02, fixtures, mocking
mode: subagent
model: opencode/mimo-v2.5-free
temperature: 0.3
permission:
  edit: allow
  bash: allow
---

# Test Writer - Tests Pytest

Escribes tests en `tests/test_excel_manager.py` y `tests/test_file_explorer.py` basados en `spec.json` y README.

## ENTRADA
- `TaskSpec` con `context.spec`, `acceptance_criteria` (TC01, TC02 del README)

## SALIDA
- `TaskResult` con `files_changed`: ["tests/test_excel_manager.py", "tests/test_file_explorer.py"]

## REQUISITOS

```python
# tests/test_excel_manager.py
import pytest
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock
from rpa_excel_ui_automation.excel_manager import ExcelManager

class TestExcelManager:
    """Tests para ExcelManager - Caso de Prueba 01."""
    
    @pytest.fixture
    def excel_manager(self):
        return ExcelManager()
    
    @pytest.fixture
    def mock_uia(self):
        with patch('rpa_excel_ui_automation.excel_manager.auto') as mock:
            yield mock
    
    def test_open_file_success(self, excel_manager, mock_uia):
        """TC01: ExcelManager.open_file() abre Excel y dialogo Abrir."""
        # Setup mocks
        mock_app = MagicMock()
        mock_app.Exists.return_value = True
        mock_app.WaitForExist.return_value = True
        mock_app.Name = "Excel"
        mock_uia.WindowControl.return_value = mock_app
        
        mock_dlg = MagicMock()
        mock_dlg.WaitForExist.return_value = True
        # Segundo WindowControl call para dialogo
        mock_uia.WindowControl.side_effect = [mock_app, mock_dlg]
        
        # Execute
        result = excel_manager.open_file()
        
        # Assert
        assert result is True
        mock_app.SendKeys.assert_called_with("{Ctrl}o", waitTime=0.5)
        mock_dlg.WaitForExist.assert_called()
    
    def test_open_file_excel_not_running_launches(self, excel_manager, mock_uia):
        """TC01: Si Excel no corre, lo lanza."""
        mock_app = MagicMock()
        mock_app.Exists.return_value = False
        mock_app.WaitForExist.return_value = True
        mock_uia.WindowControl.return_value = mock_app
        
        mock_dlg = MagicMock()
        mock_dlg.WaitForExist.return_value = True
        mock_uia.WindowControl.side_effect = [mock_app, mock_dlg]
        
        with patch('subprocess.Popen') as mock_popen:
            result = excel_manager.open_file()
            mock_popen.assert_called_once_with(["excel.exe"])
            assert result is True
    
    def test_open_file_dialog_timeout_fails(self, excel_manager, mock_uia):
        """TC01: Fallo si dialogo no aparece."""
        mock_app = MagicMock()
        mock_app.Exists.return_value = True
        mock_uia.WindowControl.return_value = mock_app
        
        mock_dlg = MagicMock()
        mock_dlg.WaitForExist.return_value = False
        mock_uia.WindowControl.side_effect = [mock_app, mock_dlg]
        
        result = excel_manager.open_file()
        assert result is False
    
    def test_save_as_success(self, excel_manager, mock_uia):
        """TC02: ExcelManager.save_as() abre dialogo Guardar como."""
        mock_app = MagicMock()
        mock_app.Exists.return_value = True
        excel_manager.app = mock_app
        
        mock_dlg = MagicMock()
        mock_dlg.WaitForExist.return_value = True
        mock_uia.WindowControl.return_value = mock_dlg
        
        result = excel_manager.save_as()
        
        assert result is True
        mock_app.SendKeys.assert_called_with("{F12}", waitTime=0.5)
        mock_dlg.WaitForExist.assert_called()
    
    def test_save_as_no_active_excel_fails(self, excel_manager):
        """TC02: Fallo si no hay Excel activo."""
        excel_manager.app = None
        result = excel_manager.save_as()
        assert result is False


# tests/test_file_explorer.py
import pytest
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock
from rpa_excel_ui_automation.file_explorer import FileExplorer

class TestFileExplorer:
    """Tests para FileExplorer - Caso de Prueba 01 y 02."""
    
    @pytest.fixture
    def file_explorer(self):
        return FileExplorer()
    
    @pytest.fixture
    def mock_uia(self):
        with patch('rpa_excel_ui_automation.file_explorer.auto') as mock:
            yield mock
    
    def test_open_file_dialog_success(self, file_explorer, mock_uia):
        """TC01: FileExplorer inyecta ruta y click Abrir."""
        mock_dlg = MagicMock()
        mock_dlg.WaitForExist.return_value = True
        mock_uia.WindowControl.return_value = mock_dlg
        
        mock_edit = MagicMock()
        mock_edit.Exists.return_value = True
        mock_dlg.EditControl.return_value = mock_edit
        
        mock_btn = MagicMock()
        mock_btn.Exists.return_value = True
        mock_dlg.ButtonControl.return_value = mock_btn
        
        test_path = Path(".data/input/origen.xlsx")
        result = file_explorer.open_file_dialog(test_path)
        
        assert result is True
        mock_edit.SetValue.assert_called_once_with(str(test_path.resolve()))
        mock_btn.Click.assert_called_once()
    
    def test_open_file_dialog_edit_not_found_fails(self, file_explorer, mock_uia):
        """TC01: Fallo si control Edit no existe."""
        mock_dlg = MagicMock()
        mock_dlg.WaitForExist.return_value = True
        mock_uia.WindowControl.return_value = mock_dlg
        
        mock_edit = MagicMock()
        mock_edit.Exists.return_value = False
        mock_dlg.EditControl.return_value = mock_edit
        
        result = file_explorer.open_file_dialog(Path("test.xlsx"))
        assert result is False
    
    def test_save_file_dialog_success(self, file_explorer, mock_uia):
        """TC02: FileExplorer inyecta ruta, click Guardar, sin modal."""
        mock_dlg = MagicMock()
        mock_dlg.WaitForExist.return_value = True
        mock_uia.WindowControl.return_value = mock_dlg
        
        mock_edit = MagicMock()
        mock_edit.Exists.return_value = True
        mock_dlg.EditControl.return_value = mock_edit
        
        mock_btn = MagicMock()
        mock_btn.Exists.return_value = True
        mock_dlg.ButtonControl.return_value = mock_btn
        
        # Mock handle_replace_modal returns False (no modal)
        with patch.object(file_explorer, 'handle_replace_modal', return_value=False):
            result = file_explorer.save_file_dialog(Path(".data/output/destino.xlsx"))
        
        assert result is True
        mock_edit.SetValue.assert_called_once()
        mock_btn.Click.assert_called_once()
    
    def test_save_file_dialog_with_replace_modal(self, file_explorer, mock_uia):
        """TC02: FileExplorer maneja modal reemplazo (click Si)."""
        mock_dlg = MagicMock()
        mock_dlg.WaitForExist.return_value = True
        mock_uia.WindowControl.return_value = mock_dlg
        
        mock_edit = MagicMock()
        mock_edit.Exists.return_value = True
        mock_dlg.EditControl.return_value = mock_edit
        
        mock_btn = MagicMock()
        mock_btn.Exists.return_value = True
        mock_dlg.ButtonControl.return_value = mock_btn
        
        with patch.object(file_explorer, 'handle_replace_modal', return_value=True) as mock_handle:
            result = file_explorer.save_file_dialog(Path(".data/output/destino.xlsx"))
        
        assert result is True
        mock_handle.assert_called_once()
    
    def test_handle_replace_modal_detects_and_clicks(self, file_explorer, mock_uia):
        """TC02: handle_replace_modal detecta modal y click Si."""
        mock_modal = MagicMock()
        mock_modal.Exists.return_value = True
        mock_modal.Name = "Confirmar guardado"
        mock_uia.WindowControl.return_value = mock_modal
        
        mock_yes = MagicMock()
        mock_yes.Exists.return_value = True
        mock_modal.ButtonControl.return_value = mock_yes
        
        result = file_explorer.handle_replace_modal()
        
        assert result is True
        mock_yes.Click.assert_called_once()
    
    def test_handle_replace_modal_not_present_returns_false(self, file_explorer, mock_uia):
        """TC02: Sin modal retorna False."""
        mock_modal = MagicMock()
        mock_modal.Exists.return_value = False
        mock_uia.WindowControl.return_value = mock_modal
        
        result = file_explorer.handle_replace_modal()
        assert result is False
```

## REGLAS
- Coverage objetivo: >=90%
- Mock `uiautomation` completamente (no tests de integracion real)
- Tests parametrizados para variaciones (idiomas, rutas)
- Fixtures reutilizables
- Nombres descriptivos: `test_<metodo>_<escenario>_<resultado>`
- AAA pattern: Arrange, Act, Assert