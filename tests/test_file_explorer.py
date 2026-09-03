"""Tests para FileExplorer - Caso de Prueba 01 y 02: Dialogos de archivos Windows."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Generator
from unittest.mock import MagicMock, patch

import pytest

from rpa_excel_ui_automation.file_explorer import FileExplorer


class TestFileExplorerInit:
    """Tests para inicializacion de FileExplorer."""

    def test_init_with_default_logger(self) -> None:
        """FileExplorer inicializa con logger por defecto."""
        # Arrange & Act
        explorer = FileExplorer()

        # Assert
        assert explorer._logger is not None

    def test_init_with_custom_logger(self, mock_logger: MagicMock) -> None:
        """FileExplorer inicializa con logger personalizado."""
        # Arrange & Act
        explorer = FileExplorer(logger=mock_logger)

        # Assert
        assert explorer._logger is mock_logger
        mock_logger.info.assert_called_with("FileExplorer inicializado")


class TestOpenFileDialog:
    """Tests para open_file_dialog()."""

    def test_open_file_dialog_success(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
        mock_open_dialog: MagicMock,
        mock_file_edit: MagicMock,
        mock_button: MagicMock,
        sample_input_path: Path,
    ) -> None:
        """TC01: FileExplorer inyecta ruta y click Abrir."""
        # Arrange
        mock_file_edit.Exists.return_value = True
        mock_button.Exists.return_value = True
        mock_open_dialog.EditControl.return_value = mock_file_edit
        mock_open_dialog.ButtonControl.return_value = mock_button
        # auto.WaitForExist retorna False = dialogo se cerro (exito)
        mock_uia_file_explorer.WaitForExist.return_value = False

        # Mock _find_file_dialog para que devuelva el dialogo mock
        with patch.object(file_explorer, '_find_file_dialog', return_value=mock_open_dialog):
            # Act
            result = file_explorer.open_file_dialog(sample_input_path)

        # Assert
        assert result is True
        mock_file_edit.Click.assert_called_once()
        mock_file_edit.SendKeys.assert_any_call("{Ctrl}a", waitTime=0.1)
        mock_file_edit.SendKeys.assert_any_call(
            str(sample_input_path.resolve()), waitTime=0.1
        )
        assert mock_file_edit.SendKeys.call_count == 2
        mock_button.Click.assert_called_once()

    def test_open_file_dialog_edit_not_found(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
        mock_open_dialog: MagicMock,
    ) -> None:
        """TC01: Fallo si control Edit no existe."""
        # Arrange
        mock_uia_file_explorer.WindowControl.return_value = mock_open_dialog
        mock_edit = MagicMock()
        mock_edit.Exists.return_value = False
        mock_open_dialog.EditControl.return_value = mock_edit

        # Act
        result = file_explorer.open_file_dialog(Path("test.xlsx"))

        # Assert
        assert result is False

    def test_open_file_dialog_button_not_found(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
        mock_open_dialog: MagicMock,
        mock_file_edit: MagicMock,
    ) -> None:
        """TC01: Fallo si boton Abrir no existe."""
        # Arrange
        mock_uia_file_explorer.WindowControl.return_value = mock_open_dialog
        mock_open_dialog.EditControl.return_value = mock_file_edit
        mock_button = MagicMock()
        mock_button.Exists.return_value = False
        mock_open_dialog.ButtonControl.return_value = mock_button

        # Act
        result = file_explorer.open_file_dialog(Path("test.xlsx"))

        # Assert
        assert result is False

    def test_open_file_dialog_not_appears(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
    ) -> None:
        """TC01: Fallo si dialogo Abrir no aparece."""
        # Arrange - _find_file_dialog retorna None (no encontro dialogo)
        with patch.object(file_explorer, '_find_file_dialog', return_value=None):
            # Act
            result = file_explorer.open_file_dialog(Path("test.xlsx"))

        # Assert
        assert result is False

    def test_open_file_dialog_dialog_not_closes(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
        mock_open_dialog: MagicMock,
        mock_file_edit: MagicMock,
        mock_button: MagicMock,
    ) -> None:
        """TC01: Fallo si dialogo no se cierra tras click."""
        # Arrange
        mock_uia_file_explorer.WindowControl.return_value = mock_open_dialog
        mock_open_dialog.EditControl.return_value = mock_file_edit
        mock_open_dialog.ButtonControl.return_value = mock_button
        # Dialogo sigue abierto despues del click
        mock_uia_file_explorer.WaitForExist.side_effect = [True, True]

        # Act
        result = file_explorer.open_file_dialog(Path("test.xlsx"))

        # Assert
        assert result is False

    def test_open_file_dialog_exception_handling(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
    ) -> None:
        """TC01: Manejo de excepciones inesperadas."""
        # Arrange
        mock_uia_file_explorer.WindowControl.side_effect = Exception("Unexpected")

        # Act
        result = file_explorer.open_file_dialog(Path("test.xlsx"))

        # Assert
        assert result is False

    @pytest.mark.parametrize(
        "path,expected",
        [
            (Path("file.xlsx"), True),
            (Path("data/report.xlsx"), True),
            (Path("C:/absolute/path/file.xlsx"), True),
        ],
    )
    def test_open_file_dialog_various_paths(
        self,
        path: Path,
        expected: bool,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
        mock_open_dialog: MagicMock,
        mock_file_edit: MagicMock,
        mock_button: MagicMock,
    ) -> None:
        """TC01: Diferentes rutas funcionan correctamente."""
        # Arrange
        mock_file_edit.Exists.return_value = True
        mock_button.Exists.return_value = True
        mock_open_dialog.EditControl.return_value = mock_file_edit
        mock_open_dialog.ButtonControl.return_value = mock_button
        # auto.WaitForExist retorna False = dialogo se cerro (exito)
        mock_uia_file_explorer.WaitForExist.return_value = False

        # Mock _find_file_dialog para que devuelva el dialogo mock
        with patch.object(file_explorer, '_find_file_dialog', return_value=mock_open_dialog):
            # Act
            result = file_explorer.open_file_dialog(path)

        # Assert
        assert result is expected
        mock_file_edit.Click.assert_called_once()
        mock_file_edit.SendKeys.assert_any_call("{Ctrl}a", waitTime=0.1)
        mock_file_edit.SendKeys.assert_any_call(str(path.resolve()), waitTime=0.1)
        assert mock_file_edit.SendKeys.call_count == 2


class TestSaveFileDialog:
    """Tests para save_file_dialog()."""

    def test_save_file_dialog_success(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
        mock_save_dialog: MagicMock,
        mock_file_edit: MagicMock,
        mock_button: MagicMock,
        sample_output_path: Path,
    ) -> None:
        """TC02: FileExplorer inyecta ruta, click Guardar, sin modal."""
        # Arrange
        mock_file_edit.Exists.return_value = True
        mock_button.Exists.return_value = True
        mock_save_dialog.EditControl.return_value = mock_file_edit
        mock_save_dialog.ButtonControl.return_value = mock_button
        # auto.WaitForExist retorna False = dialogo se cerro (exito)
        mock_uia_file_explorer.WaitForExist.return_value = False

        # Mock _find_file_dialog para que devuelva el dialogo mock
        with patch.object(file_explorer, '_find_file_dialog', return_value=mock_save_dialog), \
             patch.object(file_explorer, "handle_replace_modal", return_value=False):

            # Act
            result = file_explorer.save_file_dialog(sample_output_path)

            # Assert
            assert result is True
            mock_file_edit.Click.assert_called_once()
            mock_file_edit.SendKeys.assert_any_call("{Ctrl}a", waitTime=0.1)
            mock_file_edit.SendKeys.assert_any_call(
                str(sample_output_path.resolve()), waitTime=0.1
            )
            assert mock_file_edit.SendKeys.call_count == 2
            mock_button.Click.assert_called_once()

    def test_save_file_dialog_with_replace_modal(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
        mock_save_dialog: MagicMock,
        mock_file_edit: MagicMock,
        mock_button: MagicMock,
        sample_output_path: Path,
    ) -> None:
        """TC02: FileExplorer maneja modal reemplazo (click Si)."""
        # Arrange
        mock_file_edit.Exists.return_value = True
        mock_button.Exists.return_value = True
        mock_save_dialog.EditControl.return_value = mock_file_edit
        mock_save_dialog.ButtonControl.return_value = mock_button
        # auto.WaitForExist retorna False = dialogo se cerro (exito)
        mock_uia_file_explorer.WaitForExist.return_value = False

        # Mock _find_file_dialog para que devuelva el dialogo mock
        with patch.object(file_explorer, '_find_file_dialog', return_value=mock_save_dialog), \
             patch.object(file_explorer, "handle_replace_modal", return_value=True) as mock_handle:

            # Act
            result = file_explorer.save_file_dialog(sample_output_path)

            # Assert
            assert result is True
            mock_handle.assert_called_once()

    def test_save_file_dialog_edit_not_found(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
        mock_save_dialog: MagicMock,
    ) -> None:
        """TC02: Fallo si control Edit no existe."""
        # Arrange
        mock_uia_file_explorer.WindowControl.return_value = mock_save_dialog
        mock_edit = MagicMock()
        mock_edit.Exists.return_value = False
        mock_save_dialog.EditControl.return_value = mock_edit

        # Act
        result = file_explorer.save_file_dialog(Path("test.xlsx"))

        # Assert
        assert result is False

    def test_save_file_dialog_button_not_found(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
        mock_save_dialog: MagicMock,
        mock_file_edit: MagicMock,
    ) -> None:
        """TC02: Fallo si boton Guardar no existe."""
        # Arrange
        mock_uia_file_explorer.WindowControl.return_value = mock_save_dialog
        mock_save_dialog.EditControl.return_value = mock_file_edit
        mock_button = MagicMock()
        mock_button.Exists.return_value = False
        mock_save_dialog.ButtonControl.return_value = mock_button

        # Act
        result = file_explorer.save_file_dialog(Path("test.xlsx"))

        # Assert
        assert result is False

    def test_save_file_dialog_not_appears(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
    ) -> None:
        """TC02: Fallo si dialogo Guardar como no aparece."""
        # Arrange - _find_file_dialog retorna None (no encontro dialogo)
        with patch.object(file_explorer, '_find_file_dialog', return_value=None):
            # Act
            result = file_explorer.save_file_dialog(Path("test.xlsx"))

        # Assert
        assert result is False

    def test_save_file_dialog_dialog_not_closes(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
        mock_save_dialog: MagicMock,
        mock_file_edit: MagicMock,
        mock_button: MagicMock,
    ) -> None:
        """TC02: Guarda exitosamente aunque el dialogo tarde en cerrarse."""
        # Arrange
        mock_uia_file_explorer.WindowControl.return_value = mock_save_dialog
        mock_save_dialog.EditControl.return_value = mock_file_edit
        mock_save_dialog.ButtonControl.return_value = mock_button
        # Dialogo sigue abierto (no se cierra rapido)
        mock_uia_file_explorer.WaitForExist.side_effect = [True, True]

        with patch.object(
            file_explorer, "handle_replace_modal", return_value=False
        ):

            # Act
            result = file_explorer.save_file_dialog(Path("test.xlsx"))

            # Assert - ahora retorna True aunque el dialogo tarde en cerrarse
            assert result is True

    def test_save_file_dialog_exception_handling(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
    ) -> None:
        """TC02: Manejo de excepciones inesperadas."""
        # Arrange
        mock_uia_file_explorer.WindowControl.side_effect = Exception("Unexpected")

        # Act
        result = file_explorer.save_file_dialog(Path("test.xlsx"))

        # Assert
        assert result is False

    @pytest.mark.parametrize(
        "path,expected",
        [
            (Path("output.xlsx"), True),
            (Path("data/results/report.xlsx"), True),
            (Path("C:/absolute/output.xlsx"), True),
        ],
    )
    def test_save_file_dialog_various_paths(
        self,
        path: Path,
        expected: bool,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
        mock_save_dialog: MagicMock,
        mock_file_edit: MagicMock,
        mock_button: MagicMock,
    ) -> None:
        """TC02: Diferentes rutas funcionan correctamente."""
        # Arrange
        mock_file_edit.Exists.return_value = True
        mock_button.Exists.return_value = True
        mock_save_dialog.EditControl.return_value = mock_file_edit
        mock_save_dialog.ButtonControl.return_value = mock_button
        # auto.WaitForExist retorna False = dialogo se cerro (exito)
        mock_uia_file_explorer.WaitForExist.return_value = False

        # Mock _find_file_dialog para que devuelva el dialogo mock
        with patch.object(file_explorer, '_find_file_dialog', return_value=mock_save_dialog), \
             patch.object(file_explorer, "handle_replace_modal", return_value=False):

            # Act
            result = file_explorer.save_file_dialog(path)

            # Assert
            assert result is expected
            mock_file_edit.Click.assert_called_once()
            mock_file_edit.SendKeys.assert_any_call("{Ctrl}a", waitTime=0.1)
            mock_file_edit.SendKeys.assert_any_call(str(path.resolve()), waitTime=0.1)
            assert mock_file_edit.SendKeys.call_count == 2


class TestHandleReplaceModal:
    """Tests para handle_replace_modal()."""

    def test_handle_replace_modal_success(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
        mock_replace_modal: MagicMock,
        mock_button: MagicMock,
    ) -> None:
        """TC02: handle_replace_modal detecta modal y click Si."""
        # Arrange
        mock_uia_file_explorer.WindowControl.return_value = mock_replace_modal
        mock_replace_modal.ButtonControl.return_value = mock_button
        # auto.WaitForExist: True (modal detectado), False (modal cerro tras click)
        mock_uia_file_explorer.WaitForExist.side_effect = [True, False]

        # Act
        result = file_explorer.handle_replace_modal()

        # Assert
        assert result is True
        mock_button.Click.assert_called_once()

    def test_handle_replace_modal_not_present(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
    ) -> None:
        """TC02: Sin modal retorna False."""
        # Arrange
        mock_modal = MagicMock()
        mock_uia_file_explorer.WaitForExist.return_value = False
        mock_uia_file_explorer.WindowControl.return_value = mock_modal

        # Act
        result = file_explorer.handle_replace_modal()

        # Assert
        assert result is False

    def test_handle_replace_modal_yes_button_not_found(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
        mock_replace_modal: MagicMock,
    ) -> None:
        """TC02: Fallo si boton Si no existe."""
        # Arrange
        mock_uia_file_explorer.WindowControl.return_value = mock_replace_modal
        mock_button = MagicMock()
        mock_button.Exists.return_value = False
        mock_replace_modal.ButtonControl.return_value = mock_button

        # Act
        result = file_explorer.handle_replace_modal()

        # Assert
        assert result is False

    def test_handle_replace_modal_not_closes(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
        mock_replace_modal: MagicMock,
        mock_button: MagicMock,
    ) -> None:
        """TC02: Fallo si modal no se cierra tras click."""
        # Arrange
        mock_uia_file_explorer.WindowControl.return_value = mock_replace_modal
        mock_replace_modal.ButtonControl.return_value = mock_button
        # Modal sigue abierto
        mock_uia_file_explorer.WaitForExist.side_effect = [True, True]

        # Act
        result = file_explorer.handle_replace_modal()

        # Assert
        assert result is False

    def test_handle_replace_modal_exception_handling(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
    ) -> None:
        """TC02: Manejo de excepciones inesperadas."""
        # Arrange
        mock_uia_file_explorer.WindowControl.side_effect = Exception("Unexpected")

        # Act
        result = file_explorer.handle_replace_modal()

        # Assert
        assert result is False


class TestWaitForWindow:
    """Tests para wait_for_window()."""

    def test_wait_for_window_with_dialog(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
    ) -> None:
        """wait_for_window retorna True si current_dialog existe y aparece."""
        # Arrange
        mock_dialog = MagicMock()
        mock_uia_file_explorer.WaitForExist.return_value = True
        file_explorer.current_dialog = mock_dialog

        # Act
        result = file_explorer.wait_for_window(timeout=5)

        # Assert
        assert result is True
        mock_uia_file_explorer.WaitForExist.assert_called_once_with(mock_dialog, 5)

    def test_wait_for_window_with_dialog_timeout(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
    ) -> None:
        """wait_for_window retorna False si current_dialog no aparece."""
        # Arrange
        mock_dialog = MagicMock()
        mock_uia_file_explorer.WaitForExist.return_value = False
        file_explorer.current_dialog = mock_dialog

        # Act
        result = file_explorer.wait_for_window(timeout=3)

        # Assert
        assert result is False
        mock_uia_file_explorer.WaitForExist.assert_called_once_with(mock_dialog, 3)

    def test_wait_for_window_no_dialog(
        self,
        file_explorer: FileExplorer,
    ) -> None:
        """wait_for_window retorna False si no hay current_dialog."""
        # Arrange
        file_explorer.current_dialog = None

        # Act
        result = file_explorer.wait_for_window()

        # Assert
        assert result is False


class TestFindFileDialog:
    """Tests para _find_file_dialog()."""

    def test_find_file_dialog_by_classname(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
    ) -> None:
        """_find_file_dialog encuentra dialogo por ClassName #32770."""
        # Arrange
        mock_dialog = MagicMock()
        mock_uia_file_explorer.WindowControl.return_value = mock_dialog
        mock_uia_file_explorer.WaitForExist.return_value = True

        # Act
        result = file_explorer._find_file_dialog(timeout=5)

        # Assert
        assert result is mock_dialog
        mock_uia_file_explorer.WindowControl.assert_called_with(
            searchDepth=2, ClassName="#32770"
        )

    def test_find_file_dialog_by_name_abrir(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
    ) -> None:
        """_find_file_dialog encuentra dialogo por Name 'Abrir'."""
        # Arrange
        mock_dialog_by_class = MagicMock()
        mock_dialog_by_name = MagicMock()

        mock_uia_file_explorer.WindowControl.side_effect = [
            mock_dialog_by_class,
            mock_dialog_by_name,
        ]
        mock_uia_file_explorer.WaitForExist.side_effect = [False, True]

        # Act
        result = file_explorer._find_file_dialog(timeout=5)

        # Assert
        assert result is mock_dialog_by_name

    def test_find_file_dialog_by_name_guardar_como(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
    ) -> None:
        """_find_file_dialog encuentra dialogo por Name 'Guardar como'."""
        # Arrange
        mock_dialog_by_class = MagicMock()
        mock_dialog_by_name_abrir = MagicMock()
        mock_dialog_by_name_save = MagicMock()

        mock_uia_file_explorer.WindowControl.side_effect = [
            mock_dialog_by_class,
            mock_dialog_by_name_abrir,
            mock_dialog_by_name_save,
        ]
        mock_uia_file_explorer.WaitForExist.side_effect = [False, False, True]

        # Act
        result = file_explorer._find_file_dialog(timeout=5)

        # Assert
        assert result is mock_dialog_by_name_save

    def test_find_file_dialog_not_found(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
    ) -> None:
        """_find_file_dialog retorna None si no encuentra dialogo."""
        # Arrange
        mock_uia_file_explorer.WaitForExist.return_value = False

        # Act
        result = file_explorer._find_file_dialog(timeout=2)

        # Assert
        assert result is None


class TestFileExplorerIntegration:
    """Tests de integracion entre metodos de FileExplorer."""

    def test_open_then_save_flow(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
        mock_open_dialog: MagicMock,
        mock_save_dialog: MagicMock,
        mock_file_edit: MagicMock,
        mock_button: MagicMock,
        sample_input_path: Path,
        sample_output_path: Path,
    ) -> None:
        """TC01+TC02: Flujo completo abrir y guardar."""
        # Arrange - Primer dialogo (Abrir)
        mock_file_edit.Exists.return_value = True
        mock_button.Exists.return_value = True
        mock_open_dialog.EditControl.return_value = mock_file_edit
        mock_open_dialog.ButtonControl.return_value = mock_button

        # Arrange - Segundo dialogo (Guardar como)
        mock_save_dialog.EditControl.return_value = mock_file_edit
        mock_save_dialog.ButtonControl.return_value = mock_button
        # auto.WaitForExist retorna False = dialogo se cerro (exito)
        mock_uia_file_explorer.WaitForExist.return_value = False

        # Mock _find_file_dialog para devolver el dialogo correspondiente
        def mock_find_file_dialog(timeout=5):
            if not hasattr(mock_find_file_dialog, 'call_count'):
                mock_find_file_dialog.call_count = 0
            mock_find_file_dialog.call_count += 1
            if mock_find_file_dialog.call_count == 1:
                return mock_open_dialog
            return mock_save_dialog

        with patch.object(file_explorer, '_find_file_dialog', side_effect=mock_find_file_dialog), \
             patch.object(file_explorer, "handle_replace_modal", return_value=False):

            # Act
            open_result = file_explorer.open_file_dialog(sample_input_path)
            save_result = file_explorer.save_file_dialog(sample_output_path)

            # Assert
            assert open_result is True
            assert save_result is True
            assert mock_file_edit.Click.call_count == 2
            assert mock_file_edit.SendKeys.call_count == 4
            assert mock_button.Click.call_count == 2

    def test_save_with_modal_and_retry(
        self,
        file_explorer: FileExplorer,
        mock_uia_file_explorer: MagicMock,
        mock_save_dialog: MagicMock,
        mock_replace_modal: MagicMock,
        mock_file_edit: MagicMock,
        mock_button: MagicMock,
        sample_output_path: Path,
    ) -> None:
        """TC02: Guardar con modal de reemplazo y reintento."""
        # Arrange - Dialogo Guardar
        mock_file_edit.Exists.return_value = True
        mock_button.Exists.return_value = True
        mock_save_dialog.EditControl.return_value = mock_file_edit
        mock_save_dialog.ButtonControl.return_value = mock_button
        # auto.WaitForExist retorna False = dialogo se cerro (exito)
        mock_uia_file_explorer.WaitForExist.return_value = False

        # Mock _find_file_dialog para devolver el dialogo mock
        with patch.object(file_explorer, '_find_file_dialog', return_value=mock_save_dialog), \
             patch.object(file_explorer, "handle_replace_modal", return_value=True):

            # Act
            result = file_explorer.save_file_dialog(sample_output_path)

            # Assert
            assert result is True
