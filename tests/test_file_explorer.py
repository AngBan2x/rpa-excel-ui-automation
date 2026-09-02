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
        mock_uia_file_explorer.WindowControl.return_value = mock_open_dialog
        mock_open_dialog.EditControl.return_value = mock_file_edit
        mock_open_dialog.ButtonControl.return_value = mock_button

        # El dialogo se cierra despues del click
        mock_open_dialog.WaitForExist.side_effect = [True, False]

        # Act
        result = file_explorer.open_file_dialog(sample_input_path)

        # Assert
        assert result is True
        mock_file_edit.SetValue.assert_called_once_with(
            str(sample_input_path.resolve())
        )
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
        # Arrange
        mock_dialog = MagicMock()
        mock_dialog.WaitForExist.return_value = False
        mock_uia_file_explorer.WindowControl.return_value = mock_dialog

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
        mock_open_dialog.WaitForExist.side_effect = [True, True]

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
        mock_uia_file_explorer.WindowControl.return_value = mock_open_dialog
        mock_open_dialog.EditControl.return_value = mock_file_edit
        mock_open_dialog.ButtonControl.return_value = mock_button
        mock_open_dialog.WaitForExist.side_effect = [True, False]

        # Act
        result = file_explorer.open_file_dialog(path)

        # Assert
        assert result is expected
        mock_file_edit.SetValue.assert_called_once_with(str(path.resolve()))


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
        mock_uia_file_explorer.WindowControl.return_value = mock_save_dialog
        mock_save_dialog.EditControl.return_value = mock_file_edit
        mock_save_dialog.ButtonControl.return_value = mock_button

        # Dialogo se cierra despues del click
        mock_save_dialog.WaitForExist.side_effect = [True, False]

        # Mock handle_replace_modal retorna False (no hay modal)
        with patch.object(
            file_explorer, "handle_replace_modal", return_value=False
        ):

            # Act
            result = file_explorer.save_file_dialog(sample_output_path)

            # Assert
            assert result is True
            mock_file_edit.SetValue.assert_called_once_with(
                str(sample_output_path.resolve())
            )
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
        mock_uia_file_explorer.WindowControl.return_value = mock_save_dialog
        mock_save_dialog.EditControl.return_value = mock_file_edit
        mock_save_dialog.ButtonControl.return_value = mock_button
        mock_save_dialog.WaitForExist.side_effect = [True, False]

        # Mock handle_replace_modal retorna True (modal aparecio)
        with patch.object(
            file_explorer, "handle_replace_modal", return_value=True
        ) as mock_handle:

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
        # Arrange
        mock_dialog = MagicMock()
        mock_dialog.WaitForExist.return_value = False
        mock_uia_file_explorer.WindowControl.return_value = mock_dialog

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
        """TC02: Fallo si dialogo no se cierra tras click."""
        # Arrange
        mock_uia_file_explorer.WindowControl.return_value = mock_save_dialog
        mock_save_dialog.EditControl.return_value = mock_file_edit
        mock_save_dialog.ButtonControl.return_value = mock_button
        # Dialogo sigue abierto
        mock_save_dialog.WaitForExist.side_effect = [True, True]

        with patch.object(
            file_explorer, "handle_replace_modal", return_value=False
        ):

            # Act
            result = file_explorer.save_file_dialog(Path("test.xlsx"))

            # Assert
            assert result is False

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
        mock_uia_file_explorer.WindowControl.return_value = mock_save_dialog
        mock_save_dialog.EditControl.return_value = mock_file_edit
        mock_save_dialog.ButtonControl.return_value = mock_button
        mock_save_dialog.WaitForExist.side_effect = [True, False]

        with patch.object(
            file_explorer, "handle_replace_modal", return_value=False
        ):

            # Act
            result = file_explorer.save_file_dialog(path)

            # Assert
            assert result is expected
            mock_file_edit.SetValue.assert_called_once_with(str(path.resolve()))


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
        # Modal se cierra despues del click
        mock_replace_modal.WaitForExist.side_effect = [True, False]

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
        mock_modal.WaitForExist.return_value = False
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
        mock_replace_modal.WaitForExist.side_effect = [True, True]

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
        mock_open_dialog.EditControl.return_value = mock_file_edit
        mock_open_dialog.ButtonControl.return_value = mock_button
        mock_open_dialog.WaitForExist.side_effect = [True, False]

        # Arrange - Segundo dialogo (Guardar como)
        mock_save_dialog.EditControl.return_value = mock_file_edit
        mock_save_dialog.ButtonControl.return_value = mock_button
        mock_save_dialog.WaitForExist.side_effect = [True, False]

        mock_uia_file_explorer.WindowControl.side_effect = [
            mock_open_dialog,
            mock_save_dialog,
        ]

        with patch.object(
            file_explorer, "handle_replace_modal", return_value=False
        ):

            # Act
            open_result = file_explorer.open_file_dialog(sample_input_path)
            save_result = file_explorer.save_file_dialog(sample_output_path)

            # Assert
            assert open_result is True
            assert save_result is True
            assert mock_file_edit.SetValue.call_count == 2
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
        mock_save_dialog.EditControl.return_value = mock_file_edit
        mock_save_dialog.ButtonControl.return_value = mock_button
        mock_save_dialog.WaitForExist.side_effect = [True, False]

        # Arrange - Modal Confirmar
        mock_replace_modal.ButtonControl.return_value = mock_button
        mock_replace_modal.WaitForExist.side_effect = [True, False]

        mock_uia_file_explorer.WindowControl.side_effect = [
            mock_save_dialog,
            mock_replace_modal,
        ]

        with patch.object(
            file_explorer, "handle_replace_modal", return_value=True
        ):

            # Act
            result = file_explorer.save_file_dialog(sample_output_path)

            # Assert
            assert result is True
