"""Tests para ExcelManager - Caso de Prueba 01: Abrir dialogo y delegar a FileExplorer."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Generator
from unittest.mock import MagicMock, patch, PropertyMock

import pytest

from rpa_excel_ui_automation.excel_manager import ExcelManager


class TestExcelManagerInit:
    """Tests para inicializacion de ExcelManager."""

    def test_init_with_default_logger(self) -> None:
        """ExcelManager inicializa con logger por defecto."""
        # Arrange & Act
        manager = ExcelManager()

        # Assert
        assert manager.app is None
        assert manager._logger is not None
        assert isinstance(manager._file_explorer, type(manager._file_explorer))

    def test_init_with_custom_logger(self, mock_logger: MagicMock) -> None:
        """ExcelManager inicializa con logger personalizado."""
        # Arrange & Act
        manager = ExcelManager(logger=mock_logger)

        # Assert
        assert manager._logger is mock_logger
        mock_logger.info.assert_called_with("ExcelManager inicializado")

    def test_init_sets_app_none(self) -> None:
        """ExcelManager inicia app como None."""
        # Arrange & Act
        manager = ExcelManager()

        # Assert
        assert manager.app is None


class TestConnectOrLaunchExcel:
    """Tests para _connect_or_launch_excel()."""

    def test_connect_existing_excel(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
        mock_app_window: MagicMock,
    ) -> None:
        """TC01: Conecta a Excel existente (XLMAIN encontrado)."""
        # Arrange
        mock_uia_excel_manager.WindowControl.return_value = mock_app_window
        mock_app_window.Exists.return_value = True

        # Act
        result = excel_manager._connect_or_launch_excel()

        # Assert
        assert result is True
        assert excel_manager.app is mock_app_window
        mock_uia_excel_manager.WindowControl.assert_called_with(
            searchDepth=1, ClassName="XLMAIN"
        )

    def test_launch_new_excel(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
        mock_subprocess: MagicMock,
    ) -> None:
        """TC01: Lanza nueva instancia si Excel no esta corriendo."""
        # Arrange
        mock_app_before = MagicMock()
        mock_app_before.Exists.return_value = False

        mock_app_after = MagicMock()
        mock_app_after.WaitForExist.return_value = True
        mock_app_after.Name = "Microsoft Excel"

        mock_uia_excel_manager.WindowControl.side_effect = [
            mock_app_before,
            mock_app_after,
        ]

        # Act
        result = excel_manager._connect_or_launch_excel()

        # Assert
        assert result is True
        assert excel_manager.app is mock_app_after
        mock_subprocess.assert_called_once_with(["excel.exe"])

    def test_launch_excel_file_not_found(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
        mock_subprocess: MagicMock,
    ) -> None:
        """TC01: Falla si excel.exe no existe en PATH."""
        # Arrange
        mock_app_before = MagicMock()
        mock_app_before.Exists.return_value = False

        mock_uia_excel_manager.WindowControl.return_value = mock_app_before
        mock_subprocess.side_effect = FileNotFoundError("excel.exe not found")

        # Act
        result = excel_manager._connect_or_launch_excel()

        # Assert
        assert result is False

    def test_launch_excel_timeout(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
        mock_subprocess: MagicMock,
    ) -> None:
        """TC01: Falla si Excel no aparece tras 15 segundos."""
        # Arrange
        mock_app_before = MagicMock()
        mock_app_before.Exists.return_value = False

        mock_app_after = MagicMock()

        mock_uia_excel_manager.WindowControl.side_effect = [
            mock_app_before,
            mock_app_after,
        ]
        mock_uia_excel_manager.WaitForExist.return_value = False

        # Act
        result = excel_manager._connect_or_launch_excel()

        # Assert
        assert result is False
        mock_uia_excel_manager.WaitForExist.assert_called_with(mock_app_after, 15)

    def test_launch_excel_generic_exception(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
        mock_subprocess: MagicMock,
    ) -> None:
        """TC01: Falla con excepcion generica al lanzar Excel."""
        # Arrange
        mock_app_before = MagicMock()
        mock_app_before.Exists.return_value = False

        mock_uia_excel_manager.WindowControl.return_value = mock_app_before
        mock_subprocess.side_effect = OSError("Permission denied")

        # Act
        result = excel_manager._connect_or_launch_excel()

        # Assert
        assert result is False


class TestOpenFile:
    """Tests para open_file()."""

    def test_open_file_solo_dialogo(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
        mock_app_window: MagicMock,
        mock_open_dialog: MagicMock,
    ) -> None:
        """TC01: open_file() abre Excel y dialogo Abrir (sin file_path)."""
        # Arrange
        mock_uia_excel_manager.WindowControl.side_effect = [
            mock_app_window,
            mock_open_dialog,
        ]
        mock_app_window.Exists.return_value = True
        mock_uia_excel_manager.WaitForExist.return_value = True

        # Act
        result = excel_manager.open_file()

        # Assert
        assert result is True
        mock_app_window.SendKeys.assert_called_with("{Ctrl}o", waitTime=0.5)
        mock_uia_excel_manager.WaitForExist.assert_called_with(mock_open_dialog, 5)

    def test_open_file_con_file_path(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
        mock_app_window: MagicMock,
        mock_open_dialog: MagicMock,
        sample_input_path: Path,
    ) -> None:
        """TC01: open_file() con file_path delega a FileExplorer."""
        # Arrange
        mock_uia_excel_manager.WindowControl.side_effect = [
            mock_app_window,
            mock_open_dialog,
        ]
        mock_app_window.Exists.return_value = True

        with patch.object(
            excel_manager._file_explorer, "open_file_dialog", return_value=True
        ) as mock_fe:

            # Act
            result = excel_manager.open_file(file_path=sample_input_path)

            # Assert
            assert result is True
            mock_fe.assert_called_once_with(sample_input_path)

    def test_open_file_no_excel_fails(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
    ) -> None:
        """TC01: open_file() falla si no puede conectar/lanzar Excel."""
        # Arrange
        mock_app = MagicMock()
        mock_app.Exists.return_value = False
        mock_uia_excel_manager.WindowControl.return_value = mock_app

        # Act
        result = excel_manager.open_file()

        # Assert
        assert result is False

    def test_open_file_dialog_timeout_fails(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
        mock_app_window: MagicMock,
    ) -> None:
        """TC01: open_file() falla si dialogo Abrir no aparece."""
        # Arrange
        mock_dialog = MagicMock()

        mock_uia_excel_manager.WindowControl.side_effect = [
            mock_app_window,
            mock_dialog,
        ]
        mock_app_window.Exists.return_value = True
        mock_uia_excel_manager.WaitForExist.return_value = False

        # Act
        result = excel_manager.open_file()

        # Assert
        assert result is False
        mock_uia_excel_manager.WaitForExist.assert_called_with(mock_dialog, 5)

    def test_open_file_file_explorer_fails(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
        mock_app_window: MagicMock,
        mock_open_dialog: MagicMock,
        sample_input_path: Path,
    ) -> None:
        """TC01: open_file() falla si FileExplorer falla."""
        # Arrange
        mock_uia_excel_manager.WindowControl.side_effect = [
            mock_app_window,
            mock_open_dialog,
        ]
        mock_app_window.Exists.return_value = True

        with patch.object(
            excel_manager._file_explorer, "open_file_dialog", return_value=False
        ) as mock_fe:

            # Act
            result = excel_manager.open_file(file_path=sample_input_path)

            # Assert
            assert result is False
            mock_fe.assert_called_once_with(sample_input_path)

    def test_open_file_exception_handling(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
        mock_app_window: MagicMock,
    ) -> None:
        """TC01: open_file() maneja excepciones inesperadas."""
        # Arrange
        mock_uia_excel_manager.WindowControl.return_value = mock_app_window
        mock_app_window.Exists.side_effect = Exception("Unexpected error")

        # Act
        result = excel_manager.open_file()

        # Assert
        assert result is False


class TestSaveAs:
    """Tests para save_as()."""

    def test_save_as_solo_dialogo(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
        mock_save_dialog: MagicMock,
        mock_app_window: MagicMock,
    ) -> None:
        """TC02: save_as() abre dialogo Guardar como (sin file_path)."""
        # Arrange
        excel_manager.app = mock_app_window
        mock_uia_excel_manager.WindowControl.return_value = mock_save_dialog
        mock_uia_excel_manager.WaitForExist.return_value = True

        # Act
        result = excel_manager.save_as()

        # Assert
        assert result is True
        mock_app_window.SendKeys.assert_called_with("{F12}", waitTime=0.5)
        mock_uia_excel_manager.WaitForExist.assert_called_with(mock_save_dialog, 5)

    def test_save_as_con_file_path(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
        mock_save_dialog: MagicMock,
        mock_app_window: MagicMock,
        sample_output_path: Path,
    ) -> None:
        """TC02: save_as() con file_path delega a FileExplorer."""
        # Arrange
        excel_manager.app = mock_app_window
        mock_uia_excel_manager.WindowControl.return_value = mock_save_dialog

        with patch.object(
            excel_manager._file_explorer, "save_file_dialog", return_value=True
        ) as mock_fe:

            # Act
            result = excel_manager.save_as(file_path=sample_output_path)

            # Assert
            assert result is True
            mock_fe.assert_called_once_with(sample_output_path)

    def test_save_as_no_active_excel_fails(
        self,
        excel_manager: ExcelManager,
    ) -> None:
        """TC02: save_as() falla si no hay Excel activo."""
        # Arrange
        excel_manager.app = None

        # Act
        result = excel_manager.save_as()

        # Assert
        assert result is False

    def test_save_as_excel_not_exists_fails(
        self,
        excel_manager: ExcelManager,
        mock_app_window: MagicMock,
    ) -> None:
        """TC02: save_as() falla si Excel ya no existe."""
        # Arrange
        excel_manager.app = mock_app_window
        mock_app_window.Exists.return_value = False

        # Act
        result = excel_manager.save_as()

        # Assert
        assert result is False

    def test_save_as_dialog_timeout_fails(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
        mock_app_window: MagicMock,
    ) -> None:
        """TC02: save_as() falla si dialogo Guardar como no aparece."""
        # Arrange
        excel_manager.app = mock_app_window
        mock_dialog = MagicMock()
        mock_uia_excel_manager.WindowControl.return_value = mock_dialog
        mock_uia_excel_manager.WaitForExist.return_value = False

        # Act
        result = excel_manager.save_as()

        # Assert
        assert result is False
        mock_uia_excel_manager.WaitForExist.assert_called_with(mock_dialog, 5)

    def test_save_as_file_explorer_fails(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
        mock_save_dialog: MagicMock,
        mock_app_window: MagicMock,
        sample_output_path: Path,
    ) -> None:
        """TC02: save_as() falla si FileExplorer falla."""
        # Arrange
        excel_manager.app = mock_app_window
        mock_uia_excel_manager.WindowControl.return_value = mock_save_dialog

        with patch.object(
            excel_manager._file_explorer, "save_file_dialog", return_value=False
        ) as mock_fe:

            # Act
            result = excel_manager.save_as(file_path=sample_output_path)

            # Assert
            assert result is False
            mock_fe.assert_called_once_with(sample_output_path)

    def test_save_as_exception_handling(
        self,
        excel_manager: ExcelManager,
        mock_app_window: MagicMock,
    ) -> None:
        """TC02: save_as() maneja excepciones inesperadas."""
        # Arrange
        excel_manager.app = mock_app_window
        mock_app_window.Exists.side_effect = Exception("Unexpected error")

        # Act
        result = excel_manager.save_as()

        # Assert
        assert result is False


class TestClose:
    """Tests para close()."""

    def test_close_with_active_excel(
        self,
        excel_manager: ExcelManager,
        mock_app_window: MagicMock,
    ) -> None:
        """TC01: close() cierra Excel activo."""
        # Arrange
        excel_manager.app = mock_app_window
        mock_app_window.Exists.return_value = True
        mock_uia_excel_manager = MagicMock()
        mock_uia_excel_manager.WaitForExist.return_value = False  # Excel se cerro

        # Act
        with patch("rpa_excel_ui_automation.excel_manager.auto", mock_uia_excel_manager):
            result = excel_manager.close()

        # Assert
        assert result is True
        mock_app_window.SendKeys.assert_called_with("{Alt}f4", waitTime=0.5)
        mock_uia_excel_manager.WaitForExist.assert_called_with(mock_app_window, 3)

    def test_close_excel_not_closing(
        self,
        excel_manager: ExcelManager,
        mock_app_window: MagicMock,
    ) -> None:
        """TC01: close() falla si Excel no se cierra."""
        # Arrange
        excel_manager.app = mock_app_window
        mock_app_window.Exists.return_value = True
        mock_app_window.WaitForExist.return_value = True  # Excel sigue abierto

        # Act
        result = excel_manager.close()

        # Assert
        assert result is False

    def test_close_no_active_excel(self, excel_manager: ExcelManager) -> None:
        """TC01: close() retorna True si no hay Excel activo."""
        # Arrange
        excel_manager.app = None

        # Act
        result = excel_manager.close()

        # Assert
        assert result is True

    def test_close_excel_already_closed(
        self,
        excel_manager: ExcelManager,
        mock_app_window: MagicMock,
    ) -> None:
        """TC01: close() retorna True si Excel ya existe pero no esta visible."""
        # Arrange
        excel_manager.app = mock_app_window
        mock_app_window.Exists.return_value = False

        # Act
        result = excel_manager.close()

        # Assert
        assert result is True

    def test_close_exception_handling(
        self,
        excel_manager: ExcelManager,
        mock_app_window: MagicMock,
    ) -> None:
        """TC01: close() maneja excepciones inesperadas."""
        # Arrange
        excel_manager.app = mock_app_window
        mock_app_window.Exists.side_effect = Exception("Unexpected error")

        # Act
        result = excel_manager.close()

        # Assert
        assert result is False


class TestContextManager:
    """Tests para uso como context manager."""

    def test_enter_returns_self(self, excel_manager: ExcelManager) -> None:
        """TC01: __enter__ retorna la instancia."""
        # Arrange & Act
        with excel_manager as manager:

            # Assert
            assert manager is not None
            assert isinstance(manager, ExcelManager)

    def test_exit_calls_close(self, excel_manager: ExcelManager) -> None:
        """TC01: __exit__ llama a close()."""
        # Arrange
        with patch.object(excel_manager, "close", return_value=True) as mock_close:
            # Act
            with excel_manager:
                pass

            # Assert
            mock_close.assert_called_once()

    def test_exit_with_exception(self, excel_manager: ExcelManager) -> None:
        """TC01: __exit__ llama close() incluso con excepcion."""
        # Arrange
        with patch.object(excel_manager, "close", return_value=True) as mock_close:
            # Act
            try:
                with excel_manager:
                    raise ValueError("Test exception")
            except ValueError:
                pass

            # Assert
            mock_close.assert_called_once()

    def test_context_manager_with_open_file(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
        mock_app_window: MagicMock,
        mock_open_dialog: MagicMock,
    ) -> None:
        """TC01: Context manager funciona con open_file()."""
        # Arrange
        mock_uia_excel_manager.WindowControl.side_effect = [
            mock_app_window,
            mock_open_dialog,
        ]
        mock_app_window.Exists.return_value = True
        mock_app_window.WaitForExist.return_value = False  # Excel se cierra

        # Act
        with excel_manager as manager:
            result = manager.open_file()

        # Assert
        assert result is True


class TestExcelManagerIntegration:
    """Tests de integracion entre metodos de ExcelManager."""

    def test_open_file_then_save_as(
        self,
        excel_manager: ExcelManager,
        mock_uia_excel_manager: MagicMock,
        mock_app_window: MagicMock,
        mock_open_dialog: MagicMock,
        mock_save_dialog: MagicMock,
        sample_input_path: Path,
        sample_output_path: Path,
    ) -> None:
        """TC01+TC02: Abrir archivo y luego guardar como."""
        # Arrange
        mock_uia_excel_manager.WindowControl.side_effect = [
            mock_app_window,  # _connect_or_launch_excel
            mock_open_dialog,  # open_file dialog
            mock_save_dialog,  # save_as dialog
        ]
        mock_app_window.Exists.return_value = True

        with patch.object(
            excel_manager._file_explorer, "open_file_dialog", return_value=True
        ), patch.object(
            excel_manager._file_explorer, "save_file_dialog", return_value=True
        ):

            # Act
            open_result = excel_manager.open_file(file_path=sample_input_path)
            excel_manager.app = mock_app_window  # Simular que app sigue activa
            save_result = excel_manager.save_as(file_path=sample_output_path)

            # Assert
            assert open_result is True
            assert save_result is True
