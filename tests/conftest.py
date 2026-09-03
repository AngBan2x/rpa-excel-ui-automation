"""Fixtures compartidas para tests de rpa_excel_ui_automation."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Generator
from unittest.mock import MagicMock, patch

import pytest

from rpa_excel_ui_automation.excel_manager import ExcelManager
from rpa_excel_ui_automation.file_explorer import FileExplorer


@pytest.fixture
def mock_logger() -> MagicMock:
    """Mock logger para verificar llamadas de log."""
    logger = MagicMock(spec=logging.Logger)
    logger.info = MagicMock()
    logger.debug = MagicMock()
    logger.warning = MagicMock()
    logger.error = MagicMock()
    logger.exception = MagicMock()
    return logger


@pytest.fixture
def sample_input_path() -> Path:
    """Ruta de ejemplo para archivo de entrada."""
    return Path(".data/input/origen.xlsx")


@pytest.fixture
def sample_output_path() -> Path:
    """Ruta de ejemplo para archivo de salida."""
    return Path(".data/output/destino.xlsx")


@pytest.fixture
def mock_uia_excel_manager() -> Generator[MagicMock, None, None]:
    """Mock de uiautomation para ExcelManager."""
    with patch("rpa_excel_ui_automation.excel_manager.auto") as mock:
        yield mock


@pytest.fixture
def mock_uia_file_explorer() -> Generator[MagicMock, None, None]:
    """Mock de uiautomation para FileExplorer."""
    with patch("rpa_excel_ui_automation.file_explorer.auto") as mock:
        yield mock


@pytest.fixture
def mock_subprocess() -> Generator[MagicMock, None, None]:
    """Mock de subprocess para lanzar Excel."""
    with patch("subprocess.Popen") as mock:
        yield mock


@pytest.fixture
def excel_manager(mock_logger: MagicMock) -> ExcelManager:
    """Instancia de ExcelManager con logger mockeado."""
    return ExcelManager(logger=mock_logger)


@pytest.fixture
def file_explorer(mock_logger: MagicMock) -> FileExplorer:
    """Instancia de FileExplorer con logger mockeado."""
    return FileExplorer(logger=mock_logger)


@pytest.fixture
def mock_app_window() -> MagicMock:
    """Mock de ventana principal de Excel (XLMAIN)."""
    mock_app = MagicMock()
    mock_app.Exists.return_value = True
    mock_app.exists.return_value = True
    mock_app.Name = "Microsoft Excel"
    return mock_app


@pytest.fixture
def mock_open_dialog() -> MagicMock:
    """Mock de dialogo 'Abrir'."""
    mock_dlg = MagicMock()
    mock_dlg.WaitForExist.return_value = True
    mock_dlg.Name = "Abrir"
    return mock_dlg


@pytest.fixture
def mock_save_dialog() -> MagicMock:
    """Mock de dialogo 'Guardar como'."""
    mock_dlg = MagicMock()
    mock_dlg.WaitForExist.return_value = True
    mock_dlg.Name = "Guardar como"
    return mock_dlg


@pytest.fixture
def mock_replace_modal() -> MagicMock:
    """Mock de modal 'Confirmar guardado'."""
    mock_modal = MagicMock()
    mock_modal.WaitForExist.return_value = True
    mock_modal.Name = "Confirmar guardado"
    return mock_modal


@pytest.fixture
def mock_file_edit() -> MagicMock:
    """Mock de control Edit para nombre de archivo."""
    mock_edit = MagicMock()
    mock_edit.Exists.return_value = True
    return mock_edit


@pytest.fixture
def mock_button() -> MagicMock:
    """Mock de boton generico."""
    mock_btn = MagicMock()
    mock_btn.Exists.return_value = True
    return mock_btn
