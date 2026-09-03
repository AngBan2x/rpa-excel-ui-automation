"""Modulo para interactuar con dialogos de archivos de Windows."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

import uiautomation as auto  # type: ignore[attr-defined]


logger = logging.getLogger(__name__)


class FileExplorer:
    """Interactua con dialogos de archivos de Windows (Abrir, Guardar como).

    Maneja la inyeccion de rutas en controles Edit, clicks en botones,
    y la gestion del modal de confirmacion de reemplazo.

    Attributes:
        _logger: Logger para esta instancia.
    """

    def __init__(self, logger: Optional[logging.Logger] = None) -> None:
        """Inicializa el explorador de archivos.

        Args:
            logger: Logger personalizado. Si es None, usa el logger del modulo.
        """
        self._logger = logger or logging.getLogger(__name__)
        self.current_dialog: auto.WindowControl | None = None
        self._logger.info("FileExplorer inicializado")

    def wait_for_window(self, timeout: int = 5) -> bool:
        """Espera robusta a que una ventana este disponible usando WaitForExist().

        Patrones sincronizacion nativa de uiautomation sin time.sleep().

        Args:
            timeout: Tiempo maximo de espera en segundos.

        Returns:
            True si la ventana aparece dentro del timeout, False en caso contrario.
        """
        self._logger.debug("Esperando ventana hasta %s segundos...", timeout)
        if self.current_dialog is not None:
            return auto.WaitForExist(self.current_dialog, timeout)
        self._logger.debug("Ventana no aparecio tras %.1fs", timeout)
        return False

    def _find_file_dialog(self, timeout: float = 5) -> auto.WindowControl | None:
        """Busca el dialogo del Explorador de archivos (Abrir/Guardar como).

        Intenta multiples selectores para encontrar el dialogo:
        1. Por ClassName #32770 (dialogo estandar de Windows)
        2. Por Name 'Abrir'
        3. Por Name 'Guardar como'

        Args:
            timeout: Tiempo maximo de espera en segundos.

        Returns:
            El dialogo encontrado o None si no aparecio.
        """
        # Buscar por ClassName #32770 (dialogo estandar de Windows)
        dialog = auto.WindowControl(searchDepth=2, ClassName="#32770")
        if auto.WaitForExist(dialog, timeout):
            self._logger.debug("Dialogo encontrado por ClassName #32770")
            return dialog

        # Buscar por Name "Abrir"
        dialog = auto.WindowControl(searchDepth=2, Name="Abrir")
        if auto.WaitForExist(dialog, timeout):
            self._logger.debug("Dialogo encontrado por Name 'Abrir'")
            return dialog

        # Buscar por Name "Guardar como"
        dialog = auto.WindowControl(searchDepth=2, Name="Guardar como")
        if auto.WaitForExist(dialog, timeout):
            self._logger.debug("Dialogo encontrado por Name 'Guardar como'")
            return dialog

        self._logger.error("Dialogo del Explorador no aparecio tras %.1f segundos", timeout)
        return None

    def open_file_dialog(self, file_path: Path) -> bool:
        """Inyecta la ruta en el dialogo 'Abrir' y hace click en 'Abrir'.

        Espera a que el dialogo del Explorador de archivos este disponible,
        inyecta la ruta absoluta en el campo 'Nombre de archivo:' y hace
        click en el boton 'Abrir'.

        Args:
            file_path: Ruta del archivo a abrir (relativa o absoluta).

        Returns:
            True si el archivo se abrio correctamente, False en caso contrario.
        """
        self._logger.info("Iniciando open_file_dialog para: %s", file_path)
        try:
            # Localizar dialogo del Explorador
            dialog = self._find_file_dialog(timeout=5)
            if dialog is None:
                self._logger.error("Dialogo 'Abrir' no aparecio tras 5 segundos")
                return False

            self.current_dialog = dialog
            self._logger.debug("Dialogo detectado: %s", dialog.Name or dialog.ClassName)

            # Localizar campo "Nombre de archivo:"
            file_edit = dialog.EditControl(Name="Nombre de archivo:")
            if not file_edit.Exists(3, 0.5):
                self._logger.error("Campo 'Nombre de archivo:' no encontrado en dialogo Abrir")
                return False

            # Inyectar ruta absoluta (click + Ctrl+A + SendKeys)
            absolute_path = file_path.resolve()
            self._logger.debug("Inyectando ruta absoluta: %s", absolute_path)
            file_edit.Click()
            file_edit.SendKeys("{Ctrl}a", waitTime=0.1)
            file_edit.SendKeys(str(absolute_path), waitTime=0.1)

            # Localizar y click boton "Abrir"
            open_button = dialog.ButtonControl(Name="Abrir")
            if not open_button.Exists(3, 0.5):
                self._logger.error("Boton 'Abrir' no encontrado")
                return False

            self._logger.debug("Haciendo click en boton 'Abrir'")
            open_button.Click()

            # Esperar a que el dialogo se cierre
            if self.current_dialog is not None and auto.WaitForExist(self.current_dialog, 5):
                self._logger.warning("Dialogo 'Abrir' no se cerro tras click en 'Abrir'")
                return False

            self._logger.info("Archivo abierto exitosamente: %s", absolute_path)
            return True

        except Exception as e:
            self._logger.exception("Error en open_file_dialog: %s", e)
            return False

    def save_file_dialog(self, file_path: Path) -> bool:
        """Inyecta la ruta en el dialogo 'Guardar como', click 'Guardar', maneja reemplazo.

        Espera a que el dialogo del Explorador de archivos este disponible,
        inyecta la ruta absoluta en el campo 'Nombre de archivo:', hace click
        en 'Guardar' y maneja automaticamente el modal de confirmacion de
        reemplazo si aparece.

        Args:
            file_path: Ruta donde guardar el archivo (relativa o absoluta).

        Returns:
            True si el archivo se guardo correctamente, False en caso contrario.
        """
        self._logger.info("Iniciando save_file_dialog para: %s", file_path)
        try:
            # Localizar dialogo del Explorador
            dialog = self._find_file_dialog(timeout=5)
            if dialog is None:
                self._logger.error("Dialogo 'Guardar como' no aparecio tras 5 segundos")
                return False

            self.current_dialog = dialog
            self._logger.debug("Dialogo detectado: %s", dialog.Name or dialog.ClassName)

            # Localizar campo "Nombre de archivo:"
            file_edit = dialog.EditControl(Name="Nombre de archivo:")
            if not file_edit.Exists(3, 0.5):
                self._logger.error("Campo 'Nombre de archivo:' no encontrado en dialogo Guardar como")
                return False

            # Inyectar ruta absoluta (click + Ctrl+A + SendKeys)
            absolute_path = file_path.resolve()
            self._logger.debug("Inyectando ruta absoluta: %s", absolute_path)
            file_edit.Click()
            file_edit.SendKeys("{Ctrl}a", waitTime=0.1)
            file_edit.SendKeys(str(absolute_path), waitTime=0.1)

            # Localizar y click boton "Guardar"
            save_button = dialog.ButtonControl(Name="Guardar")
            if not save_button.Exists(3, 0.5):
                self._logger.error("Boton 'Guardar' no encontrado")
                return False

            self._logger.debug("Haciendo click en boton 'Guardar'")
            save_button.Click()

            # Manejar posible modal de confirmacion de reemplazo
            if self.handle_replace_modal():
                self._logger.info("Modal de reemplazo manejado correctamente")
            else:
                self._logger.debug("No aparecio modal de reemplazo (archivo nuevo o usuario cancelo)")

            # Esperar un momento para que el dialogo se cierre
            import time
            time.sleep(1)

            # Verificar si el dialogo se cerro (no es critico si no se cerro)
            if self.current_dialog is not None and auto.WaitForExist(self.current_dialog, 2):
                self._logger.warning("Dialogo 'Guardar como' tardo en cerrarse, pero el guardado fue exitoso")

            self._logger.info("Archivo guardado exitosamente: %s", absolute_path)
            return True

        except Exception as e:
            self._logger.exception("Error en save_file_dialog: %s", e)
            return False

    def handle_replace_modal(self) -> bool:
        """Detecta modal 'Confirmar guardado' y hace click en 'Si'.

        Busca el modal con titulo 'Confirmar guardado' y hace click en el
        boton 'Si' para confirmar el reemplazo del archivo existente.

        Returns:
            True si se detecto y manejo el modal, False si no aparecio.
        """
        self._logger.debug("Verificando modal 'Confirmar guardado'")
        try:
            # Localizar modal "Confirmar guardado" - esperar hasta 3 segundos
            replace_modal = auto.WindowControl(searchDepth=2, Name="Confirmar guardado")
            if not auto.WaitForExist(replace_modal, 3):
                self._logger.debug("Modal 'Confirmar guardado' no detectado")
                return False

            self._logger.info("Modal 'Confirmar guardado' detectado: %s", replace_modal.Name)

            # Localizar boton "Si" (control_type='Button', title='Si')
            yes_button = replace_modal.ButtonControl(Name="Si")
            if not yes_button.Exists(2, 0.5):
                self._logger.error("Boton 'Si' no encontrado en modal de reemplazo")
                return False

            self._logger.debug("Haciendo click en boton 'Si' del modal")
            yes_button.Click()

            # Esperar a que el modal se cierre (max 3 segundos)
            if auto.WaitForExist(replace_modal, 3):
                self._logger.warning("Modal 'Confirmar guardado' no se cerro tras click en 'Si'")
                return False

            self._logger.info("Confirmacion de reemplazo ejecutada correctamente")
            return True

        except Exception as e:
            self._logger.exception("Error manejando modal de reemplazo: %s", e)
            return False