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
        self._logger.info("FileExplorer inicializado")

    def open_file_dialog(self, file_path: Path) -> bool:
        """Inyecta la ruta en el dialogo 'Abrir' y hace click en 'Abrir'.

        Espera a que el dialogo 'Abrir' este disponible, inyecta la ruta
        absoluta en el campo 'Nombre de archivo:' y hace click en el boton 'Abrir'.

        Args:
            file_path: Ruta del archivo a abrir (relativa o absoluta).

        Returns:
            True si el archivo se abrio correctamente, False en caso contrario.
        """
        self._logger.info("Iniciando open_file_dialog para: %s", file_path)
        try:
            # Localizar dialogo "Abrir" - esperar hasta 5 segundos
            dialog = auto.WindowControl(searchDepth=1, Name="Abrir")
            if not dialog.WaitForExist(5, 0.5):
                self._logger.error("Dialogo 'Abrir' no aparecio tras 5 segundos")
                return False

            self._logger.debug("Dialogo 'Abrir' detectado: %s", dialog.Name)

            # Localizar campo "Nombre de archivo:" (control_type='Edit', title='Nombre de archivo:')
            file_edit = dialog.EditControl(Name="Nombre de archivo:")
            if not file_edit.Exists(3, 0.5):
                self._logger.error("Campo 'Nombre de archivo:' no encontrado en dialogo Abrir")
                return False

            # Inyectar ruta absoluta
            absolute_path = file_path.resolve()
            self._logger.debug("Inyectando ruta absoluta en campo 'Nombre de archivo:': %s", absolute_path)
            file_edit.SetValue(str(absolute_path))

            # Localizar y click boton "Abrir" (control_type='Button', title='Abrir')
            open_button = dialog.ButtonControl(Name="Abrir")
            if not open_button.Exists(3, 0.5):
                self._logger.error("Boton 'Abrir' no encontrado")
                return False

            self._logger.debug("Haciendo click en boton 'Abrir'")
            open_button.Click()

            # Esperar a que el dialogo se cierre (max 5 segundos)
            if dialog.WaitForExist(5, 0.5):
                self._logger.warning("Dialogo 'Abrir' no se cerro tras click en 'Abrir'")
                return False

            self._logger.info("Archivo abierto exitosamente via dialogo: %s", absolute_path)
            return True

        except Exception as e:
            self._logger.exception("Error en open_file_dialog: %s", e)
            return False

    def save_file_dialog(self, file_path: Path) -> bool:
        """Inyecta la ruta en el dialogo 'Guardar como', click 'Guardar', maneja reemplazo.

        Espera a que el dialogo 'Guardar como' este disponible, inyecta la ruta
        absoluta en el campo 'Nombre de archivo:', hace click en 'Guardar' y
        maneja automaticamente el modal de confirmacion de reemplazo si aparece.

        Args:
            file_path: Ruta donde guardar el archivo (relativa o absoluta).

        Returns:
            True si el archivo se guardo correctamente, False en caso contrario.
        """
        self._logger.info("Iniciando save_file_dialog para: %s", file_path)
        try:
            # Localizar dialogo "Guardar como" - esperar hasta 5 segundos
            dialog = auto.WindowControl(searchDepth=1, Name="Guardar como")
            if not dialog.WaitForExist(5, 0.5):
                self._logger.error("Dialogo 'Guardar como' no aparecio tras 5 segundos")
                return False

            self._logger.debug("Dialogo 'Guardar como' detectado: %s", dialog.Name)

            # Localizar campo "Nombre de archivo:" (control_type='Edit', title='Nombre de archivo:')
            file_edit = dialog.EditControl(Name="Nombre de archivo:")
            if not file_edit.Exists(3, 0.5):
                self._logger.error("Campo 'Nombre de archivo:' no encontrado en dialogo Guardar como")
                return False

            # Inyectar ruta absoluta
            absolute_path = file_path.resolve()
            self._logger.debug("Inyectando ruta absoluta en campo 'Nombre de archivo:': %s", absolute_path)
            file_edit.SetValue(str(absolute_path))

            # Localizar y click boton "Guardar" (control_type='Button', title='Guardar')
            save_button = dialog.ButtonControl(Name="Guardar")
            if not save_button.Exists(3, 0.5):
                self._logger.error("Boton 'Guardar' no encontrado")
                return False

            self._logger.debug("Haciendo click en boton 'Guardar'")
            save_button.Click()

            # Manejar posible modal de confirmacion de reemplazo (title='Confirmar guardado', button='Si')
            if self.handle_replace_modal():
                self._logger.info("Modal de reemplazo manejado correctamente")
            else:
                self._logger.debug("No aparecio modal de reemplazo (archivo nuevo o usuario cancelo)")

            # Esperar a que el dialogo se cierre (max 5 segundos)
            if dialog.WaitForExist(5, 0.5):
                self._logger.warning("Dialogo 'Guardar como' no se cerro tras click en 'Guardar'")
                return False

            self._logger.info("Archivo guardado exitosamente via dialogo: %s", absolute_path)
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
            replace_modal = auto.WindowControl(searchDepth=1, Name="Confirmar guardado")
            if not replace_modal.WaitForExist(3, 0.5):
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
            if replace_modal.WaitForExist(3, 0.5):
                self._logger.warning("Modal 'Confirmar guardado' no se cerro tras click en 'Si'")
                return False

            self._logger.info("Confirmacion de reemplazo ejecutada correctamente")
            return True

        except Exception as e:
            self._logger.exception("Error manejando modal de reemplazo: %s", e)
            return False