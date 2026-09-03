"""Modulo para gestionar la aplicacion Excel y sus ventanas principales."""

from __future__ import annotations

import logging
import subprocess
import shutil
from pathlib import Path
from typing import Optional
from types import TracebackType

import uiautomation as auto

from rpa_excel_ui_automation.file_explorer import FileExplorer


logger = logging.getLogger(__name__)


class ExcelManager:
    """Gestiona la instancia de Excel y eventos principales de ventana.

    Esta clase se encarga de conectar o lanzar Excel, y de invocar los
    dialogos principales (Abrir, Guardar como) delegando la interaccion
    con los dialogos a FileExplorer.

    Attributes:
        app: Ventana principal de Excel (XLMAIN) si esta conectada.
        _file_explorer: Instancia de FileExplorer para manejar dialogos.
    """

    def __init__(self, logger: Optional[logging.Logger] = None) -> None:
        """Inicializa el gestor de Excel.

        Args:
            logger: Logger personalizado. Si es None, usa el logger del modulo.
        """
        self._logger = logger or logging.getLogger(__name__)
        self.app: Optional[auto.WindowControl] = None
        self._file_explorer = FileExplorer(logger=self._logger)
        self._logger.info("ExcelManager inicializado")

    def _find_excel_path(self) -> Optional[str]:
        """Busca la ruta del ejecutable de Excel.

        Returns:
            Ruta del ejecutable de Excel o None si no se encontro.
        """
        excel_path = shutil.which("excel.exe")
        if excel_path is not None:
            return excel_path

        known_paths = [
            r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE",
            r"C:\Program Files\Microsoft Office\Office16\EXCEL.EXE",
        ]
        for path in known_paths:
            if Path(path).exists():
                return path

        return None

    def _connect_or_launch_excel(self, file_path: Optional[Path] = None) -> bool:
        """Conecta a Excel existente o lanza una nueva instancia.

        Args:
            file_path: Ruta opcional del archivo a abrir con Excel.

        Returns:
            True si se conecto o lanzo Excel correctamente, False en caso contrario.
        """
        self._logger.debug("Buscando ventana principal de Excel (XLMAIN)")
        self.app = auto.WindowControl(searchDepth=1, ClassName="XLMAIN")

        if self.app.Exists(3, 0.5):
            self._logger.debug("Ventana Excel principal encontrada: %s", self.app.Name)
            return True

        self._logger.info("Excel no esta corriendo, lanzando nueva instancia")
        try:
            excel_path = self._find_excel_path()
            if excel_path is None:
                self._logger.error("No se encontro 'excel.exe' en el PATH ni en rutas conocidas")
                return False

            if file_path is not None:
                self._logger.debug("Lanzando Excel con archivo: %s", file_path)
                subprocess.Popen([excel_path, str(file_path.resolve())])
            else:
                self._logger.debug("Lanzando Excel sin archivo")
                subprocess.Popen([excel_path])
        except Exception as e:
            self._logger.exception("Error al lanzar Excel: %s", e)
            return False

        # Esperar a que aparezca la ventana principal
        self.app = auto.WindowControl(searchDepth=1, ClassName="XLMAIN")
        if not auto.WaitForExist(self.app, 15):
            self._logger.error("Excel no se inicio a tiempo (timeout 15s)")
            return False

        self._logger.debug("Ventana Excel principal encontrada tras lanzamiento: %s", self.app.Name)
        return True

    def open_file(self, file_path: Optional[Path] = None) -> bool:
        """Abre Excel y carga el archivo especificado.

        Conecta a una instancia existente de Excel o lanza una nueva,
        y carga el archivo directamente. No usa Ctrl+O porque en Excel
        moderno eso abre la vista Backstage.

        Args:
            file_path: Ruta del archivo a abrir. Si se proporciona,
                Excel se lanza directamente con el archivo.

        Returns:
            True si el archivo se abrio correctamente, False en caso contrario.
        """
        self._logger.info("Iniciando open_file()")
        try:
            if file_path is not None:
                self._logger.debug("Abriendo Excel con archivo: %s", file_path)
                if not self._connect_or_launch_excel(file_path):
                    self._logger.error("No se pudo conectar o lanzar Excel con el archivo")
                    return False
            else:
                self._logger.debug("Abriendo Excel sin archivo")
                if not self._connect_or_launch_excel():
                    self._logger.error("No se pudo conectar o lanzar Excel")
                    return False

            self._logger.info("open_file() completado exitosamente")
            return True

        except Exception as e:
            self._logger.exception("Error inesperado en open_file: %s", e)
            return False

    def _ensure_workbook_dirty(self) -> None:
        """Asegura que el libro tenga cambios sin guardar para que F12 funcione.

        En Excel moderno (Office 365/2016+), F12 solo abre el dialogo
        'Guardar como' cuando el libro tiene cambios sin guardar (dirty state).
        Si el archivo esta limpio (ya guardado), F12 no hace nada.

        Este metodo escribe un espacio en la celda actual para crear
        ese estado, y luego se cancela con Ctrl+Z. El resultado neto es que
        Excel detecta la deshacer accion y mantiene el dirty flag activo.

        Patron: Space -> Ctrl+Z deja el workbook en estado dirty porque
        Excel registra la accion de deshacer como un cambio pendiente.
        """
        import time

        self._logger.debug("Marcando libro como 'no guardado' para que F12 funcione")

        # Asegurar que el foco este en la hoja de calculo (no en Backstage/menu)
        self.app.SendKeys("{Escape}", waitTime=0.3)
        time.sleep(0.3)

        # Escribir un espacio en la celda actual para marcar como dirty
        self.app.SendKeys(" ", waitTime=0.3)
        time.sleep(0.3)

        # Deshacer el cambio pero mantener el dirty flag
        # (Ctrl+Z deja el workbook en estado dirty en Excel moderno)
        self.app.SendKeys("{Ctrl}z", waitTime=0.3)
        time.sleep(0.3)

        self._logger.debug("Libro marcado como 'no guardado'")

    def save_as(self, file_path: Optional[Path] = None) -> bool:
        """Invoca el dialogo 'Guardar como' (F12) y opcionalmente guarda el archivo.

        Asegura primero que el libro tenga cambios sin guardar (dirty state),
        ya que en Excel moderno F12 solo funciona cuando hay cambios pendientes.
        Luego envia F12 para abrir el dialogo 'Guardar como', y opcionalmente
        delega la inyeccion de ruta y guardado a FileExplorer.

        Args:
            file_path: Ruta opcional donde guardar el archivo. Si se proporciona,
                se delega a FileExplorer.save_file_dialog().

        Returns:
            True si el dialogo 'Guardar como' se abrio correctamente (y el archivo
            se guardo si se proporciono file_path), False en caso contrario.
        """
        self._logger.info("Iniciando save_as()")
        try:
            if not self.app or not self.app.Exists(0, 0):
                self._logger.error("No hay instancia de Excel activa")
                return False

            # Asegurar que el libro tenga cambios sin guardar
            # (necesario para que F12 funcione en Excel moderno)
            self._ensure_workbook_dirty()

            # Enviar F12 para "Guardar como"
            self._logger.debug("Enviando F12 para abrir dialogo 'Guardar como'")
            assert self.app is not None
            self.app.SendKeys("{F12}", waitTime=0.5)

            import time
            time.sleep(2)

            # Esperar dialogo "Guardar como" (hijo de XLMAIN, depth=2 desde desktop)
            save_dialog = auto.WindowControl(searchDepth=2, Name="Guardar como")
            if not auto.WaitForExist(save_dialog, 5):
                self._logger.error("Dialogo 'Guardar como' no aparecio tras 5 segundos")
                return False

            self._logger.info("Dialogo 'Guardar como' detectado correctamente")

            # Si se proporciona ruta, delegar a FileExplorer
            if file_path is not None:
                self._logger.debug("Delegando guardado a FileExplorer: %s", file_path)
                result = self._file_explorer.save_file_dialog(file_path)
                if result:
                    self._logger.info("Archivo guardado exitosamente: %s", file_path)
                else:
                    self._logger.error("Falló el guardado del archivo: %s", file_path)
                return result

            self._logger.info("save_as() completado exitosamente (solo dialogo)")
            return True

        except Exception as e:
            self._logger.exception("Error inesperado en save_as: %s", e)
            return False

    def close(self) -> bool:
        """Cierra la aplicacion Excel si esta abierta.

        Returns:
            True si se cerro correctamente o no habia instancia, False si fallo.
        """
        self._logger.info("Cerrando ExcelManager")
        try:
            if self.app and self.app.Exists(0, 0):
                self._logger.debug("Cerrando ventana principal de Excel")
                assert self.app is not None
                self.app.SendKeys("{Alt}f4", waitTime=0.5)
                # Verificar que se cerro
                if auto.WaitForExist(self.app, 3):
                    self._logger.warning("Excel no se cerro tras Alt+F4")
                    return False
                self._logger.info("Excel cerrado correctamente")
            else:
                self._logger.debug("No habia instancia de Excel activa para cerrar")
            return True
        except Exception as e:
            self._logger.exception("Error al cerrar Excel: %s", e)
            return False

    def __enter__(self) -> ExcelManager:
        """Permite uso como context manager."""
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: TracebackType | None,
    ) -> None:
        """Cierra Excel al salir del context manager."""
        self.close()
