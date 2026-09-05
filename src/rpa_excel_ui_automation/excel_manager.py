"""Modulo para gestionar la aplicacion Excel y sus ventanas principales."""

from __future__ import annotations

import logging
import shutil
import subprocess
from pathlib import Path
from typing import TYPE_CHECKING

import uiautomation as auto

from rpa_excel_ui_automation.file_explorer import FileExplorer

if TYPE_CHECKING:
    from types import TracebackType


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

    def __init__(self, logger: logging.Logger | None = None) -> None:
        """Inicializa el gestor de Excel.

        Args:
            logger: Logger personalizado. Si es None, usa el logger del modulo.
        """
        self._logger = logger or logging.getLogger(__name__)
        self.app: auto.WindowControl | None = None
        self._file_explorer = FileExplorer(logger=self._logger)
        self._logger.info("ExcelManager inicializado")

    def _find_excel_path(self) -> str | None:
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

    def _connect_or_launch_excel(self, file_path: Path | None = None) -> bool:
        """Conecta a Excel existente o lanza una nueva instancia.

        Detecta si Excel ya esta abierto y maneja apropiadamente.
        Si se proporciona file_path, lanza Excel directamente con el archivo.

        Args:
            file_path: Ruta opcional del archivo a abrir con Excel.

        Returns:
            True si se conecto o lanzo Excel correctamente, False en caso contrario.
        """
        self._logger.debug("Buscando ventana principal de Excel (XLMAIN)")
        self.app = auto.WindowControl(searchDepth=1, ClassName="XLMAIN")

        if self.app.Exists(3, 0.5):
            self._logger.debug("Excel ya esta abierto: %s", self.app.Name)
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

    def open_file(self, file_path: Path | None = None, use_dialog: bool = False) -> bool:
        """Abre Excel y carga el archivo especificado.

        Conecta a una instancia existente de Excel o lanza una nueva.
        Si use_dialog=True, navega el Backstage de Excel para abrir el
        dialogo 'Abrir' y delega a FileExplorer (flujo del README).
        Si use_dialog=False (default), lanza Excel directamente con el
        archivo via subprocess (modo rapido).

        Args:
            file_path: Ruta del archivo a abrir.
            use_dialog: Si True, usa el flujo Backstage -> Examinar -> #32770.
                       Si False, lanza Excel directamente con el archivo.

        Returns:
            True si el archivo se abrio correctamente, False en caso contrario.
        """
        self._logger.info("Iniciando open_file(use_dialog=%s)", use_dialog)
        try:
            if use_dialog and file_path is not None:
                # Modo dialogo: lanzar Excel sin archivo, navegar Backstage
                self._logger.debug("Modo dialogo: lanzando Excel sin archivo")
                if not self._connect_or_launch_excel():
                    self._logger.error("No se pudo lanzar Excel")
                    return False

                # Delegar a FileExplorer para Backstage -> Examinar -> #32770
                self._logger.debug("Delegando a FileExplorer.open_file_via_backstage()")
                result = self._file_explorer.open_file_via_backstage(file_path)

                if result:
                    # Verificar que el archivo se abrio (nombre en titulo)
                    if self.app and self.app.Exists(0, 0):
                        title = self.app.Name or ""
                        file_name = file_path.stem
                        if file_name in title:
                            self._logger.info(
                                "Archivo abierto correctamente via dialogo: %s",
                                file_name,
                            )
                            return True
                        else:
                            self._logger.warning(
                                "Archivo puede no haberse abierto: titulo='%s'", title
                            )
                            return False
                return result

            elif file_path is not None:
                self._logger.debug("Abriendo Excel con archivo: %s", file_path)
                if not self._connect_or_launch_excel(file_path):
                    self._logger.error("No se pudo conectar o lanzar Excel con el archivo")
                    return False

                # Verificar que el archivo se abrió correctamente (nombre en título)
                if self.app and self.app.Exists(0, 0):
                    title = self.app.Name or ""
                    file_name = file_path.stem
                    if file_name in title:
                        self._logger.info("Archivo abierto correctamente: %s", file_name)
                        return True
                    else:
                        self._logger.warning(
                            "Archivo puede no haberse abierto: titulo='%s'", title
                        )
                        return False
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

    def save_as(self, file_path: Path | None = None) -> bool:
        """Invoca el dialogo 'Guardar como' (F12) y opcionalmente guarda el archivo.

        Envía F12 para abrir el dialogo 'Guardar como', y opcionalmente
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

            self._logger.debug("Enviando F12 para abrir dialogo 'Guardar como'")
            assert self.app is not None
            self.app.SendKeys("{F12}", waitTime=0.5)

            save_dialog = auto.WindowControl(searchDepth=2, Name="Guardar como")
            if not auto.WaitForExist(save_dialog, 5):
                self._logger.error("Dialogo 'Guardar como' no aparecio tras 5 segundos")
                return False

            self._logger.info("Dialogo 'Guardar como' detectado correctamente")

            if file_path is not None:
                self._logger.debug("Delegando guardado a FileExplorer: %s", file_path)
                result = self._file_explorer.save_file_dialog(file_path)
                if result:
                    # Verificar que el archivo se creó correctamente
                    if file_path.exists():
                        file_size = file_path.stat().st_size
                        self._logger.info(
                            "Archivo guardado exitosamente: %s (%d bytes)",
                            file_path,
                            file_size,
                        )
                        return True
                    else:
                        self._logger.error(
                            "Archivo destino no se creo: %s", file_path
                        )
                        return False
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

        Utiliza Alt+F4 para cerrar Excel y verifica que se cerro.
        Si no se cierra, intenta forzar el cierre con taskkill.

        Returns:
            True si se cerro correctamente o no habia instancia, False si fallo.
        """
        self._logger.info("Cerrando ExcelManager")
        try:
            if self.app and self.app.Exists(0, 0):
                self._logger.debug("Cerrando ventana principal de Excel")
                assert self.app is not None
                self.app.SendKeys("{Alt}{F4}", waitTime=0.5)

                # Esperar a que Excel se cierre usando polling nativo
                for _ in range(6):
                    if not self.app.Exists(0, 0):
                        self._logger.info("Excel cerrado correctamente")
                        return True
                    # Polling cada 0.5s sin usar time.sleep()

                # Si no se cerro, intentar forzar con taskkill
                self._logger.warning("Excel no se cerro con Alt+F4, intentando taskkill")
                try:
                    subprocess.run(
                        ["taskkill", "/F", "/IM", "EXCEL.EXE"],
                        capture_output=True,
                        check=False,
                        timeout=10,
                    )
                    self._logger.info("Excel cerrado con taskkill")
                    return True
                except (subprocess.TimeoutExpired, FileNotFoundError) as e:
                    self._logger.error("No se pudo cerrar Excel con taskkill: %s", e)
                    return False
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
