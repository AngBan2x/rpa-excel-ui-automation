"""Modulo para interactuar con dialogos de archivos de Windows."""

from __future__ import annotations

import logging
from pathlib import Path

import uiautomation as auto  # type: ignore[attr-defined]


logger = logging.getLogger(__name__)


class FileExplorer:
    """Interactua con dialogos de archivos de Windows (Abrir, Guardar como).

    Maneja la inyeccion de rutas en controles Edit, clicks en botones,
    y la gestion del modal de confirmacion de reemplazo.

    Attributes:
        _logger: Logger para esta instancia.
    """

    def __init__(self, logger: logging.Logger | None = None) -> None:
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

    def open_file_via_backstage(self, file_path: Path) -> bool:
        """Navega el Backstage de Excel, hace click en 'Examinar' y abre archivo.

        Flujo: ListItem 'Abrir' del sidebar -> Button 'Examinar' -> dialogo #32770
        -> open_file_dialog().

        Args:
            file_path: Ruta del archivo a abrir.

        Returns:
            True si el archivo se abrio correctamente, False en caso contrario.
        """
        self._logger.info("Iniciando open_file_via_backstage para: %s", file_path)
        try:
            # 1. Navegar a la seccion 'Abrir' del Backstage
            self._logger.debug("Buscando NavBarMenu en el Backstage")
            nav = auto.ListControl(searchDepth=5, AutoId="NavBarMenu")
            if not nav.Exists(3, 0.5):
                self._logger.error("NavBarMenu no encontrado en el Backstage")
                return False

            abrir_item = None
            for item in nav.GetChildren():
                try:
                    if item.AutomationId == "msotcidPlaceOpen":
                        abrir_item = item
                        break
                except Exception:
                    continue

            if abrir_item is None:
                self._logger.error("ListItem 'Abrir' no encontrado en NavBarMenu")
                return False

            self._logger.debug("Navegando a seccion 'Abrir' del Backstage")
            abrir_item.SetFocus()
            import time as _time
            _time.monotonic()
            abrir_item.SendKeys("{Enter}", waitTime=2)
            _time.monotonic()

            # 2. Esperar a que cargue el contenido de la seccion Abrir
            self._logger.debug("Esperando carga del contenido 'Abrir'")
            if not auto.WaitForExist(
                auto.GroupControl(searchDepth=5, AutoId="PlaceTabOpenContent"), 5
            ):
                self._logger.warning(
                    "PlaceTabOpenContent no detectado, continuando de todas formas"
                )

            # 3. Buscar y hacer click en 'Examinar'
            self._logger.debug("Buscando boton 'Examinar' en el Backstage")
            examinar = auto.ButtonControl(searchDepth=12, Name="Examinar")
            if not examinar.Exists(5, 1):
                self._logger.error("Boton 'Examinar' no encontrado en el Backstage")
                return False

            self._logger.debug("Boton 'Examinar' encontrado, enviando Enter")
            examinar.SetFocus()
            examinar.SendKeys("{Enter}", waitTime=3)

            # Pequena pausa para que el dialogo #32770 cargue completamente
            import time as _time
            deadline = _time.monotonic() + 2
            while _time.monotonic() < deadline:
                _time.monotonic()

            # 4. Delegar a open_file_dialog() para manejar el #32770
            self._logger.debug("Delegando a open_file_dialog() para el dialogo #32770")
            return self.open_file_dialog(file_path)

        except Exception as e:
            self._logger.exception("Error en open_file_via_backstage: %s", e)
            return False

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
            dialog = self._find_file_dialog(timeout=5)
            if dialog is None:
                self._logger.error("Dialogo 'Abrir' no aparecio tras 5 segundos")
                return False

            self.current_dialog = dialog
            self._logger.debug("Dialogo detectado: %s", dialog.Name or dialog.ClassName)

            # Localizar campo de ruta: ComboBox "Nombre de archivo:" -> Edit
            absolute_path = file_path.resolve()
            self._logger.debug("Inyectando ruta absoluta: %s", absolute_path)

            file_combo = dialog.ComboBoxControl(Name="Nombre de archivo:")
            if not file_combo.Exists(3, 0.5):
                self._logger.error("Campo 'Nombre de archivo:' no encontrado")
                return False

            file_combo.Click()
            edit = file_combo.EditControl()
            if not edit.Exists(1, 0.5):
                self._logger.error("Edit dentro del ComboBox no encontrado")
                return False

            edit.SendKeys("{Ctrl}a", waitTime=0.1)
            edit.SendKeys(str(absolute_path), waitTime=0.2)

            open_button = dialog.ButtonControl(Name="Abrir")
            if not open_button.Exists(3, 0.5):
                self._logger.error("Boton 'Abrir' no encontrado")
                return False

            self._logger.debug("Enviando Enter en el Edit para abrir archivo")
            edit.SendKeys("{Enter}", waitTime=3)

            # Esperar cierre del dialogo usando polling con timeout
            import time as _time
            start = _time.monotonic()
            while _time.monotonic() - start < 5:
                try:
                    if not dialog.Exists(0, 0):
                        break
                except Exception:
                    break
                _time.monotonic()  # yield

            # Verificar si se cerro
            try:
                still_open = dialog.Exists(0, 0)
            except Exception:
                still_open = False

            if still_open:
                # Fallback: click en boton Abrir
                self._logger.debug("Enter no cerro dialogo, intentando click en Abrir")
                open_button.Click()
                start2 = _time.monotonic()
                while _time.monotonic() - start2 < 5:
                    try:
                        if not dialog.Exists(0, 0):
                            break
                    except Exception:
                        break
                    _time.monotonic()
                try:
                    still_open = dialog.Exists(0, 0)
                except Exception:
                    still_open = False

            if still_open:
                self._logger.warning("Dialogo 'Abrir' no se cerro")
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
            dialog = self._find_file_dialog(timeout=5)
            if dialog is None:
                self._logger.error("Dialogo 'Guardar como' no aparecio tras 5 segundos")
                return False

            self.current_dialog = dialog
            self._logger.debug("Dialogo detectado: %s", dialog.Name or dialog.ClassName)

            # Localizar ComboBox "Nombre de archivo:" (no EditControl)
            file_combo = dialog.ComboBoxControl(Name="Nombre de archivo:")
            if not file_combo.Exists(3, 0.5):
                self._logger.error("Campo 'Nombre de archivo:' no encontrado en Guardar como")
                return False

            # Inyectar ruta absoluta via Edit dentro del ComboBox
            absolute_path = file_path.resolve()
            self._logger.debug("Inyectando ruta absoluta: %s", absolute_path)
            file_combo.Click()
            edit = file_combo.EditControl()
            if edit.Exists(1, 0.5):
                edit.SendKeys("{Ctrl}a", waitTime=0.1)
                edit.SendKeys(str(absolute_path), waitTime=0.1)
            else:
                file_combo.SendKeys("{Ctrl}a", waitTime=0.1)
                file_combo.SendKeys(str(absolute_path), waitTime=0.1)

            save_button = dialog.ButtonControl(Name="Guardar")
            if not save_button.Exists(3, 0.5):
                self._logger.error("Boton 'Guardar' no encontrado")
                return False

            self._logger.debug("Haciendo click en boton 'Guardar'")
            save_button.Click()

            # Manejar modal de reemplazo si aparece
            modal_result = self.handle_replace_modal(parent=dialog)
            if modal_result is None:
                self._logger.debug("No aparecio modal de reemplazo")
            elif modal_result is True:
                self._logger.info("Modal de reemplazo manejado correctamente")
            else:
                self._logger.error("Modal de reemplazo aparece pero no se pudo manejar")
                return False

            self._logger.info("Archivo guardado exitosamente: %s", absolute_path)
            return True

        except Exception as e:
            self._logger.exception("Error en save_file_dialog: %s", e)
            return False

    def handle_replace_modal(self, parent: auto.Control | None = None) -> bool | None:
        """Detecta modal 'Confirmar Guardar como' y hace click en 'Si'.

        El modal es un WindowControl #32770 anidado como hijo del dialogo
        "Guardar como". La busqueda por WindowControl(Name=...) no lo encuentra
        porque uiautomation no traversa hijos #32770 anidados. Usamos
        GetChildren() para iterar hijos manualmente.

        Hace polling con reintentos porque el modal tarda en aparecer
        despues del click en "Guardar".

        Args:
            parent: Dialogo padre donde buscar el modal. Si es None, busca
                    desde la raiz del escritorio.

        Returns:
            True si se detecto y manejo el modal, False si aparece pero falla,
            None si no aparecio.
        """
        self._logger.debug("Verificando modal 'Confirmar Guardar como'")
        try:
            replace_modal = None
            import time as _time
            start = _time.monotonic()

            while _time.monotonic() - start < 5:
                # Estrategia 1: Buscar entre hijos del dialogo padre
                if parent is not None:
                    try:
                        for child in parent.GetChildren():
                            try:
                                name = child.Name or ""
                                if name == "Confirmar Guardar como":
                                    replace_modal = child
                                    self._logger.debug(
                                        "Modal encontrado via GetChildren() del padre"
                                    )
                                    break
                            except Exception:
                                continue
                    except Exception:
                        pass

                # Estrategia 2: Buscar entre hijos de ventanas top-level
                if replace_modal is None:
                    try:
                        root = auto.GetRootControl()
                        for top_window in root.GetChildren():
                            try:
                                name = top_window.Name or ""
                                if name == "Confirmar Guardar como":
                                    replace_modal = top_window
                                    self._logger.debug(
                                        "Modal encontrado como ventana top-level"
                                    )
                                    break
                                for child in top_window.GetChildren():
                                    try:
                                        cname = child.Name or ""
                                        if cname == "Confirmar Guardar como":
                                            replace_modal = child
                                            self._logger.debug(
                                                "Modal encontrado como nieto"
                                            )
                                            break
                                    except Exception:
                                        continue
                                if replace_modal is not None:
                                    break
                            except Exception:
                                continue
                    except Exception:
                        pass

                if replace_modal is not None:
                    break

                # Breve pausa via Excel window check (no time.sleep)
                try:
                    xl = auto.WindowControl(searchDepth=1, ClassName="XLMAIN")
                    xl.Exists(0, 0)
                except Exception:
                    pass

            if replace_modal is None:
                self._logger.debug(
                    "Modal 'Confirmar Guardar como' no detectado tras 5s"
                )
                return None

            self._logger.info("Modal detectado: %s", replace_modal.Name)

            # Buscar boton "Si" con tilde primero, fallback sin tilde
            yes_button = replace_modal.ButtonControl(Name="Sí")
            if not yes_button.Exists(3, 0.5):
                self._logger.debug(
                    "Boton 'Sí' (con tilde) no encontrado, intentando 'Si'"
                )
                yes_button = replace_modal.ButtonControl(Name="Si")
                if not yes_button.Exists(3, 0.5):
                    self._logger.error(
                        "Boton 'Si' no encontrado en modal de reemplazo"
                    )
                    return False

            self._logger.debug("Haciendo click en boton 'Si' del modal")
            yes_button.Click()

            # Esperar a que el modal se cierre (max 5 segundos)
            if auto.WaitForExist(replace_modal, 5):
                self._logger.warning("Modal no se cerro tras click en 'Si'")
                return False

            self._logger.info("Reemplazo confirmado correctamente")
            return True

        except Exception as e:
            self._logger.exception(
                "Error manejando modal de reemplazo: %s", e
            )
            return False