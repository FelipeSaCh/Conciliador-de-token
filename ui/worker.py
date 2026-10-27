"""Ejecución de servicios en segundo plano con despacho seguro de callbacks al hilo de Tkinter."""
import queue
import threading

from core.errors import ErrorSistema, ErrorUsuario, logger


class EjecutorSegundoPlano:
    def __init__(self, root, intervalo_ms=150):
        self._root = root
        self._cola = queue.Queue()
        self._intervalo = intervalo_ms
        root.after(intervalo_ms, self._procesar)

    def en_ui(self, callback):
        """Envuelve `callback` para que, llamado desde cualquier hilo, se ejecute en el hilo de la UI."""
        return lambda *args: self._cola.put((callback, args))

    def lanzar(self, tarea, on_exito, on_error, contexto="tarea"):
        """Ejecuta `tarea()` en un hilo. on_error(mensaje, es_error_sistema)."""
        def _correr():
            try:
                resultado = tarea()
            except ErrorUsuario as e:
                logger.warning(f"Error de usuario en {contexto}: {e}")
                self._cola.put((on_error, (str(e), False)))
            except ErrorSistema as e:
                logger.error(f"Error de sistema en {contexto}: {e}")
                self._cola.put((on_error, (str(e), True)))
            except Exception as e:
                logger.exception(f"Error inesperado en {contexto}")
                self._cola.put((on_error, (f"Error inesperado: {e}", True)))
            else:
                self._cola.put((on_exito, (resultado,)))

        threading.Thread(target=_correr, daemon=True).start()

    def _procesar(self):
        try:
            while True:
                try:
                    callback, args = self._cola.get_nowait()
                except queue.Empty:
                    break
                try:
                    callback(*args)
                except Exception:
                    logger.exception("Error en callback de UI")
        finally:
            self._root.after(self._intervalo, self._procesar)
