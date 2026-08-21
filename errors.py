import logging
import traceback

from config import LOG_DIR, LOG_FILE


class ConciliacionError(Exception):
    """Excepción base del dominio de conciliación."""


class ErrorUsuario(ConciliacionError):
    """Error causado por una acción o dato ingresado por el usuario."""


class ErrorSistema(ConciliacionError):
    """Error inesperado del sistema durante el procesamiento."""


class HojaNoEncontradaError(ErrorUsuario):
    def __init__(self, nombre_hoja, hojas_disponibles):
        self.nombre_hoja = nombre_hoja
        self.hojas_disponibles = hojas_disponibles
        super().__init__(
            f"No se encontró la hoja '{nombre_hoja}' en el archivo.\n"
            f"Hojas disponibles: {', '.join(hojas_disponibles)}"
        )


class ColumnaFaltanteError(ErrorUsuario):
    def __init__(self, nombre_hoja, columna):
        self.nombre_hoja = nombre_hoja
        self.columna = columna
        super().__init__(
            f"La hoja '{nombre_hoja}' no contiene la columna requerida '{columna}'"
        )


def configurar_logger():
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("conciliador")
    logger.setLevel(logging.DEBUG)
    if logger.handlers:
        return logger

    fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    fh.setFormatter(formatter)
    logger.addHandler(fh)
    return logger


logger = configurar_logger()


def registrar_excepcion_no_controlada(exc_type, exc_value, exc_traceback):
    mensaje = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    logger.critical(f"Excepción no controlada:\n{mensaje}")
