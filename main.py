import sys

from core.errors import logger, registrar_excepcion_no_controlada


def main():
    sys.excepthook = registrar_excepcion_no_controlada
    logger.info("Iniciando aplicación")

    from ui.app import ConciliadorApp

    ConciliadorApp().mainloop()


if __name__ == "__main__":
    main()
