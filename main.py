import sys

from errors import logger, registrar_excepcion_no_controlada


def main():
    sys.excepthook = registrar_excepcion_no_controlada
    logger.info("Iniciando aplicación")

    from app import ConciliadorApp

    app = ConciliadorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
