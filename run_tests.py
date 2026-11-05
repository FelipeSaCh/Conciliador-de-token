"""Ejecuta las pruebas: python run_tests.py  |  con el listado: set AUTORRETENEDORES_JSON=ruta\\autorretenedores.json"""
import sys
import unittest

if __name__ == "__main__":
    resultado = unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.discover("tests", top_level_dir="."))
    sys.exit(not resultado.wasSuccessful())