import importlib.util
import os
import shutil
import tempfile
from contextlib import contextmanager
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from core.errors import ErrorSistema, ErrorUsuario, logger
from shared.xlsx_cache import restaurar_resultados

MOTOR_RAPIDO = (
    "calamine"
    if importlib.util.find_spec("python_calamine") is not None
    and tuple(int(p) for p in pd.__version__.split(".")[:2]) >= (2, 2)
    else None
)


def leer_excel(ruta, rapido=False, **kwargs):
    return pd.read_excel(ruta, engine=MOTOR_RAPIDO if rapido else None, **kwargs)


def hojas_de(ruta):
    with pd.ExcelFile(ruta) as xls:
        return list(xls.sheet_names)


def cargar_hoja_encabezado_variable(ruta, hoja):
    """Usa como encabezado la primera fila (de las 10 primeras) con 3+ valores no nulos."""
    try:
        df_temp = pd.read_excel(ruta, sheet_name=hoja, header=None, nrows=10)
        filas_validas = df_temp.dropna(thresh=3)
        header_idx = filas_validas.index[0] if not filas_validas.empty else 0
        return pd.read_excel(ruta, sheet_name=hoja, header=header_idx)
    except Exception as e:
        raise ErrorSistema(f"Error leyendo la hoja '{hoja}' con encabezado variable: {e}") from e


def leer_hoja_por_palabra_clave(ruta, hoja, keyword="NIT", max_filas=50):
    """Usa como encabezado la primera fila que contenga exactamente `keyword`."""
    df = pd.read_excel(ruta, sheet_name=hoja, header=None)
    keyword = keyword.upper()
    header_idx = next(
        (i for i in range(min(max_filas, len(df)))
         if keyword in df.iloc[i].astype(str).str.strip().str.upper().values),
        0,
    )
    df.columns = df.iloc[header_idx]
    return df.iloc[header_idx + 1:].reset_index(drop=True)


# ---------------------------------------------------------------- escritura segura con openpyxl
VALORES_FORMULA = "_valores_formula"  # atributo del Workbook: {(hoja, celda): valor} de fórmulas nuevas


def registrar_valor_formula(ws, celda, valor):
    """Guarda el resultado de una fórmula escrita por el programa para que el archivo lo muestre sin recalcular."""
    wb = ws.parent
    if not hasattr(wb, VALORES_FORMULA):
        setattr(wb, VALORES_FORMULA, {})
    getattr(wb, VALORES_FORMULA)[(ws.title, celda)] = valor


def _kwargs_libro(ruta):
    return {'keep_vba': True} if Path(ruta).suffix.lower() == '.xlsm' else {}


def _preparar_vista(wb, hoja_activa=None):
    """Una sola pestaña seleccionada y visible (evita libros que abren en modo [Grupo] o sobre una hoja oculta)."""
    visibles = [ws for ws in wb.worksheets if ws.sheet_state == 'visible']
    if not visibles:
        wb.worksheets[0].sheet_state = 'visible'
        visibles = [wb.worksheets[0]]
    activa = wb[hoja_activa] if hoja_activa and hoja_activa in wb.sheetnames else wb.active
    if activa is None or activa.sheet_state != 'visible':
        activa = visibles[0]
    for ws in wb.worksheets:
        for vista in ws.views.sheetView:
            vista.tabSelected = ws is activa
    wb.active = wb.worksheets.index(activa)
    wb.calculation.fullCalcOnLoad = True


def _guardar_seguro(ruta, escribir, hoja_activa=None):
    """Trabaja sobre una copia temporal y solo reemplaza el original si todo terminó bien.

    escribir(ruta_temporal) -> Workbook ya guardado en la copia. Si algo falla, el archivo original queda intacto
    (antes se guardaba un libro a medio escribir y las hojas quedaban vacías). Al final restaura los resultados de las
    fórmulas que openpyxl deja en blanco.
    """
    ruta = Path(ruta)
    fd, tmp = tempfile.mkstemp(suffix=ruta.suffix, prefix="~tmp_", dir=ruta.parent)
    os.close(fd)
    try:
        shutil.copy2(ruta, tmp)
        wb = escribir(tmp)
        try:
            restaurar_resultados(ruta, tmp, getattr(wb, VALORES_FORMULA, {}))
        except Exception as e:  # nunca impide guardar: solo afecta la vista previa de fórmulas
            logger.warning(f"No se pudieron restaurar los resultados de fórmulas: {e}")
        shutil.copymode(ruta, tmp)
        try:
            os.replace(tmp, ruta)
        except PermissionError as e:
            raise ErrorUsuario(f"No se pudo guardar '{ruta.name}': ciérralo en Excel e inténtalo de nuevo.") from e
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


@contextmanager
def escritor_excel(ruta, hoja_activa=None):
    """ExcelWriter (openpyxl, modo append, reemplaza hojas, conserva macros) con guardado seguro.

    Uso igual que antes: `with escritor_excel(ruta) as writer: df.to_excel(writer, ...)`.
    """
    ruta = Path(ruta)
    fd, tmp = tempfile.mkstemp(suffix=ruta.suffix, prefix="~tmp_", dir=ruta.parent)
    os.close(fd)
    writer = None
    try:
        shutil.copy2(ruta, tmp)
        writer = pd.ExcelWriter(tmp, engine='openpyxl', mode='a', if_sheet_exists='replace',
                                engine_kwargs=_kwargs_libro(ruta))
        yield writer
        _preparar_vista(writer.book, hoja_activa)
        writer.close()
        try:
            restaurar_resultados(ruta, tmp, getattr(writer.book, VALORES_FORMULA, {}))
        except Exception as e:
            logger.warning(f"No se pudieron restaurar los resultados de fórmulas: {e}")
        shutil.copymode(ruta, tmp)
        try:
            os.replace(tmp, ruta)
        except PermissionError as e:
            raise ErrorUsuario(f"No se pudo guardar '{ruta.name}': ciérralo en Excel e inténtalo de nuevo.") from e
    finally:
        if writer is not None:
            try:
                writer._handles.close()  # libera el archivo temporal si hubo error (no guarda)
            except Exception:
                pass
        if os.path.exists(tmp):
            os.remove(tmp)


def guardar_libro_openpyxl(ruta, modificar, hoja_activa=None):
    """Abre el libro con openpyxl, aplica modificar(wb) y lo guarda de forma segura (ver _guardar_seguro)."""
    def escribir(tmp):
        wb = load_workbook(tmp, **_kwargs_libro(ruta))
        modificar(wb)
        _preparar_vista(wb, hoja_activa)
        wb.save(tmp)
        return wb
    _guardar_seguro(ruta, escribir, hoja_activa)