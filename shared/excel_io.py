import importlib.util
from pathlib import Path

import pandas as pd

from core.errors import ErrorSistema

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


def escritor_excel(ruta):
    """ExcelWriter en modo append que reemplaza hojas y conserva macros en .xlsm."""
    ruta = Path(ruta)
    engine_kwargs = {'keep_vba': True} if ruta.suffix.lower() == '.xlsm' else {}
    return pd.ExcelWriter(ruta, engine='openpyxl', mode='a', if_sheet_exists='replace', engine_kwargs=engine_kwargs)
