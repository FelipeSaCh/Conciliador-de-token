import re
import unicodedata

from core.config import CARACTERES_ESPECIALES

PATRON_CARACTERES_ESPECIALES = f"[{re.escape(''.join(CARACTERES_ESPECIALES))}]"
VALORES_NULOS = ('nan', 'none', 'nat', 'null')
CARACTERES_INVALIDOS_ARCHIVO = '\\/:*?"<>|'

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}


def normalizar_ascii(texto):
    return unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode().strip().lower()


def clave_columna(nombre):
    return "".join(ch for ch in normalizar_ascii(nombre) if ch.isalnum())


def limpiar_codigo(serie, strip=True):
    """Texto sin el sufijo '.0' que deja Excel en números enteros."""
    s = serie.fillna('').astype(str).str.replace(r'\.0$', '', regex=True)
    return s.str.strip() if strip else s


def normalizar_nit(valor):
    s = str(valor).strip()
    if s.lower() in ('', 'nan', 'none', 'nat'):
        return ''
    s = re.sub(r'^(\d+)\.0+$', r'\1', s)
    s = s.split('-')[0]
    return re.sub(r'\D', '', s)


def normalizar_num_ext(serie):
    """Quita caracteres especiales y SOLO los ceros a la izquierda (evita falsos duplicados)."""
    return (
        serie.fillna('').astype(str)
        .str.replace(PATRON_CARACTERES_ESPECIALES, '', regex=True)
        .str.strip()
        .str.replace(r'^0+', '', regex=True)
    )


def tiene_caracter_especial(serie):
    return serie.fillna('').astype(str).str.contains(PATRON_CARACTERES_ESPECIALES, regex=True, na=False)


def mascara_duplicados(serie_num_ext):
    limpio = normalizar_num_ext(serie_num_ext)
    return limpio.duplicated(keep=False) & (limpio != '')


def nombre_archivo_seguro(nombre):
    return "".join(ch for ch in nombre if ch not in CARACTERES_INVALIDOS_ARCHIVO)
