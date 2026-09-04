from pathlib import Path
from version import __version__

APP_NAME = "Conciliador DIAN vs Contabilidad"
APP_VERSION = __version__

BASE_DIR = Path.home() / "ConciliadorDIAN"
LOG_DIR = BASE_DIR / "logs"
LOG_FILE = LOG_DIR / "conciliador.log"

DEFAULT_SHEET_NAMES = {
    "principal": "Token",
    "contabilidad": "contabilidad",
    "terceros": "TERCEROS",
    "aud_comp": "AUD-COMP",
    "autorretenedores": "AUTORRETENEDORES",
}

SHEET_LABELS = {
    "principal": "Hoja Principal (Token)",
    "contabilidad": "Hoja Contabilidad",
    "terceros": "Hoja Terceros",
    "aud_comp": "Hoja Auditoría Comprobantes (AUD-COMP)",
    "autorretenedores": "Hoja Autorretenedores (opcional)",
}

ORDEN_AUDITORIA = ["Tipo	Número", "Num.Ext", "Fecha", "Nit/C.C.", "Tercero","BASE","IVA", "Detalle"]
SHEET_ORDER = ["principal", "contabilidad", "terceros", "aud_comp", "autorretenedores"]
REQUIRED_SHEETS = ["principal", "contabilidad", "terceros", "aud_comp"]
OPTIONAL_SHEETS = ["autorretenedores","contabilidad","terceros"]

ORDEN = ['Tipo de documento', 'CUFE/CUDE', 'Folio', 'Prefijo', 'Num.Ext', 'Divisa', 'Forma de Pago',
         'Medio de Pago', 'Fecha Emisión', 'Fecha Recepción', 'NIT Emisor', 'Nombre Emisor',
         'NIT Receptor', 'Nombre Receptor', 'BASE', 'IVA', 'BASE_2', 'Total', 'Estado', 'Grupo']

COLUMNAS_DIAN_VS_CONT = [
    'Tipo de documento', 'CUFE/CUDE', 'Folio', 'Prefijo', 'Num.Ext', 'Divisa', 'Fecha',
    'Forma de Pago', 'Medio de Pago', 'Fecha Emisión', 'Fecha Recepción', 'NIT Emisor',
    'Nombre Emisor', 'NIT Receptor', 'Nombre Receptor', 'BASE', 'IVA','BASE_2', 'Total', 'Estado', 'Grupo'
]

COLUMNAS_EXCLUIDAS_AUD_COMP = ['Tipo', 'Número', 'Num.Ext', 'Fecha', 'Nit/C.C.', 'Tercero', 'Detalle','Sumas','IVA','BASE',"TIPO", "NUMERO", "CHEQ.NO",
                                "NO.EXT", "AÑO", "MES", "DIA", "NIT", "TERCERO", "SUC.PTO", "DOCRELA", "CCOSTO", "SCCOSTO", "DETALLE", "Unnamed: 17", 
                                "ELABORO", "REF1", "REF2", "REF3", "REF4", "No. AUTORIZACION(DATAFONO)"]

OUTPUT_SHEETS_TO_HIDE = ['Resultados', 'auditoria', 'resultados-auditoria']

RED_FILL_COLOR = "FFC7CE"
CURRENCY_FORMAT = '"$"#,##0'

PREVIEW_MAX_ROWS = 300
PREVIEW_MAX_COLS = 60

CATEGORY_COLORS = {
    "iva": "#EB6425",
    "base": "#16A34A",
    "base2": "#33EAEA",
    "autorretenedor": "#cc99ff",
}

CARACTERES_ESPECIALES = [
    ' ', '-', '_', '.', ',', ';', ':', '/', '\\', '|',
    '!', '"', '#', '$', '%', '&', '(', ')', '=', '?', 
    '¡', '¿', '+', '*', '~', '{', '}', '[', ']', '^', 
    '`', '´', "'", '<', '>', '@'
]