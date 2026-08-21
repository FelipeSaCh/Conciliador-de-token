from pathlib import Path

APP_NAME = "Conciliador DIAN vs Contabilidad"
APP_VERSION = "2.0.0"

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

SHEET_ORDER = ["principal", "contabilidad", "terceros", "aud_comp", "autorretenedores"]
REQUIRED_SHEETS = ["principal", "contabilidad", "terceros", "aud_comp"]
OPTIONAL_SHEETS = ["autorretenedores"]

ORDEN = ['Tipo de documento', 'CUFE/CUDE', 'Folio', 'Prefijo', 'Num.Ext', 'Divisa', 'Forma de Pago',
         'Medio de Pago', 'Fecha Emisión', 'Fecha Recepción', 'NIT Emisor', 'Nombre Emisor',
         'NIT Receptor', 'Nombre Receptor', 'BASE', 'IVA', 'Total', 'Estado', 'Grupo']

COLUMNAS_DIAN_VS_CONT = [
    'Tipo de documento', 'CUFE/CUDE', 'Folio', 'Prefijo', 'Num.Ext', 'Divisa', 'Fecha',
    'Forma de Pago', 'Medio de Pago', 'Fecha Emisión', 'Fecha Recepción', 'NIT Emisor',
    'Nombre Emisor', 'NIT Receptor', 'Nombre Receptor', 'BASE', 'IVA', 'Total', 'Estado', 'Grupo'
]

SERIALES_IVA = ['240811019', '240816019', '135530019', '531520002']
SERIALES_BASE = ['620501001', '620501019', '622501020', '52', '5305', '1540']

OUTPUT_SHEETS_TO_HIDE = ['Resultados', 'auditoria', 'resultados-auditoria']

RED_FILL_COLOR = "FFC7CE"
CURRENCY_FORMAT = '"$"#,##0'

PREVIEW_MAX_ROWS = 300
PREVIEW_MAX_COLS = 60
