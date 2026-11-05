"""Reestructuración (pivoteo) de hojas de movimientos contables hacia la hoja MOVS_AUD."""
import re
from copy import copy
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple, Union

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE, Cell
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from openpyxl.utils import get_column_letter

from core.errors import ErrorSistema, ErrorUsuario, Reportador, logger
from shared.cuentas import nombre_cuenta, ordenar_por_prefijos
from shared.excel_io import escritor_excel, leer_excel
from shared.texto import clave_columna, limpiar_codigo
from shared.totales import MARCA_TOTALES, agregar_fila_totales_ws
from shared.xlsx_rapido import LibroXlsx, texto_xml

MESES_CORTOS = {i: m for i, m in enumerate(
    ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'], start=1)}
MESES_CORTOS.update({str(k): v for k, v in list(MESES_CORTOS.items())})
MESES_CORTOS.update({f"{k:02d}": v for k, v in list(MESES_CORTOS.items()) if isinstance(k, int) and k < 10})

RENOMBRES = {
    "TIPO": "Tipo", "NUMERO": "Número", "NO.EXT": "Num.Ext", "Nit/C.C.": "NIT", "NIT": "Nit/C.C.",
    "Nombre Emisor": "Tercero", "TERCERO": "Tercero", "AÑO": "Año", "MES": "Mes", "DIA": "Dia",
}
COLUMNAS_DESCARTADAS = ["CHEQ.NO", "ELABORO", "SUC.PTO", "CCOSTO", "SCCOSTO", "REF1", "REF2", "REF3", "REF4",
                        "No. AUTORIZACION(DATAFONO)"]
ORDEN_INDICE = ["Tipo", "Número", "Num.Ext", "Fecha", "Tercero", "DETALLE", "DOCRELA", "Nit/C.C."]
PRIORIDAD_CUENTAS = ('2365', '2367', '2368', '15', '6', '5', '22')
PRIORIDAD_DETALLE = ('6', '5', '7', '15')
TIPOS_AL_FINAL = ('GC', 'GS', 'CG')
FILAS_BUSQUEDA_ENCABEZADO = 30
TIPO_GENERICO = "SIN TIPO"

# Formato genérico: destino -> alias (clave_columna). Orden = prioridad. Sufijo '*' = coincide por prefijo.
ALIAS_GENERICO: Dict[str, Tuple[str, ...]] = {
    "NOM. CUENTA": ("nomcuenta", "nombrecuenta", "descauxiliar", "nombreauxiliar", "desccuenta",
                    "descripcioncuenta", "nombrecta", "descripcionauxiliar", "nombredelacuenta"),
    "CUENTA": ("cuenta", "auxiliar", "codigocuenta", "codcuenta", "cuentacontable", "codigocontable", "cta",
               "codcta", "numcuenta", "nrocuenta", "cuentaauxiliar", "subcuenta", "codigo", "puc", "cuenta*"),
    "DEBITO": ("debito", "debitos", "dbitos", "debe", "valordebito", "movdebito", "db", "debitos*", "debito*"),
    "CREDITO": ("credito", "creditos", "crditos", "haber", "valorcredito", "movcredito", "cr", "creditos*",
                "credito*"),
    "VALOR": ("valor", "valormovimiento", "importe", "monto", "valorlocal"),
    "NATURALEZA": ("naturaleza", "nat", "dc", "debcred", "debitocredito"),
    "TIPO": ("tipo", "tipodocto", "tipodocumento", "tipocomprobante", "tipodoc", "tipocte", "clasedocumento",
             "fuente", "prefijo", "codcomprobante"),
    "COMPROBANTE": ("comprobante", "documento", "doc", "nrocomprobante", "numcomprobante", "documentocontable"),
    "NUMERO": ("numero", "numerodocumento", "numdocumento", "nrodocumento", "ndocumento", "nodocumento",
               "consecutivo", "numdoc", "nrodoc", "nro", "num"),
    "NO.EXT": ("noext", "numext", "docto", "documentoexterno", "docexterno", "factura", "numerofactura",
               "nrofactura", "numfactura", "documentosoporte", "soporte"),
    "DOCRELA": ("docrela", "docrelacionado", "documentorelacionado", "docref", "documentoreferencia"),
    "TERCERO": ("tercero", "nombretercero", "razonsocial", "raznsocial", "razonsocialtercerodocto",
                "raznsocialtercerodocto", "nombreemisor", "nombre", "beneficiario", "proveedor", "cliente",
                "nombretercero*"),
    "NIT": ("nit", "nitcc", "nittercero", "identificacion", "identificaciontercero", "cedula", "nitcedula",
                 "idtercero", "documentotercero", "nittercero*"),
    "FECHA": ("fecha", "fechadocumento", "fechadoc", "fechamovimiento", "fechacomprobante", "fechaelaboracion",
              "fechacontable", "fecha*"),
    "AÑO": ("ano", "anio", "year"),
    "MES": ("mes", "month", "periodo"),
    "DIA": ("dia", "day"),
    "DETALLE": ("detalle", "descripcion", "concepto", "observacion", "observaciones", "glosa", "nota",
                "comentario", "descripcionmovimiento", "detalle*"),
}
_RE_CONTROL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
_ALIAS_TODOS = {a.rstrip("*") for alias in ALIAS_GENERICO.values() for a in alias}
_RE_CUENTA = re.compile(r"^\s*(\d[\d.\-\s]*?)(?:\s*[-–:]\s*|\s+)?([^\d\s\-–:.].*)?$")


def _limpiar_textos(df: pd.DataFrame) -> pd.DataFrame:
    """Quita caracteres de control (p. ej. \\x02 que exportan algunos software contables) y el relleno de espacios."""
    df.columns = [_RE_CONTROL.sub("", c).strip() if isinstance(c, str) else c for c in df.columns]
    for c in df.columns:
        if pd.api.types.is_numeric_dtype(df[c]) or pd.api.types.is_datetime64_any_dtype(df[c]):
            continue
        valores = df[c]
        unicos = valores.dropna().unique()
        limpio = {v: _RE_CONTROL.sub("", v).strip() for v in unicos if isinstance(v, str)}
        if any(k != v for k, v in limpio.items()):
            df[c] = valores.map(lambda v: limpio.get(v, v) if isinstance(v, str) else v)
    return df


def _a_numero(serie: pd.Series) -> pd.Series:
    """Numérico tolerante a texto: '1.234.567,89', '$ 1,234.56', '(1.000)' -> float. Convierte sobre valores únicos."""
    if pd.api.types.is_numeric_dtype(serie):
        return serie.fillna(0).astype(float)
    num = pd.to_numeric(serie, errors="coerce")
    pendientes = num.isna() & serie.notna()
    es_texto = serie.map(lambda v: isinstance(v, str))
    if serie[es_texto].str.contains(r",\d{1,2}\s*$|\d\.\d{3}\.\d{3}", regex=True).any():
        pendientes |= es_texto  # columna en formato local (1.234,56): '12.500' son miles, no decimales
    if pendientes.any():
        def convertir(v):
            s = re.sub(r"[^\d,.\-()]", "", str(v))
            negativo = s.startswith("(") and s.endswith(")") or s.endswith("-")
            s = s.strip("()-") if negativo else s.strip("()")
            if re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
                s = s.replace(".", "")
            elif "," in s and "." in s:
                s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
            elif "," in s:
                s = s.replace(",", ".") if re.search(r",\d{1,2}$", s) and s.count(",") == 1 else s.replace(",", "")
            elif s.count(".") > 1:
                s = s.replace(".", "")
            try:
                return -float(s) if negativo else float(s)
            except ValueError:
                return np.nan
        unicos = serie[pendientes].astype(str).unique()
        num = num.astype(float)
        num[pendientes] = serie[pendientes].astype(str).map(dict(zip(unicos, map(convertir, unicos))))
    return num.fillna(0).astype(float)


def _separar_cuenta(serie: pd.Series) -> Tuple[pd.Series, pd.Series]:
    """'110505.0' / '1105-05 Caja general' -> ('110505', 'Caja general'). Sin código -> ''."""
    texto = limpiar_codigo(serie)
    unicos = texto.unique()
    codigos, nombres = {}, {}
    for u in unicos:
        m = _RE_CUENTA.match(u)
        codigos[u] = re.sub(r"\D", "", m.group(1)) if m else ""
        nombres[u] = (m.group(2) or "").strip() if m else ""
    return texto.map(codigos), texto.map(nombres)


def _derivar_tipo_numero(serie: pd.Series) -> Tuple[pd.Series, pd.Series]:
    """'FC-00123' / 'FC 123' -> ('FC', '00123')."""
    partes = limpiar_codigo(serie).str.upper().str.extract(r"^\s*([A-Z]+)?[\s\-_/.#]*(.*)$")
    return partes[0].fillna(TIPO_GENERICO).replace("", TIPO_GENERICO), partes[1].fillna("").str.strip()


def _a_fecha(serie: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    seriales = pd.to_numeric(serie, errors="coerce")
    fechas = pd.to_datetime(serie.where(seriales.isna()), errors="coerce", dayfirst=True, format="mixed")
    es_serial = seriales.between(1, 2958465)
    if es_serial.any():
        fechas[es_serial] = pd.to_datetime(seriales[es_serial], unit="D", origin="1899-12-30")
    return fechas


class TransformadorMovimientos(Reportador):
    PREFIJOS_PERMITIDOS = ("FC", "GS", "DC", "DS", "GC", "CG", "GI")
    HOJA_DESTINO = "MOVS_AUD"
    COLUMNA_FECHA_ORIGEN = "__FECHA_ORIGEN__"

    # Formatos detectados por columnas requeridas (claves sin tildes/espacios/signos) → nombres estándar.
    FORMATOS: Dict[str, dict] = {
        "estandar": {
            "descripcion": "Movimientos estándar",
            "requeridas": {
                "DEBITO": ("debito",), "CREDITO": ("credito",), "CUENTA": ("cuenta",),
                "NOM. CUENTA": ("nomcuenta",), "TIPO": ("tipo",),
            },
            "opcionales": {},
            "detalle_vacio": False,
            "eliminar": (),
        },
        "auxiliar_documento": {
            "descripcion": "Auxiliar por documento",
            "requeridas": {
                "DEBITO": ("debitos", "dbitos"), "CREDITO": ("creditos", "crditos"), "CUENTA": ("auxiliar",),
                "NOM. CUENTA": ("descauxiliar",), "TIPO": ("tipodocto",), "NO.EXT": ("docto",),
            },
            "opcionales": {
                "TERCERO": ("razonsocialtercerodocto", "raznsocialtercerodocto"),
                "NIT": ("nittercero",),
                COLUMNA_FECHA_ORIGEN: ("fecha",),
            },
            "detalle_vacio": True,
            "eliminar": ("neto", "basegravable", "comovto"),
        },
    }
    DESCRIPCION_GENERICO = "Movimientos genéricos (columnas reconocidas automáticamente)"

    def __init__(
        self,
        ruta_entrada: Union[str, Path],
        hoja_origen: str,
        prefijos_permitidos: Optional[List[str]] = None,
        progress_callback: Optional[Callable[[str], None]] = None,
        avance_callback: Optional[Callable[[Optional[float]], None]] = None,
    ):
        self.ruta_entrada = Path(ruta_entrada)
        self.hoja_origen = hoja_origen
        self.prefijos_permitidos = tuple(prefijos_permitidos) if prefijos_permitidos is not None else self.PREFIJOS_PERMITIDOS
        self.progress_callback = progress_callback
        self.avance_callback = avance_callback

    def _avance(self, valor: Optional[float]) -> None:
        """0-100; None = fase sin progreso medible (barra indeterminada)."""
        if self.avance_callback:
            self.avance_callback(valor)

    # ------------------------------------------------------------------ lectura / detección
    @staticmethod
    def _muestra(ruta, hoja) -> List[tuple]:
        """Primeras filas sin cargar la hoja completa (openpyxl en streaming); respaldo: pandas."""
        try:
            wb = load_workbook(ruta, read_only=True, data_only=True)
            try:
                ws = wb[hoja]
                ws.reset_dimensions()
                return list(ws.iter_rows(max_row=FILAS_BUSQUEDA_ENCABEZADO, values_only=True))
            finally:
                wb.close()
        except Exception:
            muestra = leer_excel(ruta, rapido=True, sheet_name=hoja, header=None, nrows=FILAS_BUSQUEDA_ENCABEZADO)
            return list(muestra.itertuples(index=False))

    @classmethod
    def encabezado(cls, ruta, hoja) -> Tuple[int, List[str]]:
        """(fila 0-based, nombres) de la fila con más columnas reconocibles entre las primeras (salta títulos de reporte)."""
        filas = cls._muestra(ruta, hoja)
        mejor, puntaje_max = 0, 1
        for i, fila in enumerate(filas):
            puntaje = sum(1 for v in fila if v is not None and pd.notna(v) and clave_columna(v) in _ALIAS_TODOS)
            if puntaje > puntaje_max:
                mejor, puntaje_max = i, puntaje
        vistos, nombres = {}, []
        for j, v in enumerate(filas[mejor] if filas else ()):
            n = f"Unnamed: {j}" if v is None or (isinstance(v, float) and np.isnan(v)) or str(v).strip() == "" else str(v)
            k = vistos.get(n, 0)
            vistos[n] = k + 1
            nombres.append(f"{n}.{k}" if k else n)
        return mejor, nombres

    @classmethod
    def leer_hoja(cls, ruta, hoja, **kwargs) -> pd.DataFrame:
        return leer_excel(ruta, rapido=True, sheet_name=hoja, header=cls.encabezado(ruta, hoja)[0], **kwargs)

    @staticmethod
    def _mapeo_generico(claves: Dict[str, object]) -> Dict[str, str]:
        usadas, destino_de = set(), {}
        for exacto in (True, False):
            for destino, alias in ALIAS_GENERICO.items():
                if destino in destino_de.values():
                    continue
                for a in alias:
                    prefijo = a.endswith("*")
                    if exacto == prefijo:
                        continue
                    base = a.rstrip("*")
                    col = next((orig for k, orig in claves.items() if orig not in usadas
                                and (k == base if exacto else k.startswith(base))), None)
                    if col is not None:
                        usadas.add(col)
                        destino_de[col] = destino
                        break
        return destino_de

    @classmethod
    def detectar_formato(cls, columnas) -> Tuple[str, Dict[str, str]]:
        claves = {}
        for c in columnas:
            if not str(c).startswith("Unnamed"):
                claves.setdefault(clave_columna(c), c)

        def buscar(alias):
            return next((claves[a] for a in alias if a in claves), None)

        for nombre, fmt in cls.FORMATOS.items():
            mapeo = {buscar(alias): destino for destino, alias in fmt["requeridas"].items()}
            if None in mapeo:
                continue
            mapeo.update({o: d for d, alias in fmt["opcionales"].items() if (o := buscar(alias)) is not None})
            return nombre, mapeo

        mapeo = cls._mapeo_generico(claves)
        destinos = set(mapeo.values())
        faltan = []
        if "CUENTA" not in destinos:
            faltan.append("cuenta")
        if not ({"DEBITO", "CREDITO"} <= destinos or "VALOR" in destinos):
            faltan.append("débito y crédito (o valor)")
        if not destinos & {"TIPO", "COMPROBANTE", "NUMERO", "NO.EXT"}:
            faltan.append("tipo, comprobante o número de documento")
        if not faltan:
            return "generico", mapeo

        detalle = "; ".join(f"{f['descripcion']}: {', '.join(f['requeridas'])}" for f in cls.FORMATOS.values())
        raise ErrorUsuario(
            f"La hoja no tiene la estructura mínima de un movimiento. Falta: {', '.join(faltan)}.\n"
            f"Columnas leídas: {', '.join(map(str, columnas))}\nFormatos fijos soportados -> {detalle}"
        )

    @classmethod
    def _descripcion(cls, nombre) -> str:
        return cls.FORMATOS[nombre]["descripcion"] if nombre in cls.FORMATOS else cls.DESCRIPCION_GENERICO

    @classmethod
    def obtener_tipos(cls, ruta, hoja) -> Tuple[List[str], str]:
        fila, columnas = cls.encabezado(ruta, hoja)
        nombre, mapeo = cls.detectar_formato(columnas)
        inverso = {d: o for o, d in mapeo.items()}
        destino = "TIPO" if "TIPO" in inverso else "COMPROBANTE" if "COMPROBANTE" in inverso else None
        if destino is None:
            return [TIPO_GENERICO], cls._descripcion(nombre)
        pos = [columnas.index(inverso[destino]), columnas.index(inverso["CUENTA"])]
        df = leer_excel(ruta, rapido=True, sheet_name=hoja, header=fila, usecols=pos)
        serie, cuenta = (df.iloc[:, 0], df.iloc[:, 1]) if pos[0] < pos[1] else (df.iloc[:, 1], df.iloc[:, 0])
        serie = serie[_separar_cuenta(cuenta)[0] != ""]
        tipos = limpiar_codigo(serie).str.upper() if destino == "TIPO" else _derivar_tipo_numero(serie)[0]
        return sorted(t for t in tipos.unique() if t), cls._descripcion(nombre)

    # ------------------------------------------------------------------ ejecución
    def ejecutar(self) -> Path:
        try:
            self._avance(None)
            df = self._cargar_y_validar()
            self._avance(20)
            df_final, cols_indice = self._transformar_y_pivotar(df)
            self._avance(40)
            self._guardar_con_formato(df_final, cols_indice)
            self._avance(100)
            self._reportar(f"Guardado finalizado exitosamente en la hoja '{self.HOJA_DESTINO}'.")
            return self.ruta_entrada
        except ErrorUsuario:
            raise
        except Exception as e:
            logger.exception("Error durante la reestructuración de movimientos.")
            raise ErrorSistema(f"Error inesperado al procesar la hoja de movimientos: {e}") from e

    def _cargar_y_validar(self) -> pd.DataFrame:
        self._reportar("Leyendo la hoja de movimientos...")
        if not self.ruta_entrada.exists():
            raise ErrorUsuario(f"El archivo origen no existe: {self.ruta_entrada}")
        try:
            df = self.leer_hoja(self.ruta_entrada, self.hoja_origen)
        except ValueError as e:
            raise ErrorUsuario(f"La hoja '{self.hoja_origen}' no existe en el archivo.") from e
        except Exception as e:
            raise ErrorSistema(f"No se pudo leer el archivo de Excel: {e}") from e
        return self._normalizar_formato(df)

    def _normalizar_formato(self, df: pd.DataFrame) -> pd.DataFrame:
        df = _limpiar_textos(df)
        df = df.loc[:, [not str(c).startswith("Unnamed") or df[c].notna().any() for c in df.columns]]
        nombre, mapeo = self.detectar_formato(df.columns)
        self._reportar(f"Formato detectado: {self._descripcion(nombre)}")
        if nombre == "generico":
            logger.info("Mapeo genérico de movimientos: %s", {d: o for o, d in mapeo.items()})
            return self._normalizar_generico(df, mapeo)
        if nombre == "estandar":
            return df

        fmt = self.FORMATOS[nombre]
        df = df.drop(columns=[c for c in df.columns if clave_columna(c) in fmt["eliminar"]]).rename(columns=mapeo)
        if self.COLUMNA_FECHA_ORIGEN in df.columns:
            self._fecha_a_partes(df, df.pop(self.COLUMNA_FECHA_ORIGEN))
        if fmt["detalle_vacio"] and "DETALLE" not in df.columns:
            df["DETALLE"] = ""
        return df

    @staticmethod
    def _fecha_a_partes(df, serie):
        fechas = _a_fecha(serie)
        df["AÑO"] = fechas.dt.year.astype("Int64")
        df["MES"] = fechas.dt.month.astype("Int64")
        df["DIA"] = fechas.dt.day.astype("Int64")

    def _normalizar_generico(self, df: pd.DataFrame, mapeo: Dict[str, str]) -> pd.DataFrame:
        df = df[list(mapeo)].rename(columns=mapeo)

        if "VALOR" in df.columns and not {"DEBITO", "CREDITO"} <= set(df.columns):
            valor = _a_numero(df.pop("VALOR"))
            if "NATURALEZA" in df.columns:
                es_credito = df.pop("NATURALEZA").astype(str).str.strip().str.upper().str.startswith("C")
                df["DEBITO"], df["CREDITO"] = valor.where(~es_credito, 0), valor.where(es_credito, 0)
            else:
                df["DEBITO"], df["CREDITO"] = valor.clip(lower=0), (-valor).clip(lower=0)
        df = df.drop(columns=[c for c in ("VALOR", "NATURALEZA") if c in df.columns])

        if "COMPROBANTE" in df.columns:
            tipo, numero = _derivar_tipo_numero(df.pop("COMPROBANTE"))
            if "TIPO" not in df.columns:
                df["TIPO"] = tipo
            if "NUMERO" not in df.columns:
                df["NUMERO"] = numero
        if "TIPO" not in df.columns:
            df["TIPO"] = TIPO_GENERICO

        if "FECHA" in df.columns and not {"AÑO", "MES", "DIA"} <= set(df.columns):
            self._fecha_a_partes(df, df.pop("FECHA"))
        df = df.drop(columns=[c for c in ("FECHA",) if c in df.columns])

        if "NOM. CUENTA" not in df.columns:
            df["NOM. CUENTA"] = ""
        if "DETALLE" not in df.columns:
            df["DETALLE"] = ""
        return df

    # ------------------------------------------------------------------ transformación
    def _filtrar(self, df):
        if "Tipo" not in df.columns:
            raise ErrorUsuario("El archivo no contiene la columna 'TIPO' o 'Tipo'.")
        if not self.prefijos_permitidos:
            raise ErrorUsuario("No se seleccionó ningún tipo de documento permitido.")

        df["Tipo"] = limpiar_codigo(df["Tipo"])
        df = df[df["Tipo"].str.upper().str.startswith(self.prefijos_permitidos)]
        if df.empty:
            raise ErrorUsuario(
                f"No se encontraron registros que inicien con los tipos seleccionados: {', '.join(self.prefijos_permitidos)}"
            )
        if "DETALLE" in df.columns:
            df = df[~df["DETALLE"].astype(str).str.contains("ANULADO", case=False, na=False)]
        return df.drop(columns=[c for c in COLUMNAS_DESCARTADAS if c in df.columns]).copy()

    @staticmethod
    def _consolidar_fecha(df):
        if all(c in df.columns for c in ("Año", "Mes", "Dia")):
            mes = df["Mes"].map(MESES_CORTOS).fillna(df["Mes"].astype(str))
            dia = limpiar_codigo(df["Dia"].astype(str), strip=False).str.zfill(2)
            anio = limpiar_codigo(df["Año"].astype(str), strip=False)
            df["Fecha"] = mes + "-" + dia + "-" + anio
            df.drop(columns=["Año", "Mes", "Dia"], inplace=True)

    @staticmethod
    def _depurar_cuentas(df):
        """Código limpio, nombre (columna, texto combinado o PUC) y descarte de totales/subtítulos/filas vacías."""
        codigo, nombre_embebido = _separar_cuenta(df["CUENTA"])
        df = df[codigo != ""].copy()
        codigo, nombre_embebido = codigo[df.index], nombre_embebido[df.index]
        nombre = limpiar_codigo(df["NOM. CUENTA"]) if "NOM. CUENTA" in df.columns else pd.Series("", index=df.index)
        nombre = nombre.where(nombre != "", nombre_embebido)
        faltantes = nombre == ""
        if faltantes.any():
            nombre[faltantes] = codigo[faltantes].map({c: nombre_cuenta(c) for c in codigo[faltantes].unique()})
        df["CUENTA"], df["NOM. CUENTA"] = codigo, nombre
        return df

    def _transformar_y_pivotar(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
        self._reportar("Estandarizando nombres de columnas y limpiando basura...")
        df = self._depurar_cuentas(df)
        if df.empty:
            raise ErrorUsuario("La hoja no contiene filas con código de cuenta válido.")
        df = self._filtrar(df.rename(columns=RENOMBRES))
        self._consolidar_fecha(df)

        if "Num.Ext" in df.columns:
            df["Num.Ext"] = limpiar_codigo(df["Num.Ext"]).replace(['nan', 'None', '<NA>'], '')

        self._reportar("Aplicando cálculos numéricos y unificando registros duplicados...")
        df["DEBITO"] = _a_numero(df["DEBITO"])
        df["CREDITO"] = _a_numero(df["CREDITO"])
        df["VALOR_NETO"] = df["DEBITO"] - df["CREDITO"]
        df["CUENTA_CONCAT"] = df["CUENTA"] + " - " + df["NOM. CUENTA"].astype(str)

        claves = [c for c in ("Tipo", "Número", "Num.Ext") if c in df.columns]
        cols_sistema = {"CUENTA", "NOM. CUENTA", "CUENTA_CONCAT", "DEBITO", "CREDITO", "VALOR_NETO", "DETALLE"}
        cols_desc = [c for c in df.columns if c not in claves and c not in cols_sistema]

        df_final = (
            df.groupby(claves + ["CUENTA_CONCAT"], dropna=False)["VALOR_NETO"].sum()
            .unstack("CUENTA_CONCAT", fill_value=0).reset_index()
        )
        df_final.columns.name = None
        self._avance(30)

        if cols_desc:
            df_final = df_final.merge(df.groupby(claves, as_index=False, dropna=False)[cols_desc].first(), on=claves, how="left")
        if "DETALLE" in df.columns:
            # DETALLE del documento: prioriza la línea de cuenta de gasto/costo/activo
            det = df[claves + ["CUENTA", "DETALLE"]].copy()
            det["_p"] = 2 - det["CUENTA"].str.startswith(PRIORIDAD_DETALLE).astype(int)
            det = det.sort_values(by=claves + ["_p"], kind="stable").drop_duplicates(subset=claves, keep="first")
            df_final = df_final.merge(det[claves + ["DETALLE"]], on=claves, how="left")
            cols_desc.append("DETALLE")

        cols_indice = [c for c in ORDEN_INDICE if c in df_final.columns]
        cols_indice += [c for c in df_final.columns if c in cols_desc and c not in cols_indice]
        cols_cuentas = ordenar_por_prefijos([c for c in df_final.columns if c not in cols_indice], PRIORIDAD_CUENTAS)
        df_final = df_final[cols_indice + cols_cuentas]

        if "Tipo" in df_final.columns and "Tercero" in df_final.columns:
            self._reportar("Ordenando filas: GC, GS, CG al final y agrupados por Tercero...")
            al_final = df_final["Tipo"].astype(str).str.strip().str.upper().str.startswith(TIPOS_AL_FINAL)
            df_final = pd.concat(
                [df_final[~al_final], df_final[al_final].sort_values(by="Tercero", kind="stable")], ignore_index=True
            )

        return df_final, cols_indice

    # ------------------------------------------------------------------ escritura
    @staticmethod
    def _dimensiones(df_final: pd.DataFrame, n_indice: int) -> Tuple[List[float], float]:
        anchos, max_lineas = [], 1
        for idx, col in enumerate(df_final.columns):
            largo_titulo = len(str(col))
            if idx < n_indice:
                largo = df_final[col].astype(str).str.len().max() if len(df_final) else 0
                largo = int(largo) if pd.notna(largo) else 0
                anchos.append(min(max(max(largo, largo_titulo) + 3, 12), 45))
                max_lineas = max(max_lineas, largo_titulo // 35 + 1)
            else:
                anchos.append(30)
                max_lineas = max(max_lineas, largo_titulo // 30 + 1)
        return anchos, min(max(24, max_lineas * 14), 32)

    def _guardar_con_formato(self, df_final: pd.DataFrame, cols_indice: List[str]) -> None:
        self._reportar(f"Guardando resultados en la hoja '{self.HOJA_DESTINO}' del archivo original...")
        try:
            self._guardar_rapido(df_final, cols_indice)
            return
        except PermissionError as e:
            raise ErrorUsuario("No se pudo guardar: cierra el archivo en Excel e inténtalo de nuevo.") from e
        except Exception:
            logger.exception("Escritura directa de MOVS_AUD falló; se usa openpyxl.")
        try:
            self._guardar_openpyxl(df_final, cols_indice)
        except PermissionError as e:
            raise ErrorUsuario("No se pudo guardar: cierra el archivo en Excel e inténtalo de nuevo.") from e

    def _guardar_rapido(self, df_final: pd.DataFrame, cols_indice: List[str]) -> None:
        """Escribe MOVS_AUD directo al paquete xlsx: solo celdas de texto y valores ≠ 0, sin bordes ni rellenos en el cuerpo
        (el usuario los personaliza) y fila final de TOTALES con SUBTOTAL."""
        n_filas, n_cols = df_final.shape
        n_indice = len(cols_indice)
        cols_cuentas = list(df_final.columns[n_indice:])
        letras = [get_column_letter(i) for i in range(1, max(n_cols, 1) + 1)]
        anchos, alto_hdr = self._dimensiones(df_final, n_indice)
        fino = '<left style="thin"><color rgb="FFE2E8F0"/></left><right style="thin"><color rgb="FFE2E8F0"/></right>'
        borde_cuerpo = '<border><left/><right/><top/><bottom/><diagonal/></border>'
        font_body = '<font><sz val="10"/><color rgb="FF000000"/><name val="Arial"/></font>'
        sin_relleno = '<fill><patternFill patternType="none"/></fill>'

        self._avance(None)
        with LibroXlsx(self.ruta_entrada) as libro:
            s_hdr = libro.estilo(
                '<font><b/><sz val="11"/><color rgb="FF000000"/><name val="Arial"/></font>',
                '<fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/><bgColor rgb="FFFFCC00"/></patternFill></fill>',
                f'<border>{fino}<top style="medium"><color rgb="FF0F172A"/></top><bottom style="medium"><color rgb="FF0F172A"/></bottom><diagonal/></border>',
                '<alignment horizontal="center" vertical="center" wrapText="1"/>')
            s_txt = libro.estilo(font_body, sin_relleno, borde_cuerpo, '<alignment horizontal="left" vertical="center"/>')
            s_num = libro.estilo(font_body, sin_relleno, borde_cuerpo, '<alignment horizontal="right" vertical="center"/>',
                                 '"$" #,##0.00')
            s_fecha = libro.estilo(font_body, sin_relleno, borde_cuerpo, '<alignment horizontal="left" vertical="center"/>',
                                   'dd/mm/yyyy')
            font_total = '<font><b/><sz val="10"/><color rgb="FF000000"/><name val="Arial"/></font>'
            borde_total = ('<border><left/><right/><top style="thin"><color rgb="FF000000"/></top>'
                           '<bottom style="double"><color rgb="FF000000"/></bottom><diagonal/></border>')
            s_tot_lbl = libro.estilo(font_total, sin_relleno, borde_cuerpo, '<alignment horizontal="left" vertical="center"/>')
            s_tot = libro.estilo(font_total, sin_relleno, borde_total, '<alignment horizontal="right" vertical="center"/>',
                                 '"$" #,##0.00')

            # Celdas de texto pre-serializadas por fila
            bloque_txt = df_final[cols_indice].astype(object)
            valores_txt = bloque_txt.where(bloque_txt.notna(), None).to_numpy()
            matriz = np.nan_to_num(df_final[cols_cuentas].to_numpy(dtype=float)) if cols_cuentas else np.zeros((n_filas, 0))
            matriz_txt = matriz.tolist()
            totales_cols = matriz.sum(axis=0).tolist()  # resultado guardado de las fórmulas SUBTOTAL
            nz_filas, nz_cols = np.nonzero(matriz)
            cortes = np.searchsorted(nz_filas, np.arange(n_filas + 1))
            nz_cols_txt = [letras[n_indice + c] for c in range(len(cols_cuentas))]
            base = pd.Timestamp("1899-12-30")

            def celda_txt(ref, v):
                if isinstance(v, (bool, np.bool_)):
                    return f'<c r="{ref}" s="{s_txt}" t="b"><v>{int(v)}</v></c>'
                if isinstance(v, (int, float, np.integer, np.floating)):
                    if not np.isfinite(v):
                        return f'<c r="{ref}" s="{s_txt}"/>'
                    numero = int(v) if float(v).is_integer() else float(v)  # NIT/número sin ".0"
                    return f'<c r="{ref}" s="{s_txt}"><v>{numero!r}</v></c>'
                if isinstance(v, (pd.Timestamp,)) or hasattr(v, "year") and hasattr(v, "month"):
                    serial = (pd.Timestamp(v) - base) / pd.Timedelta(days=1)
                    return f'<c r="{ref}" s="{s_fecha}"><v>{float(serial)!r}</v></c>'
                t = texto_xml(v)
                return f'<c r="{ref}" s="{s_txt}" t="inlineStr"><is><t xml:space="preserve">{t}</t></is></c>' if t \
                    else f'<c r="{ref}" s="{s_txt}"/>'

            def partes():
                ultima = f"{letras[-1]}{max(n_filas + 1, 1)}"
                yield (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                       f'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                       f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                       f'<dimension ref="A1:{ultima}"/>'
                       f'<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" '
                       f'state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
                       f'<sheetFormatPr defaultRowHeight="20" customHeight="1"/><cols>')
                yield "".join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(anchos, 1))
                yield f'</cols><sheetData><row r="1" ht="{alto_hdr}" customHeight="1">'
                yield "".join(f'<c r="{letras[i]}1" s="{s_hdr}" t="inlineStr"><is><t xml:space="preserve">{texto_xml(c)}</t></is></c>'
                              for i, c in enumerate(df_final.columns))
                yield "</row>"
                paso = max(n_filas // 50, 1)
                for r in range(n_filas):
                    if r % paso == 0:
                        self._avance(50 + 45 * r / max(n_filas, 1))
                    fila = r + 2
                    txt = "".join(celda_txt(f"{letras[c]}{fila}", v) for c, v in enumerate(valores_txt[r]))
                    a, b = cortes[r], cortes[r + 1]
                    num = "".join(f'<c r="{nz_cols_txt[c]}{fila}" s="{s_num}"><v>{matriz_txt[r][c]!r}</v></c>'
                                  for c in nz_cols[a:b].tolist())
                    yield f'<row r="{fila}">{txt}{num}</row>'
                if n_filas and cols_cuentas:
                    ft, ultima_fila = n_filas + 2, n_filas + 1
                    yield (f'<row r="{ft}"><c r="A{ft}" s="{s_tot_lbl}" t="inlineStr"><is><t>{MARCA_TOTALES}</t></is></c>'
                           + "".join(f'<c r="{l}{ft}" s="{s_tot}"><f>SUBTOTAL(9,{l}2:{l}{ultima_fila})</f>'
                                     f'<v>{float(t)!r}</v></c>' for l, t in zip(nz_cols_txt, totales_cols)) + '</row>')
                yield ('</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" '
                       'header="0.3" footer="0.3"/></worksheet>')

            self._reportar("Escribiendo archivo en disco...")
            libro.escribir_hoja(self.HOJA_DESTINO, partes())

    def _guardar_openpyxl(self, df_final: pd.DataFrame, cols_indice: List[str]) -> None:

        borde_fino = Side(border_style="thin", color="E2E8F0")
        borde_enc = Side(border_style="medium", color="0F172A")
        cuerpo = Border()
        font_body = Font(name="Arial", size=10, color="000000")
        sin_relleno = PatternFill(fill_type=None)
        estilos = {
            "movs_aud_hdr": (PatternFill(start_color="ffcc00", end_color="ffcc00", fill_type="solid"),
                             Font(name="Arial", size=11, bold=True, color="000000"),
                             Border(left=borde_fino, right=borde_fino, top=borde_enc, bottom=borde_enc),
                             Alignment(horizontal="center", vertical="center", wrap_text=True), "General"),
            "movs_aud_txt": (sin_relleno, font_body, cuerpo, Alignment(horizontal="left", vertical="center"), "General"),
            "movs_aud_num": (sin_relleno, font_body, cuerpo, Alignment(horizontal="right", vertical="center"),
                             '"$" #,##0.00'),
        }

        n_filas, n_cols = df_final.shape
        n_indice = len(cols_indice)
        cols_cuentas = list(df_final.columns[n_indice:])

        self._avance(None)
        with escritor_excel(self.ruta_entrada) as writer:
            book = writer.book
            indice_hoja = None
            if self.HOJA_DESTINO in book.sheetnames:
                indice_hoja = book.sheetnames.index(self.HOJA_DESTINO)
                del book[self.HOJA_DESTINO]
            ws = book.create_sheet(self.HOJA_DESTINO, indice_hoja)

            existentes = set(book.named_styles)
            for nombre, (fill, font, border, align, num_fmt) in estilos.items():
                if nombre not in existentes:
                    book.add_named_style(NamedStyle(
                        name=nombre, fill=fill, font=font, border=border, alignment=align, number_format=num_fmt))
            plantillas = {}
            for nombre in estilos:
                celda = Cell(ws)
                celda.style = nombre
                plantillas[nombre] = copy(celda._style)

            def escribir(fila, columna, valor, estilo):
                celda = ws.cell(row=fila, column=columna, value=valor)
                celda._style = copy(estilo)

            # Encabezados
            anchos, alto_hdr = self._dimensiones(df_final, n_indice)
            for idx, (col, ancho) in enumerate(zip(df_final.columns, anchos), start=1):
                escribir(1, idx, str(col), plantillas["movs_aud_hdr"])
                ws.column_dimensions[get_column_letter(idx)].width = ancho
            ws.row_dimensions[1].height = alto_hdr
            self._avance(50)

            # Columnas descriptivas (densas)
            txt = plantillas["movs_aud_txt"]
            bloque_txt = df_final[cols_indice].astype(object)
            for c in bloque_txt.columns:
                if bloque_txt[c].map(type).eq(str).any():
                    bloque_txt[c] = bloque_txt[c].map(lambda v: ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v)
            valores_txt = bloque_txt.where(bloque_txt.notna(), None).to_numpy()
            for r, fila in enumerate(valores_txt, start=2):
                for c, valor in enumerate(fila, start=1):
                    if valor is None or valor == "":
                        ws.cell(row=r, column=c)._style = copy(txt)
                    else:
                        escribir(r, c, valor, txt)
            self._avance(60)

            # Columnas de cuentas (dispersas): solo se escriben valores distintos de cero
            if cols_cuentas and n_filas:
                num = plantillas["movs_aud_num"]
                matriz = np.nan_to_num(df_final[cols_cuentas].to_numpy(dtype=float))
                filas, columnas = np.nonzero(matriz)
                total = max(len(filas), 1)
                paso = max(total // 50, 1)
                for k, (r, c) in enumerate(zip(filas.tolist(), columnas.tolist())):
                    if k % paso == 0:
                        self._avance(60 + 35 * k / total)
                    escribir(r + 2, n_indice + c + 1, float(matriz[r, c]), num)

            agregar_fila_totales_ws(ws, range(n_indice + 1, n_cols + 1), 2, n_filas + 1)
            ws.sheet_format.defaultRowHeight = 20
            ws.sheet_format.customHeight = True
            ws.freeze_panes = "A2"
            self._reportar("Escribiendo archivo en disco...")
            self._avance(None)