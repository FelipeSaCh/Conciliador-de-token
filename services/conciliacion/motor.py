from pathlib import Path

import pandas as pd

from core.errors import ErrorSistema, ErrorUsuario, HojaNoEncontradaError, Reportador, logger
from data.autorretenedores_repo import GestorAutorretenedores
from shared.dataframes import limpiar_encabezados, quitar_filas_vacias
from shared.excel_io import cargar_hoja_encabezado_variable, hojas_de
from services.conciliacion import preparacion as prep
from services.conciliacion.emparejamiento import unificar_dian_vs_cont
from services.conciliacion.escritura import escribir_resultados


class ConciliadorAuditoria(Reportador):
    def __init__(self, file_path, sheet_names, seriales_iva=None, seriales_base=None,
                 seriales_base2=None, seriales_ret=None, seriales_iva_dc=None,
                 seriales_base_dc=None, seriales_base2_dc=None, filtro_grupo="Ambos", progress_callback=None):
        self.file_path = Path(file_path)
        self.sheet_names = sheet_names
        self.progress_callback = progress_callback
        self.seriales_iva = seriales_iva or []
        self.seriales_base = seriales_base or []
        self.seriales_base2 = seriales_base2 or []
        self.seriales_ret = seriales_ret or []
        self.seriales_iva_dc = seriales_iva_dc or []
        self.seriales_base_dc = seriales_base_dc or []
        self.seriales_base2_dc = seriales_base2_dc or []
        self.filtro_grupo = filtro_grupo

    def _validar_archivo(self):
        if not self.file_path.exists():
            raise ErrorUsuario(f"El archivo no existe: {self.file_path}")
        try:
            hojas = hojas_de(self.file_path)
        except Exception as e:
            raise ErrorSistema(f"No se pudo abrir el archivo Excel: {e}") from e
        for clave in ('principal', 'aud_comp'):
            nombre = self.sheet_names.get(clave)
            if not nombre:
                raise ErrorUsuario(f"No se especificó el nombre de la hoja '{clave}'.")
            if nombre not in hojas:
                raise HojaNoEncontradaError(nombre, hojas)
        return hojas

    def _cargar_token(self):
        try:
            df = pd.read_excel(self.file_path, sheet_name=self.sheet_names['principal'])
            ultimo = df.dropna(how='all').index.max()
            if pd.notna(ultimo):
                df = df.loc[:ultimo].copy()
            return limpiar_encabezados(df)
        except Exception as e:
            raise ErrorSistema(f"Error leyendo la hoja principal: {e}") from e

    def _cargar_auditoria(self, nombre):
        return prep.limpiar_hoja_auditoria(
            limpiar_encabezados(cargar_hoja_encabezado_variable(self.file_path, nombre))
        )

    def _cargar_devoluciones(self, nombre, hojas):
        if not nombre or nombre not in hojas:
            return None
        try:
            return self._cargar_auditoria(nombre)
        except Exception as e:
            logger.warning(f"No se pudo leer la hoja de auditoría de devoluciones '{nombre}': {e}")
            return None

    def ejecutar(self):
        hojas = self._validar_archivo()

        self._reportar("Cargando hoja de Token procesado...")
        df_token = prep.completar_columnas_token(self._cargar_token())
        mapas_impuestos = prep.mapas_impuestos(df_token)
        df_token = quitar_filas_vacias(
            df_token, [c for c in ('CONCEPTO', 'TERCERO', 'TIPO', 'BASE', 'Num.Ext') if c in df_token.columns]
        )

        self._reportar("Cargando hoja de auditoría de comprobantes...")
        df_aud_comp = self._cargar_auditoria(self.sheet_names['aud_comp'])
        df_autorretenedores = GestorAutorretenedores.como_dataframe()

        self._reportar("Filtrando registros y mapeando motivos de diferencias...")
        nombre_dc = self.sheet_names.get('Aud_dc')
        df_aud_dc = self._cargar_devoluciones(nombre_dc, hojas)

        df_token, personales, mask_ignorar, n_duplicados = prep.marcar_novedades_token(df_token, self.filtro_grupo)
        self._reportar(f"Se detectaron {n_duplicados} registros duplicados en Token (excluyendo Personales)")
        df_resultado = prep.construir_resultado(df_token, personales, mask_ignorar)
        df_auditoria = prep.construir_auditoria(df_token)

        self._reportar("Procesando hoja de auditoría de comprobantes...")
        df_res_auditoria = prep.procesar_seriales(
            df_aud_comp, self.seriales_iva, self.seriales_base, self.seriales_base2, mapas_impuestos
        )

        self._reportar("Unificando y conciliando DIAN vs Contabilidad...")
        res = unificar_dian_vs_cont(
            df_auditoria, df_aud_comp, df_res_auditoria, df_autorretenedores,
            df_aud_dc=df_aud_dc, seriales_iva_dc=self.seriales_iva_dc,
            seriales_base_dc=self.seriales_base_dc, seriales_base2_dc=self.seriales_base2_dc,
            mapas_impuestos=mapas_impuestos,
        )

        self._reportar("Escribiendo resultados en el archivo Excel...")
        escribir_resultados(
            self.file_path, self.sheet_names, nombre_dc, df_resultado,
            df_auditoria.drop(columns=['Novedad'], errors='ignore'), df_res_auditoria, df_autorretenedores, res,
        )

        self._reportar("Proceso completado con éxito.")
        return {
            "filas_procesadas": len(df_token),
            "filas_personales": int(personales.sum()),
            "filas_sin_pareja": int(sum(res.incompletas)),
            "df_completo": res.df_completo,
        }
