# Conciliador de Tokens y Facturación (DIAN vs Contabilidad)

Este proyecto es una herramienta automatizada desarrollada en Python para realizar la conciliación de información fiscal de la DIAN frente a los registros contables y terceros de una compañía. El procesamiento se realiza a partir de un libro de Excel y genera reportes cruzados, formatos de moneda, resaltados condicionales y tablas automáticas.

## Estructura del Proyecto

La arquitectura del proyecto está organizada de la siguiente manera:

```text
Conciliador de token/
│
├── venv/                 # Entorno virtual de Python (excluido en git)
├── .gitignore            # Archivos y carpetas excluidos del control de versiones
├── CHANGELOG.md          # Historial detallado de cambios y versiones
├── README.md             # Documentación general del proyecto (este archivo)
├── requirements.txt      # Dependencias del proyecto para su ejecución
└── main.py               # Script principal con la lógica de procesamiento
```

---

## Requisitos Previos

Antes de ejecutar el proyecto, asegúrate de tener instalado:
* **Python 3.8 o superior**
* **Pip** (administrador de paquetes de Python)

---

## Instalación y Configuración

1. **Clonar el repositorio:**
   ```bash
   git clone https://github.com/FelipeSaCh/Conciliador-de-token.git
   cd Conciliador-de-token
   ```

2. **Crear y activar el entorno virtual (Recomendado):**
   * En Windows (PowerShell/CMD):
     ```powershell
     python -m venv venv
     venv\Scripts\activate
     ```
   * En macOS/Linux:
     ```bash
     python3 -m venv venv
     source venv/bin/activate
     ```

3. **Instalar dependencias:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Configurar la ruta del archivo de Excel:**
   Abre el archivo [main.py](file:///c:/Users/USUARIO/Desktop/Proyectos/Conciliador%20de%20token/main.py) y edita la línea 10 para apuntar a la ruta del archivo de Excel local que deseas procesar:
   ```python
   file_path = r"C:\Ruta\A\Tu\Archivo\De\Excel.xlsx"
   ```

---

## Funcionamiento del Proceso

El script realiza la carga, limpieza, cruce y formato del archivo de Excel definido en `file_path`. 

### Hojas de Entrada Esperadas
El libro de Excel de entrada debe contener las siguientes hojas de cálculo:

1. **`Sheet1`**: Información fiscal de facturas recibidas. Requiere columnas como `NIT Emisor`, `NIT Receptor`, `Prefijo`, `Folio`, `Total`, `IVA`.
2. **`contabilidad`**: Registros contables internos. Requiere columnas como `NIT`, `TIPO-DETALLE`, `TIPO`.
3. **`TERCEROS`**: Maestro de terceros creados en el sistema. Requiere columna `NIT`.
4. **`AUD-COMP`**: Auxiliar contable para comprobación de auditoría (contiene cuentas de IVA y base de compras).
5. **`AUTORRETENEDORES`**: Lista de NITs catalogados como autorretenedores, con columnas `NIT` y `COMENTARIO` (o variantes similares).

### Hojas de Salida Generadas
El script escribe y modifica directamente el archivo de Excel original, agregando/reemplazando las siguientes hojas:

* **`DIAN VS CONT` (Visible)**: Reporte principal de conciliación. Contiene una tabla con estilo estructurado y autoajuste de columnas donde se comparan los documentos de la DIAN con los registros contables. Las filas sin correspondencia (pareja) se resaltan en color **rojo suave**.
* **`Sheet1` (Modificada)**: Se añaden las columnas `CONCEPTO` y `TERCERO` a cada registro. Los gastos catalogados con tipo `PERSONALES` se resaltan en **rojo suave**.
* **Hojas Intermedias (Ocultas automáticamente)**:
  * `Resultados`: Hoja de procesamiento intermedio para cruces.
  * `auditoria`: Registros filtrados excluyendo los de tipo personal.
  * `resultados-auditoria`: Sumatoria de base e IVA por cada número externo (`Num.Ext`).

---

## Ejecución

Una vez configurada la ruta en `main.py` y con el entorno virtual activo, ejecuta el script principal:

```bash
python main.py
```

Al finalizar, se mostrará el mensaje en consola:
`Proceso completado con éxito.`

---

## Licencia

Este proyecto es de uso privado e interno. Todos los derechos reservados.
