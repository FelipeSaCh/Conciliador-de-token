# Conciliador de Tokens y Facturación (DIAN vs Contabilidad)

Este proyecto es una herramienta automatizada desarrollada en Python para realizar la conciliación de información fiscal de la DIAN frente a los registros contables y terceros de una compañía. El procesamiento se realiza a partir de un libro de Excel y genera reportes cruzados, formatos de moneda, resaltados condicionales y tablas automáticas.

## Estructura del Proyecto

La arquitectura del proyecto está organizada de la siguiente manera:

```text
Conciliador de token/
│
├── assets/               # Recursos gráficos (icono de la aplicación)
├── venv/                 # Entorno virtual de Python (excluido en git)
├── .gitignore            # Archivos y carpetas excluidos del control de versiones
├── app.py                # Interfaz gráfica de usuario (GUI) en Tkinter
├── config.py             # Configuración general y constantes de la aplicación
├── errors.py             # Excepciones personalizadas y sistema de logs
├── excel_engine.py       # Motor lógico de conciliación de datos de Excel
├── main.py               # Script de entrada para lanzar la aplicación
├── preview_widget.py     # Componente visual para la vista previa de hojas Excel
├── version.py            # Centralización de la versión del programa
├── CHANGELOG.md          # Historial detallado de cambios y versiones
├── README.md             # Documentación general del proyecto (este archivo)
└── requirements.txt      # Dependencias del proyecto para su ejecución
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

4. **Ejecutar la aplicación:**
   ```bash
   python main.py
   ```
   *Nota: No es necesario configurar rutas en el código. Toda la configuración del archivo, hojas y filtros se realiza directamente a través de la interfaz gráfica.*

---

## Guía de Uso de la Interfaz Gráfica

La aplicación cuenta con una interfaz organizada en pestañas para facilitar el proceso de conciliación:

### 1. Cargar Archivo
* Haz clic en **Examinar...** en la parte superior derecha y selecciona tu archivo de Excel (`.xlsx` o `.xlsm`).
* Una vez cargado, se leerán automáticamente las pestañas del archivo.

### 2. Configuración
* **Asignación de Estructura de Hojas**: Asocia las hojas requeridas del proceso con la pestaña correspondiente del archivo importado mediante los menús desplegables. Las pestañas opcionales se pueden vaciar usando el botón **✖**.
* **Configuración de Columnas de AUD-COMP**: Asocia las columnas numéricas de la hoja `AUD-COMP` con las categorías **IVA**, **BASE** y **BASE 2** utilizando los interruptores (chips) de color. Existe exclusión mutua automática entre las columnas marcadas como IVA y BASE. Permite asignación masiva de "Todo" o "Ninguno" por categoría.

### 3. Vista Previa
* Puedes previsualizar el contenido de cualquiera de las hojas cargadas en el archivo de Excel usando la pestaña **Vista Previa** antes de ejecutar el proceso.

### 4. Ejecución y Logs
* Ve a la pestaña **Ejecución y Logs** y haz clic en **Iniciar Conciliación**.
* La consola integrada mostrará en tiempo real los mensajes de diagnóstico y el avance del proceso.
* Podrás verificar el estado a través de la barra de progreso e indicador visual de estado.
* Si deseas revisar el historial completo de mensajes, puedes hacer clic en **Abrir Carpeta de Logs**.

---

## Funcionamiento del Proceso

El script realiza la carga, limpieza, cruce y formato del archivo de Excel importado.

### Hojas de Entrada Esperadas
El libro de Excel de entrada debe contener las siguientes hojas (cuyos nombres se pueden mapear de forma interactiva en la pestaña de configuración):

1. **`Token`** (Principal): Información fiscal de facturas recibidas de la DIAN. Requiere columnas como `NIT Emisor`, `NIT Receptor`, `Prefijo`, `Folio`, `Total`, `IVA`. Se omiten automáticamente los registros cuyo tipo de documento contenga "application response" o su grupo contenga "emitido".
2. **`contabilidad`** (Opcional): Registros contables internos. Requiere columnas como `NIT`, `TIPO-DETALLE`, `TIPO`. Si no se especifica, se asume vacía y se continúa la conciliación con los demás datos.
3. **`TERCEROS`** (Opcional): Maestro de terceros creados en el sistema. Requiere columna `NIT`. Si no se especifica, se asume vacía.
4. **`AUD-COMP`**: Auxiliar contable para comprobación de auditoría (contiene cuentas de IVA y base de compras).
5. **`AUTORRETENEDORES`** (Opcional): Lista de NITs catalogados como autorretenedores, con columnas `NIT` y `COMENTARIO` (o variantes similares).

### Hojas de Salida Generadas
El motor escribe y modifica directamente el archivo de Excel original, agregando/reemplazando las siguientes hojas y datos:

* **`DIAN VS CONT` (Visible)**: Reporte principal de conciliación. Contiene una tabla con estilo estructurado y autoajuste de columnas donde se comparan los documentos de la DIAN con los registros contables. Las filas sin correspondencia se resaltan en color **rojo suave**. Concilia las columnas de `BASE`, `BASE_2` e `IVA`.
* **Hoja `AUD-COMP` (Modificada con Inyección)**: Se insertan dos nuevas columnas llamadas `BASE` e `IVA` directamente en la pestaña original de `AUD-COMP`, inmediatamente a la derecha de la columna `Tercero`. Los valores se calculan automáticamente y se les aplica formato de moneda preservando el formato original del documento.
* **Hoja Principal (Modificada)**: Se añaden las columnas `CONCEPTO` y `TERCERO` a cada registro. Los gastos catalogados con tipo `PERSONALES` se resaltan en **rojo suave**.
* **Hojas Intermedias (Ocultas automáticamente)**:
  * `Resultados`: Hoja de procesamiento intermedio para cruces.
  * `auditoria`: Registros filtrados excluyendo los de tipo personal.

---

## Licencia

Este proyecto es de uso privado e interno. Todos los derechos reservados.
