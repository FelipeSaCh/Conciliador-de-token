# Changelog / Historial de Cambios

Todos los cambios notables en este proyecto serán documentados en este archivo.

El formato se basa en [Keep a Changelog](https://keepachangelog.com/es-ES/1.0.0/) y este proyecto se adhiere a [SemVer](https://semver.org/lang/es/).

## [1.6.0] - 2026-08-25

### Añadido
- **Mapeo dinámico de columnas de AUD-COMP**: Reemplazo de seriales de IVA y BASE estáticos por selección directa de columnas leídas desde la pestaña `AUD-COMP` en la GUI.
- **Categoría BASE 2**: Soporte para una segunda base de cálculo (`BASE_2`) seleccionable desde la GUI e integrada en las columnas de salida.
- **Inyección directa en AUD-COMP**: La base y el IVA calculados se inyectan en columnas nuevas directamente en la hoja original `AUD-COMP` justo después de la columna "Tercero", conservando su formato original.
- **Botón de limpieza para hojas opcionales**: Botón "✖" en la GUI para limpiar rápidamente los campos opcionales.

### Modificado
- **Hojas opcionales**: `contabilidad` y `terceros` ahora son opcionales. La aplicación puede ejecutar la conciliación utilizándolas como DataFrames vacíos si no se suministran.
- **Filtro de exclusión en la hoja principal**: Se filtran y descartan registros con tipo de documento "application response" y grupo "emitido".
- **Mejoras en GUI**:
  - Exclusión mutua automática entre la asignación de columnas IVA y BASE en la GUI.
  - Paleta de colores personalizada con categorización visual (IVA en Naranja, BASE en Verde, BASE 2 en Morado).
  - Control de scroll mejorado en `ScrollableChecklist` para prevenir interferencias con el scroll de la ventana principal.
  - Cambio del motor de temas a `clam` con estilos personalizados.
- Se eliminó la hoja intermedia `resultados-auditoria` al inyectarse los datos directamente en `AUD-COMP`.

## [1.5.0] - 2026-08-24

### Añadido
- **Nueva Interfaz Gráfica de Usuario (GUI)** en [`app.py`](file:///c:/Users/USUARIO/Desktop/Proyectos/Conciliador%20de%20token/app.py) basada en `tkinter` y el tema moderno `sv_ttk` (modo claro).
- **Mapeo dinámico de hojas**: Interfaz para asociar cada pestaña requerida (Token, Contabilidad, Terceros, AUD-COMP y la opcional Autorretenedores) con el nombre exacto de la pestaña del libro cargado.
- **Configuración dinámica de seriales**: Paneles de checklist con barras de control ("Todos"/"Ninguno") en la GUI para seleccionar qué seriales de IVA y de base se desean incluir/excluir en el procesamiento.
- **Vista previa integrada**: Nueva pestaña en la GUI que permite previsualizar las hojas del archivo de Excel antes de ejecutar el cruce.
- **Consola de ejecución en tiempo real**: Registro visual de actividad (logs de procesamiento) en un widget de texto con barra de progreso, indicador de estado de color y botón rápido para abrir la carpeta de logs.
- Centralización del número de versión de la aplicación en [`version.py`](file:///c:/Users/USUARIO/Desktop/Proyectos/Conciliador%20de%20token/version.py).
- Carpeta [`assets`](file:///c:/Users/USUARIO/Desktop/Proyectos/Conciliador%20de%20token/assets) con los recursos gráficos del icono de la aplicación (`icon.ico` y `icon.png`).
- Ignorado de las carpetas de compilados `build/`, `dist/` y archivos `.spec` en el control de versiones `.gitignore`.

### Modificado
- [`main.py`](file:///c:/Users/USUARIO/Desktop/Proyectos/Conciliador%20de%20token/main.py): Ahora inicializa y lanza el bucle principal de la interfaz gráfica (`ConciliadorApp`).
- [`config.py`](file:///c:/Users/USUARIO/Desktop/Proyectos/Conciliador%20de%20token/config.py): Ahora importa la versión centralizada desde `version.py`.
- [`excel_engine.py`](file:///c:/Users/USUARIO/Desktop/Proyectos/Conciliador%20de%20token/excel_engine.py): Adaptado para recibir y utilizar dinámicamente los seriales de IVA y BASE seleccionados desde la GUI en lugar de importar listas estáticas.

## [1.0.0] - 2026-08-19

### Añadido
- Archivo [`.gitignore`](file:///c:/Users/USUARIO/Desktop/Proyectos/Conciliador%20de%20token/.gitignore) para ignorar carpetas de entornos virtuales (`venv/`), compilados de Python, configuraciones de editores (`.vscode/`, `.idea/`) y archivos de Excel temporales o de datos.
- Archivo [`requirements.txt`](file:///c:/Users/USUARIO/Desktop/Proyectos/Conciliador%20de%20token/requirements.txt) con las dependencias mínimas requeridas (`pandas`, `numpy`, `openpyxl`) para el procesamiento de datos y formateo de hojas de cálculo.
- Documentación principal en [`README.md`](file:///c:/Users/USUARIO/Desktop/Proyectos/Conciliador%20de%20token/README.md), explicando el funcionamiento, la estructura de carpetas, dependencias, cómo configurar la ruta del archivo Excel en `main.py` y las hojas/columnas esperadas.
- Este archivo [`CHANGELOG.md`](file:///c:/Users/USUARIO/Desktop/Proyectos/Conciliador%20de%20token/CHANGELOG.md) para el control del historial del proyecto.

### Conservado
- Archivo [`main.py`](file:///c:/Users/USUARIO/Desktop/Proyectos/Conciliador%20de%20token/main.py) original intacto, sin alterar su lógica funcional de procesamiento, cruce y formateo de Excel según las especificaciones del usuario.
