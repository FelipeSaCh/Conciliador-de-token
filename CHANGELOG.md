# Changelog / Historial de Cambios

Todos los cambios notables en este proyecto serán documentados en este archivo.

El formato se basa en [Keep a Changelog](https://keepachangelog.com/es-ES/1.0.0/) y este proyecto se adhiere a [SemVer](https://semver.org/lang/es/).

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
