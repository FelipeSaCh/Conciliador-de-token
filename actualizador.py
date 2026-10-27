# actualizador.py — actualizaciones automáticas de Auditor desde el servidor.
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
from pathlib import Path

import version as _version

__version__ = str(getattr(_version, "__version__", None) or getattr(_version, "VERSION", None)
                  or getattr(_version, "APP_VERSION", "0.0.0"))

APP = "Auditor"
PREFIJO_INSTALADOR = "Auditor_Setup"
RUTA_ACTUALIZACIONES = r"\\SRV-CONTAB\Servidor de Archivos\2025 SERVIDOR\APLICATIVO FELIPE\Auditor"
ARCHIVO_VERSION = "version.json"
ARCHIVO_RUTA_LOCAL = "actualizacion.json"
ARGS_SILENCIOSO = ("/SILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/CLOSEAPPLICATIONS")


def tupla(v: str) -> tuple:
    return tuple(int(p) if p.isdigit() else 0 for p in str(v).strip().split("."))


def es_compilado() -> bool:
    return bool(getattr(sys, "frozen", False) or getattr(sys, "_MEIPASS", None) or "__compiled__" in globals())


def sha256(ruta) -> str:
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def _ruta_local() -> str | None:
    try:
        with open(Path(sys.argv[0]).resolve().parent / ARCHIVO_RUTA_LOCAL, encoding="utf-8-sig") as f:
            return json.load(f).get("ruta") or None
    except (OSError, ValueError, AttributeError):
        return None


def ruta_servidor() -> Path:
    return Path(_ruta_local() or RUTA_ACTUALIZACIONES)


def leer_publicada(carpeta: Path) -> dict | None:
    try:
        with open(carpeta / ARCHIVO_VERSION, encoding="utf-8-sig") as f:
            return json.load(f)
    except FileNotFoundError:
        return None


def buscar() -> dict | None:
    carpeta = ruta_servidor()
    info = leer_publicada(carpeta)
    if not info or tupla(info.get("version", "0")) <= tupla(__version__):
        return None
    instalador = carpeta / info["instalador"]
    if not instalador.exists():
        raise FileNotFoundError(f"No se encontró el instalador publicado: {instalador}")
    return {**info, "ruta_instalador": str(instalador)}


def instalar(info: dict) -> None:
    carpeta = Path(os.getenv("LOCALAPPDATA") or tempfile.gettempdir()) / APP / "actualizaciones"
    carpeta.mkdir(parents=True, exist_ok=True)
    for viejo in carpeta.glob("*.exe"):
        try:
            viejo.unlink()
        except OSError:
            pass
    destino = carpeta / Path(info["ruta_instalador"]).name
    shutil.copy2(info["ruta_instalador"], destino)
    if info.get("sha256") and sha256(destino) != info["sha256"]:
        destino.unlink(missing_ok=True)
        raise IOError("El instalador copiado está incompleto o dañado. Verifica la conexión a la red e intenta de nuevo.")
    flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen([str(destino), *ARGS_SILENCIOSO], creationflags=flags, close_fds=True, cwd=str(carpeta))


# --- Integración con Tkinter ------------------------------------------------
def verificar(ventana, manual: bool = False, al_cerrar=None) -> None:
    """Busca actualizaciones en segundo plano y pregunta al usuario. al_cerrar: función que guarda y cierra la app."""
    from tkinter import messagebox
    res = {}

    def trabajo():
        try:
            res["info"] = buscar()
        except Exception as e:
            res["error"] = e

    hilo = threading.Thread(target=trabajo, daemon=True)
    hilo.start()

    def esperar():
        if hilo.is_alive():
            ventana.after(300, esperar)
            return
        if "error" in res:
            if manual:
                messagebox.showwarning("Actualizaciones", f"No se pudo consultar el servidor:\n{res['error']}", parent=ventana)
            return
        info = res.get("info")
        if not info:
            if manual:
                messagebox.showinfo("Actualizaciones", f"Tienes la última versión (v{__version__}).", parent=ventana)
            return
        notas = f"\n\nNovedades:\n{info['notas']}" if info.get("notas") else ""
        msg = (f"Hay una nueva versión disponible: v{info['version']} (instalada: v{__version__}).{notas}\n\n"
               "¿Actualizar ahora? La aplicación se cerrará y se abrirá de nuevo.")
        if info.get("obligatoria"):
            messagebox.showinfo("Actualización obligatoria", msg.replace("¿Actualizar ahora?", "Esta actualización es obligatoria."), parent=ventana)
        elif not messagebox.askyesno("Actualización disponible", msg, parent=ventana):
            return
        try:
            instalar(info)
        except Exception as e:
            messagebox.showerror("Actualizaciones", f"No se pudo iniciar la actualización:\n{e}", parent=ventana)
            return
        (al_cerrar or ventana.destroy)()

    esperar()


def programar(ventana, al_cerrar=None, demora_ms: int = 3000) -> None:
    """Revisión automática al abrir; solo en la app instalada (no durante el desarrollo)."""
    if es_compilado():
        ventana.after(demora_ms, lambda: verificar(ventana, False, al_cerrar))