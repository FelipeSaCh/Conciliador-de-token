# publicar.py — publica la versión compilada de Auditor en el servidor para que los usuarios la reciban.
import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
sys.path.insert(0, str(RAIZ))
import actualizador as act  # noqa: E402
from actualizador import __version__  # noqa: E402


def _salir(msg: str) -> None:
    print(f"\n*** {msg}")
    sys.exit(1)


def main() -> None:
    instalador = RAIZ / "instaladores" / f"{act.PREFIJO_INSTALADOR}_{__version__}.exe"
    if not instalador.exists():
        _salir(f"No existe {instalador}. Ejecuta primero build.bat.")

    destino = act.ruta_servidor()
    print(f"Servidor: {destino}")
    try:
        destino.mkdir(parents=True, exist_ok=True)
        prueba = destino / ".prueba_escritura"
        prueba.write_text("ok", encoding="utf-8")
        prueba.unlink()
    except OSError as e:
        _salir(f"No se puede escribir en el servidor ({e}).\n"
               "Verifica que el equipo esté conectado a la red empresarial (o VPN) y que tengas permiso de escritura.")

    publicada = act.leer_publicada(destino)
    if publicada:
        print(f"Versión publicada actualmente: v{publicada.get('version')}")
        if act.tupla(publicada.get("version", "0")) >= act.tupla(__version__):
            if input(f"La v{__version__} no es mayor que la publicada. ¿Publicar de todas formas? (s/N): ").strip().lower() != "s":
                _salir("Publicación cancelada. Sube el número en version.py y vuelve a ejecutar build.bat.")

    notas = input(f"Novedades de la versión {__version__} (Enter para omitir): ").strip()
    obligatoria = input("¿Actualización obligatoria? (s/N): ").strip().lower() == "s"

    print("Copiando instalador al servidor...")
    huella = act.sha256(instalador)
    final = destino / instalador.name
    parcial = destino / f"{instalador.name}.parcial"
    shutil.copy2(instalador, parcial)
    if act.sha256(parcial) != huella:
        parcial.unlink(missing_ok=True)
        _salir("La copia al servidor quedó incompleta. Intenta de nuevo.")
    os.replace(parcial, final)

    info = {
        "version": __version__,
        "instalador": final.name,
        "notas": notas,
        "obligatoria": obligatoria,
        "sha256": huella,
        "publicado": datetime.now().isoformat(timespec="seconds"),
        "publicado_por": os.getenv("USERNAME", ""),
    }
    temporal = destino / f"{act.ARCHIVO_VERSION}.tmp"
    with open(temporal, "w", encoding="utf-8") as f:
        json.dump(info, f, ensure_ascii=False, indent=2)
    os.replace(temporal, destino / act.ARCHIVO_VERSION)
    print(f"\nVersión {__version__} publicada en {destino}")


if __name__ == "__main__":
    main()