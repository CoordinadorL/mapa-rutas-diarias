"""
Antes de que publicar_datos.bat sobreescriba latest.json con la ruta
nueva, guarda una copia de la ruta que estaba publicada en
historial/<fecha-vieja>.json (para que quede el historico). Ademas guarda
de una vez la ruta nueva en historial/<fecha-nueva>.json
(guardar_en_historial).

No hace nada si:
- Todavia no hay ningun latest.json publicado (primera vez).
- La fecha no cambio (se esta republicando el mismo dia, ej. una
  correccion), para no pisar el historial con una version a medias.
- Esa fecha ya estaba archivada (no se sobreescribe).

Uso:
    python archivar_dia.py <carpeta_repo_datos> <archivo_nuevo_latest_json>
"""
import json
import shutil
import sys
from pathlib import Path


def archivar(repo: Path, nuevo: Path) -> str:
    """Hace el archivado y devuelve un mensaje de lo que paso (tambien lo
    usa generar_rutas.py)."""
    actual = repo / "latest.json"
    if not actual.exists():
        return "No hay latest.json previo todavia, nada que archivar."

    with open(actual, encoding="utf-8") as f:
        fecha_actual = json.load(f).get("fecha")

    with open(nuevo, encoding="utf-8") as f:
        fecha_nueva = json.load(f).get("fecha")

    if not fecha_actual or fecha_actual == fecha_nueva:
        return f"No hace falta archivar (la fecha no cambio: {fecha_actual})."

    carpeta_historial = repo / "historial"
    carpeta_historial.mkdir(exist_ok=True)
    destino = carpeta_historial / f"{fecha_actual}.json"

    if destino.exists():
        return f"historial/{fecha_actual}.json ya existia, no se toca."

    shutil.copy(actual, destino)
    return f"Archivado: historial/{fecha_actual}.json"


def guardar_en_historial(repo: Path, nuevo: Path) -> str:
    """Guarda YA la ruta que se esta publicando en historial/<fecha>.json,
    sin esperar a que se publique el dia siguiente. Asi el historico no
    depende del maestro (que se vacia y se vuelve a llenar cada dia). Si
    se republica el mismo dia (una correccion), se reemplaza."""
    with open(nuevo, encoding="utf-8") as f:
        fecha_nueva = json.load(f).get("fecha")
    if not fecha_nueva:
        return "La ruta nueva no tiene fecha, no se guardo en el historial."
    carpeta_historial = repo / "historial"
    carpeta_historial.mkdir(exist_ok=True)
    shutil.copy(nuevo, carpeta_historial / f"{fecha_nueva}.json")
    return f"Guardada en el historial: historial/{fecha_nueva}.json"


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        raise SystemExit(1)

    repo, nuevo = Path(sys.argv[1]), Path(sys.argv[2])
    print(archivar(repo, nuevo))
    print(guardar_en_historial(repo, nuevo))


if __name__ == "__main__":
    main()
