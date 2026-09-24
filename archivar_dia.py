"""
Antes de que publicar_datos.bat sobreescriba latest.json con la ruta
nueva, guarda una copia de la ruta que estaba publicada en
historial/<fecha-vieja>.json (para que quede el historico).

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


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        raise SystemExit(1)

    repo = Path(sys.argv[1])
    nuevo = Path(sys.argv[2])

    actual = repo / "latest.json"
    if not actual.exists():
        print("No hay latest.json previo todavia, nada que archivar.")
        return

    with open(actual, encoding="utf-8") as f:
        fecha_actual = json.load(f).get("fecha")

    with open(nuevo, encoding="utf-8") as f:
        fecha_nueva = json.load(f).get("fecha")

    if not fecha_actual or fecha_actual == fecha_nueva:
        print(f"No hace falta archivar (la fecha no cambio: {fecha_actual}).")
        return

    carpeta_historial = repo / "historial"
    carpeta_historial.mkdir(exist_ok=True)
    destino = carpeta_historial / f"{fecha_actual}.json"

    if destino.exists():
        print(f"historial/{fecha_actual}.json ya existia, no se toca.")
        return

    shutil.copy(actual, destino)
    print(f"Archivado: historial/{fecha_actual}.json")


if __name__ == "__main__":
    main()
