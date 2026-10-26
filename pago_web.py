"""
Pago a transportes: lo que esta PC le manda a la web (el backend Apps Script
de la app OPAUSTRO, seccion "Pago a transportes").

- Al publicar la ruta del dia (generar_rutas.subir_pago) se manda el resumen
  por camion: kilos, N de facturas y $ de cada uno. En la web queda como el
  "dia" (lo que antes era una pestana a mano en el Excel
  CAMIONES-AYUDANTES-FRECUENCIAS).
- La sincronizacion con ese Excel (pago_excel.py) usa llamar() para traer y
  mandar los cambios.

La PC entra con una "clave de PC" (no con usuario y contrasena): la pagina
"Pago a transportes", abierta en esta PC con sesion de coordinador o admin,
la pide al backend y se la pasa al servicio, que la guarda en
data/pago_config.json (data/ no se sube a GitHub).

Si no hay internet, el dia queda en data/pago_pendientes.json y se manda en
la siguiente sincronizacion.
"""
import json
import os
import threading
import traceback
import urllib.error
import urllib.request
from datetime import date
from pathlib import Path

CARPETA = Path(__file__).resolve().parent
# PAGO_CARPETA_DATOS: solo para pruebas (config, pendientes y foto en otra carpeta).
CARPETA_DATOS = Path(os.environ.get("PAGO_CARPETA_DATOS") or CARPETA / "data")
ARCHIVO_CONFIG = CARPETA_DATOS / "pago_config.json"
ARCHIVO_PENDIENTES = CARPETA_DATOS / "pago_pendientes.json"
ARCHIVO_ERRORES = CARPETA / "data" / "generador_error.log"
# El mismo Web App que usa la app (cargue.html, historial-cargues.html).
API_URL = ("https://script.google.com/macros/s/AKfycbwzQRnCJHFXni_9Qvs9kh_6BoikPDtb2sAMNChS3gSk9gAn5e31z5PF"
           "vKRGRzeoFqK5og/exec")

_BLOQUEO_ARCHIVOS = threading.Lock()


class ErrorPago(Exception):
    """Error con un mensaje ya pensado para mostrarse tal cual."""


def registrar_error(detalle: str) -> None:
    try:
        from datetime import datetime
        ARCHIVO_ERRORES.parent.mkdir(exist_ok=True)
        with open(ARCHIVO_ERRORES, "a", encoding="utf-8") as f:
            f.write(f"--- {datetime.now():%Y-%m-%d %H:%M:%S} (pago a transportes)\n{detalle}\n")
    except OSError:
        pass


# ---------------------------------------------------------------------------
# Archivos locales
# ---------------------------------------------------------------------------

def _leer_json(ruta: Path, por_defecto):
    try:
        return json.loads(ruta.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return por_defecto


def _escribir_json(ruta: Path, datos) -> None:
    ruta.parent.mkdir(exist_ok=True)
    temporal = ruta.with_suffix(".tmp")
    temporal.write_text(json.dumps(datos, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(temporal, ruta)


def cargar_config() -> dict:
    return _leer_json(ARCHIVO_CONFIG, {})


def guardar_config(config: dict) -> None:
    with _BLOQUEO_ARCHIVOS:
        _escribir_json(ARCHIVO_CONFIG, config)


def guardar_clave(clave: str) -> None:
    clave = str(clave or "").strip()
    if len(clave) < 20:
        raise ErrorPago("La clave de PC que llegó no es válida.")
    config = cargar_config()
    config["clave"] = clave
    guardar_config(config)


def tiene_clave() -> bool:
    return bool(cargar_config().get("clave"))


# ---------------------------------------------------------------------------
# Backend
# ---------------------------------------------------------------------------

def llamar(tipo: str, datos: dict, timeout=90) -> dict:
    """POST al backend con la clave de PC. Devuelve la respuesta (ok=True) o
    lanza ErrorPago con un mensaje para mostrar."""
    config = cargar_config()
    clave = config.get("clave")
    if not clave:
        raise ErrorPago("Esta PC todavía no tiene la clave de Pago a transportes: abre \"Pago a transportes\" "
                        "en la app, en esta PC, con usuario coordinador o admin (se configura solo).")
    cuerpo = json.dumps(dict(datos, tipo=tipo, clave=clave), ensure_ascii=False, default=str).encode("utf-8")
    pedido = urllib.request.Request(config.get("api_url") or API_URL, data=cuerpo,
                                    headers={"Content-Type": "application/json"})
    try:
        # Apps Script contesta el POST con un 302; urllib lo sigue con GET,
        # que es justo lo que hace falta para leer la respuesta.
        with urllib.request.urlopen(pedido, timeout=timeout) as r:
            texto = r.read().decode("utf-8", "replace")
    except (urllib.error.URLError, OSError) as e:
        raise ErrorPago(f"No hay conexión con la web ({getattr(e, 'reason', e)}).")
    try:
        respuesta = json.loads(texto)
    except ValueError:
        raise ErrorPago("La web respondió algo inesperado (¿se publicó la versión nueva del backend?): "
                        + texto[:150])
    if not respuesta.get("ok"):
        if respuesta.get("auth") is False:
            raise ErrorPago("La web no aceptó la clave de esta PC. Abre \"Pago a transportes\" en la app, en esta "
                            "PC, con usuario coordinador o admin para renovarla.")
        raise ErrorPago(respuesta.get("error") or "La web no aceptó el pedido.")
    return respuesta


# ---------------------------------------------------------------------------
# Dia publicado -> web
# ---------------------------------------------------------------------------

def entrega_de(fecha_liq: date) -> date:
    import mymaps
    return mymaps.entrega_sugerida(fecha_liq)


def resumen_del_dia(seleccion, fechas) -> dict:
    """Lo que se manda por camion (o por cargue, si va separado)."""
    fecha_liq = max(fechas)
    con_facturas = any(c.get("facturas") for c in seleccion)
    camiones = []
    for c in seleccion:
        clave = str(c["id"])
        cod, _, cargue = clave.partition(" CARGUE ")
        camiones.append({
            "clave_carga": clave,
            "cod_camion": cod.strip(),
            "cargue": ("CARGUE " + cargue) if cargue else "",
            "chofer": c.get("chofer", ""),
            "kgs": round(sum(f.get("Suma de Peso") or 0 for f in c["filas"]), 2),
            # Una base sin las columnas Numero/Tipo no trae facturas: se deja
            # vacio (para escribirlo a mano) en vez de mandar 0.
            "facturas": c.get("facturas", 0) if con_facturas else "",
            "monto": round(c.get("total") or 0, 2),
        })
    return {"fecha_liq": fecha_liq.isoformat(), "fecha_entrega": entrega_de(fecha_liq).isoformat(),
            "camiones": camiones}


def _pendientes() -> dict:
    return _leer_json(ARCHIVO_PENDIENTES, {})


def _anotar_pendiente(datos: dict) -> None:
    with _BLOQUEO_ARCHIVOS:
        pendientes = _pendientes()
        pendientes[datos["fecha_liq"]] = datos
        _escribir_json(ARCHIVO_PENDIENTES, pendientes)


def _quitar_pendiente(datos: dict) -> None:
    with _BLOQUEO_ARCHIVOS:
        pendientes = _pendientes()
        # Solo si sigue siendo ESE envio (pudo publicarse otra vez mientras tanto).
        if pendientes.get(datos["fecha_liq"]) == datos:
            del pendientes[datos["fecha_liq"]]
            _escribir_json(ARCHIVO_PENDIENTES, pendientes)


def enviar_pendientes(log=None) -> int:
    """Manda los dias que quedaron sin enviar. Devuelve cuantos faltan."""
    pendientes = _pendientes()
    for fecha in sorted(pendientes):
        datos = pendientes[fecha]
        try:
            r = llamar("pago_subir_dia", datos)
        except ErrorPago as e:
            if log:
                log(f"Pago a transportes: el día {fecha} quedó pendiente de enviar — {e}")
            continue
        _quitar_pendiente(datos)
        if log and r.get("ignorado"):
            log(f"Pago a transportes: el día {fecha} no se cargó — {r.get('aviso')}")
        elif log:
            log(f"Pago a transportes: día {fecha} enviado ({r.get('nuevas', 0)} camiones nuevos, "
                f"{r.get('actualizadas', 0)} actualizados).")
    return len(_pendientes())


def subir_dia(seleccion, fechas, log=None) -> None:
    """Anota el dia como pendiente (al instante) y lo manda en segundo plano:
    si el envio no llega (sin internet, se cerro la ventana), la siguiente
    sincronizacion lo vuelve a intentar. Nunca lanza: la ruta ya quedo
    publicada y esto no debe convertirlo en un error."""
    try:
        _anotar_pendiente(resumen_del_dia(seleccion, fechas))
    except Exception:
        registrar_error(traceback.format_exc())
        return

    def enviar():
        try:
            enviar_pendientes(log)
        except Exception:
            registrar_error(traceback.format_exc())

    threading.Thread(target=enviar, daemon=True).start()
