"""
Servicio local para usar "Actualizar ruta" y "My Maps" desde la app web
OPAUSTRO (coordinadorl.github.io/Mapas-Diarios: secciones "Generar mapas" y
"My Maps"), en vez de las ventanas.

Hace exactamente lo mismo que las ventanas (generar_rutas.py y
ventana_mymaps.py): usa las mismas funciones de generar_rutas.py,
control_rutas.py y mymaps.py, asi que no cambia como salen los mapas. Solo
cambia la pantalla: en vez de una ventana de Windows, la pagina web le pide
a este servicio lo que hay que hacer y le pregunta como va.

- Se abre con doble clic en "SERVICIO MAPAS.bat" (queda escondido).
- Solo escucha en esta PC (127.0.0.1:5178) y solo atiende a la app
  (encabezado Origin + X-Opaustro): otra pagina web no puede usarlo.
- No usar al mismo tiempo que las ventanas viejas (ACTUALIZAR RUTA.bat /
  MY MAPS.bat): las dos publicarian o moverian el robot a la vez.

Prueba sin publicar ni usar el robot (todo lo que escribe va a una carpeta
temporal):  python servicio_web.py --prueba
"""
import json
import os
import socket
import subprocess
import sys
import tempfile
import threading
import time
import traceback
from datetime import date, datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs

import generar_rutas as gr
import mymaps

PUERTO = 5178
VERSION = "2"
ORIGENES_PERMITIDOS = {
    "https://coordinadorl.github.io",
    "http://localhost:8080",        # pruebas de la app en esta PC
    "http://127.0.0.1:8080",
}
LINEAS_LOG = 300
MODO_PRUEBA = "--prueba" in sys.argv
INSTANCIA = f"{os.getpid()}-{int(time.time())}"


def ahora():
    return datetime.now().strftime("%H:%M:%S")


def fecha_iso(d):
    return d.isoformat() if d else None


# ---------------------------------------------------------------------------
# Base comun: registro, avisos y preguntas para la pagina
#
# Las ventanas muestran mensajes (messagebox) y preguntas (askyesno). Aqui
# quedan guardados como "avisos" y "pregunta": la pagina los muestra y
# contesta (aviso_visto / responder).
# ---------------------------------------------------------------------------

class Controlador:
    def __init__(self):
        self.lock = threading.RLock()
        self.log_lineas = []
        self.log_n = 0
        self.avisos = []
        self.aviso_n = 0
        self.pregunta = None
        self.pregunta_n = 0
        self.version = 0

    def cambio(self):
        self.version += 1

    def log(self, mensaje):
        with self.lock:
            for linea in str(mensaje).rstrip().splitlines() or [""]:
                self.log_n += 1
                self.log_lineas.append({"n": self.log_n, "t": linea})
            del self.log_lineas[:-LINEAS_LOG]

    def avisar(self, tipo, titulo, texto, **extra):
        with self.lock:
            self.aviso_n += 1
            self.avisos.append({"id": self.aviso_n, "tipo": tipo, "titulo": titulo, "texto": texto, **extra})

    def aviso_visto(self, cuerpo):
        with self.lock:
            self.avisos = [a for a in self.avisos if a["id"] != cuerpo.get("id")]
        return {"ok": True}

    def preguntar(self, titulo, texto, accion, datos=None):
        """Una sola pregunta a la vez (como un askyesno): si ya hay otra
        abierta no se reemplaza (su 'Si' se perderia), se avisa."""
        with self.lock:
            if self.pregunta is not None:
                self.avisar("aviso", titulo, "No se pudo preguntar porque había otra pregunta abierta: contéstala y "
                                             "vuelve a intentar.")
                return {"ok": False, "error": "Primero contesta la pregunta que está abierta.", "ya_avisado": True}
            self.pregunta_n += 1
            self.pregunta = {"id": self.pregunta_n, "titulo": titulo, "texto": texto,
                             "accion": accion, "datos": datos, "version": self.version}
        return {"ok": True, "pregunta": True}

    def responder(self, cuerpo):
        with self.lock:
            p = self.pregunta
            if not p or p["id"] != cuerpo.get("id"):
                return {"ok": False, "error": "Esa pregunta ya no está vigente."}
            self.pregunta = None
        try:
            self.al_responder(p, bool(cuerpo.get("si")))
        except gr.ErrorRuta as e:
            self.fallo(str(e))
        except PermissionError as e:
            self.fallo(f"No pude guardar:\n{e.filename}\n\n¿Está abierto en Excel? Ciérralo y vuelve a intentar.")
        except Exception:
            detalle = traceback.format_exc()
            gr.registrar_error("(servicio web)\n" + detalle)
            self.fallo("Algo falló inesperadamente:\n\n" + detalle.strip().splitlines()[-1])
        finally:
            self.despues_de_pregunta()
        return {"ok": True}

    def al_responder(self, pregunta, si):
        raise NotImplementedError

    def fallo(self, texto):
        """Un error al contestar una pregunta: siempre le llega al usuario."""
        self.log("ERROR: " + texto)
        self.avisar("error", "No se pudo completar", texto)

    def despues_de_pregunta(self):
        pass

    def base_estado(self):
        p = self.pregunta
        # instancia: cambia en cada arranque del servicio; la pagina la usa para
        # darse cuenta de que el servicio se reinicio (los id vuelven a 1).
        return {"log": list(self.log_lineas), "avisos": list(self.avisos),
                "pregunta": {k: p[k] for k in ("id", "titulo", "texto")} if p else None,
                "version": self.version, "instancia": INSTANCIA}


# ---------------------------------------------------------------------------
# Actualizar ruta (lo mismo que la ventana de generar_rutas.py)
# ---------------------------------------------------------------------------

class Generador(Controlador):
    def __init__(self):
        super().__init__()
        self.config = gr.cargar_config()
        self.archivo_base = None
        self.base_info = ""
        self.filas = []
        self.fechas = []                    # las que se muestran, la mas nueva primero
        self.marcadas = set()
        self.camiones_por_fecha = {}
        self.camiones = []
        self.incluidos = {}
        self.separados = {}
        self.editables = {"ubicaciones": [], "cambios": [], "foto": None}
        self.reglas = []
        self.estados_cambios = []
        self.cambios_ubic = []
        self.estados_ubic = {}
        self.carpeta_ultima = None
        self.carpeta_ultima_dia = None
        self.nota_sabado = False
        self.ocupado = False
        self.mensaje = ""
        self.overrides_base = gr.overrides_locales()
        self.overrides_al_dia = False
        self.overrides_listos = False
        self.sin_georef = []
        self.leyendo_control = False
        self.fallos_vigilancia = 0
        self.control_pendiente = None
        self.iniciado = False
        self.tabla = []
        self.resumen = ""
        self.mymaps_sugerido = None         # tras publicar: ir a My Maps con esa carpeta
        self.base_mtime = None              # cuando se guardo el archivo base que se leyo
        self.lectura_dia = None             # que dia se leyo (otro dia = empezar de cero)
        self.lectura_n = 0                  # para descartar una lectura que ya no es la vigente
        # My Maps lo cambia por "el robot esta trabajando" (usa los MAPA).
        self.otra_ocupada = lambda: False
        threading.Thread(target=self._vigilar_control, daemon=True).start()

    # -- utilidades --------------------------------------------------------
    def ocupar(self, ocupado, mensaje=""):
        with self.lock:
            self.ocupado = ocupado
            self.mensaje = mensaje if (mensaje or ocupado) else self.mensaje_listo()
            self.cambio()

    def pausada(self):
        return self.pregunta is not None

    def en_segundo_plano(self, funcion, *argumentos):
        def correr():
            try:
                funcion(*argumentos)
            except gr.ErrorRuta as e:
                self.error(str(e))
            except PermissionError as e:
                self.error(f"No pude guardar:\n{e.filename}\n\n¿Está abierto en Excel? Ciérralo y vuelve a "
                           "intentar.")
            except Exception:
                detalle = traceback.format_exc()
                gr.registrar_error(detalle)
                self.error("Algo falló inesperadamente:\n\n" + detalle.strip().splitlines()[-1])
        threading.Thread(target=correr, daemon=True).start()

    def error(self, texto):
        self.ocupar(False, "Hubo un problema, revisa el mensaje.")
        self.log("ERROR: " + texto)
        self.avisar("error", "No se pudo completar", texto)

    def fallo(self, texto):
        self.error(texto)

    def mensaje_listo(self):
        if not self.fechas:
            return ""
        if self.sin_georef:
            return (f"Hay {len(self.sin_georef)} clientes sin ubicación: corrígelos en la lista roja y luego "
                    "\"Generar y publicar en el mapa\".")
        if self.camiones and not self.overrides_listos:
            return "Revisando las ubicaciones corregidas en GitHub..."
        return "Revisa la fecha y los camiones, y dale a \"Generar y publicar en el mapa\"."

    # -- archivo base --------------------------------------------------------
    def iniciar(self, cuerpo=None):
        """Al abrir la pagina. El servicio queda abierto de un dia para otro,
        asi que:
        - otro dia (o nunca se leyo): como abrir la ventana de cero, con la
          configuracion recordada, camiones segun lo guardado y el archivo
          base mas reciente;
        - mismo dia y el archivo base se guardo despues de leerlo (se pegaron
          ventas): se vuelve a leer ESE archivo (aunque se haya elegido otro
          con "Cambiar archivo...");
        - mismo dia y nada cambio (ir a My Maps y volver, F5): se deja todo
          como estaba, igual que la ventana que seguia abierta.
        Si se esta trabajando o hay una pregunta abierta, no se toca nada."""
        with self.lock:
            if self.ocupado or self.pregunta:
                return {"ok": True}
            if self.archivo_base and self.lectura_dia == date.today():
                try:
                    mtime = self.archivo_base.stat().st_mtime
                except OSError:
                    mtime = None
                if mtime is not None and mtime == self.base_mtime:
                    return {"ok": True}
                if mtime is not None:
                    self.cargar_base(self.archivo_base)
                    return {"ok": True}
            self.iniciado = True
            self.config = gr.cargar_config()
            self.incluidos, self.separados = {}, {}
            base = gr.buscar_base_reciente(self.config)
            if base:
                self.cargar_base(base)
            else:
                self.archivo_base = None
                self.filas, self.camiones, self.fechas, self.marcadas = [], [], [], set()
                self.tabla, self.resumen, self.sin_georef = [], "", []
                self.base_info = "Elige el archivo con \"Cambiar archivo...\"."
                self.mensaje = "No encontré el archivo base."
                self.cambio()
        return {"ok": True}

    def releer(self, cuerpo=None):
        with self.lock:
            no = self._ocupado_o_pregunta()
            if no:
                return no
            if self.archivo_base:
                self.cargar_base(self.archivo_base)
        return {"ok": True}

    def elegir_base(self, cuerpo=None):
        with self.lock:
            no = self._ocupado_o_pregunta()
            if no:
                return no
            inicial = self.archivo_base.parent if self.archivo_base else gr.CARPETA_MAPAS_DIARIOS
            # Mientras el cuadro de Windows esta abierto nada mas puede empezar
            # (en la ventana el cuadro era modal).
            self.ocupar(True, "Elige el archivo base en el cuadro de Windows (puede haber quedado detrás del "
                              "navegador)...")
        try:
            ruta = elegir_archivo_windows(inicial)
        finally:
            self.ocupar(False)
        if not ruta:
            return {"ok": True, "cancelado": True}
        with self.lock:
            no = self._ocupado_o_pregunta()
            if no:
                return no
            self.cargar_base(Path(ruta))
        return {"ok": True}

    def base_cambio(self) -> bool:
        """Si el archivo base se guardo despues de leerlo (se pegaron ventas
        sin "Volver a leer"), se vuelve a leer y se pide revisar antes de
        seguir: nunca se publica una lectura vieja."""
        if not self.archivo_base:
            return False
        try:
            mtime = self.archivo_base.stat().st_mtime
        except OSError:
            return False
        if mtime == self.base_mtime:
            return False
        self.cargar_base(self.archivo_base)
        self.avisar("info", "Archivo base", "El archivo base cambió desde que se leyó (se guardaron datos nuevos), "
                    "así que se volvió a leer.\n\nRevisa la fecha y los camiones y vuelve a darle al botón.")
        return True

    def cargar_base(self, ruta: Path):
        with self.lock:
            self.lectura_n += 1
            n = self.lectura_n
            self.lectura_dia = date.today()
            try:
                self.base_mtime = ruta.stat().st_mtime
            except OSError:
                self.base_mtime = None
            self.archivo_base = ruta
            carpeta = ruta.parent
            if gr.CARPETA_MAPAS_DIARIOS in carpeta.parents:
                carpeta = Path("…") / carpeta.relative_to(gr.CARPETA_MAPAS_DIARIOS)
            try:
                guardado = datetime.fromtimestamp(ruta.stat().st_mtime)
                self.base_info = (f"Carpeta: {carpeta}\nGuardado por última vez: {gr.formato_fecha(guardado.date())} "
                                  f"{guardado:%H:%M} — si acabas de pegar datos nuevos, guarda el Excel y dale a "
                                  "\"Volver a leer\".")
            except OSError:
                self.base_info = f"Carpeta: {carpeta}"
            # Nada de la lectura anterior debe quedar: si esta lectura falla, no
            # se puede publicar con datos viejos sin darse cuenta.
            self.filas, self.camiones, self.separados = [], [], {}
            self.fechas, self.marcadas = [], set()
            self.nota_sabado = False
            self.tabla, self.resumen, self.sin_georef = [], "", []
            self.overrides_listos = False
            self.ocupar(True, "Leyendo el archivo base...")
        threading.Thread(target=self._traer_overrides, daemon=True).start()
        self.en_segundo_plano(self._leer, ruta, n)

    def _leer(self, ruta, n):
        inicio = time.time()
        try:
            filas = gr.leer_base(ruta)
        except Exception:
            if n != self.lectura_n:
                return                      # lectura vieja: su error ya no importa (la nueva sigue)
            raise
        self.control_leido(self.leer_control_seguro())
        self.base_leida(n, filas, time.time() - inicio)

    def _traer_overrides(self):
        try:
            overrides, al_dia = gr.overrides_vigentes()
        except Exception:
            gr.registrar_error(traceback.format_exc())
            overrides, al_dia = gr.overrides_locales(), False
        self.overrides_llegaron(overrides, al_dia)

    def overrides_llegaron(self, overrides, al_dia):
        with self.lock:
            ya_estaban = self.overrides_listos
            self.overrides_base, self.overrides_al_dia = overrides, al_dia
            self.overrides_listos = True
            if not al_dia:
                self.log("Aviso: sin conexión con GitHub; para revisar la georreferencia se usan las correcciones "
                         "de ubicación descargadas la última vez.")
            if self.fechas:
                self.pintar()
                if not ya_estaban and not self.ocupado:
                    self.mensaje = self.mensaje_listo()
            self.cambio()

    def leer_control_seguro(self, abrir_enlaces=True):
        import control_rutas
        try:
            return control_rutas.leer_editables(abrir_enlaces=abrir_enlaces)
        except Exception as e:
            if not isinstance(e, gr.ErrorRuta):
                gr.registrar_error(traceback.format_exc())
            motivo = str(e) if isinstance(e, gr.ErrorRuta) else "no se pudo leer"
            self.log(f"Aviso: Excel de control: {motivo} Se sigue sin sus cambios.")
            return {"ubicaciones": [], "cambios": [], "foto": None, "ilegible": motivo,
                    "mtime": control_rutas.mtime_control()}

    def control_leido(self, editables):
        import control_rutas
        reglas = control_rutas.reglas_desde(editables)
        cambios_ubic, estados_ubic = control_rutas.cambios_de_ubicacion(editables)
        with self.lock:
            if self.pausada():
                # Hay una pregunta abierta: no se cambia lo que se esta
                # confirmando; se aplica al contestarla.
                self.control_pendiente = editables
                return
            self.editables, self.reglas = editables, reglas
            self.cambios_ubic, self.estados_ubic = cambios_ubic, estados_ubic
            activas = sum(1 for r in self.reglas if r["activa"])
            if activas or self.cambios_ubic:
                self.log(f"Excel de control: {activas} cambios de ruta activos, "
                         f"{len(self.cambios_ubic)} ubicaciones por aplicar.")
            if self.fechas:
                self.refrescar_camiones()
            self.cambio()

    def despues_de_pregunta(self):
        pendiente, self.control_pendiente = self.control_pendiente, None
        if pendiente is not None and not self.ocupado:
            self.control_leido(pendiente)

    def base_leida(self, n, filas, segundos):
        with self.lock:
            if n != self.lectura_n:
                return                      # ya se pidio otra lectura despues: esa es la que vale
            self.filas = filas
            fechas = sorted({f["Fecha Liq"] for f in filas}, reverse=True)
            sugeridas = gr.fechas_sugeridas(fechas)
            self.camiones_por_fecha = {}
            for f in filas:
                self.camiones_por_fecha.setdefault(f["Fecha Liq"], set()).add(gr.id_camion(gr.texto(f["Chofer"])))
            self.fechas = fechas[:gr.FECHAS_A_MOSTRAR]
            self.marcadas = {d for d in self.fechas if d in sugeridas}
            self.nota_sabado = len(sugeridas) > 1
            self.log(f"Leídas {len(filas):,} filas de la base en {segundos:.0f} s.")
            self.refrescar_camiones()
            self.ocupar(False)

    def nota_fechas(self):
        if not self.fechas:
            return ""
        notas = []
        if self.nota_sabado:
            notas.append("Se juntó viernes + sábado porque la última fecha es sábado.")
        publicada = gr.fecha_publicada()
        if publicada:
            d = datetime.strptime(publicada, "%Y-%m-%d").date()
            notas.append(f"En el mapa está publicada la ruta del {gr.formato_fecha(d)}.")
        return "  ".join(notas)

    # -- camiones --------------------------------------------------------------
    def refrescar_camiones(self):
        fechas = set(self.marcadas)
        filas, self.estados_cambios = self.filas, []
        if fechas and self.reglas:
            import control_rutas
            filas, self.estados_cambios = control_rutas.aplicar_movimientos(self.filas, self.reglas, fechas)
        self.camiones = gr.agrupar(filas, fechas) if fechas else []
        excluidos = set(self.config.get("excluidos", []))
        separados_guardados = set(self.config.get("separados", {}).get(
            max(fechas).isoformat() if fechas else "", []))
        for camion in self.camiones:
            if camion["id"] not in self.incluidos:
                self.incluidos[camion["id"]] = camion["id"] not in excluidos
            if camion["id"] not in self.separados:
                self.separados[camion["id"]] = camion["id"] in separados_guardados
        self.pintar()

    def overrides_actuales(self):
        return gr.con_pendientes(self.overrides_base,
                                 gr.pendientes_aplicables(self.overrides_base, self.cambios_ubic)[0])

    def pintar(self):
        conocidos = set(self.config.get("conocidos", []))
        marcados = clientes = 0
        total = 0.0
        overrides = self.overrides_actuales()
        tabla = []
        for camion in self.camiones:
            cid = camion["id"]
            dentro = self.incluidos[cid]
            nuevo = bool(conocidos) and cid not in conocidos
            varios = len(camion["liquidaciones"]) > 1
            sin_geo = len(gr.clientes_sin_georef([camion], overrides))
            if not varios:
                cargues = "1"
            elif self.separados.get(cid):
                cargues = f"{len(camion['liquidaciones'])} · SEPARADOS"
            else:
                cargues = f"{len(camion['liquidaciones'])} · Unificado"
            tabla.append({"id": cid, "incluido": dentro, "nombre": gr.nombre_chofer(camion["chofer"]),
                          "nuevo": nuevo, "movidos": camion.get("movidos") or 0, "sin_geo": sin_geo,
                          "liq": ", ".join(str(liq) for liq in camion["liquidaciones"]),
                          "clientes": len(camion["filas"]), "total": f"${camion['total']:,.2f}",
                          "cargues": cargues, "varios": varios})
            if dentro:
                marcados += 1
                clientes += len(camion["filas"])
                total += camion["total"]
        texto = f"{marcados} de {len(self.camiones)} camiones · {clientes} clientes · ${total:,.2f}"
        aplicados = sum(1 for nivel, _ in self.estados_cambios if nivel == "ok")
        problemas = sum(1 for nivel, _ in self.estados_cambios if nivel == "error")
        if aplicados or problemas:
            texto += f" · {aplicados} cambios de ruta"
            if problemas:
                texto += f" ({problemas} con problema)"
        seleccion = gr.armar_seleccion(self.camiones, self.incluidos, self.separados)
        self.sin_georef = gr.clientes_sin_georef(seleccion, overrides)
        if seleccion:
            texto += (f" · {len(self.sin_georef)} SIN georreferencia" if self.sin_georef
                      else " · todos con georreferencia ✓")
        self.tabla, self.resumen = tabla, texto
        if not self.ocupado:
            self.mensaje = self.mensaje_listo()
        self.cambio()

    def errores_de_fila_ubicacion(self):
        import control_rutas
        filas = self.editables.get("ubicaciones", [])
        salida = {}
        for i, estado in (self.estados_ubic or {}).items():
            if isinstance(i, int) and i < len(filas):
                codigo = control_rutas.normalizar_codigo(filas[i].get("Código"))
                if codigo:
                    salida.setdefault(codigo, estado)
        _, conflicto = gr.pendientes_aplicables(self.overrides_base, self.cambios_ubic)
        for c in conflicto:
            salida.setdefault(c["codigo"], "choca con un cambio hecho en el mapa (mira 'Corrección vigente' "
                                           "y vuelve a escribir la tuya)")
        for c in self.cambios_ubic:
            if c["accion"] == "quitar":
                salida.setdefault(c["codigo"], "pide QUITAR la corrección (bórrala si ya no hace falta)")
        return salida

    def datos_sin_georef(self):
        if not self.sin_georef:
            return None
        n = len(self.sin_georef)
        rutas = sorted({r for c in self.sin_georef for r in c["rutas"]})
        aviso = ("Todos deben tener ubicación (mapa y My Maps). Elige un cliente → \"1. Buscar en Google Maps\" → "
                 "en el local, clic derecho → clic en los números (se copian) → pégalos aquí y \"2. Guardar "
                 "ubicación\". No inventes un punto: si no se sabe, pide la ubicación al vendedor (\"Copiar "
                 "lista\"); si el camión no sale a ruta hoy, desmárcalo.")
        if not self.overrides_al_dia:
            aviso += "  (Sin conexión con GitHub: revisado con las correcciones de esta PC.)"
        errores = self.errores_de_fila_ubicacion()
        return {"titulo": f"⚠ {n} cliente{'s' if n > 1 else ''} SIN ubicación (georreferencia) en "
                          f"{', '.join(rutas)}: corrígelos antes de publicar",
                "aviso": aviso,
                "clientes": [{"codigo": c["codigo"], "nombre": c["nombre"], "direccion": c["direccion"],
                              "rutas": ", ".join(c["rutas"]), "vendedores": ", ".join(c["vendedores"]),
                              "motivo": (f"tu fila en el Excel: {errores[c['codigo']]}" if c["codigo"] in errores
                                         else c["motivo"])} for c in self.sin_georef]}

    def _ocupado_o_pregunta(self):
        if self.ocupado:
            return {"ok": False, "error": "Espera a que termine lo que se está haciendo."}
        if self.pregunta:
            return {"ok": False, "error": "Primero contesta la pregunta que está abierta."}
        return None

    def marcar_fecha(self, cuerpo):
        with self.lock:
            no = self._ocupado_o_pregunta()
            if no:
                return no
            d = next((d for d in self.fechas if d.isoformat() == cuerpo.get("fecha")), None)
            if d is None:
                return {"ok": False, "error": "Esa fecha ya no está en la lista."}
            if cuerpo.get("marcada"):
                self.marcadas.add(d)
            else:
                self.marcadas.discard(d)
            self.refrescar_camiones()
        return {"ok": True}

    def marcar_camion(self, cuerpo):
        with self.lock:
            no = self._ocupado_o_pregunta()
            if no:
                return no
            cid = cuerpo.get("id")
            camion = next((c for c in self.camiones if c["id"] == cid), None)
            if camion is None:
                return {"ok": False, "error": "Ese camión ya no está en la lista."}
            # La pagina manda el valor que quiere (no "alternar"): un doble clic
            # no lo marca y desmarca (la ventana bloqueaba el doble clic).
            if cuerpo.get("campo") == "separar":
                if len(camion["liquidaciones"]) <= 1:
                    return {"ok": False, "error": f"{cid} ya no tiene 2 cargues (la lista cambió). Revísala."}
                self.separados[cid] = (bool(cuerpo["valor"]) if "valor" in cuerpo
                                       else not self.separados.get(cid, False))
            else:
                self.incluidos[cid] = bool(cuerpo["valor"]) if "valor" in cuerpo else not self.incluidos[cid]
            self.pintar()
        return {"ok": True}

    def marcar_todos(self, cuerpo):
        with self.lock:
            no = self._ocupado_o_pregunta()
            if no:
                return no
            for camion in self.camiones:
                self.incluidos[camion["id"]] = bool(cuerpo.get("valor"))
            self.pintar()
        return {"ok": True}

    # -- vigilancia del Excel de control -----------------------------------------
    def _vigilar_control(self):
        import control_rutas
        while True:
            time.sleep(gr.SEGUNDOS_VIGILANCIA)
            try:
                with self.lock:
                    mirar = (self.fechas and not self.ocupado and not self.leyendo_control and not self.pausada()
                             and self.editables.get("mtime") != control_rutas.mtime_control())
                    if mirar:
                        self.leyendo_control = True
                if not mirar:
                    continue
                try:
                    editables = control_rutas.leer_editables()
                except Exception:
                    with self.lock:
                        self.leyendo_control = False
                        self.fallos_vigilancia += 1
                        reintentar = self.fallos_vigilancia >= 3 and not self.ocupado
                        if reintentar:
                            self.fallos_vigilancia = 0
                    if reintentar:
                        self._control_vigilado(self.leer_control_seguro())
                    continue
                self._control_vigilado(editables)
            except Exception:
                gr.registrar_error(traceback.format_exc())

    def _control_vigilado(self, editables):
        with self.lock:
            self.leyendo_control = False
            self.fallos_vigilancia = 0
            if self.ocupado or self.pausada():
                return
            antes = len(self.sin_georef)
            self.control_leido(editables)
            if antes != len(self.sin_georef):
                self.log(f"Excel de control releído: {len(self.sin_georef)} clientes sin georreferencia.")

    # -- ubicaciones: lista roja --------------------------------------------------
    def lista_vendedores(self, cuerpo=None):
        por_vendedor = {}
        for c in self.sin_georef:
            for v in c["vendedores"] or ["(sin vendedor)"]:
                por_vendedor.setdefault(v, []).append(c)
        lineas = ["Clientes SIN ubicación en el mapa. Por favor envíen la ubicación (pin) del local:"]
        for v in sorted(por_vendedor):
            lineas += ["", f"Vendedor {v}:"]
            lineas += [f"• {c['codigo']} {c['nombre']} — {c['direccion']} — camión {', '.join(c['rutas'])}"
                       for c in por_vendedor[v]]
        return {"ok": True, "texto": "\n".join(lineas), "cantidad": len(self.sin_georef)}

    def ubicacion(self, cuerpo):
        """'2. Pegar ubicacion copiada' de la ventana: la pagina manda lo que
        el usuario pego (coordenada o enlace de Google Maps)."""
        import control_rutas
        # Esto usa git (puede tardar): se mira antes de tomar el lock, para que
        # la revision y la pregunta de abajo queden juntas, sin hueco.
        sin_subir = gr.ruta_sin_subir()
        copiado = str(cuerpo.get("texto") or "").strip()
        with self.lock:
            no = self._ocupado_o_pregunta()
            if no:
                return no
            c = next((x for x in self.sin_georef if x["codigo"] == cuerpo.get("codigo")), None)
            if not c:
                return {"ok": False, "error": "Ese cliente ya no está en la lista roja."}
            if any(x["codigo"] == c["codigo"] for x in self.cambios_ubic):
                return {"ok": False, "error": f"El cliente {c['codigo']} ya tiene una fila pendiente en la hoja "
                                              "UBICACIONES del Excel de control (por ejemplo, QUITAR u otra "
                                              "coordenada). Corrígela o bórrala allí, para que una no pise a la otra."}
            if sin_subir:
                return {"ok": False, "error": "Hay una ruta de una publicación anterior que no llegó a GitHub (se cortó "
                                              "el internet), así que ahora no se puede guardar una ubicación sola.\n\n"
                                              "Pega esta coordenada en el Excel de control (\"Varios: Excel de "
                                              "control\", columna 'Nueva coordenada' del cliente), guarda y ciérralo: "
                                              "con eso se habilita \"Generar y publicar\", que sube la ruta y las "
                                              "ubicaciones juntas."}
            if not control_rutas.es_enlace_corto(copiado):
                return self._preguntar_ubicacion(c, copiado)
            # El enlace que comparte el celular: se abre en segundo plano
            # (necesita internet) y despues sigue igual.
            self.ocupar(True, "Abriendo el enlace de Google Maps...")

        def resolver():
            try:
                largo = control_rutas.resolver_enlace_corto(copiado)
            except ValueError as e:
                self.ocupar(False)
                self.avisar("aviso", "Pegar ubicación", f"El enlace no sirve como ubicación: {e}")
                return
            except Exception:
                gr.registrar_error(traceback.format_exc())
                self.ocupar(False)
                self.avisar("aviso", "Pegar ubicación", "No se pudo abrir el enlace (¿hay internet?).")
                return
            with self.lock:
                self.ocupar(False)
                self._preguntar_ubicacion(c, largo)
        threading.Thread(target=resolver, daemon=True).start()
        return {"ok": True}

    def _preguntar_ubicacion(self, c, copiado):
        import control_rutas
        try:
            lat, lng = control_rutas.parsear_coordenada(copiado, abrir_enlaces=False)
        except ValueError as e:
            texto = (f"Lo pegado no sirve como ubicación: {e}\n\nEn Google Maps: clic derecho sobre el local → clic "
                     "en los números de arriba (se copian solos, ej. -1.671234, -78.654321) → pégalos aquí.")
            self.avisar("aviso", "Pegar ubicación", texto)
            return {"ok": False, "error": texto, "ya_avisado": True}
        cambio = {"fila": f"ventana:{c['codigo']}", "codigo": c["codigo"], "accion": "poner",
                  "lat": lat, "lng": lng, "antes": control_rutas._como_par(self.overrides_base.get(c["codigo"]))}
        return self.preguntar(
            "Guardar ubicación",
            f"Cliente {c['codigo']} {c['nombre']}\n({', '.join(c['rutas'])})\n\nUbicación: {lat:.6f}, {lng:.6f}\n\n"
            "Se guarda ahora en el mapa (para siempre, igual que arrastrar el punto en el mapa). ¿Continuar?",
            "guardar_ubicacion", {"cambio": cambio, "cliente": c})

    def _guardar_ubicacion(self, cambio, cliente):
        estados = publicar(self.log, con_ruta=False, ubicaciones=[cambio])
        self.overrides_llegaron(gr.overrides_locales(), True)
        estado = estados.get(cambio["fila"], "")
        if estado.startswith("Conflicto"):
            raise gr.ErrorRuta(f"No se guardó: la ubicación de {cliente['codigo']} se cambió en el mapa mientras "
                               "tanto. Revisa la lista (ya se actualizó) y vuelve a intentar si hace falta.")
        mensaje = f"Ubicación guardada: {cliente['codigo']} {cliente['nombre']}."
        self.ocupar(False, mensaje)
        self.log(mensaje)

    # -- generar / publicar --------------------------------------------------------
    def control_cambio(self) -> bool:
        import control_rutas
        if self.editables.get("mtime") == control_rutas.mtime_control():
            return False
        editables = self.leer_control_seguro(abrir_enlaces=False)
        self.control_leido(editables)
        if control_rutas.enlaces_sin_abrir(editables):
            threading.Thread(target=lambda: self.control_leido(self.leer_control_seguro()), daemon=True).start()
        if editables.get("ilegible"):
            self.avisar("aviso", "Excel de control", f"El Excel de control cambió, pero: {editables['ilegible']}\n\n"
                        "Se sigue SIN sus cambios de ruta ni ubicaciones. Revisa la lista y vuelve a darle al botón.")
        else:
            self.avisar("info", "Excel de control", "El Excel de control cambió desde que se leyó, así que se volvió "
                        "a leer con tus cambios.\n\nRevisa la lista de camiones y vuelve a darle al botón.")
        return True

    def robot_trabajando(self):
        if self.otra_ocupada():
            return {"ok": False, "error": "El robot de My Maps está trabajando con los archivos MAPA: espera a que "
                                          "termine (míralo en My Maps) y vuelve a intentar."}
        return None

    def generar(self, cuerpo):
        publicar_mapa = bool(cuerpo.get("publicar"))
        # Todo con el lock tomado (revisar, releer si cambio algo, preguntar):
        # un doble clic o dos pestanas no pueden colarse en medio.
        with self.lock:
            no = self._ocupado_o_pregunta() or self.robot_trabajando()
            if no:
                return no
            if not self.archivo_base or not self.camiones:
                return {"ok": False, "error": "Todavía no hay camiones leídos del archivo base."}
            if publicar_mapa and not self.overrides_listos:
                return {"ok": False, "error": "Todavía se están revisando las ubicaciones corregidas en GitHub. "
                                              "Espera unos segundos."}
            if self.base_cambio() or self.control_cambio():
                return {"ok": True}
            seleccion = gr.armar_seleccion(self.camiones, self.incluidos, self.separados)
            if not seleccion:
                return {"ok": False, "error": "Marca al menos un camión."}
            fechas = set(self.marcadas)
            if not publicar_mapa:
                self._arrancar_generar(False)
                return {"ok": True}
            if self.sin_georef:
                return {"ok": False, "error": f"{len(self.sin_georef)} clientes no tienen georreferencia y todos "
                                              "deben tenerla antes de publicar:\n\n"
                                              f"{gr.texto_sin_georef(self.sin_georef)}\n\nEn la lista en rojo: elige "
                                              "cada cliente → \"1. Buscar en Google Maps\" → pega la ubicación y "
                                              "\"2. Guardar ubicación\" (o varios a la vez con \"Varios: Excel de "
                                              "control\")."}
            clientes = sum(len(c["filas"]) for c in seleccion)
            pregunta = (f"Se va a publicar en el mapa la ruta de:\n{gr.formato_fechas(fechas)}\n\n"
                        f"{len(seleccion)} rutas, {clientes} clientes.")
            partidos = [c["id"] for c in self.camiones if self.incluidos[c["id"]]
                        and self.separados.get(c["id"]) and len(c["liquidaciones"]) > 1]
            if partidos:
                pregunta += f"\nCargues separados: {', '.join(partidos)}."
            fuera = [c["id"] for c in self.camiones if not self.incluidos[c["id"]]]
            if fuera:
                pregunta += f"\nNo van: {', '.join(fuera)}."
            aplicados = [t for nivel, t in self.estados_cambios if nivel == "ok"]
            problemas = [t for nivel, t in self.estados_cambios if nivel == "error"]
            if aplicados:
                pregunta += f"\nCambios de ruta del Excel de control: {len(aplicados)}."
            if problemas:
                pregunta += (f"\nOJO: {len(problemas)} cambios de ruta NO se pudieron aplicar: "
                             + "; ".join(problemas[:3]) + ("..." if len(problemas) > 3 else "."))
            reciben_fuera = [c["id"] for c in self.camiones if c.get("movidos") and not self.incluidos[c["id"]]]
            if reciben_fuera:
                pregunta += (f"\nOJO: {', '.join(reciben_fuera)} recibe clientes de otra ruta pero NO está marcado: "
                             "esos clientes no saldrán en el mapa.")
            if self.cambios_ubic:
                pregunta += f"\nUbicaciones del Excel de control por aplicar: {len(self.cambios_ubic)}."
            problemas_ubic = list(self.estados_ubic.values())
            if problemas_ubic:
                pregunta += (f"\nOJO: {len(problemas_ubic)} filas de UBICACIONES tienen un problema y no se "
                             "aplicarán: " + "; ".join(problemas_ubic[:2]) + ("..." if len(problemas_ubic) > 2 else "."))
            if self.editables.get("ilegible"):
                pregunta += "\nOJO: no se pudo leer el Excel de control; se publica SIN sus cambios de ruta ni ubicaciones."
            publicada = gr.fecha_publicada()
            if publicada and max(fechas).isoformat() < publicada:
                pregunta += (f"\n\nOJO: en el mapa ya está publicada una fecha más nueva ({publicada}). "
                             "Esta la reemplazaría.")
            return self.preguntar("Publicar ruta", pregunta + "\n\n¿Continuar?", "publicar")

    def _arrancar_generar(self, publicar_mapa):
        fechas = set(self.marcadas)
        seleccion = gr.armar_seleccion(self.camiones, self.incluidos, self.separados)
        if publicar_mapa:
            excluidos = set(self.config.get("excluidos", []))
            for camion in self.camiones:
                if self.incluidos[camion["id"]]:
                    excluidos.discard(camion["id"])
                else:
                    excluidos.add(camion["id"])
            self.config["excluidos"] = sorted(excluidos)
            self.config["conocidos"] = sorted(set(self.config.get("conocidos", [])) | {c["id"] for c in self.camiones})
            separados = self.config.get("separados", {})
            separados[max(fechas).isoformat()] = sorted(cid for cid, si in self.separados.items() if si)
            self.config["separados"] = {f: separados[f] for f in sorted(separados)[-15:]}
        self.config["archivo_base"] = str(self.archivo_base)
        if not MODO_PRUEBA:
            gr.guardar_config(self.config)
        carpeta = carpeta_de_salida(self.archivo_base, fechas)
        self.carpeta_ultima = carpeta
        self.carpeta_ultima_dia = date.today()
        self.mymaps_sugerido = None
        self.ocupar(True, "Publicando..." if publicar_mapa else "Generando Excel...")
        self.en_segundo_plano(self._generar, seleccion, fechas, carpeta, publicar_mapa, list(self.cambios_ubic),
                              self.editables, {i: t for i, (_, t) in enumerate(self.estados_cambios) if t})

    def _generar(self, seleccion, fechas, carpeta, publicar_mapa, cambios_ubic=(), editables=None,
                 estados_cambios=None):
        carpeta.mkdir(parents=True, exist_ok=True)
        nuevos = [carpeta / f"MAPA {camion['id']}.xlsx" for camion in seleccion]
        viejos = [p for p in gr.mapas_viejos(carpeta) if p not in nuevos] if publicar_mapa else []
        gr.revisar_cerrados(nuevos + viejos)
        if publicar_mapa:
            self.log("Revisando la georreferencia con las correcciones vigentes...")
            base, al_dia = gr.overrides_vigentes()
            self.overrides_llegaron(base, al_dia)
        else:
            base = self.overrides_base
        overrides = gr.con_pendientes(base, gr.pendientes_aplicables(base, cambios_ubic)[0])
        datos, avisos = gr.armar_json(seleccion, fechas, overrides)
        if publicar_mapa and datos.get("sin_ubicacion"):
            faltan = gr.clientes_sin_georef(seleccion, overrides)
            raise gr.ErrorRuta(f"{len(faltan)} clientes no tienen georreferencia y todos deben tenerla:\n\n"
                               f"{gr.texto_sin_georef(faltan)}\n\nCorrígelos en la lista en rojo y vuelve a "
                               f"publicar.\n{gr.NO_PUBLICADO}")
        if publicar_mapa and not datos["camiones"]:
            raise gr.ErrorRuta("Ningún camión tiene clientes con ubicación válida.\n" + gr.NO_PUBLICADO)
        for viejo in viejos:
            viejo.unlink()
        for camion in seleccion:
            gr.escribir_excel(camion, carpeta)
        self.log(f"{len(seleccion)} archivos MAPA guardados en {carpeta}")
        salida_json = SALIDA_JSON()
        salida_json.parent.mkdir(exist_ok=True)
        salida_json.write_text(json.dumps(datos, ensure_ascii=False), encoding="utf-8")
        aviso = gr.texto_avisos(avisos)
        if aviso:
            self.log(aviso)
        if publicar_mapa:
            estados_ubic = publicar(self.log, con_ruta=True, ubicaciones=cambios_ubic)
            self.overrides_llegaron(gr.overrides_locales(), True)
            gr.anotar_publicado(carpeta, nuevos, fechas)
            fin = "Listo. Los choferes ya ven la ruta actualizada en el mapa."
            fin += self._actualizar_control(editables, estados_cambios, estados_ubic)
        else:
            fin = f"Listo. Se generaron {len(seleccion)} archivos MAPA en:\n{carpeta}"
        self.fin(fin + ("\n\n" + aviso if aviso else ""), carpeta if publicar_mapa else None)

    def fin(self, texto, carpeta_mymaps=None, ofrecer_actualizar=False):
        self.ocupar(False, texto.splitlines()[0])
        self.log(texto)
        extra = {}
        if carpeta_mymaps:
            # Recien publicada la ruta: lo siguiente es subirla a My Maps.
            self.mymaps_sugerido = {"carpeta": str(carpeta_mymaps)}
            extra["mymaps"] = {"carpeta": str(carpeta_mymaps)}
        elif ofrecer_actualizar:
            self.mymaps_sugerido = {"ofrecer": True}
            extra["mymaps"] = {"ofrecer": True}
        self.avisar("info", "Listo", texto, **extra)

    def _actualizar_control(self, editables=None, estados_cambios=None, estados_ubic=None, sin_georef=()):
        import control_rutas
        try:
            if not MODO_PRUEBA:
                control_rutas.construir_control(self.log, editables, estados_cambios, estados_ubic,
                                                sin_georef=sin_georef)
        except gr.ErrorRuta as e:
            self.log(f"Aviso: {e}")
            return f"\n\nEl Excel de control no se pudo actualizar: {e}"
        except Exception:
            gr.registrar_error(traceback.format_exc())
            self.log("Aviso: no se pudo actualizar el Excel de control (detalle en data/generador_error.log).")
            return "\n\nOJO: el Excel de control no se pudo actualizar."
        finally:
            self._releer_control_sin_fallar()
        problemas = [e for e in (estados_ubic or {}).values() if not e.startswith(control_rutas.RESUELTOS)]
        if editables:
            problemas += list(control_rutas.cambios_de_ubicacion(editables)[1].values())
        if problemas:
            return (f"\n\nOJO: {len(problemas)} filas de UBICACIONES no se aplicaron "
                    "(mira la columna Estado en el Excel de control).")
        return ""

    def _releer_control_sin_fallar(self):
        """Releer el Excel de control despues de publicar/armarlo. Si falla,
        no debe tapar lo que ya se hizo (una publicacion exitosa no puede
        terminar mostrandose como error)."""
        try:
            self.control_leido(self.leer_control_seguro())
        except Exception:
            gr.registrar_error(traceback.format_exc())
            self.log("Aviso: no se pudo volver a leer el Excel de control (detalle en data/generador_error.log).")

    # -- Excel de control / aplicar ubicaciones / abrir cosas en la PC ------------------
    def abrir_control(self, cuerpo=None):
        with self.lock:
            no = self._ocupado_o_pregunta()
            if no:
                return no
            self.ocupar(True, "Armando el Excel de control...")
            sin_georef = list(self.sin_georef)
        self.en_segundo_plano(self._abrir_control, sin_georef)
        return {"ok": True}

    def _abrir_control(self, sin_georef=()):
        import control_rutas
        aviso = None
        try:
            if MODO_PRUEBA:
                raise gr.ErrorRuta("(modo prueba: no se arma el Excel de control)")
            ruta = control_rutas.construir_control(self.log, sin_georef=sin_georef)
            mensaje = "Excel de control listo."
            if sin_georef:
                mensaje += f" Los {len(sin_georef)} clientes sin georreferencia están arriba en UBICACIONES."
        except gr.ErrorRuta as e:
            if not control_rutas.ARCHIVO_CONTROL.exists():
                raise
            ruta = control_rutas.ARCHIVO_CONTROL
            mensaje = f"Se abre sin actualizar: {e}"
            self.log(mensaje)
            if sin_georef:
                abierto = str(e).startswith("Cierra")
                aviso = ((f"El Excel de control está abierto, así que no se pudieron poner arriba los "
                          f"{len(sin_georef)} clientes sin georreferencia.\n\nCiérralo y vuelve a darle a "
                          "\"Varios: Excel de control\".") if abierto else
                         (f"No se pudieron poner arriba los {len(sin_georef)} clientes sin georreferencia: {e}")) \
                    + ("\n\nTambién puedes corregirlos uno por uno en la lista en rojo.")
        self._releer_control_sin_fallar()
        self.ocupar(False, mensaje)
        abrir_en_pc(ruta, self)
        if aviso:
            self.avisar("info", "Excel de control", aviso)

    def aplicar_ubicaciones(self, cuerpo=None):
        with self.lock:
            no = self._ocupado_o_pregunta()
            if no:
                return no
            self.ocupar(True, "Leyendo el Excel de control...")
        self.en_segundo_plano(self._preparar_ubicaciones)
        return {"ok": True}

    def _preparar_ubicaciones(self):
        import control_rutas
        editables = control_rutas.leer_editables()
        cambios, estados = control_rutas.cambios_de_ubicacion(editables)
        self.ocupar(False)
        self.control_leido(editables)
        errores = list(estados.values())
        if not cambios:
            texto = "No hay ubicaciones por aplicar en la hoja UBICACIONES del Excel de control."
            if errores:
                texto += (f"\n\n{len(errores)} filas tienen un problema: " + "; ".join(errores[:3])
                          + "\n(Abre el Excel de control para verlas.)")
            self.avisar("info", "Aplicar ubicaciones", texto)
            return
        poner = sum(1 for c in cambios if c["accion"] == "poner")
        quitar = len(cambios) - poner
        pregunta = (f"Se van a aplicar en el mapa {poner} ubicaciones corregidas" + (f" y quitar {quitar}" if quitar else "")
                    + ".\nLos choferes las ven al volver a abrir el mapa. Un cliente que NO salía en el mapa (sin "
                      "ubicación) aparece recién al volver a \"Generar y publicar\".")
        if errores:
            pregunta += f"\n\nOJO: {len(errores)} filas tienen un problema y no se aplicarán."
        self.preguntar("Aplicar ubicaciones", pregunta + "\n\n¿Continuar?", "aplicar_ubicaciones",
                       {"editables": editables, "cambios": cambios})

    def _aplicar_ubicaciones(self, editables, cambios, sin_base=()):
        estados = publicar(self.log, con_ruta=False, ubicaciones=cambios)
        self.overrides_llegaron(gr.overrides_locales(), True)
        aplicadas = sum(1 for e in estados.values() if e.startswith(("Aplicada", "Quitada")))
        fin = f"Listo. Se aplicaron {aplicadas} de {len(cambios)} cambios de ubicación."
        sin_geo = {c["codigo"] for c in sin_base}
        if any(c["codigo"] in sin_geo and c["accion"] == "poner"
               and str(estados.get(c["fila"], "")).startswith("Aplicada") for c in cambios):
            fin += ("\n\nOJO: corregiste clientes que NO salían en el mapa (sin ubicación). Para que salgan, vuelve "
                    "a darle a \"Generar y publicar en el mapa\" (y después a My Maps).")
        fin += self._actualizar_control(editables, None, estados, sin_georef=list(sin_base))
        self.fin(fin, ofrecer_actualizar=aplicadas > 0)

    def abrir_carpeta(self, cuerpo=None):
        if self.carpeta_ultima and self.carpeta_ultima.is_dir():
            abrir_en_pc(self.carpeta_ultima, self)
        elif self.archivo_base and gr.es_base_fija(self.archivo_base):
            gr.CARPETA_MAPAS_FIJA.mkdir(parents=True, exist_ok=True)
            abrir_en_pc(gr.CARPETA_MAPAS_FIJA, self)
        elif self.archivo_base:
            abrir_en_pc(self.archivo_base.parent, self)
        return {"ok": True}

    def abrir_base(self, cuerpo=None):
        if self.archivo_base and self.archivo_base.is_file():
            abrir_en_pc(self.archivo_base, self)
        return {"ok": True}

    # -- preguntas contestadas ------------------------------------------------------
    def al_responder(self, p, si):
        if not si:
            return
        if self.ocupado:
            self.avisar("aviso", p["titulo"], "Se estaba haciendo otra cosa: vuelve a intentarlo.")
            return
        if p["accion"] == "publicar":
            with self.lock:
                if p["version"] != self.version:
                    self.avisar("aviso", "Publicar ruta", "La lista cambió mientras se preguntaba. Revísala y vuelve "
                                "a darle a \"Generar y publicar en el mapa\".")
                    return
                no = self.robot_trabajando()
                if no:
                    self.avisar("aviso", "Publicar ruta", no["error"])
                    return
                if self.base_cambio():
                    return
                self._arrancar_generar(True)
        elif p["accion"] == "guardar_ubicacion":
            d = p["datos"]
            self.ocupar(True, f"Guardando la ubicación de {d['cliente']['codigo']}...")
            self.en_segundo_plano(self._guardar_ubicacion, d["cambio"], d["cliente"])
        elif p["accion"] == "aplicar_ubicaciones":
            d = p["datos"]
            seleccion = gr.armar_seleccion(self.camiones, self.incluidos, self.separados)
            sin_base = gr.clientes_sin_georef(seleccion, self.overrides_base)
            self.ocupar(True, "Aplicando ubicaciones...")
            self.en_segundo_plano(self._aplicar_ubicaciones, d["editables"], d["cambios"], sin_base)

    # -- lo que ve la pagina ----------------------------------------------------------
    def estado(self, consulta=None):
        with self.lock:
            hay = bool(self.camiones)
            datos = {
                "ocupado": self.ocupado, "mensaje": self.mensaje,
                "base": {"nombre": self.archivo_base.name, "info": self.base_info} if self.archivo_base else
                        {"nombre": "No encontré el archivo base.", "info": self.base_info} if self.iniciado else None,
                "fechas": [{"iso": d.isoformat(), "texto": f"{gr.formato_fecha(d)}  "
                                                          f"({len(self.camiones_por_fecha.get(d, ()))} camiones)",
                            "marcada": d in self.marcadas} for d in self.fechas],
                "nota_fechas": self.nota_fechas(),
                "camiones": self.tabla, "resumen": self.resumen, "resumen_alerta": bool(self.sin_georef),
                "sin_georef": self.datos_sin_georef(),
                "botones": {"excel": hay, "publicar": hay and not self.sin_georef and self.overrides_listos},
                "mymaps_sugerido": self.mymaps_sugerido,
                "robot_trabajando": bool(self.otra_ocupada()),
                "prueba": MODO_PRUEBA,
            }
            datos.update(self.base_estado())
            return datos


# ---------------------------------------------------------------------------
# My Maps (lo mismo que ventana_mymaps.py)
# ---------------------------------------------------------------------------

import ventana_mymaps as vm  # noqa: E402  (funciones sin ventana: carpetas_del_dia, revisar_carpeta, ...)


class MyMaps(Controlador):
    def __init__(self, generador):
        super().__init__()
        self.generador = generador
        self.cancelar = threading.Event()
        self.trabajando = False
        self.trabajo = None
        self.mensaje = ""
        self.carpetas = []
        self.carpeta_i = -1
        self.fechas_entrega = []
        self.entrega_i = -1
        self.planes = []
        self.aviso_planes = ""
        self.georef = {"carpeta": None, "estado": "revisando", "faltan": [], "al_dia": True}
        self.revision_n = 0
        self.por_actualizar = []
        self.avisos_act = []
        self.ofrecer_actualizar = False
        self.ofrecer_programado = False
        self.carpetas_texto = []
        self._links = []
        self._links_mtime = None

    def otra_ocupada(self):
        return bool(self.motivo_otra_ocupada())

    def motivo_otra_ocupada(self):
        """Lo que en la ventana era ocupada_para_mymaps: en Generar mapas se
        esta trabajando, hay una pregunta abierta o un mensaje sin ver (por
        ejemplo, el error de una publicacion que fallo)."""
        g = self.generador
        if g.ocupado:
            return "En Generar mapas se está trabajando (leyendo o publicando)."
        if g.pregunta:
            return "En Generar mapas hay una pregunta abierta sin contestar."
        if any(a.get("tipo") in ("error", "aviso") for a in g.avisos):
            return "En Generar mapas hay un mensaje sin ver."
        return ""

    def ocupar(self, ocupado, mensaje=""):
        with self.lock:
            self.trabajando = ocupado
            self.mensaje = mensaje
            self.cambio()

    def fallo(self, texto):
        with self.lock:
            self.trabajando = False
            self.trabajo = None
            self.mensaje = "No se completó, revisa el mensaje."
            self.cambio()
        super().fallo(texto)

    def iniciar(self, cuerpo=None):
        """Al abrir la pagina, como al abrir la ventana: vuelve a listar las
        rutas del dia y busca mapas con cambios. Queda elegida la ruta que
        llega en el link (la recien publicada), si no la ultima generada en
        Generar mapas, y si no la mas nueva. Con 'ofrecer' (despues de
        aplicar ubicaciones) no se cambia la que ya estaba elegida."""
        cuerpo = cuerpo or {}
        with self.lock:
            self.ofrecer_actualizar = self.ofrecer_actualizar or bool(cuerpo.get("ofrecer"))
            if self.trabajando:
                return {"ok": True}
            if cuerpo.get("carpeta"):
                preferida = Path(cuerpo["carpeta"])
            elif cuerpo.get("ofrecer") and self.carpeta_actual():
                preferida = self.carpeta_actual()
            else:
                # La generada hoy en Generar mapas (el servicio queda abierto de
                # un dia para otro: una de ayer no cuenta); si no, la mas nueva.
                g = self.generador
                preferida = g.carpeta_ultima if g.carpeta_ultima_dia == date.today() else None
        self.cargar_carpetas(preferida)
        self.buscar_actualizaciones()
        return {"ok": True}

    # -- ruta y fecha ------------------------------------------------------------
    def cargar_carpetas(self, preferida=None):
        with self.lock:
            self.carpetas = vm.carpetas_del_dia()
            if preferida and Path(preferida).is_dir() and Path(preferida) not in self.carpetas \
                    and mymaps.fecha_de_carpeta(Path(preferida)):
                self.carpetas.insert(0, Path(preferida))
            self.carpetas_texto = []
            for p in self.carpetas:
                d = mymaps.fecha_de_carpeta(p)
                self.carpetas_texto.append(f"Liquidación {mymaps.DIAS_LARGOS[d.weekday()].capitalize()} "
                                           f"{d:%d/%m/%Y}  ({len(mymaps.archivos_mapa(p))} archivos)")
            if not self.carpetas:
                self.carpeta_i = -1
                self.aviso_planes = ("Todavía no hay rutas generadas en BASE\\MAPAS. Primero genera la ruta en "
                                     "\"Generar mapas\".")
                self.planes = []
                self.cambio()
                return
            self.carpeta_i = (self.carpetas.index(Path(preferida))
                              if preferida and Path(preferida) in self.carpetas else 0)
        self.carpeta_elegida()

    def carpeta_actual(self):
        return self.carpetas[self.carpeta_i] if 0 <= self.carpeta_i < len(self.carpetas) else None

    def entrega_actual(self):
        return self.fechas_entrega[self.entrega_i] if 0 <= self.entrega_i < len(self.fechas_entrega) else None

    def carpeta_elegida(self):
        with self.lock:
            carpeta = self.carpeta_actual()
            if not carpeta:
                return
            liq = mymaps.fecha_de_carpeta(carpeta)
            self.fechas_entrega = mymaps.opciones_de_entrega(liq)
            self.entrega_i = self.fechas_entrega.index(mymaps.entrega_sugerida(liq))
            self.refrescar_planes()
        self.revisar_georef()

    def elegir_carpeta(self, cuerpo):
        if self.trabajando:
            return {"ok": False, "error": "Espera a que termine el trabajo en curso."}
        i = int(cuerpo.get("i", -1))
        if not 0 <= i < len(self.carpetas):
            return {"ok": False, "error": "Esa ruta ya no está en la lista."}
        self.carpeta_i = i
        self.carpeta_elegida()
        return {"ok": True}

    def elegir_entrega(self, cuerpo):
        if self.trabajando:
            return {"ok": False, "error": "Espera a que termine el trabajo en curso."}
        i = int(cuerpo.get("i", -1))
        if not 0 <= i < len(self.fechas_entrega):
            return {"ok": False, "error": "Esa fecha ya no está en la lista."}
        with self.lock:
            self.entrega_i = i
            self.refrescar_planes()
        return {"ok": True}

    def revisar_georef(self, cuerpo=None):
        with self.lock:
            carpeta = self.carpeta_actual()
            if not carpeta or self.trabajando:
                return {"ok": True}
            self.revision_n += 1
            n = self.revision_n
            self.georef = {"carpeta": carpeta, "estado": "revisando", "faltan": [], "al_dia": True}
            self.cambio()

        def revisar():
            try:
                faltan, al_dia, cambiados = vm.revisar_carpeta(carpeta)
                error = None
            except Exception as e:
                vm.registrar_error(traceback.format_exc())
                faltan, al_dia, cambiados, error = None, False, None, str(e) or type(e).__name__
            with self.lock:
                if n == self.revision_n:
                    self.georef = {"carpeta": carpeta, "faltan": faltan or [], "al_dia": al_dia, "error": error,
                                   "cambiados": cambiados or [],
                                   "estado": "error" if error else ("faltan" if faltan else
                                                                    ("cambiados" if cambiados else "ok"))}
                    self.cambio()
        threading.Thread(target=revisar, daemon=True).start()
        return {"ok": True}

    def texto_georef(self):
        g = self.georef
        if g["estado"] == "revisando":
            texto, color = "Revisando que todos los clientes tengan ubicación...", "gris"
        elif g["estado"] == "error":
            texto, color = (f"No se pudo revisar la ubicación de los clientes: {g.get('error')}. Revisa que los "
                            "archivos MAPA no estén abiertos y dale a \"Revisar de nuevo\"."), "rojo"
        elif g["estado"] == "faltan":
            faltan = g["faltan"]
            texto = (f"✗ {len(faltan)} clientes SIN ubicación (todos deben tenerla para subir):\n"
                     f"{gr.texto_sin_georef(faltan, maximo=6)}\nCorrígelos en \"Generar mapas\" (lista en rojo), "
                     "vuelve a \"Generar y publicar\" y después súbelo aquí.")
            color = "rojo"
        elif g["estado"] == "cambiados":
            texto, color = "✗ " + mymaps.texto_cambiados(g["cambiados"]), "rojo"
        else:
            texto, color = "✓ Todos los clientes tienen ubicación.", "verde"
        if g["estado"] in ("ok", "faltan") and not g.get("al_dia", True):
            texto += "  (Sin conexión con GitHub: revisado con las correcciones de esta PC.)"
        return {"texto": texto, "color": color}

    def refrescar_planes(self):
        carpeta, entrega = self.carpeta_actual(), self.entrega_actual()
        self.planes = mymaps.armar_mapas(mymaps.archivos_mapa(carpeta), entrega) if carpeta and entrega else []
        avisos = []
        if entrega:
            ya = mymaps.activos(mymaps.leer_registro(), entrega.isoformat())
            if ya:
                avisos.append(f"Ya hay {len(ya)} mapas en My Maps para el {mymaps.fecha_larga(entrega)}: si los vuelves "
                              "a subir, los anteriores van a la papelera y sus links dejan de servir.")
            if entrega < date.today():
                avisos.append("OJO: la fecha de entrega ya pasó.")
        self.aviso_planes = "  ".join(avisos)
        self.cambio()

    def puede_subir(self):
        return (not self.trabajando and bool(self.planes) and self.georef["estado"] == "ok"
                and self.georef["carpeta"] == self.carpeta_actual())

    # -- subir / actualizar / limpiar -----------------------------------------------
    def _no_se_puede(self):
        if self.trabajando:
            return {"ok": False, "error": "Espera a que termine el trabajo en curso."}
        if self.pregunta:
            return {"ok": False, "error": "Primero contesta la pregunta que está abierta."}
        motivo = self.motivo_otra_ocupada()
        if motivo:
            return {"ok": False, "error": motivo + " Ve allí y termina eso primero (el robot usaría los MAPA "
                                                   "mientras se reescriben)."}
        return None

    def subir(self, cuerpo=None):
        with self.lock:
            no = self._no_se_puede()
            if no:
                return no
            if not self.puede_subir():
                return {"ok": False, "error": "Todavía no se puede subir: revisa el aviso de ubicaciones."}
            carpeta, entrega = self.carpeta_actual(), self.entrega_actual()
            registro = mymaps.leer_registro()
            ya = mymaps.activos(registro, entrega.isoformat())
            capas = sum(len(p["archivos"]) for p in self.planes)
            pregunta = (f"Se van a crear {len(self.planes)} mapas en My Maps para el {mymaps.fecha_larga(entrega)} "
                        f"({capas} capas):\n\n"
                        + "\n".join(f"• {p['titulo']}: {'-'.join(p['transportistas'])}" for p in self.planes)
                        + f"\n\nTarda unos {vm.minutos(self.planes)} minutos. Se abre un Chrome aparte (el del "
                          "robot): no lo uses mientras trabaja.")
            if ya:
                pregunta += (f"\n\nOJO: ya hay {len(ya)} mapas para esa fecha. Se mandan a la papelera y sus links "
                             "dejan de servir (habrá que mandar los nuevos).")
            vencidos = [m for m in mymaps.por_mandar_a_papelera(registro, date.today()) if m.get("estado") == "activo"]
            if vencidos:
                pregunta += f"\n\nTambién van a la papelera {len(vencidos)} mapas de hace más de {mymaps.DIAS_HISTORIAL} días."
            return self.preguntar("Subir a My Maps", pregunta + "\n\n¿Continuar?", "subir",
                                  {"carpeta": carpeta, "entrega": entrega, "reemplazar": bool(ya)})

    def actualizar(self, cuerpo=None):
        with self.lock:
            no = self._no_se_puede()
            if no:
                return no
            if not self.por_actualizar:
                return {"ok": False, "error": "No hay cambios para pasar a My Maps."}
            return self._preguntar_actualizar()

    def _preguntar_actualizar(self):
        capas = sum(len(c) for _, c in self.por_actualizar)
        pregunta = ("Se van a pasar a My Maps los cambios (ubicaciones corregidas o ruta vuelta a publicar). "
                    "Los links NO cambian:\n\n" + "\n".join(f"• {t}: {', '.join(c)}" for t, c in self.por_actualizar)
                    + f"\n\nTarda unos {max(1, round(capas * 1.5))} minutos. Se abre el Chrome del robot: no lo uses "
                      "mientras trabaja.")
        return self.preguntar("Actualizar My Maps", pregunta + "\n\n¿Continuar?", "actualizar")

    def limpiar(self, cuerpo=None):
        with self.lock:
            no = self._no_se_puede()
            if no:
                return no
            return self.preguntar(
                "Limpiar mapas viejos",
                "Se revisan tus mapas de My Maps (también los que hiciste a mano, por su fecha de entrega) y se "
                f"mandan a la papelera de Drive los de entrega de hace más de {mymaps.DIAS_HISTORIAL} días.\n"
                "Se abre el Chrome del robot unos minutos.\n\n¿Continuar?", "limpiar")

    def cancelar_trabajo(self, cuerpo=None):
        """La pagina ya pregunto si se cancela (con el texto de 'textos')."""
        with self.lock:
            if not self.trabajando:
                return {"ok": True, "ya_termino": True}
            self.cancelar.set()
            self.mensaje = "Cancelando (termina el paso que está haciendo)..."
            self.cambio()
        return {"ok": True}

    def al_responder(self, p, si):
        if not si:
            return
        with self.lock:
            no = self._no_se_puede()
            if no:
                self.avisar("aviso", p["titulo"], no["error"])
                return
            self.cancelar.clear()
            self.trabajo = p["accion"]
            if p["accion"] == "subir":
                self.log_lineas, self.log_n = [], self.log_n
                self.ocupar(True, "Subiendo a My Maps...")
                d = p["datos"]
                hilo = lambda: self._robot(lambda log: self._subir(log, d["carpeta"], d["entrega"], d["reemplazar"]))
            elif p["accion"] == "actualizar":
                self.ocupar(True, "Actualizando My Maps...")
                hilo = lambda: self._robot(lambda log: self._fin_actualizar(
                    robot("actualizar_ubicaciones", log, self.cancelar)))
            else:
                self.ocupar(True, "Limpiando mapas viejos...")
                hilo = lambda: self._robot(lambda log: self._fin_limpiar(robot("limpiar", log, self.cancelar)))
        threading.Thread(target=hilo, daemon=True).start()

    def _subir(self, log, carpeta, entrega, reemplazar):
        creados = robot("subir", carpeta, entrega, log, self.cancelar, reemplazar=reemplazar)
        self._terminar()
        with self.lock:
            self.refrescar_planes()
            self.mensaje = f"Listo: {len(creados)} mapas en My Maps."
        texto = mymaps.texto_para_compartir(creados)
        self.avisar("info", "My Maps", f"Listo. Se crearon {len(creados)} mapas en My Maps para el "
                    f"{mymaps.fecha_larga(entrega)}.\n\nCopia el mensaje con los links y pégalo en WhatsApp.",
                    copiar=texto, dia=entrega.isoformat())

    def _fin_actualizar(self, hechos):
        self._terminar()
        if not hechos:
            self.mensaje = "No había nada para actualizar."
            self.avisar("info", "My Maps", "No había cambios para pasar a My Maps.")
            return
        self.mensaje = f"Listo: {len(hechos)} mapas actualizados (mismos links)."
        self.avisar("info", "My Maps", "Listo. Se actualizaron en My Maps:\n\n"
                    + "\n".join(f"• {t}: {', '.join(c)}" for t, c in hechos)
                    + "\n\nLos links son los mismos: los choferes solo vuelven a abrir el mapa.")

    def _fin_limpiar(self, enviados):
        self._terminar()
        self.mensaje = f"Limpieza lista: {len(enviados)} mapas a la papelera."
        self.avisar("info", "My Maps", (f"Se mandaron a la papelera {len(enviados)} mapas:\n\n"
                                        + "\n".join(f"• {m.get('titulo')}" for m in enviados[:20])
                                        + ("\n..." if len(enviados) > 20 else "")) if enviados
                    else f"No hay mapas con entrega de hace más de {mymaps.DIAS_HISTORIAL} días.")

    def _terminar(self):
        with self.lock:
            self.trabajando = False
            self.trabajo = None
            self.cambio()
        self.buscar_actualizaciones()

    def _robot(self, funcion):
        """Corre un trabajo del robot con los mismos mensajes de error que la ventana."""
        try:
            funcion(self.log)
        except mymaps.Cancelado:
            self._error(vm.TEXTOS_TRABAJO.get(self.trabajo, vm.TEXTOS_TRABAJO["subir"])["cancelado"])
        except mymaps.ErrorMyMaps as e:
            self._error(str(e))
        except Exception as e:
            vm.registrar_error(traceback.format_exc())
            if type(e).__name__ == "ErrorRuta":
                self._error(str(e))
            elif type(e).__module__.startswith("playwright"):
                # "Subelo a mano" solo para la subida: tras actualizar o limpiar
                # crearia mapas con links nuevos (lo mismo que la ventana).
                if self.trabajo == "subir":
                    consejo = "Vuelve a intentar; si se repite, súbelo a mano con los archivos de la carpeta del día."
                else:
                    consejo = "Vuelve a intentar."
                self._error("My Maps no respondió como se esperaba (puede que Google haya cambiado algo en la página "
                            f"o que se haya cerrado el Chrome del robot).\n\n{consejo}\n\n"
                            f"Detalle: {str(e).splitlines()[0][:200]}")
            else:
                self._error("Algo falló inesperadamente:\n\n" + traceback.format_exc().strip().splitlines()[-1])

    def _error(self, texto):
        with self.lock:
            self.trabajando = False
            self.trabajo = None
            self.mensaje = "No se completó, revisa el mensaje."
            self.cambio()
            self.refrescar_planes()
        self.log("ERROR: " + texto)
        self.avisar("error", "My Maps", texto)
        self.revisar_georef()
        self.buscar_actualizaciones()

    # -- mapas ya subidos con cambios ----------------------------------------------
    def buscar_actualizaciones(self):
        def buscar():
            try:
                resumen, avisos = vm.revisar_actualizaciones()
            except Exception:
                vm.registrar_error(traceback.format_exc())
                resumen, avisos = [], []
            with self.lock:
                if self.trabajando:
                    return
                self.por_actualizar = resumen or []
                self.avisos_act = avisos or []
                if not self.por_actualizar:
                    self.ofrecer_actualizar = False
                self.cambio()
                programado = self.ofrecer_programado
            if not programado:
                self.ofrecer_si_se_puede()
        threading.Thread(target=buscar, daemon=True).start()

    def ofrecer_si_se_puede(self):
        """Despues de corregir ubicaciones: preguntar enseguida si se pasan a
        My Maps (cuando no haya otra pregunta ni trabajo, aqui ni en Generar).
        Un solo reintento programado a la vez (ofrecer_programado, como la
        ventana)."""
        with self.lock:
            self.ofrecer_programado = False
            if not self.ofrecer_actualizar or self.trabajando or not self.por_actualizar:
                return
            if self.pregunta or self.otra_ocupada():
                self.ofrecer_programado = True
                threading.Timer(1.0, self.ofrecer_si_se_puede).start()
                return
            self.ofrecer_actualizar = False
            self._preguntar_actualizar()

    # -- links ------------------------------------------------------------------------
    def links(self):
        """Links de los ultimos dias. Mientras el robot trabaja NO se lee el
        registro: el robot lo reescribe muchas veces y en Windows, si justo
        se esta leyendo, su os.replace falla y corta el trabajo. Fuera de eso
        se relee solo si el archivo cambio (os.stat no estorba al robot)."""
        if self.trabajando:
            return self._links
        try:
            mtime = mymaps.ARCHIVO_REGISTRO.stat().st_mtime_ns
        except (OSError, AttributeError):
            mtime = None
        # El dia tambien cuenta: "ultimos N dias" cambia aunque no se escriba nada.
        clave = (mtime, date.today())
        if mtime is not None and clave == self._links_mtime:
            return self._links
        self._links = self._leer_links()
        self._links_mtime = clave
        return self._links

    def _leer_links(self):
        registro = mymaps.leer_registro()
        limite = (date.today() - timedelta(days=mymaps.DIAS_HISTORIAL)).isoformat()
        por_dia = {}
        for m in mymaps.activos(registro):
            if m.get("fecha_entrega", "") >= limite and m.get("url"):
                por_dia.setdefault(m["fecha_entrega"], []).append(m)
        salida = []
        for dia in sorted(por_dia, reverse=True):
            vistos, mapas = set(), []
            for m in sorted(por_dia[dia], key=lambda m: m.get("numero", 0)):
                if m["mid"] in vistos:
                    continue
                vistos.add(m["mid"])
                mapas.append({"mid": m["mid"], "titulo": m["titulo"], "url": m["url"],
                              "transportistas": "-".join(m.get("transportistas", []))
                              or ("(hecho a mano)" if m.get("origen") == "manual" else "")})
            salida.append({"dia": dia, "texto": mymaps.fecha_larga(date.fromisoformat(dia)), "mapas": mapas})
        return salida

    def mensaje_dia(self, cuerpo):
        if self.trabajando:
            return {"ok": False, "error": "Espera a que termine el robot (está escribiendo el registro de mapas) y "
                                          "vuelve a copiar el mensaje."}
        dia = cuerpo.get("dia") or ""
        mapas = sorted(mymaps.activos(mymaps.leer_registro(), dia), key=lambda m: m.get("numero", 0))
        if not mapas:
            return {"ok": False, "error": "No hay mapas de esa fecha."}
        return {"ok": True, "texto": mymaps.texto_para_compartir(mapas), "cantidad": len(mapas)}

    def estado(self, consulta=None):
        with self.lock:
            capas = sum(len(p["archivos"]) for p in self.planes)
            textos = vm.TEXTOS_TRABAJO.get(self.trabajo, vm.TEXTOS_TRABAJO["subir"])
            act = ""
            if self.por_actualizar:
                capas_act = sum(len(c) for _, c in self.por_actualizar)
                act = (f"⟳ Cambios que aún no están en My Maps (ubicaciones corregidas o ruta vuelta a publicar): "
                       f"{len(self.por_actualizar)} mapas, {capas_act} capas ("
                       + "; ".join(f"{t.split('-MAPA')[0]}: {', '.join(c)}" for t, c in self.por_actualizar) + ").")
            if self.avisos_act:
                act = (act + "\n" if act else "") + "\n".join(self.avisos_act)
            datos = {
                "trabajando": self.trabajando, "mensaje": self.mensaje,
                "pregunta_cancelar": textos["pregunta"] if self.trabajando else None,
                "carpetas": [{"i": i, "texto": t} for i, t in enumerate(self.carpetas_texto)],
                "carpeta_i": self.carpeta_i,
                "entregas": [{"i": i, "texto": mymaps.fecha_larga(d)} for i, d in enumerate(self.fechas_entrega)],
                "entrega_i": self.entrega_i,
                "planes": [{"titulo": p["titulo"], "capas": len(p["archivos"]),
                            "transportistas": "-".join(p["transportistas"])} for p in self.planes],
                "tiempo": (f"{len(self.planes)} mapas, {capas} capas · tarda unos {vm.minutos(self.planes)} min"
                           if self.planes else ""),
                "aviso_planes": self.aviso_planes,
                "georef": self.texto_georef(),
                "puede_subir": self.puede_subir(),
                "actualizables": {"texto": act, "hay": bool(self.por_actualizar)},
                "links": self.links(),
                "dias_historial": mymaps.DIAS_HISTORIAL,
                "otra_ocupada": self.otra_ocupada(),
                "otra_motivo": self.motivo_otra_ocupada(),
                "prueba": MODO_PRUEBA,
            }
            datos.update(self.base_estado())
            return datos


# ---------------------------------------------------------------------------
# Cosas de esta PC: abrir archivos/carpetas, elegir archivo, publicar y robot
# (en --prueba no se publica ni se usa el robot)
# ---------------------------------------------------------------------------

_SALIDA_PRUEBA = None


def carpeta_prueba():
    global _SALIDA_PRUEBA
    if _SALIDA_PRUEBA is None:
        _SALIDA_PRUEBA = Path(tempfile.mkdtemp(prefix="servicio_web_prueba_"))
    return _SALIDA_PRUEBA


def SALIDA_JSON():
    return carpeta_prueba() / "latest.json" if MODO_PRUEBA else gr.SALIDA_JSON


def carpeta_de_salida(archivo_base, fechas):
    if MODO_PRUEBA:
        ultima = max(fechas)
        return carpeta_prueba() / "MAPAS" / f"{ultima:%Y-%m-%d} {gr.DIAS_LARGOS[ultima.weekday()]}"
    return gr.carpeta_de_salida(archivo_base, fechas)


def publicar(log, con_ruta=True, ubicaciones=None):
    if MODO_PRUEBA:
        log("(modo prueba: NO se publica en GitHub)")
        time.sleep(1)
        return {c["fila"]: "Aplicada (prueba)" for c in (ubicaciones or [])}
    return gr.publicar(log, con_ruta=con_ruta, ubicaciones=ubicaciones)


def robot(nombre, *argumentos, **opciones):
    if MODO_PRUEBA:
        log = next((a for a in argumentos if callable(a)), print)
        log(f"(modo prueba: el robot NO se usa — {nombre})")
        time.sleep(2)
        return []
    return getattr(mymaps, nombre)(*argumentos, **opciones)


def abrir_en_pc(ruta, controlador=None):
    try:
        os.startfile(str(ruta))
    except OSError as e:
        if controlador:
            controlador.avisar("error", "Abrir", f"No pude abrir:\n{ruta}\n\n{e}")


def elegir_archivo_windows(inicial):
    """El 'Cambiar archivo...' de la ventana: el cuadro de Windows para elegir
    el archivo base sale en esta PC (encima de todo)."""
    script = (
        "Add-Type -AssemblyName System.Windows.Forms;"
        "$f = New-Object System.Windows.Forms.Form; $f.TopMost = $true; $f.ShowInTaskbar = $false;"
        "$d = New-Object System.Windows.Forms.OpenFileDialog;"
        "$d.Title = 'Elige el archivo base'; $d.Filter = 'Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Todos (*.*)|*.*';"
        f"$d.InitialDirectory = '{str(inicial).replace(chr(39), chr(39) * 2)}';"
        "if ($d.ShowDialog($f) -eq 'OK') { [Console]::OutputEncoding = [Text.Encoding]::UTF8; $d.FileName }")
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-STA", "-Command", script], capture_output=True,
                           timeout=600, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        ruta = r.stdout.decode("utf-8", "replace").strip()
        return ruta if ruta and Path(ruta).is_file() else None
    except (OSError, subprocess.TimeoutExpired):
        return None


# ---------------------------------------------------------------------------
# HTTP
# ---------------------------------------------------------------------------

GENERADOR = Generador()
MYMAPS = MyMaps(GENERADOR)
# Coordinacion en los dos sentidos: My Maps no arranca el robot si Generar esta
# trabajando, y Generar no reescribe los MAPA mientras el robot los usa.
GENERADOR.otra_ocupada = lambda: MYMAPS.trabajando

RUTAS_POST = {
    "/generador/iniciar": GENERADOR.iniciar,
    "/generador/releer": GENERADOR.releer,
    "/generador/elegir_base": GENERADOR.elegir_base,
    "/generador/abrir_base": GENERADOR.abrir_base,
    "/generador/fecha": GENERADOR.marcar_fecha,
    "/generador/camion": GENERADOR.marcar_camion,
    "/generador/todos": GENERADOR.marcar_todos,
    "/generador/ubicacion": GENERADOR.ubicacion,
    "/generador/lista_vendedores": GENERADOR.lista_vendedores,
    "/generador/generar": GENERADOR.generar,
    "/generador/control": GENERADOR.abrir_control,
    "/generador/aplicar_ubicaciones": GENERADOR.aplicar_ubicaciones,
    "/generador/carpeta": GENERADOR.abrir_carpeta,
    "/generador/responder": GENERADOR.responder,
    "/generador/aviso_visto": GENERADOR.aviso_visto,
    "/mymaps/iniciar": MYMAPS.iniciar,
    "/mymaps/carpeta": MYMAPS.elegir_carpeta,
    "/mymaps/entrega": MYMAPS.elegir_entrega,
    "/mymaps/revisar": MYMAPS.revisar_georef,
    "/mymaps/subir": MYMAPS.subir,
    "/mymaps/actualizar": MYMAPS.actualizar,
    "/mymaps/limpiar": MYMAPS.limpiar,
    "/mymaps/cancelar": MYMAPS.cancelar_trabajo,
    "/mymaps/mensaje_dia": MYMAPS.mensaje_dia,
    "/mymaps/responder": MYMAPS.responder,
    "/mymaps/aviso_visto": MYMAPS.aviso_visto,
}
RUTAS_GET = {
    "/generador/estado": GENERADOR.estado,
    "/mymaps/estado": MYMAPS.estado,
}


class Manejador(BaseHTTPRequestHandler):
    server_version = "ServicioMapas/" + VERSION

    def log_message(self, formato, *args):
        pass                                # pythonw no tiene consola

    def _origen_ok(self):
        return self.headers.get("Origin") in ORIGENES_PERMITIDOS

    def _cabeceras_cors(self):
        origen = self.headers.get("Origin")
        if origen in ORIGENES_PERMITIDOS:
            self.send_header("Access-Control-Allow-Origin", origen)
            self.send_header("Vary", "Origin")
            self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
            self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Opaustro")
            self.send_header("Access-Control-Allow-Private-Network", "true")
            self.send_header("Access-Control-Max-Age", "600")

    def _responder(self, codigo, datos):
        cuerpo = json.dumps(datos, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(codigo)
        self._cabeceras_cors()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(cuerpo)))
        self.end_headers()
        try:
            self.wfile.write(cuerpo)
        except (ConnectionAbortedError, ConnectionResetError, BrokenPipeError):
            pass                            # la pagina se cerro o corto la consulta: no es un error

    def do_OPTIONS(self):
        self.send_response(204 if self._origen_ok() else 403)
        self._cabeceras_cors()
        self.send_header("Content-Length", "0")
        self.end_headers()

    def _permitido(self):
        # Origin + un encabezado propio: el navegador obliga a preguntar antes
        # (preflight), asi otra pagina web no puede mandar acciones a escondidas.
        if not self._origen_ok() or self.headers.get("X-Opaustro") != "1":
            self._responder(403, {"ok": False, "error": "No permitido."})
            return False
        return True

    def do_GET(self):
        url = urlparse(self.path)
        if url.path == "/salud":
            if not self._permitido():
                return
            self._responder(200, {"ok": True, "servicio": "mapas", "version": VERSION, "prueba": MODO_PRUEBA})
            return
        funcion = RUTAS_GET.get(url.path)
        if not funcion:
            self._responder(404, {"ok": False, "error": "No existe."})
            return
        if not self._permitido():
            return
        self._ejecutar(funcion, parse_qs(url.query))

    def do_POST(self):
        url = urlparse(self.path)
        funcion = RUTAS_POST.get(url.path)
        if not funcion:
            self._responder(404, {"ok": False, "error": "No existe."})
            return
        if not self._permitido():
            return
        if not (self.headers.get("Content-Type") or "").startswith("application/json"):
            self._responder(415, {"ok": False, "error": "Se espera JSON."})
            return
        try:
            largo = int(self.headers.get("Content-Length") or 0)
            cuerpo = json.loads(self.rfile.read(largo).decode("utf-8") or "{}") if largo else {}
        except (ValueError, UnicodeDecodeError):
            self._responder(400, {"ok": False, "error": "Pedido mal armado."})
            return
        self._ejecutar(funcion, cuerpo)

    def _ejecutar(self, funcion, argumento):
        try:
            self._responder(200, funcion(argumento) or {"ok": True})
        except (ConnectionAbortedError, ConnectionResetError, BrokenPipeError):
            pass
        except gr.ErrorRuta as e:
            self._responder(200, {"ok": False, "error": str(e)})
        except Exception:
            detalle = traceback.format_exc()
            gr.registrar_error("(servicio web)\n" + detalle)
            self._responder(500, {"ok": False, "error": "Algo falló inesperadamente: "
                                                        + detalle.strip().splitlines()[-1]})


class Servidor(ThreadingHTTPServer):
    # En Windows, allow_reuse_address (que HTTPServer trae encendido) deja
    # abrir el MISMO puerto dos veces: quedaban dos servicios con estados
    # distintos (doble clic al .bat, o .bat + arranque con Windows). El puerto
    # tiene que ser exclusivo para que el segundo vea que ya hay uno y se cierre.
    allow_reuse_address = False
    daemon_threads = True

    def server_bind(self):
        if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()


def avisar_si_hay_otro_viejo():
    """El puerto ya esta ocupado. Si es este mismo servicio (misma version),
    no pasa nada: se usa ese. Si es una version anterior (o no responde), se
    avisa: si no, los arreglos nuevos no entrarian en uso y nadie lo notaria
    (con pythonw no hay consola donde decirlo)."""
    import urllib.request
    version = None
    try:
        abridor = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        pedido = urllib.request.Request(f"http://127.0.0.1:{PUERTO}/salud",
                                        headers={"Origin": "https://coordinadorl.github.io", "X-Opaustro": "1"})
        with abridor.open(pedido, timeout=5) as r:
            version = json.loads(r.read().decode("utf-8")).get("version")
    except Exception:
        pass
    if version == VERSION or MODO_PRUEBA:
        return
    texto = ("Ya hay otro programa usando el servicio de mapas"
             + (f" (una versión anterior, la {version})" if version else "")
             + ", así que esta versión no pudo abrirse.\n\nReinicia la PC (el servicio arranca solo al prenderla) o "
               "cierra 'pythonw' en el Administrador de tareas y vuelve a abrir \"SERVICIO MAPAS.bat\".")
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, texto, "Servicio de mapas", 0x30)
    except Exception:
        gr.registrar_error("(servicio web) " + texto)


def main():
    try:
        servidor = Servidor(("127.0.0.1", PUERTO), Manejador)
    except OSError:
        # Ya hay otro servicio abierto (doble clic dos veces, o ya arrancó con
        # Windows): si es este mismo, no pasa nada, se usa ese.
        print(f"El servicio ya estaba funcionando en http://127.0.0.1:{PUERTO}")
        avisar_si_hay_otro_viejo()
        return
    print(f"Servicio de mapas en http://127.0.0.1:{PUERTO}" + ("  (MODO PRUEBA)" if MODO_PRUEBA else ""))
    servidor.serve_forever()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        gr.registrar_error("(servicio web)\n" + traceback.format_exc())
        raise
