"""
Ventana de My Maps: sube la ruta del dia a Google My Maps (con el robot de
mymaps.py) y muestra los links de los ultimos 30 dias para copiarlos.

Se abre desde "Actualizar ruta" (boton My Maps, y sola despues de publicar),
o directo con doble clic en "MY MAPS.bat".
"""
import ctypes
import os
import queue
import threading
import traceback
import webbrowser
from datetime import date, datetime, timedelta
from pathlib import Path

import mymaps

CARPETA = Path(__file__).resolve().parent
CARPETA_MAPAS = CARPETA / "BASE" / "MAPAS"
ARCHIVO_ERRORES = CARPETA / "data" / "generador_error.log"
CARPETAS_A_MOSTRAR = 15
SEGUNDOS_POR_CAPA = 50
SEGUNDOS_POR_MAPA = 40

_abierta = None


# Textos de cancelar segun el trabajo en curso.
TEXTOS_TRABAJO = {
    "subir": {
        "haciendo": "Se está subiendo a My Maps",
        "pregunta": "¿Cancelar la subida a My Maps?\nEl mapa que se está armando va a la papelera.",
        "cancelado": "Se canceló la subida a My Maps.\nLo que quedó a medias se mandó a la papelera.",
    },
    "actualizar": {
        "haciendo": "Se están actualizando ubicaciones en My Maps",
        "pregunta": "¿Cancelar la actualización?\nLos mapas y sus links siguen; el que se está "
                    "actualizando puede quedar a medias.",
        "cancelado": "Se canceló la actualización. Los mapas y sus links siguen, pero el que se estaba "
                     "actualizando puede haber quedado a medias: vuelve a darle a «Actualizar My Maps» para "
                     "completarlo.",
    },
    "limpiar": {
        "haciendo": "Se están limpiando los mapas viejos",
        "pregunta": "¿Cancelar la limpieza de mapas viejos?\nLo que ya fue a la papelera se queda allí.",
        "cancelado": "Se canceló la limpieza. Lo que ya fue a la papelera se queda allí; "
                     "lo demás se limpia la próxima vez.",
    },
}


def registrar_error(detalle: str) -> None:
    try:
        ARCHIVO_ERRORES.parent.mkdir(exist_ok=True)
        with open(ARCHIVO_ERRORES, "a", encoding="utf-8") as f:
            f.write(f"--- {datetime.now():%Y-%m-%d %H:%M:%S} (My Maps)\n{detalle}\n")
    except OSError:
        pass


def minutos(planes) -> int:
    """Lo que tarda el robot, mas o menos (medido en la prueba real)."""
    capas = sum(len(p["archivos"]) for p in planes)
    return max(1, round((capas * SEGUNDOS_POR_CAPA + len(planes) * SEGUNDOS_POR_MAPA) / 60))


def revisar_carpeta(carpeta: Path):
    """Revision antes de subir: MAPA que cambiaron despues de publicar, y
    clientes sin georreferencia (con las correcciones vigentes en GitHub).
    Devuelve (faltan, al_dia, cambiados)."""
    import generar_rutas
    cambiados = mymaps.cambiados_desde_publicar(carpeta)
    overrides, al_dia = generar_rutas.overrides_vigentes()
    return mymaps.revisar_georef(mymaps.archivos_mapa(carpeta), overrides), al_dia, cambiados


def carpetas_del_dia():
    """Carpetas BASE\\MAPAS\\<fecha DIA> con archivos MAPA, la mas nueva primero."""
    if not CARPETA_MAPAS.is_dir():
        return []
    salida = [p for p in CARPETA_MAPAS.iterdir()
              if p.is_dir() and mymaps.fecha_de_carpeta(p) and mymaps.archivos_mapa(p)]
    return sorted(salida, key=lambda p: p.name, reverse=True)[:CARPETAS_A_MOSTRAR]


def revisar_actualizaciones():
    """Mapas ya subidos (de hoy en adelante) con ubicaciones corregidas
    despues de subirlos. Devuelve (resumen [(titulo, [capas])], avisos)."""
    import tempfile
    import shutil
    import generar_rutas
    registro = mymaps.leer_registro()
    if not mymaps.actualizables(registro, date.today()):
        return [], []                   # nada subido de hoy en adelante: ni se consulta GitHub
    overrides, al_dia = generar_rutas.overrides_vigentes()
    if not al_dia:
        # Con lo descargado la ultima vez se podria proponer volver a una
        # ubicacion vieja: no se propone nada.
        return [], ["Sin conexión con GitHub: no se puede revisar si hay ubicaciones corregidas para My Maps."]
    temporal = Path(tempfile.mkdtemp(prefix="mymaps_rev_"))
    try:
        planes, avisos = mymaps.mapas_por_actualizar(registro, overrides, date.today(), temporal)
        return [(p["mapa"]["titulo"], sorted(p["capas"])) for p in planes], avisos
    finally:
        shutil.rmtree(temporal, ignore_errors=True)


def abrir(padre=None, carpeta: Path = None, ofrecer_actualizar=False):
    """Abre la ventana (o la trae al frente si ya esta abierta).
    ofrecer_actualizar: si hay mapas ya subidos con ubicaciones corregidas,
    preguntar enseguida si se actualizan (se usa despues de corregir)."""
    global _abierta
    if _abierta is not None and _abierta.viva():
        _abierta.traer(carpeta, ofrecer_actualizar)
        return _abierta
    import tkinter as tk

    if padre is None:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
        raiz = tk.Tk()
        _abierta = VentanaMyMaps(raiz, carpeta, ofrecer_actualizar)
        raiz.mainloop()
    else:
        _abierta = VentanaMyMaps(tk.Toplevel(padre), carpeta, ofrecer_actualizar)
    return _abierta


class VentanaMyMaps:
    def __init__(self, ventana, carpeta=None, ofrecer_actualizar=False):
        import tkinter as tk
        from tkinter import messagebox, ttk

        self.tk, self.ttk, self.messagebox = tk, ttk, messagebox
        self.v = ventana
        self.cola = queue.Queue()
        self.cancelar = threading.Event()
        self.trabajando = False
        self.cerrar_al_terminar = None      # ventana a cerrar cuando termine el trabajo en curso
        self.cerrar_pendiente = None        # X de "Actualizar ruta" con una pregunta abierta aqui
        self.preguntando = 0                # preguntas (askyesno) abiertas
        self.ofrecer_programado = False
        # "Actualizar ruta" lo cambia por una funcion que dice si alla hay una
        # pregunta abierta o se esta publicando (no se ofrece encima).
        self.otra_ocupada = lambda: False
        self.carpetas = []
        self.fechas_entrega = []
        self.planes = []
        # Revision de georreferencia de la carpeta elegida (en segundo plano):
        # hasta que diga "ok" no se puede subir.
        self.georef = {"carpeta": None, "estado": "revisando", "faltan": [], "al_dia": True}
        self.revision_n = 0
        # Mapas ya subidos con ubicaciones corregidas (para actualizarlos).
        self.por_actualizar = []
        self.ofrecer_actualizar = ofrecer_actualizar
        self.trabajo = None             # "subir", "actualizar" o "limpiar" (para los textos de cancelar)
        self.escala = max(1.0, ventana.winfo_fpixels("1i") / 96)
        self.construir()
        self.cargar_carpetas(carpeta)
        self.pintar_links()
        ventana.protocol("WM_DELETE_WINDOW", self.al_cerrar)
        ventana.after(100, self.revisar_cola)
        self.buscar_actualizaciones()

    def px(self, n):
        return int(n * self.escala)

    def viva(self):
        try:
            return bool(self.v.winfo_exists())
        except Exception:
            return False

    def traer(self, carpeta=None, ofrecer_actualizar=False):
        self.v.deiconify()
        self.v.lift()
        self.v.focus_force()
        # Aunque este trabajando se recuerda que hay que ofrecer actualizar: se
        # busca al terminar el trabajo en curso.
        self.ofrecer_actualizar = self.ofrecer_actualizar or ofrecer_actualizar
        if not self.trabajando:
            if carpeta:
                self.cargar_carpetas(carpeta)
            self.buscar_actualizaciones()

    # -- armado --------------------------------------------------------------
    def construir(self):
        tk, ttk = self.tk, self.ttk
        v = self.v
        v.title("My Maps — rutas del día")
        alto = min(self.px(760), v.winfo_screenheight() - self.px(80))
        v.geometry(f"{self.px(860)}x{alto}+{self.px(80)}+{self.px(20)}")
        v.minsize(self.px(720), self.px(560))
        estilo = ttk.Style()
        estilo.configure("Accion.TButton", font=("Segoe UI", 10, "bold"), padding=(12, 6))
        estilo.configure("Titulo.TLabel", font=("Segoe UI", 10, "bold"))
        estilo.configure("Ayuda.TLabel", foreground="#666666")
        estilo.configure("Treeview", rowheight=self.px(24))

        cont = ttk.Frame(v, padding=12)
        cont.pack(fill="both", expand=True)

        m1 = ttk.LabelFrame(cont, text=" 1. Ruta que se sube ", padding=8)
        m1.pack(fill="x")
        m1.columnconfigure(1, weight=1)
        ttk.Label(m1, text="Ruta del día:").grid(row=0, column=0, sticky="w")
        self.cmb_carpeta = ttk.Combobox(m1, state="readonly", width=46)
        self.cmb_carpeta.grid(row=0, column=1, sticky="w", padx=(6, 0))
        self.cmb_carpeta.bind("<<ComboboxSelected>>", lambda e: self.carpeta_elegida())
        ttk.Label(m1, text="Fecha de entrega:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.cmb_entrega = ttk.Combobox(m1, state="readonly", width=34)
        self.cmb_entrega.grid(row=1, column=1, sticky="w", padx=(6, 0), pady=(6, 0))
        self.cmb_entrega.bind("<<ComboboxSelected>>", lambda e: self.refrescar_planes())
        ttk.Label(m1, text="Va en el título de cada mapa. Normalmente es el siguiente día hábil (viernes y sábado → "
                           "lunes); cámbiala si ese día no se trabaja.",
                  style="Ayuda.TLabel", wraplength=self.px(780)).grid(row=2, column=0, columnspan=2, sticky="w",
                                                                     pady=(4, 0))

        marco = ttk.Frame(m1)
        marco.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        marco.columnconfigure(0, weight=1)
        self.lista_planes = ttk.Treeview(marco, columns=("titulo", "capas", "transportistas"), show="headings",
                                         height=4, selectmode="none")
        for col, titulo, ancho, estira in [("titulo", "Mapa", 330, False), ("capas", "Capas", 56, False),
                                           ("transportistas", "Transportistas", 360, True)]:
            self.lista_planes.heading(col, text=titulo, anchor="w")
            self.lista_planes.column(col, width=self.px(ancho), anchor="w" if col != "capas" else "center",
                                     stretch=estira)
        self.lista_planes.grid(row=0, column=0, sticky="ew")
        self.lbl_aviso = ttk.Label(m1, text="", foreground="#a15c00", wraplength=self.px(780))
        self.lbl_aviso.grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))
        marco_geo = ttk.Frame(m1)
        marco_geo.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        marco_geo.columnconfigure(0, weight=1)
        self.lbl_georef = tk.Label(marco_geo, text="", anchor="w", justify="left", wraplength=self.px(680),
                                   font=("Segoe UI", 9, "bold"))
        self.lbl_georef.grid(row=0, column=0, sticky="w")
        self.btn_revisar = ttk.Button(marco_geo, text="Revisar de nuevo", command=lambda: self.revisar_georef())
        self.btn_revisar.grid(row=0, column=1, sticky="ne", padx=(6, 0))

        barra = ttk.Frame(m1)
        barra.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.btn_subir = ttk.Button(barra, text="Subir a My Maps", style="Accion.TButton", command=self.subir)
        self.btn_subir.pack(side="right")
        self.btn_cancelar = ttk.Button(barra, text="Cancelar", command=self.pedir_cancelar, state="disabled")
        self.btn_cancelar.pack(side="right", padx=(0, 8))
        self.lbl_tiempo = ttk.Label(barra, text="", style="Ayuda.TLabel")
        self.lbl_tiempo.pack(side="left")

        m2 = ttk.LabelFrame(cont, text=f" 2. Links de My Maps (últimos {mymaps.DIAS_HISTORIAL} días) ", padding=8)
        m2.pack(fill="both", expand=True, pady=(10, 0))
        # Mapas ya subidos con cambios despues (ubicaciones corregidas o ruta
        # vuelta a publicar): se actualizan sin cambiar el link. Solo aparece
        # cuando hay alguno.
        self.marco_act = tk.Frame(m2, background="#fff4e0", padx=8, pady=6)
        self.lbl_act = tk.Label(self.marco_act, text="", background="#fff4e0", foreground="#8a4b00", anchor="w",
                                justify="left", wraplength=self.px(600), font=("Segoe UI", 9, "bold"))
        self.lbl_act.pack(side="left", fill="x", expand=True)
        self.btn_actualizar = ttk.Button(self.marco_act, text="Actualizar My Maps (mismo link)",
                                         command=self.actualizar)
        self.btn_actualizar.pack(side="right", padx=(8, 0))
        marco2 = ttk.Frame(m2)
        marco2.pack(fill="both", expand=True)
        self.marco_links = marco2
        self.lista_links = ttk.Treeview(marco2, columns=("transportistas", "link"), show="tree headings",
                                        selectmode="browse")
        self.lista_links.heading("#0", text="Entrega / mapa", anchor="w")
        self.lista_links.heading("transportistas", text="Transportistas", anchor="w")
        self.lista_links.heading("link", text="Link", anchor="w")
        self.lista_links.column("#0", width=self.px(390), stretch=False)
        self.lista_links.column("transportistas", width=self.px(250), stretch=False)
        self.lista_links.column("link", width=self.px(220), stretch=True)
        self.lista_links.tag_configure("dia", font=("Segoe UI", 9, "bold"), background="#eef3fb")
        barra_v = ttk.Scrollbar(marco2, orient="vertical", command=self.lista_links.yview)
        self.lista_links.configure(yscrollcommand=barra_v.set)
        self.lista_links.pack(side="left", fill="both", expand=True)
        barra_v.pack(side="right", fill="y")
        self.lista_links.bind("<Double-Button-1>", lambda e: self.abrir_link())

        botones = ttk.Frame(m2)
        botones.pack(fill="x", pady=(6, 0))
        ttk.Button(botones, text="Copiar link", command=self.copiar_link).pack(side="left")
        ttk.Button(botones, text="Copiar mensaje del día", command=self.copiar_dia).pack(side="left", padx=(6, 0))
        ttk.Button(botones, text="Abrir en el navegador", command=self.abrir_link).pack(side="left", padx=(6, 0))
        self.btn_limpiar = ttk.Button(botones, text="Limpiar mapas viejos", command=self.limpiar)
        self.btn_limpiar.pack(side="right")
        ttk.Label(m2, text="\"Copiar mensaje del día\" copia título, transportistas y link de todos los mapas de esa "
                           f"entrega, listo para pegar en WhatsApp. \"Limpiar mapas viejos\" manda a la papelera los "
                           f"de entrega de hace más de {mymaps.DIAS_HISTORIAL} días (también los hechos a mano); "
                           "se hace solo en cada subida.",
                  style="Ayuda.TLabel", wraplength=self.px(780)).pack(anchor="w", pady=(4, 0))

        m3 = ttk.Frame(cont)
        m3.pack(fill="x", pady=(8, 0))
        self.barra_progreso = ttk.Progressbar(m3, mode="indeterminate", length=120)
        self.lbl_estado = ttk.Label(m3, text="", style="Ayuda.TLabel")
        self.lbl_estado.pack(side="left")
        self.txt = tk.Text(cont, height=6, wrap="word", font=("Consolas", 9), state="disabled",
                           background="#f6f6f6", relief="flat", padx=6, pady=4)
        self.txt.pack(fill="x", pady=(4, 0))

    # -- utilidades -------------------------------------------------------------
    def log(self, mensaje):
        self.txt.config(state="normal")
        self.txt.insert("end", mensaje.rstrip() + "\n")
        self.txt.see("end")
        self.txt.config(state="disabled")

    def estado_subir(self):
        """Subir solo si hay mapas, no se esta subiendo, y la revision de la
        carpeta elegida dijo que TODOS los clientes tienen georreferencia."""
        listo = (not self.trabajando and bool(self.planes) and self.georef["estado"] == "ok"
                 and self.georef["carpeta"] == self.carpeta_actual())
        self.btn_subir.config(state="normal" if listo else "disabled")
        for boton in (self.btn_revisar, self.btn_limpiar, self.btn_actualizar):
            boton.config(state="disabled" if self.trabajando else "normal")

    def ocupar(self, ocupado, mensaje=""):
        self.trabajando = ocupado
        self.estado_subir()
        self.btn_cancelar.config(state="normal" if ocupado else "disabled")
        for cmb in (self.cmb_carpeta, self.cmb_entrega):
            cmb.config(state="disabled" if ocupado else "readonly")
        self.lbl_estado.config(text=mensaje)
        if ocupado:
            self.barra_progreso.pack(side="right")
            self.barra_progreso.start(12)
        else:
            self.barra_progreso.stop()
            self.barra_progreso.pack_forget()

    def revisar_cola(self):
        try:
            while True:
                try:
                    tipo, *datos = self.cola.get_nowait()
                except queue.Empty:
                    break
                try:
                    self.procesar(tipo, datos)
                except Exception:
                    detalle = traceback.format_exc()
                    registrar_error(detalle)
                    if self.viva():
                        self.ocupar(False, "Hubo un problema, revisa el mensaje.")
                        self.messagebox.showerror("My Maps", "Algo falló inesperadamente:\n\n"
                                                  + detalle.strip().splitlines()[-1], parent=self.v)
        finally:
            if self.viva():
                self.v.after(150, self.revisar_cola)

    def procesar(self, tipo, datos):
        if tipo == "log":
            self.log(datos[0])
            return
        if tipo == "georef":
            # Resultado de la revision en segundo plano: no es el fin de una
            # subida (no toca "trabajando"). Se ignora si ya es de otra carpeta
            # o hubo una revision mas nueva.
            n, carpeta, faltan, al_dia, error, cambiados = datos
            if n == self.revision_n:
                self.georef = {"carpeta": carpeta, "faltan": faltan or [], "al_dia": al_dia, "error": error,
                               "cambiados": cambiados or [],
                               "estado": "error" if error else ("faltan" if faltan else
                                                                ("cambiados" if cambiados else "ok"))}
                self.pintar_georef()
                self.estado_subir()
            return
        if tipo == "actualizables":
            # Resultado de buscar mapas con cambios para My Maps (en segundo plano).
            resumen, avisos = datos
            if self.trabajando:
                return              # se vuelve a buscar al terminar el trabajo en curso
            self.por_actualizar = resumen or []
            self.pintar_actualizables(avisos or [])
            if not self.por_actualizar:
                self.ofrecer_actualizar = False     # nada que ofrecer
            if not self.ofrecer_programado:
                self.ofrecer_si_se_puede()
            return
        self.ocupar(False)
        self.trabajo = None
        self.pintar_links()
        cerrar, self.cerrar_al_terminar = self.cerrar_al_terminar, None
        if cerrar is not None and tipo == "error":
            # Se pidio cancelar y cerrar. Queda anotado por si al final no se
            # cierra (p. ej. "Actualizar ruta" estaba publicando y se dijo No).
            self.log("ERROR: " + datos[0])
            self.lbl_estado.config(text="Cancelado.")
            cerrar.destroy()
            return
        if cerrar is None:
            # Al terminar cualquier trabajo se vuelve a mirar si hay mapas
            # con cambios (pudieron aplicarse mientras tanto).
            self.buscar_actualizaciones()
        # Si se pidio cerrar pero el trabajo alcanzo a terminar bien, primero
        # se avisa lo que quedo hecho (y se copian los links) y luego se cierra.
        previo = "No alcanzó a cancelarse: ya había terminado.\n\n" if cerrar is not None else ""
        self.terminar(tipo, datos, previo)
        if cerrar is not None:
            cerrar.destroy()

    def terminar(self, tipo, datos, previo=""):
        if tipo == "fin_actualizar":
            hechos = datos[0]
            if not hechos:
                self.lbl_estado.config(text="No había nada para actualizar.")
                self.messagebox.showinfo("My Maps", previo + "No había cambios para pasar a My Maps.", parent=self.v)
                return
            self.lbl_estado.config(text=f"Listo: {len(hechos)} mapas actualizados (mismos links).")
            self.messagebox.showinfo("My Maps", previo + "Listo. Se actualizaron en My Maps:\n\n"
                                     + "\n".join(f"• {t}: {', '.join(c)}" for t, c in hechos)
                                     + "\n\nLos links son los mismos: los choferes solo vuelven a abrir el mapa.",
                                     parent=self.v)
            return
        if tipo == "fin_limpiar":
            enviados = datos[0]
            self.lbl_estado.config(text=f"Limpieza lista: {len(enviados)} mapas a la papelera.")
            self.messagebox.showinfo("My Maps", previo + ((f"Se mandaron a la papelera {len(enviados)} mapas:\n\n"
                                                           + "\n".join(f"• {m.get('titulo')}" for m in enviados[:20])
                                                           + ("\n..." if len(enviados) > 20 else ""))
                                                          if enviados else f"No hay mapas con entrega de hace más "
                                                                           f"de {mymaps.DIAS_HISTORIAL} días."),
                                     parent=self.v)
            return
        if tipo == "fin":
            creados, fecha = datos
            self.refrescar_planes()
            texto = mymaps.texto_para_compartir(creados)
            self.copiar(texto)
            self.seleccionar_dia(fecha)
            self.lbl_estado.config(text=f"Listo: {len(creados)} mapas en My Maps.")
            self.messagebox.showinfo("My Maps", previo + f"Listo. Se crearon {len(creados)} mapas en My Maps para el "
                                     f"{mymaps.fecha_larga(fecha)}.\n\nEl mensaje con los links ya está copiado: "
                                     "pégalo en WhatsApp.", parent=self.v)
        elif tipo == "error":
            self.refrescar_planes()
            self.lbl_estado.config(text="No se completó, revisa el mensaje.")
            self.log("ERROR: " + datos[0])
            self.revisar_georef()       # por si el problema fue de ubicaciones: que se vea abajo
            self.messagebox.showerror("My Maps", datos[0], parent=self.v)

    def preguntar(self, titulo, texto):
        """askyesno sin que la oferta automatica de actualizar salga encima
        (Tk sigue atendiendo la cola mientras la pregunta esta abierta). Si
        llego una oferta mientras tanto, se vuelve a buscar al cerrarla."""
        self.preguntando += 1
        try:
            return self.messagebox.askyesno(titulo, texto, parent=self.v)
        finally:
            self.preguntando -= 1
            if not self.preguntando and self.ofrecer_actualizar and self.viva():
                self.v.after_idle(self.ofrecer_si_se_puede)

    def ofrecer_si_se_puede(self):
        """Ofrece actualizar (si se pidio y hay algo) cuando no hay otra
        pregunta ni trabajo, aqui ni en "Actualizar ruta"; si no, lo vuelve
        a intentar en un segundo."""
        self.ofrecer_programado = False
        if not self.viva() or not self.ofrecer_actualizar or self.trabajando or not self.por_actualizar:
            return
        if self.preguntando or self.otra_ocupada():
            self.ofrecer_programado = True
            self.v.after(1000, self.ofrecer_si_se_puede)
            return
        self.ofrecer_actualizar = False
        self.actualizar()

    def copiar(self, texto):
        self.v.clipboard_clear()
        self.v.clipboard_append(texto)

    # -- ruta y fecha -------------------------------------------------------------
    def cargar_carpetas(self, preferida=None):
        self.carpetas = carpetas_del_dia()
        if preferida and Path(preferida).is_dir() and Path(preferida) not in self.carpetas \
                and mymaps.fecha_de_carpeta(Path(preferida)):
            self.carpetas.insert(0, Path(preferida))
        nombres = []
        for p in self.carpetas:
            d = mymaps.fecha_de_carpeta(p)
            nombres.append(f"Liquidación {mymaps.DIAS_LARGOS[d.weekday()].capitalize()} {d:%d/%m/%Y}  "
                           f"({len(mymaps.archivos_mapa(p))} archivos)")
        self.cmb_carpeta.config(values=nombres)
        if not self.carpetas:
            self.cmb_carpeta.set("")
            self.lbl_aviso.config(text="Todavía no hay rutas generadas en BASE\\MAPAS. Primero genera la ruta en "
                                       "\"Actualizar ruta\".")
            self.planes = []
            self.ocupar(False)
            return
        i = self.carpetas.index(Path(preferida)) if preferida and Path(preferida) in self.carpetas else 0
        self.cmb_carpeta.current(i)
        self.carpeta_elegida()

    def carpeta_actual(self):
        i = self.cmb_carpeta.current()
        return self.carpetas[i] if 0 <= i < len(self.carpetas) else None

    def carpeta_elegida(self):
        carpeta = self.carpeta_actual()
        if not carpeta:
            return
        liq = mymaps.fecha_de_carpeta(carpeta)
        self.fechas_entrega = mymaps.opciones_de_entrega(liq)
        self.cmb_entrega.config(values=[mymaps.fecha_larga(d) for d in self.fechas_entrega])
        self.cmb_entrega.current(self.fechas_entrega.index(mymaps.entrega_sugerida(liq)))
        self.refrescar_planes()
        self.revisar_georef()

    def revisar_georef(self):
        carpeta = self.carpeta_actual()
        if not carpeta or self.trabajando:
            return
        self.revision_n += 1
        n = self.revision_n
        self.georef = {"carpeta": carpeta, "estado": "revisando", "faltan": [], "al_dia": True}
        self.pintar_georef()
        self.estado_subir()

        def revisar():
            try:
                faltan, al_dia, cambiados = revisar_carpeta(carpeta)
                self.cola.put(("georef", n, carpeta, faltan, al_dia, None, cambiados))
            except Exception as e:
                registrar_error(traceback.format_exc())
                self.cola.put(("georef", n, carpeta, None, False, str(e) or type(e).__name__, None))
        threading.Thread(target=revisar, daemon=True).start()

    def pintar_georef(self):
        g = self.georef
        if g["estado"] == "revisando":
            texto, color = "Revisando que todos los clientes tengan ubicación...", "#666666"
        elif g["estado"] == "error":
            texto, color = (f"No se pudo revisar la ubicación de los clientes: {g.get('error')}. "
                            "Revisa que los archivos MAPA no estén abiertos y dale a \"Revisar de nuevo\"."), "#b00020"
        elif g["estado"] == "faltan":
            faltan = g["faltan"]
            from generar_rutas import texto_sin_georef
            texto = (f"✗ {len(faltan)} clientes SIN ubicación (todos deben tenerla para subir):\n"
                     f"{texto_sin_georef(faltan, maximo=6)}\nCorrígelos en \"Actualizar ruta\" (lista en rojo), "
                     "vuelve a \"Generar y publicar\" y después súbelo aquí.")
            color = "#b00020"
        elif g["estado"] == "cambiados":
            texto, color = "✗ " + mymaps.texto_cambiados(g["cambiados"]), "#b00020"
        else:
            texto, color = "✓ Todos los clientes tienen ubicación.", "#1b7a2f"
        if g["estado"] in ("ok", "faltan") and not g.get("al_dia", True):
            texto += "  (Sin conexión con GitHub: revisado con las correcciones de esta PC.)"
        self.lbl_georef.config(text=texto, fg=color)

    def entrega_actual(self):
        i = self.cmb_entrega.current()
        return self.fechas_entrega[i] if 0 <= i < len(self.fechas_entrega) else None

    def refrescar_planes(self):
        carpeta, entrega = self.carpeta_actual(), self.entrega_actual()
        self.lista_planes.delete(*self.lista_planes.get_children())
        self.planes = mymaps.armar_mapas(mymaps.archivos_mapa(carpeta), entrega) if carpeta and entrega else []
        for plan in self.planes:
            self.lista_planes.insert("", "end", values=(plan["titulo"], len(plan["archivos"]),
                                                        "-".join(plan["transportistas"])))
        self.lista_planes.config(height=max(3, min(len(self.planes), 6)))
        capas = sum(len(p["archivos"]) for p in self.planes)
        self.lbl_tiempo.config(text=f"{len(self.planes)} mapas, {capas} capas · tarda unos "
                                    f"{minutos(self.planes)} min" if self.planes else "")
        avisos = []
        if entrega:
            ya = mymaps.activos(mymaps.leer_registro(), entrega.isoformat())
            if ya:
                avisos.append(f"Ya hay {len(ya)} mapas en My Maps para el {mymaps.fecha_larga(entrega)}: si los "
                              "vuelves a subir, los anteriores van a la papelera y sus links dejan de servir.")
            if entrega < date.today():
                avisos.append("OJO: la fecha de entrega ya pasó.")
        self.lbl_aviso.config(text="  ".join(avisos))
        self.estado_subir()

    # -- subir ----------------------------------------------------------------------
    def subir(self):
        carpeta, entrega = self.carpeta_actual(), self.entrega_actual()
        if not carpeta or not entrega or not self.planes:
            return
        registro = mymaps.leer_registro()
        ya = mymaps.activos(registro, entrega.isoformat())
        capas = sum(len(p["archivos"]) for p in self.planes)
        pregunta = (f"Se van a crear {len(self.planes)} mapas en My Maps para el {mymaps.fecha_larga(entrega)} "
                    f"({capas} capas):\n\n"
                    + "\n".join(f"• {p['titulo']}: {'-'.join(p['transportistas'])}" for p in self.planes)
                    + f"\n\nTarda unos {minutos(self.planes)} minutos. Se abre un Chrome "
                      "aparte (el del robot): no lo uses mientras trabaja.")
        if ya:
            pregunta += (f"\n\nOJO: ya hay {len(ya)} mapas para esa fecha. Se mandan a la papelera y sus links "
                         "dejan de servir (habrá que mandar los nuevos).")
        vencidos = [m for m in mymaps.por_mandar_a_papelera(registro, date.today())
                    if m.get("estado") == "activo"]
        if vencidos:
            pregunta += f"\n\nTambién van a la papelera {len(vencidos)} mapas de hace más de {mymaps.DIAS_HISTORIAL} días."
        if not self.preguntar("Subir a My Maps", pregunta + "\n\n¿Continuar?") or self.trabajando:
            return
        self.cancelar.clear()
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.config(state="disabled")
        self.trabajo = "subir"
        self.ocupar(True, "Subiendo a My Maps...")
        threading.Thread(target=self._subir, args=(carpeta, entrega, bool(ya)), daemon=True).start()

    def _subir(self, carpeta, entrega, reemplazar):
        def log(mensaje):
            self.cola.put(("log", mensaje))

        try:
            creados = mymaps.subir(carpeta, entrega, log, self.cancelar, reemplazar=reemplazar)
            self.cola.put(("fin", creados, entrega))
        except mymaps.Cancelado:
            self.cola.put(("error", TEXTOS_TRABAJO["subir"]["cancelado"]))
        except mymaps.ErrorMyMaps as e:
            self.cola.put(("error", str(e)))
        except Exception as e:
            detalle = traceback.format_exc()
            registrar_error(detalle)
            if type(e).__name__ == "ErrorRuta":
                # Mensaje ya pensado para el usuario (de generar_rutas).
                self.cola.put(("error", str(e)))
            elif type(e).__module__.startswith("playwright"):
                self.cola.put(("error", "My Maps no respondió como se esperaba (puede que Google haya cambiado "
                                        "algo en la página o que se haya cerrado el Chrome del robot).\n\n"
                                        "Vuelve a intentar; si se repite, súbelo a mano con los archivos de la "
                                        f"carpeta del día.\n\nDetalle: {str(e).splitlines()[0][:200]}"))
            else:
                self.cola.put(("error", "Algo falló inesperadamente:\n\n" + detalle.strip().splitlines()[-1]))

    # -- actualizar ubicaciones / limpiar --------------------------------------------
    def buscar_actualizaciones(self):
        """En segundo plano: mapas ya subidos con cambios despues (ubicaciones
        corregidas o ruta vuelta a publicar)."""
        def buscar():
            try:
                resumen, avisos = revisar_actualizaciones()
            except Exception:
                registrar_error(traceback.format_exc())
                resumen, avisos = [], []
            self.cola.put(("actualizables", resumen, avisos))
        threading.Thread(target=buscar, daemon=True).start()

    def pintar_actualizables(self, avisos=()):
        if not self.por_actualizar and not avisos:
            self.marco_act.pack_forget()
            return
        if self.por_actualizar:
            capas = sum(len(c) for _, c in self.por_actualizar)
            texto = (f"⟳ Cambios que aún no están en My Maps (ubicaciones corregidas o ruta vuelta a publicar): "
                     f"{len(self.por_actualizar)} mapas, "
                     f"{capas} capas (" + "; ".join(f"{t.split('-MAPA')[0]}: {', '.join(c)}"
                                                    for t, c in self.por_actualizar) + ").")
        else:
            texto = ""
        if avisos:
            texto = (texto + "\n" if texto else "") + "\n".join(avisos)
        self.lbl_act.config(text=texto)
        self.btn_actualizar.pack(side="right", padx=(8, 0)) if self.por_actualizar else self.btn_actualizar.pack_forget()
        if self.marco_act.winfo_manager() != "pack":
            self.marco_act.pack(fill="x", pady=(0, 6), before=self.marco_links)

    def actualizar(self):
        if self.trabajando or not self.por_actualizar:
            return
        capas = sum(len(c) for _, c in self.por_actualizar)
        pregunta = ("Se van a pasar a My Maps los cambios (ubicaciones corregidas o ruta vuelta a publicar). "
                    "Los links NO cambian:\n\n"
                    + "\n".join(f"• {t}: {', '.join(c)}" for t, c in self.por_actualizar)
                    + f"\n\nTarda unos {max(1, round(capas * 1.5))} minutos. Se abre el Chrome del robot: no lo "
                      "uses mientras trabaja.")
        if not self.preguntar("Actualizar My Maps", pregunta + "\n\n¿Continuar?") or self.trabajando:
            return
        self.cancelar.clear()
        self.trabajo = "actualizar"
        self.ocupar(True, "Actualizando My Maps...")
        threading.Thread(target=self._trabajo_robot, args=(self._actualizar,), daemon=True).start()

    def _actualizar(self, log):
        self.cola.put(("fin_actualizar", mymaps.actualizar_ubicaciones(log, self.cancelar)))

    def limpiar(self):
        if self.trabajando:
            return
        if not self.preguntar(
                "Limpiar mapas viejos",
                "Se revisan tus mapas de My Maps (también los que hiciste a mano, por su fecha de entrega) y se "
                f"mandan a la papelera de Drive los de entrega de hace más de {mymaps.DIAS_HISTORIAL} días.\n"
                "Se abre el Chrome del robot unos minutos.\n\n¿Continuar?") or self.trabajando:
            return
        self.cancelar.clear()
        self.trabajo = "limpiar"
        self.ocupar(True, "Limpiando mapas viejos...")
        threading.Thread(target=self._trabajo_robot, args=(self._limpiar,), daemon=True).start()

    def _limpiar(self, log):
        self.cola.put(("fin_limpiar", mymaps.limpiar(log, self.cancelar)))

    def _trabajo_robot(self, funcion):
        """Corre un trabajo del robot en segundo plano con los mismos mensajes
        de error que la subida."""
        def log(mensaje):
            self.cola.put(("log", mensaje))
        try:
            funcion(log)
        except mymaps.Cancelado:
            self.cola.put(("error", TEXTOS_TRABAJO.get(self.trabajo, TEXTOS_TRABAJO["subir"])["cancelado"]))
        except mymaps.ErrorMyMaps as e:
            self.cola.put(("error", str(e)))
        except Exception as e:
            registrar_error(traceback.format_exc())
            if type(e).__name__ == "ErrorRuta":
                self.cola.put(("error", str(e)))
            elif type(e).__module__.startswith("playwright"):
                self.cola.put(("error", "My Maps no respondió como se esperaba (puede que Google haya cambiado "
                                        "algo en la página o que se haya cerrado el Chrome del robot).\n\nVuelve a "
                                        f"intentar.\n\nDetalle: {str(e).splitlines()[0][:200]}"))
            else:
                self.cola.put(("error", "Algo falló inesperadamente:\n\n"
                                        + traceback.format_exc().strip().splitlines()[-1]))

    def pedir_cancelar(self):
        textos = TEXTOS_TRABAJO.get(self.trabajo, TEXTOS_TRABAJO["subir"])
        if not self.trabajando:
            return
        si = self.preguntar("Cancelar", textos["pregunta"])
        pendiente, self.cerrar_pendiente = self.cerrar_pendiente, None
        if not si:
            return
        if not self.trabajando:         # termino mientras se preguntaba
            if pendiente is not None:
                pendiente.destroy()
            return
        self.cancelar.set()
        if pendiente is not None:
            # Mientras se preguntaba se cerro "Actualizar ruta": se cierra al terminar.
            self.cerrar_al_terminar = pendiente
        self.lbl_estado.config(text="Cancelando (termina el paso que está haciendo)...")

    def al_cerrar(self, cerrar=None):
        """cerrar: la ventana que se cierra al final (esta, o la de
        "Actualizar ruta" si se cerro aquella mientras el robot trabajaba)."""
        cerrar = cerrar or self.v
        if not self.trabajando:
            cerrar.destroy()
            return
        if self.cerrar_al_terminar is not None:
            # Ya se pidio cancelar y cerrar: no se vuelve a preguntar. Si ahora
            # se pide cerrar tambien "Actualizar ruta", se cierra todo.
            if cerrar is not self.v:
                self.cerrar_al_terminar = cerrar
            return
        if self.preguntando:
            # Ya hay una pregunta abierta aqui: no se apila otra. Si es la X de
            # "Actualizar ruta", se recuerda y vale si se contesta que si.
            if cerrar is not self.v:
                self.cerrar_pendiente = cerrar
            return
        textos = TEXTOS_TRABAJO.get(self.trabajo, TEXTOS_TRABAJO["subir"])
        si = self.preguntar("My Maps", f"{textos['haciendo']}. ¿Cancelar y cerrar?")
        pendiente, self.cerrar_pendiente = self.cerrar_pendiente, None
        if not si:
            return
        cerrar = pendiente or cerrar
        if not self.trabajando:         # termino mientras se preguntaba
            cerrar.destroy()
            return
        self.cancelar.set()
        self.cerrar_al_terminar = cerrar
        self.lbl_estado.config(text="Cancelando; se cierra solo al terminar el paso en curso...")

    # -- links ------------------------------------------------------------------------
    def pintar_links(self):
        registro = mymaps.leer_registro()
        limite = (date.today() - timedelta(days=mymaps.DIAS_HISTORIAL)).isoformat()
        por_dia = {}
        for m in mymaps.activos(registro):
            if m.get("fecha_entrega", "") >= limite and m.get("url"):
                por_dia.setdefault(m["fecha_entrega"], []).append(m)
        self.lista_links.delete(*self.lista_links.get_children())
        for i, dia in enumerate(sorted(por_dia, reverse=True)):
            d = date.fromisoformat(dia)
            padre = self.lista_links.insert("", "end", iid=f"dia:{dia}", text=mymaps.fecha_larga(d),
                                            values=(f"{len(por_dia[dia])} mapas", ""), open=(i < 2), tags=("dia",))
            for m in sorted(por_dia[dia], key=lambda m: m.get("numero", 0)):
                if self.lista_links.exists(f"mapa:{m['mid']}"):
                    continue
                self.lista_links.insert(padre, "end", iid=f"mapa:{m['mid']}", text=m["titulo"],
                                        values=("-".join(m.get("transportistas", []))
                                                or ("(hecho a mano)" if m.get("origen") == "manual" else ""),
                                                m["url"]))
        if not por_dia:
            self.lista_links.insert("", "end", text="Todavía no hay mapas subidos con el robot.")

    def seleccionar_dia(self, fecha):
        iid = f"dia:{fecha.isoformat()}"
        if self.lista_links.exists(iid):
            self.lista_links.item(iid, open=True)
            self.lista_links.selection_set(iid)
            self.lista_links.see(iid)

    def mapa_elegido(self):
        sel = self.lista_links.selection()
        if not sel or not sel[0].startswith("mapa:"):
            return None
        mid = sel[0][5:]
        return next((m for m in mymaps.activos(mymaps.leer_registro()) if m.get("mid") == mid), None)

    def dia_elegido(self):
        sel = self.lista_links.selection()
        if not sel:
            return None
        iid = sel[0] if sel[0].startswith("dia:") else self.lista_links.parent(sel[0])
        return iid[4:] if iid.startswith("dia:") else None

    def copiar_link(self):
        m = self.mapa_elegido()
        if not m:
            self.messagebox.showinfo("Copiar link", "Elige un mapa de la lista (no la fecha).", parent=self.v)
            return
        self.copiar(m["url"])
        self.lbl_estado.config(text=f"Link copiado: {m['titulo']}")

    def copiar_dia(self):
        dia = self.dia_elegido()
        if not dia:
            self.messagebox.showinfo("Copiar mensaje", "Elige una fecha (o un mapa de esa fecha) en la lista.",
                                     parent=self.v)
            return
        mapas = sorted(mymaps.activos(mymaps.leer_registro(), dia), key=lambda m: m.get("numero", 0))
        self.copiar(mymaps.texto_para_compartir(mapas))
        self.lbl_estado.config(text=f"Mensaje copiado: {len(mapas)} mapas del "
                                    f"{mymaps.fecha_larga(date.fromisoformat(dia))}.")

    def abrir_link(self):
        m = self.mapa_elegido()
        if m:
            webbrowser.open(m["url"])


if __name__ == "__main__":
    try:
        abrir()
    except Exception:
        registrar_error(traceback.format_exc())
        raise
