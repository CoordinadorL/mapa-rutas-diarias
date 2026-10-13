"""
Sube las rutas del dia a Google My Maps (lo que antes se hacia a mano) y
lleva el registro de los links.

My Maps no tiene API: un "robot" (Playwright) maneja un Chrome aparte, con su
propio perfil en data/chrome-mymaps, donde la cuenta de Opaustro inicio
sesion a mano (el robot nunca escribe contrasenas). Por cada grupo de
camiones crea un mapa:

  titulo       "1-MAPA DE ENTREGAS LUNES 12 DE OCTUBRE"  (fecha de ENTREGA)
  descripcion  "Transportistas F150-F151-F153"
  capas        una por archivo MAPA (max. 10 por mapa), "Agrupar sitios por:
               vendedor" con un color distinto por vendedor, etiquetas con
               la Razon, y todas desmarcadas al abrir
  compartir    visible para cualquiera con el enlace, sin busqueda publica

Grupos: F1xx+F2xx, F3xx, FO (y cualquier otro prefijo aparte); si un grupo
pasa de 10 capas se parte en mapas parejos.

Las capas se suben desde una copia de cada MAPA con las ubicaciones
corregidas (las mismas que usa el mapa web). Los mapas con fecha de entrega
de hace mas de 30 dias se mandan a la papelera de Drive (solo en My Maps: el
mapa web guarda su historial completo).
"""
import json
import math
import os
import re
import shutil
import subprocess
import tempfile
import time
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import parse_qs, urlparse

CARPETA = Path(__file__).resolve().parent
PERFIL_CHROME = CARPETA / "data" / "chrome-mymaps"
ARCHIVO_REGISTRO = CARPETA / "data" / "mymaps_registro.json"
CAPTURA_ERROR = CARPETA / "data" / "mymaps_error.png"
# La cuenta de Google de los mapas se anota aqui (data/ nunca va a GitHub):
# {"cuenta": "...@gmail.com"}. Si no esta, se toma la primera con la que se
# inicie sesion en el Chrome del robot.
ARCHIVO_CONFIG = CARPETA / "data" / "mymaps_config.json"
PUERTO = 9333
INICIO = "https://www.google.com/maps/d/?hl=es"
MAX_CAPAS = 10
DIAS_HISTORIAL = 30
MINUTOS_PARA_INICIAR_SESION = 10

DIAS_LARGOS = ["LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO", "DOMINGO"]
MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO",
         "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

# Mismos nombres que escribe generar_rutas.escribir_excel.
RE_ARCHIVO = re.compile(r"MAPA ([A-Z]{1,3})(\d{1,4})( CARGUE (\d+) - Liq [0-9A-Za-z]+)?\.xlsx")

# My Maps solo deja elegir entre sus 30 colores fijos. Este orden alterna
# tono y claridad, para que dos vendedores seguidos nunca se parezcan.
PALETA = [
    (255, 82, 82), (2, 136, 209), (15, 157, 88), (245, 124, 0), (156, 39, 176),
    (255, 234, 0), (0, 0, 0), (0, 151, 167), (194, 24, 91), (124, 179, 66),
    (26, 35, 126), (121, 85, 72), (189, 189, 189), (165, 39, 20), (103, 58, 183),
    (175, 180, 43), (0, 96, 100), (249, 168, 37), (136, 14, 79), (57, 73, 171),
    (9, 113, 56), (230, 81, 0), (117, 117, 117), (1, 87, 155), (129, 119, 23),
    (85, 139, 47), (255, 214, 0), (251, 192, 45), (78, 52, 46), (66, 66, 66),
]


class ErrorMyMaps(Exception):
    """Error con un mensaje ya pensado para mostrarse tal cual."""


class Cancelado(ErrorMyMaps):
    pass


# ---------------------------------------------------------------------------
# Fechas, grupos y textos
# ---------------------------------------------------------------------------

def fecha_larga(d: date) -> str:
    return f"{DIAS_LARGOS[d.weekday()]} {d.day:02d} DE {MESES[d.month - 1]}"


def entrega_sugerida(fecha_liq: date) -> date:
    """El siguiente dia habil: viernes (o viernes + sabado) se entrega el lunes."""
    d = fecha_liq + timedelta(days=1)
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d


def opciones_de_entrega(fecha_liq: date, cuantas=8):
    """Dias que se pueden elegir como entrega (sin domingos), por si el
    siguiente dia habil no se trabaja."""
    dias, d = [], fecha_liq
    while len(dias) < cuantas:
        d += timedelta(days=1)
        if d.weekday() != 6:
            dias.append(d)
    return dias


def fecha_de_carpeta(carpeta: Path):
    """BASE\\MAPAS\\2026-10-09 VIERNES -> date(2026, 10, 9)."""
    try:
        return datetime.strptime(carpeta.name[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


ARCHIVO_PUBLICADO = "_publicado.json"     # lo escribe generar_rutas.anotar_publicado


def archivos_mapa(carpeta: Path):
    """Los MAPA de la carpeta del dia. Si la ruta se publico, solo los que se
    publicaron (asi My Maps sube lo mismo que el mapa web, aunque despues se
    haya sacado algun Excel suelto con "Solo generar Excel")."""
    todos = sorted(p for p in carpeta.iterdir() if p.is_file() and RE_ARCHIVO.fullmatch(p.name))
    try:
        publicados = set(json.loads((carpeta / ARCHIVO_PUBLICADO).read_text(encoding="utf-8"))["archivos"])
    except (OSError, ValueError, KeyError, TypeError):
        return todos
    return [p for p in todos if p.name in publicados]


def cambiados_desde_publicar(carpeta: Path):
    """MAPA publicados que ya no son iguales a lo que se publico (o que
    faltan). Lista vacia si no hay anotacion de lo publicado."""
    from generar_rutas import huella_archivo
    try:
        huellas = json.loads((carpeta / ARCHIVO_PUBLICADO).read_text(encoding="utf-8")).get("huellas") or {}
    except (OSError, ValueError, AttributeError):
        return []
    cambiados = []
    for nombre, huella in sorted(huellas.items()):
        ruta = carpeta / nombre
        try:
            if huella_archivo(ruta) != huella:
                cambiados.append(nombre)
        except OSError:
            cambiados.append(nombre)
    return cambiados


def texto_cambiados(cambiados) -> str:
    return (f"Estos archivos cambiaron después de publicar la ruta (por ejemplo con \"Solo generar Excel\"): "
            f"{', '.join(cambiados)}.\nPara que My Maps tenga lo mismo que el mapa web, vuelve a \"Generar y "
            "publicar\" en \"Actualizar ruta\".")


def datos_archivo(ruta: Path):
    m = RE_ARCHIVO.fullmatch(ruta.name)
    letras, numero, cargue = m.group(1), m.group(2), m.group(4)
    return {"camion": letras + numero, "letras": letras, "numero": int(numero),
            "cargue": int(cargue) if cargue else 0}


def grupo_de(letras: str, numero: str):
    """(orden, nombre) del mapa al que va un camion."""
    if letras == "F" and numero[0] in "12":
        return (1, "F1-F2")
    if letras == "F":
        return (int(numero[0]), f"F{numero[0]}xx")
    if letras == "FO":
        return (50, "FO")
    return (90, letras)


def partir(camiones, maximo=MAX_CAPAS):
    """Reparte camiones (cada uno con 1+ capas) en mapas de hasta `maximo`
    capas, parejos, sin separar los cargues de un mismo camion."""
    total = sum(len(capas) for _, capas in camiones)
    if total <= maximo:
        return [camiones]
    objetivo = math.ceil(total / math.ceil(total / maximo))
    partes, actual, n = [], [], 0
    for camion in camiones:
        capas = len(camion[1])
        if actual and (n + capas > objetivo or n + capas > maximo):
            partes.append(actual)
            actual, n = [], 0
        actual.append(camion)
        n += capas
    if actual:
        partes.append(actual)
    return partes


def armar_mapas(archivos, fecha_entrega: date):
    """Lista de mapas a crear: numero, titulo, descripcion, transportistas y
    archivos (capas), en el orden F1-F2, F3, ..., FO, otros."""
    camiones = {}
    for ruta in archivos:
        d = datos_archivo(ruta)
        clave = (grupo_de(d["letras"], str(d["numero"])), d["letras"], d["numero"])
        camiones.setdefault(clave, (d["camion"], []))[1].append((d["cargue"], ruta))
    grupos = {}
    for clave in sorted(camiones):
        camion, capas = camiones[clave]
        grupos.setdefault(clave[0], []).append((camion, [r for _, r in sorted(capas)]))
    mapas = []
    for grupo in sorted(grupos):
        for parte in partir(grupos[grupo]):
            n = len(mapas) + 1
            ids = [camion for camion, _ in parte]
            mapas.append({
                "numero": n,
                "grupo": grupo[1],
                "titulo": f"{n}-MAPA DE ENTREGAS {fecha_larga(fecha_entrega)}",
                "descripcion": "Transportistas " + "-".join(ids),
                "transportistas": ids,
                "archivos": [r for _, capas in parte for r in capas],
            })
    return mapas


def texto_para_compartir(mapas) -> str:
    """Para pegar en WhatsApp: titulo, transportistas y link de cada mapa."""
    bloques = [f"{m['titulo']}\n{m['descripcion']}\n{m['url']}" for m in mapas if m.get("url")]
    return "\n\n".join(bloques)


# ---------------------------------------------------------------------------
# Registro de mapas creados (data/mymaps_registro.json)
# ---------------------------------------------------------------------------

def leer_registro(ruta: Path = ARCHIVO_REGISTRO) -> dict:
    try:
        datos = json.loads(ruta.read_text(encoding="utf-8"))
        if isinstance(datos, dict) and isinstance(datos.get("mapas"), list):
            return datos
    except (OSError, ValueError):
        pass
    return {"mapas": []}


def guardar_registro(registro: dict, ruta: Path = ARCHIVO_REGISTRO) -> None:
    ruta.parent.mkdir(parents=True, exist_ok=True)
    temporal = ruta.with_suffix(".tmp")
    temporal.write_text(json.dumps(registro, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(temporal, ruta)


def activos(registro: dict, fecha_entrega: str = None):
    return [m for m in registro["mapas"] if m.get("estado") == "activo"
            and (fecha_entrega is None or m.get("fecha_entrega") == fecha_entrega)]


def por_mandar_a_papelera(registro: dict, hoy: date, reemplazar_fecha: str = None):
    """Activos con entrega de hace mas de 30 dias, los de la fecha que se
    reemplaza, y los que quedaron a medias en un intento anterior."""
    limite = (hoy - timedelta(days=DIAS_HISTORIAL)).isoformat()
    salida = []
    for m in registro["mapas"]:
        if not m.get("mid"):
            continue
        if m.get("estado") in ("creando", "fallido"):
            salida.append(m)
        elif m.get("estado") == "activo" and (m.get("fecha_entrega", "") < limite
                                              or m.get("fecha_entrega") == reemplazar_fecha):
            salida.append(m)
    return salida


def limpiar_registro(registro: dict, hoy: date) -> None:
    """Lo que ya esta en la papelera hace mas de 60 dias deja de anotarse."""
    limite = (hoy - timedelta(days=2 * DIAS_HISTORIAL)).isoformat()
    registro["mapas"] = [m for m in registro["mapas"]
                         if not (m.get("estado") in ("papelera", "sin_crear")
                                 and m.get("fecha_entrega", "") < limite)]


# ---------------------------------------------------------------------------
# Copias con ubicaciones corregidas
# ---------------------------------------------------------------------------

COLUMNAS_UBICACION = ("CodigoCliente", "longitud_x", "latitud_y")


def _encabezado(ws, ruta: Path):
    posiciones = {str(c.value or "").strip(): i for i, c in enumerate(next(ws.iter_rows(max_row=1)))}
    if not all(c in posiciones for c in COLUMNAS_UBICACION):
        raise ErrorMyMaps(f"\"{ruta.name}\" no tiene las columnas CodigoCliente, longitud_x y latitud_y.")
    return posiciones


def camion_de_archivo(ruta: Path) -> str:
    """'MAPA F152 CARGUE 1 - Liq 1557.xlsx' -> 'F152 CARGUE 1 - Liq 1557'."""
    return ruta.name[len("MAPA "):-len(".xlsx")]


def revisar_georef(archivos, overrides: dict):
    """Clientes de los MAPA sin georreferencia valida, con la MISMA regla del
    mapa web (generar_rutas.clientes_sin_georef). Lista vacia = todo bien."""
    import openpyxl
    from generar_rutas import clientes_sin_georef

    camiones = []
    for ruta in archivos:
        wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
        try:
            ws = wb.worksheets[0]
            pos = _encabezado(ws, ruta)
            filas = []
            for valores in ws.iter_rows(min_row=2, values_only=True):
                if not valores or all(v in (None, "") for v in valores):
                    continue
                filas.append({c: (valores[i] if i < len(valores) else None) for c, i in pos.items()})
        finally:
            wb.close()
        chofer = next((f.get("Chofer") for f in filas if f.get("Chofer")), "")
        camiones.append({"id": camion_de_archivo(ruta), "chofer": str(chofer or ""), "filas": filas})
    return clientes_sin_georef(camiones, overrides)


def preparar_copias(archivos, destino: Path, overrides: dict):
    """Copia cada MAPA (mismo nombre: es el nombre de la capa) con la
    ubicacion corregida de cada cliente (la misma que usa el mapa web).
    Antes hay que revisar con revisar_georef: aqui un cliente sin ubicacion
    valida es un error (un 0,0 aleja el mapa, y una sola celda vacia hace que
    My Maps ya no reconozca la columna como latitud).
    Devuelve (copias, cuantas corregidas)."""
    import openpyxl
    from generar_rutas import a_coordenada, codigo_texto, correccion_valida, problema_de_ubicacion

    copias, corregidas = [], 0
    for ruta in archivos:
        wb = openpyxl.load_workbook(ruta)
        try:
            ws = wb.worksheets[0]
            pos = _encabezado(ws, ruta)
            i_cod, i_lng, i_lat = (pos[c] for c in COLUMNAS_UBICACION)
            vacias = []
            for fila in ws.iter_rows(min_row=2):
                if all(c.value in (None, "") for c in fila):
                    vacias.append(fila[0].row)
                    continue
                ajuste = correccion_valida(overrides, fila[i_cod].value)
                if ajuste:
                    fila[i_lat].value, fila[i_lng].value = ajuste
                    corregidas += 1
                elif problema_de_ubicacion(fila[i_lat].value, fila[i_lng].value):
                    raise ErrorMyMaps(f"El cliente {codigo_texto(fila[i_cod].value)} de {ruta.name} no tiene "
                                      "georreferencia válida.")
                else:
                    fila[i_lat].value = a_coordenada(fila[i_lat].value)
                    fila[i_lng].value = a_coordenada(fila[i_lng].value)
            for numero in reversed(vacias):
                ws.delete_rows(numero)
            copia = destino / ruta.name
            wb.save(copia)
        finally:
            wb.close()
        copias.append(copia)
    return copias, corregidas


# ---------------------------------------------------------------------------
# Robot
# ---------------------------------------------------------------------------

def cuenta_esperada():
    try:
        cuenta = json.loads(ARCHIVO_CONFIG.read_text(encoding="utf-8")).get("cuenta")
        return cuenta.strip().lower() if isinstance(cuenta, str) and "@" in cuenta else None
    except (OSError, ValueError, AttributeError):
        return None


def anotar_cuenta(cuenta: str) -> None:
    ARCHIVO_CONFIG.parent.mkdir(parents=True, exist_ok=True)
    ARCHIVO_CONFIG.write_text(json.dumps({"cuenta": cuenta}, indent=1), encoding="utf-8")


def ruta_chrome() -> Path:
    for variable in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
        base = os.environ.get(variable)
        if base:
            exe = Path(base) / "Google" / "Chrome" / "Application" / "chrome.exe"
            if exe.exists():
                return exe
    raise ErrorMyMaps("No encontré Google Chrome instalado en esta computadora.")


def chrome_escuchando() -> bool:
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{PUERTO}/json/version", timeout=2):
            return True
    except Exception:
        return False


def esperar(condicion, segundos, paso=0.4) -> bool:
    fin = time.time() + segundos
    while True:
        try:
            if condicion():
                return True
        except Exception:
            pass
        if time.time() >= fin:
            return False
        time.sleep(paso)


class Robot:
    """Chrome del robot abierto con su perfil, y los pasos en My Maps.

    Se abre Chrome normal (con un puerto de depuracion local) y Playwright
    se conecta a el: asi es el mismo Chrome donde se inicio sesion a mano."""

    def __init__(self, log, cancelar=None):
        self.log = log
        self.cancelar = cancelar
        self.pw = self.navegador = self.pg = None
        self.mid_en_curso = None
        # Se llama con el mid apenas existe el mapa (para anotarlo aunque
        # despues se corte todo).
        self.al_crear = None

    def __enter__(self):
        from playwright.sync_api import sync_playwright

        if not chrome_escuchando():
            PERFIL_CHROME.mkdir(parents=True, exist_ok=True)
            subprocess.Popen([str(ruta_chrome()), f"--user-data-dir={PERFIL_CHROME}",
                              f"--remote-debugging-port={PUERTO}", "--no-first-run",
                              "--no-default-browser-check", "--lang=es", "--window-size=1300,880", "about:blank"])
            if not esperar(chrome_escuchando, 30, 0.5):
                raise ErrorMyMaps("El Chrome del robot no respondió.\nSi hay una ventana de Chrome del robot "
                                  "abierta (la de My Maps), ciérrala y vuelve a intentar.")
        self.pw = sync_playwright().start()
        try:
            self.navegador = self.pw.chromium.connect_over_cdp(f"http://127.0.0.1:{PUERTO}")
            # El robot trabaja en una pestana propia y nunca toca las demas
            # (puede haber otra cosa abierta en ese Chrome).
            self.pg = self.navegador.contexts[0].new_page()
            self.pg.set_default_timeout(30000)
        except Exception:
            self.pw.stop()
            raise
        return self

    def __exit__(self, *exc):
        # Se cierra la pestana del robot, y el Chrome del robot entero si no
        # le queda nada mas abierto (asi no queda abierto con la sesion).
        try:
            self.pg.close()
            vacias = ("about:blank", "chrome://newtab/", "chrome://new-tab-page/", "")
            if all(p.url in vacias for c in self.navegador.contexts for p in c.pages):
                self.navegador.new_browser_cdp_session().send("Browser.close")
            else:
                self.log("(El Chrome del robot queda abierto porque tiene otras pestañas.)")
        except Exception:
            pass
        try:
            self.pw.stop()
        except Exception:
            pass
        return False

    def revisar_cancelado(self):
        if self.cancelar is not None and self.cancelar.is_set():
            raise Cancelado("Se canceló la subida a My Maps.")

    # -- sesion -----------------------------------------------------------
    def cuenta_visible(self):
        """Correo de la cuenta con sesion iniciada en My Maps (None si no hay)."""
        etiqueta = self.pg.locator('[aria-label^="Cuenta de Google"]').first
        if not etiqueta.count():
            return None
        m = re.search(r"\(([^)]+@[^)]+)\)", etiqueta.get_attribute("aria-label") or "")
        return m.group(1).strip().lower() if m else None

    def asegurar_sesion(self):
        pg = self.pg
        pg.goto(INICIO, wait_until="domcontentloaded")
        esperar(lambda: self.cuenta_visible() or "accounts.google" in pg.url, 15)
        cuenta, esperada = self.cuenta_visible(), cuenta_esperada()
        if cuenta is None:
            self.log(f"Inicia sesión con {esperada or 'la cuenta de Google de los mapas'} en la ventana de Chrome "
                     f"que se abrió (te espero hasta {MINUTOS_PARA_INICIAR_SESION} minutos)...")
            pg.bring_to_front()
            fin = time.time() + MINUTOS_PARA_INICIAR_SESION * 60
            while time.time() < fin:
                self.revisar_cancelado()
                time.sleep(3)
                if "google.com/maps/d" in pg.url:
                    cuenta = self.cuenta_visible()
                    if cuenta:
                        break
                elif "accounts.google" not in pg.url and "google.com" in pg.url:
                    # Termino el inicio de sesion en otra pagina de Google.
                    pg.goto(INICIO, wait_until="domcontentloaded")
            if not cuenta:
                raise ErrorMyMaps("No se inició sesión en Google a tiempo. Vuelve a intentar cuando puedas "
                                  "iniciar sesión (puede pedir el código de verificación).")
        if esperada is None:
            anotar_cuenta(cuenta)
            self.log(f"Los mapas se crean en la cuenta {cuenta} (anotada en {ARCHIVO_CONFIG.name}).")
        elif cuenta != esperada:
            raise ErrorMyMaps(f"El Chrome del robot tiene abierta la cuenta {cuenta}, pero los mapas van en "
                              f"{esperada}.\nCierra esa sesión en el Chrome del robot e inicia con la correcta.")

    # -- crear un mapa ------------------------------------------------------
    def crear_mapa(self, titulo, descripcion, archivos):
        """Crea el mapa completo y devuelve {mid, url}. Si algo falla, el
        mapa a medias queda en self.mid_en_curso (para mandarlo a la papelera)."""
        pg = self.pg
        self.mid_en_curso = None
        pg.goto(INICIO, wait_until="domcontentloaded")
        boton = pg.locator('div[role="button"][aria-label="Crear un nuevo mapa"]').filter(visible=True).first
        boton.wait_for(state="visible", timeout=30000)
        boton.click()

        def en_editor():
            if "maps/d/edit" in pg.url and "mid=" in pg.url:
                return True
            # Aviso de Drive que a veces sale antes de crear el mapa.
            for texto in ("CREATE", "CREAR", "Crear"):
                aviso = pg.get_by_text(texto, exact=True).filter(visible=True)
                if aviso.count():
                    aviso.first.click()
                    break
            return False

        if not esperar(en_editor, 45, 1):
            raise ErrorMyMaps("My Maps no abrió el mapa nuevo.")
        self.mid_en_curso = parse_qs(urlparse(pg.url).query)["mid"][0]
        if self.al_crear:
            self.al_crear(self.mid_en_curso)
        if not esperar(lambda: self.capas(), 30):
            raise ErrorMyMaps("El mapa nuevo no terminó de abrir.")

        # Titulo y descripcion
        pg.get_by_text("Mapa sin nombre", exact=True).filter(visible=True).first.click()
        campo = pg.locator('input[type="text"][placeholder="Mapa sin nombre"]').filter(visible=True).first
        campo.wait_for(state="visible")
        campo.fill(titulo)
        pg.locator("textarea").filter(visible=True).first.fill(descripcion)
        pg.get_by_role("button", name="Guardar").filter(visible=True).first.click()
        if not esperar(lambda: pg.get_by_text(titulo, exact=True).filter(visible=True).count(), 15):
            raise ErrorMyMaps("No se pudo poner el título del mapa.")

        for i, archivo in enumerate(archivos):
            self.revisar_cancelado()
            self.log(f"   capa {i + 1}/{len(archivos)}: {archivo.name}")
            if i == 0:
                ly = self.capas()[0]
            else:
                antes = self.capas()
                pg.get_by_text("Añadir capa", exact=True).filter(visible=True).first.click()
                if not esperar(lambda: len(self.capas()) > len(antes), 20):
                    raise ErrorMyMaps("No se pudo añadir una capa nueva.")
                ly = [c for c in self.capas() if c not in antes][0]
            self.importar(ly, archivo)
            self.dar_estilo(ly)
            nombre = pg.locator(f"#{ly}-layer-header").inner_text().strip()
            if nombre != archivo.name:
                self.log(f"   Aviso: la capa quedó con el nombre \"{nombre}\".")

        self.revisar_cancelado()
        for ly in self.capas():
            casilla = pg.locator(f'#{ly}-layer-header [role="checkbox"]').first
            if casilla.get_attribute("aria-checked") == "true":
                casilla.click()
                esperar(lambda: casilla.get_attribute("aria-checked") == "false", 5)
        url = self.compartir()
        self.esperar_guardado()
        self.verificar(len(archivos))
        return {"mid": self.mid_en_curso, "url": url}

    def capas(self):
        return self.pg.evaluate("""() => Array.from(document.querySelectorAll('[id$="-layer-items-container"]'))
            .map(e => e.id.split('-')[0])""")

    def selector_de_archivos(self):
        """El cuadro "Selecciona un archivo que quieras importar" (un iframe
        de Google Picker), cuando ya muestra la opcion "Subir"."""
        for marco in self.pg.frames:
            if "docs.google.com/picker" in marco.url:
                try:
                    if marco.get_by_text("Subir", exact=True).filter(visible=True).count():
                        return marco
                except Exception:
                    pass
        return None

    def abrir_selector(self, ly):
        pg = self.pg
        enlace = pg.locator(f"#{ly}-layerview-import-link")
        if not enlace.count():
            enlace = pg.get_by_text("Importar", exact=True).filter(visible=True).last
        enlace.click()
        # La primera vez el cuadro tarda bastante (carga la lista de Drive).
        # Solo si ni siquiera existe su iframe, el clic se perdio y se repite.
        if esperar(self.selector_de_archivos, 90, 1):
            return
        if not any("docs.google.com/picker" in f.url for f in pg.frames):
            try:
                enlace.click(timeout=5000)
            except Exception:
                pass
            if esperar(self.selector_de_archivos, 60, 1):
                return
        raise ErrorMyMaps("No se abrió la ventana para importar el archivo.")

    def cuadro_visible(self, texto, segundos):
        cuadros = self.pg.locator('div[role="dialog"]').filter(has_text=texto).filter(visible=True)
        if not esperar(cuadros.count, segundos, 0.5):
            raise ErrorMyMaps(f"No apareció el cuadro \"{texto}...\" al importar.")
        return cuadros.first

    def importar(self, ly, archivo: Path):
        self.abrir_selector(ly)
        self.subir_archivo(ly, archivo)

    def subir_archivo(self, ly, archivo: Path):
        pg = self.pg
        selector = self.selector_de_archivos()
        selector.get_by_text("Subir", exact=True).filter(visible=True).first.click()
        entrada = selector.locator("input[type=file]")
        entrada.wait_for(state="attached", timeout=20000)
        entrada.set_input_files(str(archivo))

        # Columnas de la ubicacion
        # My Maps deja ocultos los cuadros de importaciones anteriores: hay
        # que usar siempre el que esta visible.
        dlg = self.cuadro_visible("Selecciona columnas", 90)
        for columna, valor in (("latitud_y", "latitude"), ("longitud_x", "longitude")):
            id_etiqueta = dlg.locator('span[id^="upload-checkbox-label-"]', has_text=re.compile(rf"^{columna}$")) \
                .first.get_attribute("id")
            numero = id_etiqueta.rsplit("-", 1)[1]
            casilla = dlg.locator(f'[role="checkbox"][aria-labelledby="{id_etiqueta}"]')
            casilla.scroll_into_view_if_needed()
            if casilla.get_attribute("aria-checked") != "true":
                casilla.click()
            pg.locator(f'[role="radio"][data-value="{valor}"]').filter(visible=True).first.click()
            if not esperar(lambda: dlg.locator(f"#upload-checkbox-mapped-{numero}").inner_text().strip(), 5):
                raise ErrorMyMaps(f"No se pudo marcar la columna {columna} al importar {archivo.name}.")
        dlg.get_by_role("button", name="Continuar").click()

        # Titulo de cada punto: la Razon
        dlg2 = self.cuadro_visible("Selecciona una columna", 30)
        dlg2.locator('[role="radio"][data-value="Razon"]').click()
        dlg2.get_by_role("button", name="Finalizar").click()

        sitios = pg.locator(f"#{ly}-layer-items-container div.un1lmc-pbTTYe-ibnC6b-KDwhZb")
        if not esperar(lambda: sitios.count() > 0 and not pg.locator('div[role="dialog"]').filter(visible=True).count(),
                       180, 1):
            abiertos = pg.locator('div[role="dialog"]').filter(visible=True)
            detalle = abiertos.first.inner_text()[:300] if abiertos.count() else "no aparecieron los clientes"
            raise ErrorMyMaps(f"No se pudo importar {archivo.name}: {detalle}")
        aviso = pg.locator(f"#{ly}-layerview-error-infobox-text")
        if aviso.count() and aviso.is_visible():
            self.log(f"   Aviso de My Maps: {aviso.inner_text().strip()}")

    def elegir_del_menu(self, texto):
        opcion = self.pg.locator("div.VIpgJd-j7LFlb-bN97Pc").filter(visible=True).filter(
            has_text=re.compile(rf"^{re.escape(texto)}$"))
        if not opcion.count():
            opcion = self.pg.get_by_text(texto, exact=True).filter(visible=True)
        opcion.last.click()

    def dar_estilo(self, ly):
        pg = self.pg
        pg.locator(f"#{ly}-layerview-stylepopup-link").click()
        popup = pg.locator("#layer-style-popup")
        popup.wait_for(state="visible")
        listas = popup.locator('[role="listbox"]').filter(visible=True)
        listas.first.click()
        self.elegir_del_menu("vendedor")
        if not esperar(lambda: listas.first.inner_text().strip() == "vendedor", 10):
            raise ErrorMyMaps("No se pudo agrupar los sitios por vendedor.")
        listas.last.click()
        self.elegir_del_menu("Razon")
        if not esperar(lambda: listas.last.inner_text().strip() == "Razon", 10):
            raise ErrorMyMaps("No se pudo poner las etiquetas con la Razon.")
        pg.keyboard.press("Escape")
        if not esperar(lambda: not popup.is_visible(), 3):
            popup.locator(".TvD9Pc, [aria-label='Cerrar']").first.click()

        categorias = pg.locator(f"#{ly}-layer-items-container div.un1lmc-pbTTYe-ibnC6b-JNdkSc")
        for i in range(categorias.count()):
            fila = categorias.nth(i)
            fila.hover()
            fila.locator("div.un1lmc-pbTTYe-ibnC6b-DyVDA").click()
            pg.locator("#stylepopup-color").wait_for(state="visible")
            r, g, b = PALETA[i % len(PALETA)]
            pg.locator(f'#stylepopup-color [aria-label="RGB ({r}, {g}, {b})"]').click()
            hexa = f"{r:02x}{g:02x}{b:02x}"
            esperar(lambda: hexa in (fila.locator("[iconcode]").first.get_attribute("iconcode") or ""), 5)
            pg.locator("#stylepopup-close").click()

    def compartir(self) -> str:
        pg = self.pg
        pg.get_by_text("Compartir", exact=True).filter(visible=True).first.click()
        enlace, busqueda = pg.locator("#linksharing"), pg.locator("#discoverability")
        enlace.wait_for(state="visible", timeout=20000)
        if not enlace.is_checked():
            enlace.click()
            if not esperar(enlace.is_checked, 10):
                raise ErrorMyMaps("No se pudo activar \"Visible para cualquier persona con el enlace\".")
        if busqueda.is_checked():
            busqueda.click()
            esperar(lambda: not busqueda.is_checked(), 10)
        dialogo = pg.locator('div[role="dialog"]').filter(has=enlace).first
        url = ""
        for _ in range(20):
            campos = dialogo.locator("input[type=text], input:not([type])")
            for i in range(campos.count()):
                valor = campos.nth(i).input_value()
                if valor.startswith("https://www.google.com/maps/d/"):
                    url = valor
            if url:
                break
            time.sleep(0.5)
        if not url:
            url = f"https://www.google.com/maps/d/edit?mid={self.mid_en_curso}&usp=sharing"
        cerrar = dialogo.get_by_role("button", name="Cerrar").filter(visible=True)
        if cerrar.count():
            cerrar.first.click()
        else:
            pg.keyboard.press("Escape")
        return url

    def esperar_guardado(self):
        pg = self.pg
        time.sleep(2)
        esperar(lambda: not pg.get_by_text(re.compile("Guardando")).filter(visible=True).count(), 30, 1)
        time.sleep(1)

    def verificar(self, cuantas):
        """Vuelve a abrir el mapa y revisa que haya quedado guardado."""
        pg = self.pg
        pg.reload(wait_until="domcontentloaded")
        if not esperar(lambda: len(self.capas()) >= cuantas, 40, 1):
            raise ErrorMyMaps(f"Al volver a abrir el mapa no están las {cuantas} capas.")
        for ly in self.capas():
            if pg.locator(f'#{ly}-layer-header [role="checkbox"]').first.get_attribute("aria-checked") == "true":
                raise ErrorMyMaps("Al volver a abrir el mapa, alguna capa quedó marcada.")

    # -- papelera -----------------------------------------------------------
    def a_la_papelera(self, mid) -> None:
        pg = self.pg
        pg.goto(f"https://www.google.com/maps/d/edit?mid={mid}&hl=es", wait_until="domcontentloaded")
        menu = pg.locator("#map-action-menu")
        if not esperar(lambda: menu.count() or "/maps/d/home" in pg.url, 30, 1):
            raise ErrorMyMaps("No se pudo abrir el mapa para mandarlo a la papelera.")
        if not menu.count():
            # My Maps manda al inicio si el mapa ya esta en la papelera (o no existe).
            return
        menu.click()
        item = pg.locator("#mapmenu-trashmap").filter(visible=True).first
        item.wait_for(state="visible")
        item.click()
        dialogo = pg.locator('div[role="dialog"]').filter(has_text="papelera").filter(visible=True).first
        dialogo.wait_for(state="visible")
        dialogo.locator('[name="trash"]').click()
        if not esperar(lambda: "/maps/d/home" in pg.url or mid not in pg.url, 30, 1):
            raise ErrorMyMaps("My Maps no confirmó que el mapa fue a la papelera.")


# ---------------------------------------------------------------------------
# Todo el proceso
# ---------------------------------------------------------------------------

def subir(carpeta: Path, fecha_entrega: date, log, cancelar=None, reemplazar=False,
          registro_ruta: Path = ARCHIVO_REGISTRO, overrides=None, hoy: date = None):
    """Crea en My Maps los mapas de la carpeta del dia. Antes manda a la
    papelera los vencidos (y, si reemplazar, los de la misma entrega).
    Devuelve los mapas creados (con url)."""
    hoy = hoy or date.today()
    archivos = archivos_mapa(carpeta)
    if not archivos:
        raise ErrorMyMaps(f"No hay archivos MAPA en:\n{carpeta}")
    cambiados = cambiados_desde_publicar(carpeta)
    if cambiados:
        raise ErrorMyMaps(texto_cambiados(cambiados) + "\n(No se subió nada.)")
    if overrides is None:
        from generar_rutas import overrides_vigentes
        log("Revisando la georreferencia de los clientes...")
        overrides, al_dia = overrides_vigentes()
        if not al_dia:
            log("Aviso: sin conexión con GitHub; se usan las ubicaciones corregidas descargadas la última vez.")
    faltan = revisar_georef(archivos, overrides)
    if faltan:
        from generar_rutas import texto_sin_georef
        raise ErrorMyMaps(f"{len(faltan)} clientes no tienen georreferencia y todos deben tenerla:\n\n"
                          f"{texto_sin_georef(faltan)}\n\nCorrígelos en \"Actualizar ruta\" (lista en rojo: \"1. "
                          "Buscar en Google Maps\" → \"2. Pegar ubicación copiada\"), vuelve a \"Generar y publicar\" "
                          "y luego sube a My Maps.\n(No se subió nada.)")

    registro = leer_registro(registro_ruta)
    entrega = fecha_entrega.isoformat()
    if activos(registro, entrega) and not reemplazar:
        raise ErrorMyMaps(f"Ya hay mapas en My Maps para el {fecha_larga(fecha_entrega)}.")

    temporal = Path(tempfile.mkdtemp(prefix="mymaps_"))
    creados = []
    try:
        copias, corregidas = preparar_copias(archivos, temporal, overrides)
        if corregidas:
            log(f"{corregidas} clientes van con la ubicación corregida.")
        planes = armar_mapas(copias, fecha_entrega)

        with Robot(log, cancelar) as robot:
            log("Revisando la sesión de Google...")
            robot.asegurar_sesion()

            viejos = por_mandar_a_papelera(registro, hoy, entrega if reemplazar else None)
            if viejos:
                log(f"Mandando a la papelera {len(viejos)} mapas (vencidos o reemplazados)...")
            for m in viejos:
                robot.revisar_cancelado()
                try:
                    robot.a_la_papelera(m["mid"])
                    m["estado"], m["papelera"] = "papelera", datetime.now().isoformat(timespec="seconds")
                except Cancelado:
                    raise
                except Exception as e:
                    log(f"   Aviso: no se pudo mandar a la papelera \"{m.get('titulo')}\" ({e}). "
                        "Se intenta de nuevo la próxima vez.")
                guardar_registro(registro, registro_ruta)

            for plan in planes:
                robot.revisar_cancelado()
                log(f"Creando {plan['titulo']} ({len(plan['archivos'])} capas)...")
                entrada = {
                    "titulo": plan["titulo"], "descripcion": plan["descripcion"],
                    "transportistas": plan["transportistas"], "capas": [a.name for a in plan["archivos"]],
                    "numero": plan["numero"], "fecha_entrega": entrega,
                    "fecha_liq": (fecha_de_carpeta(carpeta) or hoy).isoformat(),
                    "creado": datetime.now().isoformat(timespec="seconds"), "estado": "creando",
                }
                registro["mapas"].append(entrada)
                guardar_registro(registro, registro_ruta)

                def anotar_mid(mid, entrada=entrada):
                    entrada["mid"] = mid
                    guardar_registro(registro, registro_ruta)

                robot.al_crear = anotar_mid
                try:
                    resultado = robot.crear_mapa(plan["titulo"], plan["descripcion"],
                                                 plan["archivos"])
                except Exception:
                    try:
                        robot.pg.screenshot(path=str(CAPTURA_ERROR))
                        log(f"   (Captura de la pantalla del error en {CAPTURA_ERROR})")
                    except Exception:
                        pass
                    entrada["mid"] = robot.mid_en_curso
                    entrada["estado"] = "fallido" if robot.mid_en_curso else "sin_crear"
                    guardar_registro(registro, registro_ruta)
                    if robot.mid_en_curso:
                        try:
                            robot.a_la_papelera(robot.mid_en_curso)
                            entrada["estado"] = "papelera"
                            guardar_registro(registro, registro_ruta)
                        except Exception:
                            pass
                    raise
                entrada.update(resultado, estado="activo")
                guardar_registro(registro, registro_ruta)
                creados.append(entrada)
                log(f"   Listo: {resultado['url']}")
        limpiar_registro(registro, hoy)
        guardar_registro(registro, registro_ruta)
    finally:
        shutil.rmtree(temporal, ignore_errors=True)
    return creados
