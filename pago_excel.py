"""
Pago a transportes <-> el Excel bimestral de esta PC
("CAMIONES-AYUDANTES-FRECUENCIAS - <MES>-<MES> <AÑO>.xlsx", en OneDrive
...\\DOC COMPARTIDOS LOGISTICA Y CARTERA).

La web es la fuente en linea; el Excel es un espejo que tambien se puede
editar. Cada sincronizacion:
 1. Lee las pestanas de cada dia (LNS26-10, MTS27-10...). Lo que cambio en
    el Excel desde la ultima vez (comparando con data/pago_excel_foto.json)
    se manda a la web; filas escritas a mano sin ID se crean en la web (asi
    la PRIMERA sincronizacion sube todo el bimestre: carga inicial).
 2. Trae de la web todas las filas y las escribe en el Excel: un dia nuevo
    es una copia de la pestana BASE con el nombre de siempre, sus filas van
    al CONSOLIDADO y se refrescan las tablas dinamicas. La columna W
    (oculta) guarda el ID de cada fila en la web.
 3. TARIFA, BASE DE TRANSPORTISTAS y la lista de lineas (TARIFA!M) quedan
    iguales a la web (ahi se editan).
Si el mismo dato se cambio en los dos lados, gana la web y se avisa
(data/pago_avisos.log y la pagina).

Excel se maneja por COM (pywin32), no con openpyxl: asi no se pierden las
tablas dinamicas ni las validaciones, y si el libro esta abierto se escribe
en vivo en ese mismo Excel. Las columnas calculadas (ruta, valor de ruta,
nombre, % de retorno, pago, valor a cancelar) quedan con sus formulas.

Prueba sobre una copia:  python pago_excel.py --carpeta <carpeta con la copia>
"""
import json
import re
import sys
import threading
import time
import traceback
import unicodedata
from datetime import date, datetime, timedelta
from decimal import Decimal
from pathlib import Path

import pago_web
from pago_web import ErrorPago

CARPETA = Path(__file__).resolve().parent
ARCHIVO_FOTO = pago_web.CARPETA_DATOS / "pago_excel_foto.json"
ARCHIVO_AVISOS = pago_web.CARPETA_DATOS / "pago_avisos.log"
CARPETA_DOCS = Path.home() / "OneDrive" / "Documentos" / "2AVANCES DE RUTA Y MAPAS DE ENTREGA"
PATRON_ARCHIVO = "CAMIONES-AYUDANTES-FRECUENCIAS - *.xlsx"

PREFIJOS_DIA = ["LNS", "MTS", "MCL", "JVS", "VNS", "SBD", "DMG"]
# Pestanas de dia: LNS28-09, MTS29-09... (en archivos viejos tambien "MCLS23-09"
# o "MRLS 26-08"): 2 a 4 letras, dia y mes.
RE_HOJA_DIA = re.compile(r"^([A-Z]{2,4})\s*(\d{2})-(\d{2})$")
MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE",
         "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
HOJA_PLANTILLA = "BASE"
HOJA_CONSOLIDADO = "CONSOLIDADO"
HOJA_TARIFA = "TARIFA"
HOJA_TRANSPORTISTAS = "BASE DE TRANSPORTISTAS"
HOJA_RESUMEN = "RESUMEN"

FILA_ENCABEZADO = 3
PRIMERA_FILA = 4
# Columna del Excel -> campo de la web (lo que se escribe a mano o trae la carga).
COLUMNAS = {"B": "fecha_entrega", "C": "ajuste", "D": "linea", "E": "cod_frec", "H": "cod_camion",
            "J": "chofer_adicional", "K": "kgs", "L": "facturas", "M": "monto", "N": "retorno",
            "R": "sin_sancion", "T": "n_ayudantes", "U": "ayudantes", "V": "observaciones"}
COLUMNA_ID = "W"
ULTIMA_COLUMNA = 23     # W
CAMPOS_CODIGO = {"linea", "cod_frec", "cod_camion", "sin_sancion"}
# Con alguno de estos escrito, la fila de una pestana es un dato (y no una
# fila vacia de la plantilla, que ya trae MES y FECHA).
CAMPOS_DATO = ["ajuste", "cod_frec", "cod_camion", "kgs", "facturas", "monto"]
# Valor a cancelar. La formula vieja (IF anidados con <3% y >3%) da FALSO
# cuando la devolucion es exactamente 3%; esta usa <=3%.
FORMULA_S = ('=IF(C{f}=0,0,IF(OR(P{f}<=3%,R{f}="SI"),G{f},'
             'IF(OR(C{f}=1,C{f}=2),IFERROR(VLOOKUP(E{f},TARIFA!A:J,4,0),0),IFERROR(VLOOKUP(E{f},TARIFA!A:J,9,0),0))))')

_BLOQUEO = threading.Lock()


def col_num(letra: str) -> int:
    n = 0
    for c in letra:
        n = n * 26 + ord(c) - 64
    return n


def normalizar(valor) -> str:
    s = unicodedata.normalize("NFKD", str(valor or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(s.upper().split())


def a_fecha(v):
    if isinstance(v, datetime) and v.tzinfo is not None:
        # COM entrega las fechas de las celdas tal cual, marcadas como UTC.
        return v.replace(tzinfo=None).date()
    if isinstance(v, datetime):
        return v.date() if hasattr(v, "date") else None
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        for formato in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(v.strip()[:10], formato).date()
            except ValueError:
                pass
    return None


def norm(campo, v):
    """Valor comparable y guardable en JSON."""
    if v is None:
        return ""
    if campo == "fecha_entrega":
        d = a_fecha(v)
        return d.isoformat() if d else ""
    if isinstance(v, Decimal):
        v = float(v)
    if isinstance(v, bool):
        v = int(v)
    if isinstance(v, (int, float)):
        return round(float(v), 4)
    s = " ".join(str(v).split())
    if s == "":
        return ""
    if campo in CAMPOS_CODIGO:
        s = s.upper()
    if campo in ("ajuste", "kgs", "facturas", "monto", "retorno", "n_ayudantes"):
        try:
            return round(float(s.replace(",", ".")), 4)
        except ValueError:
            return s
    return s


def iguales(a, b) -> bool:
    if a == "" or b == "":
        return a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) < 0.005
    return str(a).strip().upper() == str(b).strip().upper()


def nombre_hoja(d: date) -> str:
    return f"{PREFIJOS_DIA[d.weekday()]}{d:%d-%m}"


# Cortes de pago: del 26 de un mes al 25 del siguiente, por fecha de
# entrega. Cada corte tiene su propio Excel:
# "CAMIONES-AYUDANTES-FRECUENCIAS - SEPTIEMBRE-OCTUBRE 2026.xlsx" (26/09-25/10).
# Igual que PagoTransportes.gs (_pagoCorteDe) y pago-transportes.js (corteDe).
DIA_CORTE = 25
HOY = None      # solo para pruebas: otra fecha como "hoy" (para probar el cambio de corte)


def corte_de(d: date) -> dict:
    y1, m1 = d.year, d.month
    if d.day <= DIA_CORTE:
        m1 -= 1
        if m1 == 0:
            m1, y1 = 12, y1 - 1
    y2, m2 = (y1 + 1, 1) if m1 == 12 else (y1, m1 + 1)
    return {"id": f"{y2}-{m2:02d}", "nombre": f"{MESES[m1 - 1]}-{MESES[m2 - 1]} {y2}",
            "desde": date(y1, m1, DIA_CORTE + 1), "hasta": date(y2, m2, DIA_CORTE)}


def corte_anterior(corte: dict) -> dict:
    return corte_de(corte["desde"] - timedelta(days=1))


def corte_de_archivo(ruta: Path):
    """'... - SEPTIEMBRE-OCTUBBRE 2026.xlsx' -> '2026-10' (los meses se
    reconocen por sus 3 primeras letras: un error como OCTUBBRE no importa)."""
    nombre = normalizar(ruta.stem)
    anos = re.findall(r"(20\d\d)", nombre)
    meses = [i + 1 for palabra in re.findall(r"[A-Z]+", nombre.split(" - ", 1)[-1])
             for i, m in enumerate(MESES) if len(palabra) >= 3 and palabra[:3] == m[:3]]
    return f"{int(anos[-1])}-{meses[-1]:02d}" if anos and meses else None


def avisar(texto: str) -> None:
    try:
        ARCHIVO_AVISOS.parent.mkdir(exist_ok=True)
        with open(ARCHIVO_AVISOS, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {texto}\n")
    except OSError:
        pass


# ---------------------------------------------------------------------------
# Archivos
# ---------------------------------------------------------------------------

def _orden_archivo(ruta: Path):
    """'... - SEPTIEMBRE-OCTUBBRE 2026.xlsx' -> (2026, 10): ano y ultimo mes
    del nombre (los meses se reconocen por sus 3 primeras letras, asi un
    error de tipeo como OCTUBBRE no importa). Sin meses: por fecha del archivo."""
    nombre = normalizar(ruta.stem)
    anos = re.findall(r"(20\d\d)", nombre)
    meses = [i + 1 for palabra in re.findall(r"[A-Z]+", nombre.split(" - ", 1)[-1])
             for i, m in enumerate(MESES) if len(palabra) >= 3 and palabra[:3] == m[:3]]
    if anos and meses:
        return (1, int(anos[-1]), meses[-1], ruta.stat().st_mtime)
    return (0, 0, 0, ruta.stat().st_mtime)


def archivos_bimestrales(carpeta=None):
    """Los Excel del pago, del mas nuevo al mas viejo (sin el de REVISION)."""
    config = pago_web.cargar_config()
    if carpeta is None and config.get("excel_carpeta"):
        carpeta = Path(config["excel_carpeta"])
    carpetas = [Path(carpeta)] if carpeta else [c for c in CARPETA_DOCS.glob("DOC COMPARTIDOS*") if c.is_dir()]
    archivos = []
    for c in carpetas:
        for ruta in c.glob(PATRON_ARCHIVO):
            if ruta.name.startswith("~$") or "REVISION" in normalizar(ruta.name):
                continue
            archivos.append(ruta)
    return sorted(archivos, key=_orden_archivo, reverse=True)


# ---------------------------------------------------------------------------
# Excel (COM)
# ---------------------------------------------------------------------------

def _es_excel_ocupado(e) -> bool:
    codigos = {-2147418111, -2146777998, -2147417846}   # llamada rechazada / Excel ocupado
    return getattr(e, "hresult", None) in codigos or (e.args and e.args[0] in codigos)


def ruta_local_de(full_name: str):
    """Ruta en disco de un libro abierto. Con OneDrive, Excel da una URL
    (https://d.docs.live.net/<id>/Documentos/...): se traduce a la carpeta
    local de OneDrive."""
    from urllib.parse import unquote
    if not re.match(r"^https?://", full_name or "", re.I):
        return Path(full_name)
    m = re.match(r"^https?://d\.docs\.live\.net/[^/]+/(.*)$", full_name, re.I)
    if not m:
        return None
    import os
    raiz = Path(os.environ.get("OneDriveConsumer") or os.environ.get("OneDrive") or Path.home() / "OneDrive")
    return raiz / unquote(m.group(1)).replace("/", "\\")


def misma_ruta(a: Path, b: Path) -> bool:
    try:
        return normalizar(str(Path(a).resolve())) == normalizar(str(Path(b).resolve()))
    except OSError:
        return False


class SesionExcel:
    """Abre los libros: usa el Excel del usuario si ESE MISMO archivo (misma
    ruta) ya esta abierto ahi (se escribe en vivo); si no, un Excel
    escondido propio. Mientras escribe, el calculo queda en manual (si no,
    cada celda escrita recalcula todo el libro y tarda minutos)."""

    def __init__(self, usar_excel_abierto=True):
        self.usar_excel_abierto = usar_excel_abierto
        self.app = None
        self.propios = []           # (ruta, libro) abiertos por esta sesion
        self.abiertos = {}          # ruta normalizada -> libro (abrir dos veces = el mismo)
        self.apps = []              # Excel usados (el del usuario y/o el propio)
        self.apps_pausadas = []     # (app, calculo anterior, pantalla anterior)
        self.en_pausa = False

    def _usar(self, app):
        if not any(a is app for a in self.apps):
            self.apps.append(app)

    def pausar(self):
        """Calculo manual y pantalla quieta mientras se escribe. Solo para
        escribir: leer no hace falta pausarlo (asi una sincronizacion sin
        cambios ni se nota en el Excel abierto del usuario)."""
        self.en_pausa = True
        for app in self.apps:
            if any(a is app for a, _, _ in self.apps_pausadas):
                continue
            self.apps_pausadas.append((app, app.Calculation, app.ScreenUpdating))
            app.ScreenUpdating = False
            app.Calculation = -4135     # xlCalculationManual

    def abrir(self, ruta: Path):
        clave = normalizar(str(Path(ruta).resolve()))
        if clave not in self.abiertos:
            self.abiertos[clave] = self._abrir(ruta)
            if self.en_pausa:
                self.pausar()       # un Excel que se sumo a mitad de la escritura
        return self.abiertos[clave]

    def cerrar_libro(self, ruta: Path):
        clave = normalizar(str(Path(ruta).resolve()))
        wb = self.abiertos.pop(clave, None)
        for i, (r, propio) in enumerate(self.propios):
            if propio is wb:
                try:
                    propio.Close(SaveChanges=False)
                except Exception:
                    pass
                del self.propios[i]
                break

    def _abrir(self, ruta: Path):
        import pythoncom
        import win32com.client as w
        if self.usar_excel_abierto:
            try:
                app_usuario = w.GetActiveObject("Excel.Application")
            except pythoncom.com_error:
                app_usuario = None
            if app_usuario is not None:
                for i in range(1, app_usuario.Workbooks.Count + 1):
                    wb = app_usuario.Workbooks(i)
                    local = ruta_local_de(wb.FullName)
                    if local is None or not misma_ruta(local, ruta):
                        continue
                    if wb.ReadOnly:
                        raise ErrorPago(f"\"{ruta.name}\" está abierto como solo lectura en Excel: ciérralo y "
                                        "vuelve a abrirlo normal para poder sincronizar.")
                    self._usar(app_usuario)
                    return wb
        if self.app is None:
            # Recien cerrado otro Excel, Windows a veces no deja abrir uno
            # nuevo al instante ("Error en la ejecucion de servidor").
            for intento in range(4):
                try:
                    self.app = w.DispatchEx("Excel.Application")
                    break
                except pythoncom.com_error:
                    if intento == 3:
                        raise ErrorPago("No pude abrir Excel en esta PC. Se reintenta en la próxima sincronización.")
                    time.sleep(3)
            self.app.Visible = False
            self.app.DisplayAlerts = False
            self.app.AskToUpdateLinks = False
            try:
                import win32process
                self.pid = win32process.GetWindowThreadProcessId(self.app.Hwnd)[1]
            except Exception:
                self.pid = None
        wb = self.app.Workbooks.Open(str(ruta), UpdateLinks=0)
        self.propios.append((ruta, wb))
        if wb.ReadOnly:
            raise ErrorPago(f"\"{ruta.name}\" lo tiene abierto otra persona o programa: no se puede escribir. "
                            "Se reintenta en la próxima sincronización.")
        self._usar(self.app)
        return wb

    def reanudar(self):
        """Vuelve a dejar el calculo como estaba (y recalcula)."""
        for app, calculo, pantalla in self.apps_pausadas:
            try:
                app.Calculation = calculo
                app.ScreenUpdating = pantalla
            except Exception:
                pass
        self.apps_pausadas = []
        self.en_pausa = False

    def guardar(self, wb):
        """Guardar con reintentos: si Excel esta ocupado, rechaza la llamada."""
        ultimo = None
        for _ in range(6):
            try:
                wb.Save()
                return
            except Exception as e:      # com_error o AttributeError (llamada rechazada)
                ultimo = e
                time.sleep(2)
        raise ErrorPago(f"No pude guardar el Excel (¿está ocupado?): {ultimo}")

    def cerrar(self):
        self.reanudar()
        for _, wb in self.propios:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if self.app is not None:
            try:
                self.app.Quit()
            except Exception:
                pass
        self.propios, self.app, self.apps, self.abiertos = [], None, [], {}
        self._esperar_que_termine()

    def _esperar_que_termine(self):
        """El Excel escondido de esta sesion tarda en cerrarse; si la proxima
        sincronizacion lo pide mientras tanto, COM se engancha a ese proceso
        moribundo y falla. Se espera a que termine (y si se colgo, se cierra
        ESE proceso, que es el que abrio este programa)."""
        pid, self.pid = getattr(self, "pid", None), None
        if not pid:
            return
        import gc
        gc.collect()
        try:
            import pythoncom
            import win32api
            import win32con
            import win32event
            pythoncom.CoFreeUnusedLibraries()
            proceso = win32api.OpenProcess(win32con.SYNCHRONIZE | win32con.PROCESS_TERMINATE, False, pid)
        except Exception:
            return
        try:
            if win32event.WaitForSingleObject(proceso, 15000) == win32event.WAIT_TIMEOUT:
                win32api.TerminateProcess(proceso, 1)
        finally:
            win32api.CloseHandle(proceso)


class Libro:
    """El Excel de un corte, abierto: sus pestanas de dia y donde esta cada ID."""

    def __init__(self, ruta: Path, wb):
        self.ruta, self.wb = ruta, wb
        self.corte_id = corte_de_archivo(ruta)
        self.hojas = {}         # nombre -> HojaDia
        self.modificado = False
        for i in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets(i)
            if RE_HOJA_DIA.match(ws.Name.strip()):
                self.hojas[ws.Name.strip()] = HojaDia(self, ws)

    def hoja(self, nombre):
        return self.wb.Worksheets(nombre)

    def tiene_hoja(self, nombre) -> bool:
        return any(normalizar(self.wb.Worksheets(i).Name) == normalizar(nombre)
                   for i in range(1, self.wb.Worksheets.Count + 1))


class HojaDia:
    def __init__(self, libro: Libro, ws):
        self.libro, self.ws, self.nombre = libro, ws, ws.Name.strip()
        self.leer()

    def leer(self):
        ws = self.ws
        ultima = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        ultima = max(ultima, PRIMERA_FILA)
        valores = ws.Range(ws.Cells(PRIMERA_FILA, 1), ws.Cells(ultima, ULTIMA_COLUMNA)).Value
        formulas_g = ws.Range(ws.Cells(PRIMERA_FILA, 7), ws.Cells(ultima, 7)).Formula
        if not isinstance(valores[0], tuple):
            valores = (valores,)
            formulas_g = ((formulas_g,),)
        # La fila de totales es la primera con =SUM( en VALOR DE RUTA.
        self.fila_totales = None
        for i, (f,) in enumerate(formulas_g):
            if isinstance(f, str) and f.replace(" ", "").upper().startswith("=SUM("):
                self.fila_totales = PRIMERA_FILA + i
                break
        fin = (self.fila_totales or ultima + 1) - 1
        self.filas = {}         # fila -> {"id", "datos": {campo: valor}, "con_datos"}
        for i in range(fin - PRIMERA_FILA + 1):
            celdas = valores[i]
            datos = {campo: norm(campo, celdas[col_num(letra) - 1]) for letra, campo in COLUMNAS.items()}
            ident = celdas[col_num(COLUMNA_ID) - 1]
            ident = str(ident).strip() if ident not in (None, "") else ""
            self.filas[PRIMERA_FILA + i] = {"id": ident, "datos": datos,
                                            "con_datos": any(datos[c] != "" for c in CAMPOS_DATO)}
        self.fecha = self._fecha()

    def _fecha(self):
        m = RE_HOJA_DIA.match(self.nombre)
        dia, mes = int(m.group(2)), int(m.group(3))
        for f in self.filas.values():
            d = a_fecha(f["datos"]["fecha_entrega"])
            if d and d.month == mes:
                try:
                    return date(d.year, mes, dia)
                except ValueError:
                    pass
        hoy = date.today()
        try:
            d = date(hoy.year, mes, dia)
        except ValueError:
            return None
        return d if d <= hoy + timedelta(days=60) else date(hoy.year - 1, mes, dia)

    def fila_libre(self):
        """Una fila de datos vacia (sin ID y sin datos), o None."""
        for fila, f in sorted(self.filas.items()):
            if not f["id"] and not f["con_datos"]:
                return fila
        return None


# ---------------------------------------------------------------------------
# Escritura
# ---------------------------------------------------------------------------

def _valor_excel(campo, v):
    if v in ("", None):
        return None
    if campo == "fecha_entrega":
        # Como numero de serie de Excel (la celda ya tiene formato de fecha):
        # un datetime de Python pasa por COM con la diferencia horaria y la
        # celda quedaba con 5 horas de mas.
        d = a_fecha(v)
        return float((d - date(1899, 12, 30)).days) if d else None
    return v


def _escribir(ws, fila, letra, valor):
    ws.Cells(fila, col_num(letra)).Value = valor


def escribir_fila(hoja: HojaDia, fila: int, web: dict, cambios_hechos: list):
    """Deja la fila del Excel igual a la fila de la web (solo toca las celdas
    distintas). Devuelve los datos como quedaron (para la foto)."""
    ws = hoja.ws
    actual = hoja.filas.get(fila, {"id": "", "datos": {}})
    nuevos = {}
    for letra, campo in COLUMNAS.items():
        deseado = norm(campo, web.get(campo))
        nuevos[campo] = deseado
        if not iguales(actual["datos"].get(campo, ""), deseado):
            _escribir(ws, fila, letra, _valor_excel(campo, deseado))
            cambios_hechos.append(1)
    d = a_fecha(web.get("fecha_entrega"))
    if d:
        mes = MESES[d.month - 1]
        if str(ws.Cells(fila, 1).Value or "").strip().upper() != mes:
            _escribir(ws, fila, "A", mes)
    if actual.get("id") != web["id"]:
        _escribir(ws, fila, COLUMNA_ID, web["id"])
        cambios_hechos.append(1)
    hoja.filas[fila] = {"id": web["id"], "datos": nuevos, "con_datos": True}
    hoja.libro.modificado = hoja.libro.modificado or bool(cambios_hechos)
    return nuevos


def limpiar_fila(hoja: HojaDia, fila: int):
    """Fila eliminada en la web: se vacian los datos (no se borra la fila,
    para no romper las referencias del CONSOLIDADO)."""
    ws = hoja.ws
    for letra in COLUMNAS:
        if letra != "B":
            _escribir(ws, fila, letra, None)
    _escribir(ws, fila, COLUMNA_ID, None)
    hoja.filas[fila] = {"id": "", "datos": {c: "" for c in COLUMNAS.values()}, "con_datos": False}
    hoja.libro.modificado = True


def preparar_columna_id(ws):
    celda = ws.Cells(FILA_ENCABEZADO, col_num(COLUMNA_ID))
    if str(celda.Value or "").strip() != "ID WEB":
        celda.Value = "ID WEB"
    if not ws.Columns(COLUMNA_ID).Hidden:
        ws.Columns(COLUMNA_ID).Hidden = True


def _consolidado_agregar(libro: Libro, hoja_nombre: str, filas_origen):
    """Agrega al CONSOLIDADO una fila por cada fila de la pestana (formulas
    ='HOJA'!A4 ...), despues de la ultima que ya apunta a esa pestana o, si
    no hay, despues de la ultima fila que apunta a alguna pestana."""
    if not filas_origen or not libro.tiene_hoja(HOJA_CONSOLIDADO):
        return
    ws = libro.hoja(HOJA_CONSOLIDADO)
    ultima = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    formulas = ws.Range(ws.Cells(1, 1), ws.Cells(ultima, ULTIMA_COLUMNA - 1)).Formula
    re_ref = re.compile(r"^=\+?'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)$")
    despues, de_esta = None, None
    patron = None
    for i, fila in enumerate(formulas):
        m = re_ref.match(str(fila[0] or ""))
        if not m:
            continue
        despues = i + 1
        patron = fila
        if m.group(1).strip() == hoja_nombre:
            de_esta = i + 1
    ancla = de_esta or despues or FILA_ENCABEZADO
    # Columnas que el CONSOLIDADO trae de cada pestana (todas menos las que
    # se saltean, ej. S). Un CONSOLIDADO vacio (corte nuevo) usa las de siempre.
    columnas = []
    for j, f in enumerate(patron or ()):
        m = re_ref.match(str(f or ""))
        if m:
            columnas.append((j, m.group(2)))
    if not columnas:
        columnas = [(col_num(letra) - 1, letra) for letra in "ABCDEFGHIJKLMNOPQRTUV"]
    ancho = max(j for j, _ in columnas) + 1
    n = len(filas_origen)
    ws.Range(ws.Rows(ancla + 1), ws.Rows(ancla + n)).Insert()
    nuevas = []
    for fila_origen in filas_origen:
        fila = [""] * ancho
        for j, letra in columnas:
            fila[j] = f"=+'{hoja_nombre}'!{letra}{fila_origen}"
        nuevas.append(tuple(fila))
    ws.Range(ws.Cells(ancla + 1, 1), ws.Cells(ancla + n, ancho)).Formula = tuple(nuevas)
    _consolidado_arreglar_sumas(ws)
    libro.modificado = True


def _consolidado_arreglar_sumas(ws):
    """Las sumas de abajo del CONSOLIDADO (=SUM(G4:G...) y =SUBTOTAL(9,G4:G...))
    tienen que abarcar todas las filas: al insertar justo al principio o al
    final del rango, Excel no las estira solo."""
    ultima = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    formulas = ws.Range(ws.Cells(PRIMERA_FILA, 7), ws.Cells(max(ultima, PRIMERA_FILA), 7)).Formula
    if not isinstance(formulas, tuple):
        formulas = ((formulas,),)
    fila_suma = None
    for i, (f,) in enumerate(formulas):
        f = str(f or "").replace(" ", "").upper()
        fila = PRIMERA_FILA + i
        if fila_suma is None and f.startswith("=SUM(G"):
            fila_suma = fila
            ws.Range(f"G{fila}").Formula = f"=SUM(G{PRIMERA_FILA}:G{fila - 1})"
        elif fila_suma is not None and f.startswith("=SUBTOTAL(9,G"):
            ws.Range(f"G{fila}").Formula = f"=SUBTOTAL(9,G{PRIMERA_FILA}:G{fila_suma - 1})"


def crear_archivo_corte(sesion: "SesionExcel", plantilla: Path, corte: dict, log=print) -> Path:
    """El Excel de un corte que empieza: copia del archivo del corte anterior
    (misma TARIFA, BASE, BASE DE TRANSPORTISTAS, RESUMEN...), sin sus pestanas
    de dia ni sus filas en el CONSOLIDADO. Queda en la misma carpeta."""
    destino = plantilla.parent / f"CAMIONES-AYUDANTES-FRECUENCIAS - {corte['nombre']}.xlsx"
    if destino.exists():
        return destino
    import ctypes
    import shutil
    # CopyFileW copia aunque el anterior este abierto en Excel.
    if not ctypes.windll.kernel32.CopyFileW(str(plantilla), str(destino), True):
        shutil.copyfile(plantilla, destino)
    try:
        wb = sesion.abrir(destino)
        dias = [wb.Worksheets(i) for i in range(1, wb.Worksheets.Count + 1)
                if RE_HOJA_DIA.match(wb.Worksheets(i).Name.strip())]
        nombres = {ws.Name.strip() for ws in dias}
        # Primero las filas del CONSOLIDADO (si no, quedan en #REF!): todo lo
        # que hay entre el encabezado y las sumas de abajo (las filas de cada
        # dia y las filas de ceros que separan los dias), menos la ultima fila
        # antes de las sumas, para que =SUM(G4:G4) siga teniendo un rango.
        if any(normalizar(wb.Worksheets(i).Name) == HOJA_CONSOLIDADO for i in range(1, wb.Worksheets.Count + 1)):
            ws = wb.Worksheets(HOJA_CONSOLIDADO)
            ultima = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
            formulas = ws.Range(ws.Cells(PRIMERA_FILA, 7), ws.Cells(max(ultima, PRIMERA_FILA), 7)).Formula
            if not isinstance(formulas, tuple):
                formulas = ((formulas,),)
            fila_suma = next((PRIMERA_FILA + i for i, (f,) in enumerate(formulas)
                              if str(f or "").replace(" ", "").upper().startswith("=SUM(G")), None)
            if fila_suma and fila_suma - 2 >= PRIMERA_FILA:
                ws.Range(ws.Rows(PRIMERA_FILA), ws.Rows(fila_suma - 2)).Delete()
            elif fila_suma is None:
                raise ErrorPago(f"El CONSOLIDADO de {plantilla.name} no tiene la fila de la suma (=SUM): no sé cómo vaciarlo.")
            _consolidado_arreglar_sumas(ws)
        for ws in dias:
            ws.Delete()
        sesion.guardar(wb)
    except Exception:
        # Un archivo a medio armar no debe quedar como si fuera el corte.
        sesion.cerrar_libro(destino)
        try:
            destino.unlink()
        except OSError:
            pass
        raise
    log(f"Pago a transportes: Excel del corte {corte['nombre']} creado ({destino.name}).")
    return destino


def crear_hoja_dia(libro: Libro, fecha: date, cuantas: int) -> HojaDia:
    """Copia de BASE con el nombre del dia, con 'cuantas' filas de datos,
    ubicada en orden de fecha (la mas nueva al lado del CONSOLIDADO)."""
    wb = libro.wb
    if not libro.tiene_hoja(HOJA_PLANTILLA):
        raise ErrorPago(f"El Excel \"{libro.ruta.name}\" no tiene la pestaña {HOJA_PLANTILLA} (la plantilla del día).")
    nombre = nombre_hoja(fecha)
    # Va antes de la primera pestana de un dia mas viejo; si no hay, al final de los dias.
    antes_de = None
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        h = libro.hojas.get(ws.Name.strip())
        if h and h.fecha and h.fecha < fecha:
            antes_de = ws
            break
    plantilla = wb.Worksheets(HOJA_PLANTILLA)
    antes = {wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)}
    # Copy(Before, After) por POSICION: por COM, Copy(After=hoja) se ignora y
    # la copia sale en un libro nuevo.
    if antes_de is not None:
        plantilla.Copy(antes_de)
    else:
        ancla = wb.Worksheets(HOJA_CONSOLIDADO) if libro.tiene_hoja(HOJA_CONSOLIDADO) else wb.Worksheets(1)
        ultimo_dia = max((h for h in libro.hojas.values()), key=lambda h: h.ws.Index, default=None)
        ancla = ultimo_dia.ws if ultimo_dia else ancla
        plantilla.Copy(None, ancla)
    # La copia se busca por nombre ("BASE (2)"), no por posicion: si despues
    # del ancla hay hojas ocultas (RUTA DE CONTROL...), Excel la pone despues de ellas.
    ws = next(wb.Worksheets(i) for i in range(1, wb.Worksheets.Count + 1) if wb.Worksheets(i).Name not in antes)
    ws.Name = nombre
    if ws.Visible != -1:
        ws.Visible = -1
    hoja = HojaDia(libro, ws)
    cuantas = max(cuantas, 1)
    datos = sorted(hoja.filas)
    if hoja.fila_totales is None:
        raise ErrorPago(f"La pestaña {HOJA_PLANTILLA} no tiene la fila de totales (=SUM en VALOR DE RUTA).")
    if len(datos) > cuantas:
        ws.Range(ws.Rows(PRIMERA_FILA + cuantas), ws.Rows(hoja.fila_totales - 1)).Delete()
    elif len(datos) < cuantas:
        faltan = cuantas - len(datos)
        ultima = hoja.fila_totales - 1
        ws.Range(ws.Rows(ultima), ws.Rows(ultima + faltan - 1)).Insert()
        ws.Rows(ultima + faltan).Copy(ws.Range(ws.Rows(ultima), ws.Rows(ultima + faltan - 1)))
    fin = PRIMERA_FILA + cuantas - 1
    # Plantilla limpia: sin datos viejos, valor a cancelar con la formula arreglada.
    for letra in COLUMNAS:
        ws.Range(f"{letra}{PRIMERA_FILA}:{letra}{fin}").Value = None
    ws.Range(f"S{PRIMERA_FILA}:S{fin}").Formula = FORMULA_S.format(f=PRIMERA_FILA)
    _arreglar_totales(ws, fin + 1)
    preparar_columna_id(ws)
    hoja.leer()
    libro.hojas[nombre] = hoja
    libro.modificado = True
    _consolidado_agregar(libro, nombre, list(range(PRIMERA_FILA, fin + 1)))
    return hoja


def _arreglar_totales(ws, fila_totales):
    for letra in ("G", "S"):
        celda = ws.Range(f"{letra}{fila_totales}")
        f = str(celda.Formula or "")
        if f.replace(" ", "").upper().startswith("=SUM("):
            celda.Formula = f"=SUM({letra}{PRIMERA_FILA}:{letra}{fila_totales - 1})"


def agregar_fila(hoja: HojaDia) -> int:
    """Una fila mas en una pestana que ya existe (antes de los totales), con
    las formulas de la fila de arriba y en el CONSOLIDADO."""
    ws = hoja.ws
    if hoja.fila_totales is None:
        raise ErrorPago(f"La pestaña {hoja.nombre} no tiene fila de totales: no sé dónde agregar el camión nuevo.")
    nueva = hoja.fila_totales
    ws.Rows(nueva).Insert()
    ws.Rows(nueva - 1).Copy(ws.Rows(nueva))
    for letra in list(COLUMNAS) + [COLUMNA_ID]:
        ws.Range(f"{letra}{nueva}").Value = None
    _arreglar_totales(ws, nueva + 1)
    hoja.leer()
    hoja.libro.modificado = True
    _consolidado_agregar(hoja.libro, hoja.nombre, [nueva])
    return nueva


# ---------------------------------------------------------------------------
# Catalogos (la web manda)
# ---------------------------------------------------------------------------

def tarifas_vigentes(tarifas, hoy=None):
    """La version de cada COD que rige hoy (o la mas vieja si todas son futuras)."""
    hoy = (hoy or date.today()).isoformat()
    por_cod = {}
    for t in tarifas:
        por_cod.setdefault(str(t["cod"]).strip().upper(), []).append(t)
    vigentes = {}
    for cod, versiones in por_cod.items():
        versiones.sort(key=lambda t: str(t.get("vigente_desde") or ""))
        ya = [t for t in versiones if str(t.get("vigente_desde") or "") <= hoy]
        vigentes[cod] = ya[-1] if ya else versiones[0]
    return vigentes


def escribir_catalogos(libro: Libro, cat: dict, fecha: date = None):
    wb = libro.wb
    if libro.tiene_hoja(HOJA_TARIFA) and cat.get("tarifas"):
        ws = wb.Worksheets(HOJA_TARIFA)
        ultima = max(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1, 3)
        valores = ws.Range(ws.Cells(3, 1), ws.Cells(ultima, 8)).Value
        filas = {}
        ultima_cod = 2
        for i, v in enumerate(valores):
            cod = str(v[0] or "").strip().upper()
            if cod:
                filas[cod] = (3 + i, v)
                ultima_cod = 3 + i
        for cod, t in sorted(tarifas_vigentes(cat["tarifas"], fecha).items()):
            deseado = {1: cod, 2: t.get("ruta") or None, 3: t.get("distancia") or None,
                       4: t.get("tarifa_actual"), 7: t.get("pct_nueva"), 8: t.get("pct_6")}
            if cod in filas:
                fila, v = filas[cod]
                for col, valor in deseado.items():
                    if not iguales(norm("", v[col - 1]), norm("", valor)):
                        ws.Cells(fila, col).Value = valor
                        libro.modificado = True
            else:
                ultima_cod += 1
                fila = ultima_cod
                ws.Rows(fila).Insert()
                for col, valor in deseado.items():
                    ws.Cells(fila, col).Value = valor
                ws.Range(f"E{fila}").Formula = f"=+D{fila}*(1+G{fila})"
                ws.Range(f"F{fila}").Formula = f"=+D{fila}*(1+H{fila})"
                ws.Range(f"I{fila}").Formula = f"=+D{fila}-20"
                ws.Range(f"J{fila}").Formula = f"=+E{fila}-20"
                libro.modificado = True
        # Lista de lineas: la validacion de LINEA en las pestanas usa TARIFA!M:M.
        lineas = cat.get("lineas") or []
        if lineas:
            ultima_m = max(ultima, len(lineas) + 2)
            actuales = ws.Range(ws.Cells(2, 13), ws.Cells(ultima_m, 13)).Value
            deseadas = list(lineas) + [None] * (len(actuales) - len(lineas))
            # Solo si cambio la LISTA (una linea repetida en el Excel no cuenta).
            unicas = []
            for (a,) in actuales:
                a = str(a or "").strip().upper()
                if a and a not in unicas:
                    unicas.append(a)
            if unicas != [str(l).strip().upper() for l in lineas]:
                ws.Range(ws.Cells(2, 13), ws.Cells(ultima_m, 13)).Value = tuple((d,) for d in deseadas)
                libro.modificado = True
    if libro.tiene_hoja(HOJA_TRANSPORTISTAS) and cat.get("transportistas"):
        ws = wb.Worksheets(HOJA_TRANSPORTISTAS)
        ultima = max(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1, 2)
        valores = ws.Range(ws.Cells(2, 1), ws.Cells(ultima, 2)).Value
        filas = {str(v[0] or "").strip().upper(): (2 + i, v) for i, v in enumerate(valores) if v[0]}
        siguiente = 2 + max((i for i, v in enumerate(valores) if v[0]), default=-1) + 1
        for t in cat["transportistas"]:
            tag = str(t.get("tag") or "").strip().upper()
            if not tag:
                continue
            nombre = str(t.get("nombre") or "").strip()
            if tag in filas:
                fila, v = filas[tag]
                if str(v[1] or "").strip() != nombre:
                    ws.Cells(fila, 2).Value = nombre
                    libro.modificado = True
            else:
                ws.Cells(siguiente, 1).Value = tag
                ws.Cells(siguiente, 2).Value = nombre
                siguiente += 1
                libro.modificado = True


def catalogos_del_excel(libro: Libro) -> dict:
    """Para la carga inicial (la web solo los toma si sus hojas estan vacias)."""
    wb = libro.wb
    cat = {"tarifas": [], "transportistas": [], "lineas": []}
    if libro.tiene_hoja(HOJA_TARIFA):
        ws = wb.Worksheets(HOJA_TARIFA)
        ultima = max(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1, 3)
        for v in ws.Range(ws.Cells(3, 1), ws.Cells(ultima, 8)).Value:
            if v[0] and v[3] is not None:
                cat["tarifas"].append({"cod": str(v[0]).strip(), "ruta": str(v[1] or "").strip(),
                                       "distancia": str(v[2] or "").strip(), "tarifa_actual": norm("", v[3]),
                                       "pct_nueva": norm("", v[6]), "pct_6": norm("", v[7]),
                                       "vigente_desde": "2000-01-01"})
        vistas = []
        for (v,) in ws.Range(ws.Cells(2, 13), ws.Cells(ultima, 13)).Value:
            s = str(v or "").strip().upper()
            if s and s not in vistas:
                vistas.append(s)
        cat["lineas"] = vistas
    if libro.tiene_hoja(HOJA_TRANSPORTISTAS):
        ws = wb.Worksheets(HOJA_TRANSPORTISTAS)
        ultima = max(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1, 2)
        for v in ws.Range(ws.Cells(2, 1), ws.Cells(ultima, 2)).Value:
            if v[0]:
                cat["transportistas"].append({"tag": str(v[0]).strip(), "nombre": str(v[1] or "").strip()})
    return cat


# ---------------------------------------------------------------------------
# Sincronizacion
# ---------------------------------------------------------------------------

def cargar_foto() -> dict:
    try:
        return json.loads(ARCHIVO_FOTO.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"filas": {}}


def guardar_foto(foto: dict) -> None:
    ARCHIVO_FOTO.parent.mkdir(exist_ok=True)
    temporal = ARCHIVO_FOTO.with_suffix(".tmp")
    temporal.write_text(json.dumps(foto, ensure_ascii=False), encoding="utf-8")
    temporal.replace(ARCHIVO_FOTO)


def sincronizar(log=print, carpeta=None) -> dict:
    """Una vuelta completa. Devuelve {"ok", "mensaje", "avisos"}; los errores
    esperables salen como ErrorPago."""
    if not _BLOQUEO.acquire(blocking=False):
        return {"ok": False, "mensaje": "Ya hay una sincronización en curso.", "avisos": []}
    import pythoncom
    pythoncom.CoInitialize()
    # Con --carpeta (pruebas sobre una copia) nunca se usa el Excel abierto del usuario.
    sesion = SesionExcel(usar_excel_abierto=carpeta is None)
    try:
        return _sincronizar(sesion, log, carpeta)
    except pythoncom.com_error as e:
        if _es_excel_ocupado(e):
            raise ErrorPago("Excel está ocupado (¿hay una celda en edición o un cuadro abierto?). "
                            "Se reintenta en la próxima sincronización.")
        raise
    finally:
        sesion.cerrar()
        pythoncom.CoUninitialize()
        _BLOQUEO.release()


def _sincronizar(sesion: SesionExcel, log, carpeta) -> dict:
    todos = archivos_bimestrales(carpeta)
    if not todos:
        raise ErrorPago("No encontré el Excel CAMIONES-AYUDANTES-FRECUENCIAS en la carpeta "
                        "DOC COMPARTIDOS LOGÍSTICA Y CARTERA.")
    foto = cargar_foto()
    fotos = foto.setdefault("filas", {})
    config = pago_web.cargar_config()
    hoy = HOY or date.today()
    actual = corte_de(hoy)
    anterior = corte_anterior(actual)
    # Desde que corte se suben a la web las filas escritas en el Excel sin ID:
    # el del archivo mas nuevo cuando se sincronizo por primera vez (los
    # archivos de cortes viejos no se cargan enteros).
    primer_corte = config.get("primer_corte") or corte_de_archivo(todos[0]) or actual["id"]
    por_corte = {}
    for a in todos:                         # del mas nuevo al mas viejo
        cid = corte_de_archivo(a)
        if cid and cid not in por_corte:
            por_corte[cid] = a
    # Se abren el Excel del corte actual y el del anterior (devoluciones que
    # llegan despues del 25), y los que ya tienen filas sincronizadas.
    con_filas = {f.get("_archivo") for f in fotos.values()}
    abrir = [por_corte[c["id"]] for c in (actual, anterior) if c["id"] in por_corte and c["id"] >= primer_corte]
    abrir += [a for a in todos if a.name in con_filas and a not in abrir]
    pendientes = pago_web.enviar_pendientes(log)
    libros = [Libro(r, sesion.abrir(r)) for r in abrir]
    rechazadas = foto.setdefault("rechazadas", {})

    # 1) Excel -> web
    ubicacion = {}          # id -> (HojaDia, fila)
    cambios, nuevas, por_tmp, vistas = [], [], {}, set()
    for libro in libros:
        for hoja in libro.hojas.values():
            for fila, f in hoja.filas.items():
                if f["id"]:
                    ubicacion[f["id"]] = (hoja, fila)
                    antes = fotos.get(f["id"])
                    if antes:
                        for campo, valor in f["datos"].items():
                            if not iguales(valor, antes.get(campo, "")):
                                cambios.append({"id": f["id"], "campo": campo, "valor": valor,
                                                "base": antes.get(campo, "")})
                elif f["con_datos"] and (libro.corte_id or "") >= primer_corte:
                    tmp = f"XLS|{hoja.nombre}|{fila}"
                    datos = dict(f["datos"])
                    if not datos.get("fecha_entrega") and hoja.fecha:
                        datos["fecha_entrega"] = hoja.fecha.isoformat()
                    firma = json.dumps(datos, sort_keys=True, default=str)
                    vistas.add(tmp)
                    # Una fila que la web ya rechazo (ej. corte cerrado) no se
                    # vuelve a mandar mientras no cambie.
                    if rechazadas.get(tmp) == firma:
                        continue
                    nuevas.append(dict(datos, tmp=tmp, clave_carga=tmp, orden=fila - FILA_ENCABEZADO))
                    por_tmp[tmp] = (hoja, fila, firma)
    for tmp in list(rechazadas):
        if tmp not in vistas:
            del rechazadas[tmp]
    # Filas que estaban en el Excel y ya no estan: las borraron a mano.
    eliminar = []
    for ident, antes in fotos.items():
        if ident in ubicacion:
            continue
        if any(l.ruta.name == antes.get("_archivo") for l in libros):
            eliminar.append({"id": ident, "base": {k: v for k, v in antes.items() if not k.startswith("_")}})

    fechas = [h.fecha for l in libros for h in l.hojas.values() if h.fecha]
    desde = min(fechas + [anterior["desde"]]).isoformat()
    pedido = {"cambios": cambios, "nuevas": nuevas, "eliminar": eliminar, "desde": desde}
    if not config.get("catalogos_subidos") and libros:
        pedido["catalogos_excel"] = catalogos_del_excel(libros[0])
    if cambios or nuevas or eliminar:
        log(f"Pago a transportes: subiendo del Excel {len(cambios)} cambios, {len(nuevas)} filas nuevas"
            + (f", {len(eliminar)} borradas" if eliminar else "") + "...")
    r = pago_web.llamar("pago_sync", pedido, timeout=180)
    config = pago_web.cargar_config()
    config["primer_corte"] = primer_corte
    if "catalogos_excel" in pedido:
        config["catalogos_subidos"] = True
    pago_web.guardar_config(config)
    foto["cortes_cerrados"] = sorted(r.get("cortes_cerrados") or [])

    sesion.pausar()
    avisos = []
    for c in r.get("conflictos", []):
        if c.get("corte_cerrado"):
            texto = (f"{c.get('fecha_entrega')} {c.get('cod_camion')}: el corte {c['corte_cerrado']} está cerrado; "
                     f"el cambio del Excel en {c.get('campo')} (\"{c.get('valor_excel')}\") no se aplicó.")
        else:
            texto = (f"{c.get('fecha_entrega')} {c.get('cod_camion')}: en el Excel se puso \"{c.get('valor_excel')}\" "
                     f"en {c.get('campo')}, pero en la web ya estaba \"{c.get('valor_web')}\". Se dejó el de la web.")
        avisos.append(texto)
        avisar(texto)
    for e in r.get("errores", []):
        if e.get("tmp") in por_tmp:
            rechazadas[e["tmp"]] = por_tmp[e["tmp"]][2]
        texto = f"No se pudo subir un dato del Excel ({e.get('campo') or e.get('tmp') or ''}): {e.get('error')}"
        avisos.append(texto)
        avisar(texto)

    # IDs de las filas nuevas (escribir_fila los anota en la columna W).
    for tmp, ident in (r.get("ids") or {}).items():
        if tmp in por_tmp:
            ubicacion[ident] = por_tmp[tmp][:2]

    # 2) web -> Excel
    hechos = []
    filas_web = sorted(r.get("filas", []), key=lambda o: (str(o.get("fecha_entrega")), float(o.get("orden") or 0)))
    eliminadas = [o for o in filas_web if o.get("eliminado") == "SI"]
    vivas = [o for o in filas_web if o.get("eliminado") != "SI"]
    limpiadas = 0
    for o in eliminadas:
        if o["id"] in ubicacion:
            hoja, fila = ubicacion.pop(o["id"])
            limpiar_fila(hoja, fila)
            limpiadas += 1
        fotos.pop(o["id"], None)
    sin_lugar = {}
    for o in vivas:
        if o["id"] in ubicacion:
            hoja, fila = ubicacion[o["id"]]
            fotos[o["id"]] = dict(escribir_fila(hoja, fila, o, hechos), _archivo=hoja.libro.ruta.name)
        else:
            sin_lugar.setdefault(o["fecha_entrega"], []).append(o)

    def libro_del_corte(corte):
        """El Excel del corte (abierto, o se abre, o se crea copiando el mas nuevo)."""
        libro = next((l for l in libros if l.corte_id == corte["id"]), None)
        if libro:
            return libro
        ruta = por_corte.get(corte["id"])
        if ruta is None:
            plantilla = por_corte[max(por_corte)] if por_corte else todos[0]
            ruta = crear_archivo_corte(sesion, plantilla, corte, log)
            por_corte[corte["id"]] = ruta
        libro = Libro(ruta, sesion.abrir(ruta))
        libros.append(libro)
        return libro

    nuevas_hojas = 0
    for fecha_txt, filas in sorted(sin_lugar.items()):
        fecha = a_fecha(fecha_txt)
        if not fecha:
            continue
        corte = corte_de(fecha)
        if corte["id"] < primer_corte:
            continue        # dias de cortes viejos: no se arman pestanas en esos archivos
        libro = libro_del_corte(corte)
        nombre = nombre_hoja(fecha)
        hoja = libro.hojas.get(nombre)
        if hoja is None:
            hoja = crear_hoja_dia(libro, fecha, len(filas))
            nuevas_hojas += 1
            log(f"Pago a transportes: pestaña {nombre} creada en {libro.ruta.name} ({len(filas)} camiones).")
        for o in filas:
            fila = hoja.fila_libre() or agregar_fila(hoja)
            fotos[o["id"]] = dict(escribir_fila(hoja, fila, o, hechos), _archivo=hoja.libro.ruta.name)
            ubicacion[o["id"]] = (hoja, fila)
    # Filas con ID que la web ya no devuelve (anteriores a "desde"): se dejan como estan.

    for libro in libros:
        for hoja in libro.hojas.values():
            preparar_columna_id(hoja.ws)
    # TARIFA y transportistas solo en el Excel del corte actual (los cortes
    # anteriores quedan con las tarifas con que se pagaron).
    activo = next((l for l in libros if l.corte_id == actual["id"]), None)
    if activo:
        # Una sola tarifa por COD en el Excel: la que rige hoy (o desde el
        # inicio del corte, si se programo un cambio para el corte nuevo).
        escribir_catalogos(activo, r, max(hoy, actual["desde"]))
    sesion.reanudar()       # calculo automatico otra vez (recalcula antes de guardar)
    # RESUMEN no se toca: refrescar sus tablas dinamicas por programa las
    # agranda encima de la tabla de detalle de al lado (Excel no pregunta con
    # DisplayAlerts apagado). Se refrescan a mano, como siempre.
    for libro in libros:
        if libro.modificado:
            sesion.guardar(libro.wb)
    guardar_foto(foto)

    partes = []
    if r.get("nuevas"):
        partes.append(f"{r['nuevas']} filas del Excel subidas")
    if cambios:
        partes.append(f"{len(cambios)} cambios del Excel subidos")
    if eliminar:
        partes.append(f"{len(eliminar)} filas borradas en el Excel")
    if limpiadas:
        partes.append(f"{limpiadas} filas quitadas en la web vaciadas en el Excel")
    if nuevas_hojas:
        partes.append(f"{nuevas_hojas} pestañas nuevas")
    if hechos:
        partes.append(f"{len(hechos)} celdas actualizadas en el Excel")
    if pendientes:
        partes.append(f"{pendientes} días sin enviar todavía")
    mensaje = ("Sincronizado: " + ", ".join(partes) + ".") if partes else "Sincronizado: sin cambios."
    log("Pago a transportes: " + mensaje)
    archivo = activo or (libros[0] if libros else None)
    return {"ok": True, "mensaje": mensaje, "avisos": avisos, "archivo": archivo.ruta.name if archivo else "",
            "corte": actual["nombre"]}


if __name__ == "__main__":
    carpeta = None
    if "--carpeta" in sys.argv:
        carpeta = Path(sys.argv[sys.argv.index("--carpeta") + 1])
    try:
        print(sincronizar(print, carpeta))
    except ErrorPago as e:
        print("ERROR:", e)
    except Exception:
        traceback.print_exc()
