#!/usr/bin/env python3
"""
SP-Migrar: descarga carpetas de SharePoint Online a un disco local o NAS.

- Conexión por terminal (código de dispositivo de Microsoft).
- Selección interactiva de bibliotecas / carpetas.
- Reanudable: lleva un registro (SQLite) de lo descargado y continúa donde quedó.
- Tolerante a fallos: cortes de red, limitación de SharePoint (429/503), caídas del NAS,
  disco lleno, tokens caducados, descargas interrumpidas (reanuda por rangos).
- Pensado para dejarlo corriendo toda la noche.

Uso:
    python sp_migrar.py                       -> menú interactivo
    python sp_migrar.py --trabajo NOMBRE      -> ejecuta/reanuda un trabajo guardado
    python sp_migrar.py --trabajo NOMBRE --solo-fallidos
"""
import argparse
import csv
import ctypes
import datetime as dt
import errno
import fnmatch
import glob
import json
import logging
import os
import re
import shutil
import signal
import sqlite3
import subprocess
import sys
import threading
import time
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, wait

try:
    import msal
    import requests
    import urllib3
except ImportError:
    print("Faltan dependencias. Ejecuta:  pip install -r requirements.txt")
    sys.exit(1)

# Empaquetado como .exe (PyInstaller): los datos van junto al ejecutable, no en la carpeta temporal.
BASE = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
DIR_TRABAJOS = os.path.join(BASE, "trabajos")
DIR_LOGS = os.path.join(BASE, "logs")
DIR_REPORTES = os.path.join(BASE, "reportes")
RUTA_CONFIG = os.path.join(BASE, "config.json")
RUTA_TOKENS = os.path.join(BASE, "token_cache.bin")

GRAPH = "https://graph.microsoft.com/v1.0"
# App pública de Microsoft "Microsoft Graph Command Line Tools" (la misma que usa Connect-MgGraph).
CLIENT_ID_DEFECTO = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
SCOPES = ["Sites.Read.All", "Files.Read.All"]
CAMPOS = "id,name,size,file,folder,package,fileSystemInfo,lastModifiedDateTime,parentReference"
BLOQUE = 1024 * 1024
BUFER_DISCO = 8 * 1024 * 1024  # escrituras grandes = menos viajes por la red hasta el NAS

CFG_DEFECTO = {
    "tenant": "organizations",
    "client_id": CLIENT_ID_DEFECTO,
    "client_secret": "",
    "hilos": 6,
    "hilos_auto": True,
    "hilos_min": 2,
    "hilos_max": 16,
    "user_agent": "NONISV|SP-Migrar|SP-Migrar/1.2",
    "reintentos_peticion": 8,
    "reintentos_archivo": 5,
    "pasadas_reintento": 3,
    "espera_entre_pasadas_seg": 300,
    "timeout_lectura_seg": 120,
    "intervalo_estado_seg": 60,
    "incluir_carpeta_sitio": True,
    "evitar_suspension": True,
    "comprimir_imagenes": True,
    "calidad_imagenes": 65,
    "ffmpeg_ruta": "",
}

# Formatos que se recomprimen conservando la extensión (para no romper enlaces ni nombres).
EXT_IMAGEN = {".jpg": "jpeg", ".jpeg": "jpeg", ".jfif": "jpeg", ".png": "png", ".webp": "webp"}

ERRORES_RED = (
    requests.exceptions.ConnectionError,
    requests.exceptions.Timeout,
    requests.exceptions.ChunkedEncodingError,
    urllib3.exceptions.HTTPError,
    ConnectionError,
)

log = logging.getLogger("spmigrar")


# --------------------------------------------------------------------------- excepciones
class Detenido(Exception):
    """El usuario pidió detener (Ctrl+C)."""


class ErrorPermanente(Exception):
    """Error que no se arregla reintentando (permisos, petición inválida...)."""


class NoEncontrado(ErrorPermanente):
    pass


class AgotadoReintentos(Exception):
    pass


class UrlCaducada(Exception):
    pass


class RangoInvalido(Exception):
    pass


# --------------------------------------------------------------------------- utilidades
RESERVADOS = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}


def limpiar(nombre):
    """Convierte un nombre de SharePoint en un nombre válido para Windows / NAS."""
    n = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", nombre).rstrip(" .")
    if not n:
        n = "_"
    if n.split(".")[0].upper() in RESERVADOS:
        n = "_" + n
    return n


def lp(ruta):
    """Ruta larga de Windows (\\\\?\\) para soportar más de 260 caracteres."""
    if os.name != "nt":
        return ruta
    ruta = os.path.abspath(ruta)
    if ruta.startswith("\\\\?\\"):
        return ruta
    if ruta.startswith("\\\\"):
        return "\\\\?\\UNC\\" + ruta[2:]
    return "\\\\?\\" + ruta


def fmt_bytes(n):
    n = float(n or 0)
    for u in ("B", "KB", "MB", "GB", "TB"):
        if abs(n) < 1024 or u == "TB":
            return f"{n:.0f} {u}" if u == "B" else f"{n:.2f} {u}"
        n /= 1024


def fmt_duracion(seg):
    seg = int(seg)
    h, r = divmod(seg, 3600)
    m, s = divmod(r, 60)
    return f"{h}h {m:02d}m {s:02d}s" if h else f"{m}m {s:02d}s"


def fecha_remota(it):
    return (it.get("fileSystemInfo") or {}).get("lastModifiedDateTime") or it.get("lastModifiedDateTime")


def a_timestamp(txt):
    if not txt:
        return None
    try:
        return dt.datetime.fromisoformat(txt.replace("Z", "+00:00")).timestamp()
    except ValueError:
        return None


def ahora():
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def coincide_filtro(ruta, filtro):
    """True si algún nivel de la ruta se llama como el filtro (sin distinguir mayúsculas; admite * y ?).
    Sin filtro, todo coincide."""
    if not filtro:
        return True
    return any(fnmatch.fnmatchcase(seg.strip().lower(), patron.lower())
               for seg in ruta.split("/") if seg for patron in filtro)


def es_disco_lleno(e):
    return getattr(e, "winerror", None) in (39, 112) or e.errno == errno.ENOSPC


def preguntar(texto, defecto=None):
    sufijo = f" [{defecto}]" if defecto not in (None, "") else ""
    r = input(f"{texto}{sufijo}: ").strip()
    return r if r else (defecto or "")


def si_no(texto, defecto=True):
    r = input(f"{texto} ({'S/n' if defecto else 's/N'}): ").strip().lower()
    return defecto if not r else r in ("s", "si", "sí", "y", "yes")


def cargar_config():
    cfg = dict(CFG_DEFECTO)
    if os.path.exists(RUTA_CONFIG):
        with open(RUTA_CONFIG, encoding="utf-8") as f:
            cfg.update(json.load(f))
    else:
        with open(RUTA_CONFIG, "w", encoding="utf-8") as f:
            json.dump(CFG_DEFECTO, f, indent=2, ensure_ascii=False)
    return cfg


def preparar_consola():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    if os.name == "nt":
        # Desactiva "Edición rápida": un clic en la consola congela el programa hasta pulsar una tecla.
        try:
            k = ctypes.windll.kernel32
            h = k.GetStdHandle(-10)
            modo = ctypes.c_uint32()
            if k.GetConsoleMode(h, ctypes.byref(modo)):
                k.SetConsoleMode(h, (modo.value & ~0x0040) | 0x0080)
        except Exception:
            pass


def evitar_suspension(activar):
    if os.name != "nt":
        return
    try:
        ES_CONTINUOUS, ES_SYSTEM_REQUIRED = 0x80000000, 0x00000001
        estado = ES_CONTINUOUS | ES_SYSTEM_REQUIRED if activar else ES_CONTINUOUS
        ctypes.windll.kernel32.SetThreadExecutionState(estado)
    except Exception:
        pass


def titulo_consola(texto):
    if os.name == "nt":
        try:
            ctypes.windll.kernel32.SetConsoleTitleW(texto)
        except Exception:
            pass


def configurar_log(nombre_trabajo=None):
    log.setLevel(logging.DEBUG)
    if not any(isinstance(h, logging.StreamHandler) and not isinstance(h, logging.FileHandler) for h in log.handlers):
        consola = logging.StreamHandler(sys.stdout)
        consola.setLevel(logging.INFO)
        consola.setFormatter(logging.Formatter("%(asctime)s %(levelname)-7s %(message)s", "%H:%M:%S"))
        log.addHandler(consola)
    if nombre_trabajo:
        for h in [h for h in log.handlers if isinstance(h, logging.FileHandler)]:
            log.removeHandler(h)
            h.close()
        os.makedirs(DIR_LOGS, exist_ok=True)
        ruta = os.path.join(DIR_LOGS, f"{nombre_trabajo}_{dt.datetime.now():%Y%m%d_%H%M%S}.log")
        fh = logging.FileHandler(ruta, encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)-7s [%(threadName)s] %(message)s"))
        log.addHandler(fh)
        return ruta


# --------------------------------------------------------------------------- autenticación
class Auth:
    def __init__(self, cfg, stop):
        self.cfg = cfg
        self.stop = stop
        self.lock = threading.Lock()
        self.cache = msal.SerializableTokenCache()
        if os.path.exists(RUTA_TOKENS):
            with open(RUTA_TOKENS, encoding="utf-8") as f:
                self.cache.deserialize(f.read())
        autoridad = f"https://login.microsoftonline.com/{cfg['tenant']}"
        self.app_only = bool(cfg.get("client_secret"))
        if self.app_only:
            self.app = msal.ConfidentialClientApplication(
                cfg["client_id"], authority=autoridad, client_credential=cfg["client_secret"], token_cache=self.cache)
        else:
            self.app = msal.PublicClientApplication(cfg["client_id"], authority=autoridad, token_cache=self.cache)

    def _guardar(self):
        if self.cache.has_state_changed:
            tmp = RUTA_TOKENS + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                f.write(self.cache.serialize())
            os.replace(tmp, RUTA_TOKENS)

    def _silencioso(self, forzar=False):
        cuentas = self.app.get_accounts()
        if not cuentas:
            return None
        res = self.app.acquire_token_silent(SCOPES, account=cuentas[0], force_refresh=forzar)
        self._guardar()
        return res.get("access_token") if res else None

    def _codigo_dispositivo(self):
        while True:
            flujo = self.app.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flujo:
                raise ErrorPermanente(f"No se pudo iniciar sesión: {flujo.get('error_description', flujo)}")
            log.warning("=" * 70)
            log.warning("INICIO DE SESIÓN REQUERIDO: " + flujo["message"])
            log.warning("=" * 70)
            res = self.app.acquire_token_by_device_flow(flujo, exit_condition=lambda f: self.stop.is_set())
            if self.stop.is_set():
                raise Detenido()
            if "access_token" in res:
                self._guardar()
                log.info(f"Sesión iniciada como {res.get('id_token_claims', {}).get('preferred_username', '?')}")
                return res["access_token"]
            log.error(f"Inicio de sesión fallido: {res.get('error_description', res.get('error'))}. Reintentando...")

    def token(self, forzar=False):
        with self.lock:
            if self.app_only:
                res = self.app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
                if "access_token" in res:
                    return res["access_token"]
                raise ErrorPermanente(f"Error de autenticación de aplicación: {res.get('error_description')}")
            tok = self._silencioso(forzar)
            if tok:
                return tok
            # Refresh token caducado o revocado: se pide un nuevo inicio de sesión y todo queda en pausa.
            return self._codigo_dispositivo()

    def usuario(self):
        if self.app_only:
            return "(aplicación)"
        cuentas = self.app.get_accounts()
        return cuentas[0].get("username") if cuentas else None


# --------------------------------------------------------------------------- cliente Graph
class Graph:
    def __init__(self, auth, cfg, stop):
        self.auth = auth
        self.cfg = cfg
        self.stop = stop
        self.s = requests.Session()
        # Microsoft prioriza el tráfico "decorado" con User-Agent al aplicar limitaciones.
        self.s.headers["User-Agent"] = cfg.get("user_agent") or CFG_DEFECTO["user_agent"]
        conexiones = max(16, int(cfg.get("hilos_max", 16)) * 2, int(cfg["hilos"]) * 2)
        adaptador = requests.adapters.HTTPAdapter(pool_connections=8, pool_maxsize=conexiones)
        self.s.mount("https://", adaptador)
        self.control = None  # ControlHilos del migrador en curso (si lo hay)
        self.pausa_hasta = 0.0
        self.pausa_lock = threading.Lock()
        self.red_lock = threading.Lock()

    def dormir(self, seg):
        if self.stop.wait(max(0, seg)):
            raise Detenido()

    def _respetar_pausa(self):
        while True:
            resta = self.pausa_hasta - time.time()
            if resta <= 0:
                return
            self.dormir(min(resta, 30))

    def _limitar(self, r):
        try:
            seg = int(float(r.headers.get("Retry-After", 30)))
        except ValueError:
            seg = 30
        seg = max(5, min(seg, 900))
        if self.control:
            self.control.limitado()
        with self.pausa_lock:
            nuevo = time.time() + seg
            if nuevo > self.pausa_hasta:
                self.pausa_hasta = nuevo
                log.warning(f"SharePoint pide bajar el ritmo (HTTP {r.status_code}). Pausa general de {seg}s.")

    @staticmethod
    def hay_internet():
        try:
            requests.head("https://graph.microsoft.com/v1.0/", timeout=10)
            return True
        except Exception:
            return False

    def esperar_red(self, err):
        """Si no hay internet, bloquea hasta que vuelva y devuelve True (no cuenta como intento)."""
        if self.hay_internet():
            return False
        with self.red_lock:
            if self.hay_internet():
                return True
            log.warning(f"Sin conexión a internet ({err.__class__.__name__}). Esperando a que vuelva la red...")
            t0 = aviso = time.time()
            espera = 5
            while not self.hay_internet():
                self.dormir(espera)
                espera = min(espera * 2, 60)
                if time.time() - aviso > 600:
                    log.warning(f"Sigue sin conexión ({fmt_duracion(time.time() - t0)}). Se continuará automáticamente.")
                    aviso = time.time()
            log.info(f"Conexión restablecida tras {fmt_duracion(time.time() - t0)}. Continuando.")
        return True

    @staticmethod
    def _mensaje(r):
        try:
            return r.json()["error"]["message"]
        except Exception:
            return (r.text or "")[:300]

    def request(self, metodo, url, *, headers=None, stream=False, preauth=False, allow_redirects=True):
        intentos = limitaciones = 0
        forzar = False
        maximo = self.cfg["reintentos_peticion"]
        while True:
            if self.stop.is_set():
                raise Detenido()
            self._respetar_pausa()
            try:
                h = dict(headers or {})
                if not preauth:
                    h["Authorization"] = "Bearer " + self.auth.token(forzar)
                r = self.s.request(metodo, url, headers=h, stream=stream, allow_redirects=allow_redirects,
                                   timeout=(30, self.cfg["timeout_lectura_seg"]))
            except ERRORES_RED as e:
                if not self.esperar_red(e):
                    intentos += 1
                    if intentos >= maximo:
                        raise AgotadoReintentos(f"Error de red persistente: {e}")
                    log.debug(f"Error de red ({e}); reintento {intentos}/{maximo}")
                    self.dormir(min(2 ** intentos, 120))
                continue

            sc = r.status_code
            if sc < 400:
                # SharePoint envía RateLimit-* cuando ya se consumió ~80% de la cuota: se frena antes del 429.
                if self.control and "RateLimit-Remaining" in r.headers:
                    self.control.precaucion()
                return r
            msg = self._mensaje(r)
            r.close()

            if sc in (429, 503, 509):
                limitaciones += 1
                if limitaciones > 60:
                    raise AgotadoReintentos(f"HTTP {sc} persistente: {msg}")
                self._limitar(r)
                continue
            if preauth and sc in (401, 403, 404, 410):
                raise UrlCaducada()
            if sc == 401:
                intentos += 1
                if intentos >= 3:
                    raise ErrorPermanente(f"HTTP 401 (sin autorización): {msg}")
                forzar = True
                self.dormir(3)
                continue
            if sc == 404:
                raise NoEncontrado(f"HTTP 404: {msg}")
            if sc == 416:
                raise RangoInvalido()
            if sc >= 500 or sc == 408:
                intentos += 1
                if intentos >= maximo:
                    raise AgotadoReintentos(f"HTTP {sc}: {msg}")
                log.debug(f"HTTP {sc} en {url[:120]}; reintento {intentos}/{maximo}")
                self.dormir(min(5 * 2 ** intentos, 300))
                continue
            raise ErrorPermanente(f"HTTP {sc}: {msg}")

    def get_json(self, url):
        return self.request("GET", url).json()

    def listar(self, drive_id, item_id, con_url=False):
        """con_url=True: sin $select, para que cada archivo traiga su @microsoft.graph.downloadUrl y la descarga
        no necesite otra petición a Graph (la mitad de peticiones = menos limitaciones)."""
        url = f"{GRAPH}/drives/{drive_id}/items/{item_id}/children?$top=999"
        if not con_url:
            url += f"&$select={CAMPOS}"
        while url:
            datos = self.get_json(url)
            yield from datos.get("value", [])
            url = datos.get("@odata.nextLink")

    def item(self, drive_id, item_id):
        return self.get_json(f"{GRAPH}/drives/{drive_id}/items/{item_id}?$select={CAMPOS}")

    def url_descarga(self, drive_id, item_id):
        r = self.request("GET", f"{GRAPH}/drives/{drive_id}/items/{item_id}/content",
                         allow_redirects=False, stream=True)
        ubicacion = r.headers.get("Location")
        r.close()
        if ubicacion:
            return ubicacion
        it = self.get_json(f"{GRAPH}/drives/{drive_id}/items/{item_id}")
        url = it.get("@microsoft.graph.downloadUrl")
        if not url:
            raise ErrorPermanente("El elemento no tiene URL de descarga")
        return url

    def resolver_sitio(self, url_sitio):
        u = urllib.parse.urlparse(url_sitio.strip())
        if not u.netloc:
            raise ErrorPermanente("URL inválida. Ejemplo: https://empresa.sharepoint.com/sites/MiSitio")
        ruta = urllib.parse.unquote(u.path)
        ruta = re.split(r"/(?:_layouts|Forms|SitePages|Lists|Shared%20Documents|Shared Documents|"
                        r"Documentos compartidos)(?:/|$)", ruta, flags=re.I)[0].rstrip("/")
        partes = [p for p in ruta.split("/") if p]
        # Prueba la ruta completa y va recortando (sirve para subsitios y URLs de bibliotecas).
        while True:
            if len(partes) >= 2:
                destino = f"{GRAPH}/sites/{u.netloc}:/{urllib.parse.quote('/'.join(partes))}"
            else:
                destino = f"{GRAPH}/sites/{u.netloc}"
            try:
                return self.get_json(destino + "?$select=id,displayName,webUrl,name")
            except NoEncontrado:
                if len(partes) <= 2:
                    raise
                partes.pop()


# --------------------------------------------------------------------------- estado persistente
class Estado:
    def __init__(self, ruta):
        self.c = sqlite3.connect(ruta, check_same_thread=False, timeout=60)
        self.c.row_factory = sqlite3.Row
        self.l = threading.Lock()
        with self.l:
            self.c.executescript("""
                PRAGMA journal_mode=WAL;
                PRAGMA synchronous=NORMAL;
                CREATE TABLE IF NOT EXISTS archivos(
                    drive_id TEXT, item_id TEXT, drive_nombre TEXT, ruta_remota TEXT, ruta_local TEXT,
                    tamano INTEGER, mtime_remoto TEXT, estado TEXT, tamano_local INTEGER,
                    intentos INTEGER DEFAULT 0, error TEXT, actualizado TEXT,
                    PRIMARY KEY(drive_id, item_id));
                CREATE TABLE IF NOT EXISTS carpetas_error(
                    drive_id TEXT, item_id TEXT, drive_nombre TEXT, ruta TEXT, error TEXT, actualizado TEXT,
                    PRIMARY KEY(drive_id, item_id));
            """)
            columnas = {f["name"] for f in self.c.execute("PRAGMA table_info(archivos)")}
            if "comprimida" not in columnas:  # trabajos creados antes de la compresión de imágenes
                self.c.execute("ALTER TABLE archivos ADD COLUMN comprimida INTEGER DEFAULT 0")
                self.c.commit()

    def get(self, drive_id, item_id):
        with self.l:
            return self.c.execute("SELECT * FROM archivos WHERE drive_id=? AND item_id=?",
                                  (drive_id, item_id)).fetchone()

    def archivo(self, drive_id, it, drive_nombre, ruta_remota, ruta_local, estado, tamano_local=None, error=None,
                comprimida=0):
        with self.l:
            self.c.execute("""
                INSERT INTO archivos(drive_id,item_id,drive_nombre,ruta_remota,ruta_local,tamano,mtime_remoto,
                                     estado,tamano_local,intentos,error,actualizado,comprimida)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(drive_id,item_id) DO UPDATE SET
                    drive_nombre=excluded.drive_nombre, ruta_remota=excluded.ruta_remota,
                    ruta_local=excluded.ruta_local, tamano=excluded.tamano, mtime_remoto=excluded.mtime_remoto,
                    estado=excluded.estado, tamano_local=excluded.tamano_local,
                    intentos=archivos.intentos + excluded.intentos, error=excluded.error,
                    actualizado=excluded.actualizado, comprimida=excluded.comprimida
            """, (drive_id, it["id"], drive_nombre, ruta_remota, ruta_local, it.get("size"), fecha_remota(it),
                  estado, tamano_local, 1 if estado == "error" else 0, error, ahora(), comprimida))
            self.c.commit()

    def imagenes_sin_comprimir(self):
        with self.l:
            filas = self.c.execute("SELECT * FROM archivos WHERE estado='ok' AND COALESCE(comprimida,0)=0 "
                                   "ORDER BY ruta_local").fetchall()
        return [f for f in filas if os.path.splitext(f["ruta_local"] or "")[1].lower() in EXT_IMAGEN]

    def marcar_comprimida(self, drive_id, item_id, tamano_local):
        with self.l:
            self.c.execute("UPDATE archivos SET comprimida=1, tamano_local=?, actualizado=? "
                           "WHERE drive_id=? AND item_id=?", (tamano_local, ahora(), drive_id, item_id))
            self.c.commit()

    def carpeta_error(self, drive_id, item_id, drive_nombre, ruta, error):
        with self.l:
            self.c.execute("INSERT OR REPLACE INTO carpetas_error VALUES(?,?,?,?,?,?)",
                           (drive_id, item_id, drive_nombre, ruta, error, ahora()))
            self.c.commit()

    def carpeta_ok(self, drive_id, item_id):
        with self.l:
            self.c.execute("DELETE FROM carpetas_error WHERE drive_id=? AND item_id=?", (drive_id, item_id))
            self.c.commit()

    def fallidos(self):
        with self.l:
            return self.c.execute("SELECT * FROM archivos WHERE estado='error' ORDER BY ruta_remota").fetchall()

    def carpetas_fallidas(self):
        with self.l:
            return self.c.execute("SELECT * FROM carpetas_error ORDER BY ruta").fetchall()

    def resumen(self):
        with self.l:
            filas = self.c.execute("SELECT estado, COUNT(*) n, SUM(COALESCE(tamano_local, tamano, 0)) b "
                                   "FROM archivos GROUP BY estado").fetchall()
        return {f["estado"]: (f["n"], f["b"] or 0) for f in filas}

    def cerrar(self):
        with self.l:
            self.c.close()


# --------------------------------------------------------------------------- estadísticas
class Stats:
    def __init__(self):
        self.lock = threading.Lock()
        self.inicio = time.time()
        self.ok = self.omitidos = self.errores = self.eliminados = self.descubiertos = 0
        self.bytes = 0
        self.comprimidas = self.ahorro = 0

    def sumar(self, **kw):
        with self.lock:
            for k, v in kw.items():
                setattr(self, k, getattr(self, k) + v)


# --------------------------------------------------------------------------- compresión de imágenes
def buscar_ffmpeg(cfg):
    candidatos = [cfg.get("ffmpeg_ruta") or "",
                  os.path.join(BASE, "ffmpeg.exe"),
                  os.path.join(BASE, "ffmpeg", "ffmpeg.exe"),
                  os.path.join(BASE, "ffmpeg", "bin", "ffmpeg.exe"),
                  shutil.which("ffmpeg") or ""]
    return next((c for c in candidatos if c and os.path.isfile(c)), None)


def es_animada(ruta):
    """WebP/PNG animados: recomprimirlos con un solo fotograma perdería la animación."""
    try:
        with open(lp(ruta), "rb") as f:
            cabecera = f.read(65536)
    except OSError:
        return False
    if cabecera[:4] == b"RIFF" and cabecera[8:12] == b"WEBP":
        return b"ANIM" in cabecera
    if cabecera[:8] == b"\x89PNG\r\n\x1a\n":
        idat = cabecera.find(b"IDAT")
        return b"acTL" in (cabecera[:idat] if idat > 0 else cabecera)
    return False


class Compresor:
    """Recomprime JPG/PNG/WebP con FFmpeg conservando nombre y extensión.

    La calidad (1-100) se aplica a JPEG (q:v de FFmpeg) y WebP; PNG es sin pérdida, solo se optimiza.
    Solo reemplaza el archivo si el resultado es más pequeño.
    """

    def __init__(self, ffmpeg, calidad):
        self.ffmpeg = ffmpeg
        self.calidad = max(1, min(100, int(calidad)))
        # Escala q:v de mjpeg: 2 (mejor) .. 31 (peor). Calidad 65 -> 10.
        self.q_jpeg = max(2, min(31, round((100 - self.calidad) / 3.5)))

    @staticmethod
    def soporta(ruta):
        return os.path.splitext(ruta)[1].lower() in EXT_IMAGEN

    def comprimir(self, ruta):
        """Devuelve el tamaño final. Lanza RuntimeError si FFmpeg falla (el original queda intacto)."""
        ext = os.path.splitext(ruta)[1].lower()
        tipo = EXT_IMAGEN[ext]
        antes = os.path.getsize(lp(ruta))
        if es_animada(ruta):
            return antes
        if tipo == "jpeg":
            opciones = ["-c:v", "mjpeg", "-q:v", str(self.q_jpeg)]
        elif tipo == "png":
            opciones = ["-c:v", "png", "-pred", "mixed", "-compression_level", "9"]
        else:
            opciones = ["-c:v", "libwebp", "-quality", str(self.calidad), "-compression_level", "6"]

        temporal = ruta + ".spcomp" + ext
        ent, sal = (ruta, temporal) if len(temporal) < 240 else (lp(ruta), lp(temporal))
        cmd = [self.ffmpeg, "-hide_banner", "-loglevel", "error", "-nostdin", "-y", "-i", ent,
               "-map", "0:v:0", "-frames:v", "1", *opciones, "-f", "image2", "-update", "1", sal]
        try:
            r = subprocess.run(cmd, capture_output=True, timeout=300,
                               creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            despues = os.path.getsize(lp(temporal)) if os.path.exists(lp(temporal)) else 0
            if r.returncode != 0 or despues == 0:
                detalle = r.stderr.decode("utf-8", "replace").strip().splitlines()
                raise RuntimeError(detalle[-1] if detalle else f"FFmpeg terminó con código {r.returncode}")
            if despues >= antes * 0.98:  # no vale la pena: se queda el original
                return antes
            os.replace(lp(temporal), lp(ruta))
            return despues
        except subprocess.TimeoutExpired:
            raise RuntimeError("FFmpeg tardó demasiado (más de 5 minutos)")
        finally:
            if os.path.exists(lp(temporal)):
                os.remove(lp(temporal))


# --------------------------------------------------------------------------- hilos adaptativos
class ControlHilos:
    """Cuántas descargas corren a la vez, ajustado solo (AIMD, como el control de congestión de TCP).

    - Cada 15 s sin limitaciones (y con descargas terminando) sube 1 hilo, hasta el máximo.
    - Ante un 429/503 baja a la mitad: así no vuelve a chocar enseguida con el límite y se evitan pausas repetidas.
    - Si SharePoint avisa que se acerca al límite (cabecera RateLimit-*), baja 1 antes de que llegue la pausa.
    """

    SUBIR_CADA = 15        # segundos entre subidas de 1 hilo
    CALMA = 60             # segundos sin limitaciones antes de volver a subir
    AGRUPAR_429 = 10       # 429 recibidos dentro de esta ventana cuentan como uno
    ESPACIO_PRECAUCION = 20

    def __init__(self, inicial, minimo, maximo, auto=True):
        self.minimo = max(1, int(minimo))
        self.maximo = max(self.minimo, int(maximo))
        self.auto = auto
        self.limite = max(self.minimo, min(self.maximo, int(inicial)))
        if not auto:
            self.maximo = self.limite
        self.activos = 0
        self.cond = threading.Condition()
        self.ultimo_cambio = time.time()
        self.ultima_limitacion = 0.0
        self.terminadas = 0

    def adquirir(self, stop):
        with self.cond:
            while self.activos >= self.limite:
                if stop.is_set():
                    raise Detenido()
                self.cond.wait(1)
            self.activos += 1

    def liberar(self):
        with self.cond:
            self.activos -= 1
            self.terminadas += 1
            ahora_ = time.time()
            if (self.auto and self.limite < self.maximo and self.terminadas >= 1
                    and ahora_ - self.ultimo_cambio >= self.SUBIR_CADA and ahora_ - self.ultima_limitacion >= self.CALMA):
                self.limite += 1
                self.ultimo_cambio = ahora_
                self.terminadas = 0
                log.info(f"[HILOS] SharePoint responde bien: subo a {self.limite} descargas simultáneas.")
            self.cond.notify_all()

    def limitado(self):
        with self.cond:
            ahora_ = time.time()
            # Varios hilos reciben el 429 casi a la vez: se cuenta como una sola limitación.
            if ahora_ - self.ultima_limitacion < self.AGRUPAR_429:
                return
            self.ultima_limitacion = ahora_
            if self.auto and self.limite > self.minimo:
                self.limite = max(self.minimo, self.limite // 2)
                self.ultimo_cambio = ahora_
                log.warning(f"[HILOS] Limitación de SharePoint: bajo a {self.limite} descargas simultáneas.")

    def precaucion(self):
        with self.cond:
            ahora_ = time.time()
            if not self.auto or ahora_ - self.ultimo_cambio < self.ESPACIO_PRECAUCION:
                return
            self.ultima_limitacion = ahora_  # no subir durante un rato
            if self.limite > self.minimo:
                self.limite -= 1
                self.ultimo_cambio = ahora_
                log.info(f"[HILOS] SharePoint avisa que se acerca al límite: bajo a {self.limite} para evitar la pausa.")


# --------------------------------------------------------------------------- migrador
class Migrador:
    def __init__(self, trabajo, graph, estado, cfg, stop, compresor=None):
        self.t = trabajo
        self.g = graph
        self.db = estado
        self.cfg = cfg
        self.stop = stop
        self.compresor = compresor
        self.filtro = trabajo.get("filtro_carpetas") or []
        self.stats = Stats()
        self.raiz = trabajo["destino"]
        self.base = (os.path.join(self.raiz, limpiar(trabajo["sitio_nombre"]))
                     if cfg["incluir_carpeta_sitio"] else self.raiz)
        self.control = ControlHilos(cfg["hilos"], cfg.get("hilos_min", 2), cfg.get("hilos_max", 16),
                                    cfg.get("hilos_auto", True))
        graph.control = self.control
        # El pool tiene tantos hilos como el máximo; ControlHilos decide cuántos descargan a la vez.
        self.pool = ThreadPoolExecutor(self.control.maximo, thread_name_prefix="desc")
        self.cupo = threading.BoundedSemaphore(self.control.maximo * 3)
        self.futuros = set()
        self.futuros_lock = threading.Lock()
        self.en_curso = set()
        self.nas_lock = threading.Lock()
        self.terminado = threading.Event()

    # ---------- rutas y disco
    def ruta_local(self, drive_nombre, ruta_remota):
        partes = [limpiar(p) for p in ruta_remota.split("/") if p]
        return os.path.join(self.base, limpiar(drive_nombre), *partes)

    def nas_disponible(self):
        try:
            if os.path.isdir(lp(self.raiz)):
                return True
            # La carpeta destino no existe (nueva o borrada): si el NAS responde, se crea; si no, está caído.
            os.makedirs(lp(self.raiz), exist_ok=True)
            return True
        except OSError:
            return False

    def error_disco_transitorio(self, e):
        """True si el error de disco es temporal (NAS caído / disco lleno) y ya se esperó a que se resuelva."""
        if es_disco_lleno(e):
            log.critical(f"¡DESTINO SIN ESPACIO! ({e}). Libera espacio en {self.raiz}; se reintentará en 5 minutos.")
            self.g.dormir(300)
            return True
        if self.nas_disponible():
            return False
        with self.nas_lock:
            if self.nas_disponible():
                return True
            log.warning(f"No se puede acceder al destino {self.raiz} ({e}). Esperando a que vuelva el NAS...")
            t0 = aviso = time.time()
            while not self.nas_disponible():
                self.g.dormir(30)
                if time.time() - aviso > 600:
                    log.warning(f"El destino sigue sin responder ({fmt_duracion(time.time() - t0)}).")
                    aviso = time.time()
            log.info(f"Destino disponible de nuevo tras {fmt_duracion(time.time() - t0)}. Continuando.")
        return True

    def crear_dir(self, ruta):
        while True:
            try:
                os.makedirs(lp(ruta), exist_ok=True)
                return True
            except OSError as e:
                if not self.error_disco_transitorio(e):
                    log.error(f"No se pudo crear la carpeta {ruta}: {e}")
                    return False

    def stat_local(self, ruta):
        while True:
            try:
                return os.stat(lp(ruta))
            except FileNotFoundError:
                return None
            except OSError as e:
                if not self.error_disco_transitorio(e):
                    return None

    # ---------- ciclo principal
    def ejecutar(self, solo_fallidos=False):
        monitor = threading.Thread(target=self._monitor, name="monitor", daemon=True)
        monitor.start()
        detenido = False
        try:
            if not self.nas_disponible():
                self.error_disco_transitorio(OSError("destino no disponible"))
            self.crear_dir(self.base)
            if self.filtro:
                log.info(f"Filtro activo: solo se descarga lo que está dentro de carpetas llamadas "
                         f"{', '.join(repr(f) for f in self.filtro)} (en cualquier nivel).")
            if not solo_fallidos:
                for sel in self.t["selecciones"]:
                    log.info(f"=== Recorriendo {sel['drive_nombre']}/{sel['ruta'] or '(biblioteca completa)'} ===")
                    self.recorrer(sel["drive_id"], sel["drive_nombre"], sel["item_id"], sel["ruta"])
                self.esperar_cola()

            for pasada in range(1, self.cfg["pasadas_reintento"] + 1):
                archivos = self.db.fallidos()
                carpetas = self.db.carpetas_fallidas()
                if not archivos and not carpetas:
                    break
                if not (solo_fallidos and pasada == 1):
                    espera = self.cfg["espera_entre_pasadas_seg"]
                    log.info(f"Hay {len(archivos)} archivo(s) y {len(carpetas)} carpeta(s) con error. "
                             f"Pasada de reintento {pasada} en {espera}s...")
                    self.g.dormir(espera)
                log.info(f"=== Pasada de reintento {pasada}/{self.cfg['pasadas_reintento']} ===")
                for c in carpetas:
                    self.recorrer(c["drive_id"], c["drive_nombre"], c["item_id"], c["ruta"])
                for f in archivos:
                    self.reintentar(f)
                self.esperar_cola()
        except Detenido:
            detenido = True
            log.warning("Deteniendo... esperando a que terminen los hilos (lo parcial se reanudará después).")
        finally:
            self.pool.shutdown(wait=True, cancel_futures=True)
            self.terminado.set()
            self._informe_estado()
        return detenido

    def recorrer(self, drive_id, drive_nombre, item_id, ruta):
        """Recorre en profundidad. Con filtro de carpetas, solo descarga lo que esté dentro de una carpeta que
        coincida (p. ej. '2026'); el resto se recorre únicamente para encontrarlas, sin crear carpetas vacías."""
        pila = [(item_id, ruta)]
        while pila:
            if self.stop.is_set():
                raise Detenido()
            iid, r = pila.pop()
            dentro = coincide_filtro(r, self.filtro)
            if dentro:
                self.crear_dir(self.ruta_local(drive_nombre, r))
            try:
                for it in self.g.listar(drive_id, iid, con_url=dentro):
                    sub = f"{r}/{it['name']}" if r else it["name"]
                    if "folder" in it or "package" in it:
                        pila.append((it["id"], sub))
                    elif "file" in it and dentro:
                        self.encolar(drive_id, drive_nombre, it, sub)
                self.db.carpeta_ok(drive_id, iid)
            except Detenido:
                raise
            except NoEncontrado:
                log.warning(f"La carpeta '{r}' ya no existe en SharePoint.")
                self.db.carpeta_ok(drive_id, iid)
            except Exception as e:
                log.error(f"No se pudo listar la carpeta '{r}': {e}")
                self.db.carpeta_error(drive_id, iid, drive_nombre, r, str(e))

    def reintentar(self, fila):
        try:
            it = self.g.item(fila["drive_id"], fila["item_id"])
        except Detenido:
            raise
        except NoEncontrado:
            log.warning(f"Ya no existe en SharePoint: {fila['ruta_remota']}")
            self.db.archivo(fila["drive_id"], {"id": fila["item_id"], "size": fila["tamano"]}, fila["drive_nombre"],
                            fila["ruta_remota"], fila["ruta_local"], "eliminado")
            return
        except Exception as e:
            log.error(f"No se pudo consultar {fila['ruta_remota']}: {e}")
            return
        self.encolar(fila["drive_id"], fila["drive_nombre"], it, fila["ruta_remota"])

    def encolar(self, drive_id, drive_nombre, it, ruta_remota):
        clave = (drive_id, it["id"])
        with self.futuros_lock:
            if clave in self.en_curso:
                return
            self.en_curso.add(clave)
        self.stats.sumar(descubiertos=1)
        while not self.cupo.acquire(timeout=1):
            if self.stop.is_set():
                with self.futuros_lock:
                    self.en_curso.discard(clave)
                raise Detenido()
        fut = self.pool.submit(self._tarea, drive_id, drive_nombre, it, ruta_remota)
        with self.futuros_lock:
            self.futuros.add(fut)

        def fin(f):
            with self.futuros_lock:
                self.futuros.discard(f)
                self.en_curso.discard(clave)
        fut.add_done_callback(fin)

    def esperar_cola(self):
        while True:
            with self.futuros_lock:
                pendientes = list(self.futuros)
            if not pendientes:
                return
            wait(pendientes, timeout=2)
            if self.stop.is_set():
                raise Detenido()

    def _tarea(self, drive_id, drive_nombre, it, ruta_remota):
        try:
            self.procesar(drive_id, drive_nombre, it, ruta_remota)
        except Detenido:
            pass
        except Exception as e:
            log.exception(f"Error inesperado con {ruta_remota}: {e}")
            self.db.archivo(drive_id, it, drive_nombre, ruta_remota, "", "error", error=repr(e))
            self.stats.sumar(errores=1)
        finally:
            self.cupo.release()

    # ---------- un archivo
    def procesar(self, drive_id, drive_nombre, it, ruta_remota):
        local = self.ruta_local(drive_nombre, ruta_remota)
        tam = it.get("size") or 0
        mt_txt = fecha_remota(it)
        mt = a_timestamp(mt_txt)

        st = self.stat_local(local)
        if st:
            previo = self.db.get(drive_id, it["id"])
            igual_registro = (previo is not None and previo["estado"] == "ok" and previo["mtime_remoto"] == mt_txt
                              and previo["tamano_local"] == st.st_size)
            igual_disco = st.st_size == tam and (mt is None or abs(st.st_mtime - mt) <= 2)
            if igual_registro or igual_disco:
                if not igual_registro:
                    self.db.archivo(drive_id, it, drive_nombre, ruta_remota, local, "ok", tamano_local=st.st_size)
                self.stats.sumar(omitidos=1)
                log.debug(f"Ya existe, se omite: {ruta_remota}")
                return

        intentos = 0
        maximo = self.cfg["reintentos_archivo"]
        while True:
            try:
                if not self.crear_dir(os.path.dirname(local)):
                    raise ErrorPermanente("No se pudo crear la carpeta de destino")
                t0 = time.time()
                self.control.adquirir(self.stop)
                try:
                    final = self.bajar(drive_id, it["id"], local, tam,
                                       it.pop("@microsoft.graph.downloadUrl", None))
                finally:
                    self.control.liberar()
                bajado = final
                comprimida, nota = 0, ""
                if self.compresor and self.compresor.soporta(local):
                    final, comprimida, nota = self.comprimir_local(local, ruta_remota)
                elif final != tam:
                    nota = f" (tamaño remoto {fmt_bytes(tam)}: normal en documentos Office)"
                if mt:
                    try:
                        os.utime(lp(local), (mt, mt))
                    except OSError as e:
                        log.debug(f"No se pudo fijar la fecha de {local}: {e}")
                self.db.archivo(drive_id, it, drive_nombre, ruta_remota, local, "ok", tamano_local=final,
                                comprimida=comprimida)
                self.stats.sumar(ok=1)
                log.info(f"OK  {ruta_remota}  [{fmt_bytes(bajado)} en {time.time() - t0:.1f}s]{nota}")
                return
            except Detenido:
                raise
            except NoEncontrado:
                log.warning(f"Eliminado en SharePoint durante la copia: {ruta_remota}")
                self.db.archivo(drive_id, it, drive_nombre, ruta_remota, local, "eliminado")
                self.stats.sumar(eliminados=1)
                return
            except ErrorPermanente as e:
                error, permanente = str(e), True
            except AgotadoReintentos as e:
                error, permanente = str(e), False
            except requests.RequestException as e:
                error, permanente = f"Error de red: {e}", False
            except OSError as e:
                if self.error_disco_transitorio(e):
                    continue
                error, permanente = f"Error de disco: {e}", getattr(e, "winerror", None) in (123, 206)

            intentos += 1
            if permanente or intentos >= maximo:
                log.error(f"FALLO {ruta_remota}: {error}")
                self.db.archivo(drive_id, it, drive_nombre, ruta_remota, local, "error", error=error)
                self.stats.sumar(errores=1)
                return
            log.warning(f"Reintento {intentos}/{maximo} de {ruta_remota}: {error}")
            self.g.dormir(min(15 * intentos, 120))

    def comprimir_local(self, local, nombre):
        """Comprime una imagen ya descargada. Devuelve (tamaño_final, estado_compresion, nota_para_log).

        estado_compresion: 1 = procesada (comprimida o se dejó el original), 2 = FFmpeg falló (original intacto).
        Un fallo de compresión nunca hace fallar la copia: la imagen original ya está a salvo.
        """
        antes = os.path.getsize(lp(local))
        try:
            despues = self.compresor.comprimir(local)
        except Exception as e:
            log.warning(f"No se pudo comprimir {nombre} (se conserva el original): {e}")
            return antes, 2, "  · sin comprimir (error de FFmpeg)"
        if despues < antes:
            self.stats.sumar(comprimidas=1, ahorro=antes - despues)
            return despues, 1, f"  · comprimida {fmt_bytes(antes)} -> {fmt_bytes(despues)} (-{100 - despues * 100 // antes}%)"
        return antes, 1, "  · ya estaba optimizada"

    def bajar(self, drive_id, item_id, destino, esperado, url_listado=None):
        """Descarga a un .sppart reanudable y lo renombra al terminar. Devuelve el tamaño final.
        url_listado: enlace de descarga que ya vino en el listado (ahorra una petición a Graph)."""
        parcial = destino + ".sppart"
        url = url_listado
        intentos = 0
        maximo = self.cfg["reintentos_peticion"]

        def fallo(texto):
            nonlocal intentos
            intentos += 1
            if intentos >= maximo:
                raise AgotadoReintentos(texto)

        while True:
            if self.stop.is_set():
                raise Detenido()
            if url is None:
                url = self.g.url_descarga(drive_id, item_id)
            try:
                desde = os.path.getsize(lp(parcial))
            except FileNotFoundError:
                desde = 0
            cab = {"Range": f"bytes={desde}-"} if desde else {}
            try:
                r = self.g.request("GET", url, headers=cab, stream=True, preauth=True)
            except UrlCaducada:
                if url is url_listado:
                    url_listado = url = None  # el enlace del listado venció (dura ~1 h): se pide uno nuevo sin contar intento
                    continue
                url = None
                fallo("Acceso denegado o URL de descarga caducada repetidamente")
                continue
            except RangoInvalido:
                os.remove(lp(parcial))
                fallo("Rango inválido repetidamente")
                continue

            modo = "ab" if r.status_code == 206 else "wb"
            largo = r.headers.get("Content-Length")
            largo = int(largo) if largo and largo.isdigit() else None
            recibidos = 0
            try:
                with open(lp(parcial), modo, buffering=BUFER_DISCO) as f:
                    for trozo in r.iter_content(BLOQUE):
                        if self.stop.is_set():
                            raise Detenido()
                        f.write(trozo)
                        recibidos += len(trozo)
                        self.stats.sumar(bytes=len(trozo))
            except ERRORES_RED as e:
                if not self.g.esperar_red(e):
                    fallo(f"Descarga interrumpida repetidamente: {e}")
                    self.g.dormir(min(5 * intentos, 60))
                log.debug(f"Descarga interrumpida ({e}); se reanuda desde {fmt_bytes(desde + recibidos)}")
                continue
            finally:
                r.close()

            if largo is not None and recibidos != largo:
                fallo(f"Descarga incompleta ({recibidos} de {largo} bytes)")
                continue
            final = os.path.getsize(lp(parcial))
            os.replace(lp(parcial), lp(destino))
            return final

    # ---------- informes
    def _linea_estado(self):
        s = self.stats
        trans = time.time() - s.inicio
        vel = s.bytes / trans if trans > 0 else 0
        with self.futuros_lock:
            cola = len(self.futuros)
        imgs = f" | imágenes comprimidas {s.comprimidas} (ahorro {fmt_bytes(s.ahorro)})" if self.compresor else ""
        return (f"OK {s.ok} | omitidos {s.omitidos} | errores {s.errores} | descubiertos {s.descubiertos} | "
                f"en cola {cola} | hilos {self.control.activos}/{self.control.limite} | "
                f"{fmt_bytes(s.bytes)} a {fmt_bytes(vel)}/s | {fmt_duracion(trans)}{imgs}")

    def _monitor(self):
        ultimo_titulo = 0
        ultimo_log = time.time()
        while not self.stop.wait(5):
            if time.time() - ultimo_titulo >= 5:
                titulo_consola(f"SP-Migrar {self.t['nombre']} | {self._linea_estado()}")
                ultimo_titulo = time.time()
            if time.time() - ultimo_log >= self.cfg["intervalo_estado_seg"]:
                log.info("[ESTADO] " + self._linea_estado())
                ultimo_log = time.time()
            if self.terminado.is_set():
                return

    def _informe_estado(self):
        log.info("=" * 70)
        log.info("RESUMEN DE ESTA EJECUCIÓN: " + self._linea_estado())
        res = self.db.resumen()
        for est, (n, b) in sorted(res.items()):
            log.info(f"  Total histórico '{est}': {n} archivos, {fmt_bytes(b)}")
        fallidos = self.db.fallidos()
        carpetas = self.db.carpetas_fallidas()
        if fallidos or carpetas:
            os.makedirs(DIR_REPORTES, exist_ok=True)
            ruta = os.path.join(DIR_REPORTES, f"{self.t['nombre']}_errores_{dt.datetime.now():%Y%m%d_%H%M%S}.csv")
            with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(["tipo", "biblioteca", "ruta", "tamaño", "intentos", "error", "fecha"])
                for c in carpetas:
                    w.writerow(["carpeta", c["drive_nombre"], c["ruta"], "", "", c["error"], c["actualizado"]])
                for a in fallidos:
                    w.writerow(["archivo", a["drive_nombre"], a["ruta_remota"], a["tamano"], a["intentos"],
                                a["error"], a["actualizado"]])
            log.warning(f"Quedan {len(fallidos)} archivo(s) y {len(carpetas)} carpeta(s) con error. Detalle: {ruta}")
        else:
            log.info("Sin errores pendientes.")
        log.info("=" * 70)


# --------------------------------------------------------------------------- navegador interactivo
def parse_indices(texto, n):
    res = set()
    for parte in re.split(r"[,\s]+", texto.strip()):
        if not parte:
            continue
        if "-" in parte:
            a, b = parte.split("-", 1)
            if a.isdigit() and b.isdigit():
                res.update(range(int(a), int(b) + 1))
        elif parte.isdigit():
            res.add(int(parte))
    return sorted(i for i in res if 1 <= i <= n)


def _ruta_legible(s):
    return f"{s['drive_nombre']}/{s['ruta']}" if s["ruta"] else f"{s['drive_nombre']} (biblioteca completa)"


def navegar(g, sitio):
    drives = g.get_json(f"{GRAPH}/sites/{sitio['id']}/drives?$select=id,name,driveType,webUrl").get("value", [])
    if not drives:
        print("El sitio no tiene bibliotecas de documentos visibles para tu usuario.")
        return None
    sel = {}
    drive = None
    pila = []  # [(item_id, nombre, tamaño)]
    cache = {}

    while True:
        if drive is None:
            ubic = f"{sitio['displayName']}  >  (bibliotecas de documentos)"
            entradas = [{"drive_id": d["id"], "drive_nombre": d["name"], "item_id": "root", "ruta": "",
                         "nombre": d["name"], "tamano": None, "hijos": None} for d in drives]
        else:
            iid = pila[-1][0] if pila else "root"
            clave = (drive["id"], iid)
            if clave not in cache:
                print("  cargando...")
                cache[clave] = list(g.listar(drive["id"], iid))
            hijos = cache[clave]
            ruta_actual = "/".join(p[1] for p in pila)
            carpetas = sorted((h for h in hijos if "folder" in h or "package" in h), key=lambda h: h["name"].lower())
            n_arch = sum(1 for h in hijos if "file" in h)
            ubic = f"{sitio['displayName']}  >  {drive['name']}" + (f"  >  {ruta_actual.replace('/', '  >  ')}" if pila else "")
            ubic += f"     ({n_arch} archivo(s) aquí)"
            entradas = [{"drive_id": drive["id"], "drive_nombre": drive["name"], "item_id": h["id"],
                         "ruta": f"{ruta_actual}/{h['name']}" if ruta_actual else h["name"], "nombre": h["name"],
                         "tamano": h.get("size"), "hijos": (h.get("folder") or {}).get("childCount")}
                        for h in carpetas]

        print("\n" + "=" * 78)
        print(f" ESTÁS EN:  {ubic}")
        print("=" * 78)
        if not entradas:
            print("   (no hay subcarpetas)")
        for i, e in enumerate(entradas, 1):
            extra = []
            if e["tamano"] is not None:
                extra.append(fmt_bytes(e["tamano"]))
            if e["hijos"] is not None:
                extra.append(f"{e['hijos']} elementos")
            elegida = "  <-- ELEGIDA" if (e["drive_id"], e["item_id"]) in sel else ""
            print(f"  {i:>3}) {e['nombre']}" + (f"   ({', '.join(extra)})" if extra else "") + elegida)
        if sel:
            print(f"\n  Carpetas elegidas para descargar: {len(sel)}")
        print("\n  ¿Qué quieres hacer?")
        print("    Escribe un NÚMERO        -> abrir esa carpeta y ver lo que tiene")
        print("    Escribe D y el número    -> DESCARGAR esa carpeta completa   (ej: D 3   o varias: D 1,4)")
        if drive is not None:
            print("    Escribe D sola           -> DESCARGAR la carpeta donde estás ahora")
        print("    Escribe A                -> volver Atrás")
        if sel:
            print("    Escribe L                -> Listo, ir a elegir la ruta de destino")
        print("    Escribe S                -> Salir sin descargar")
        cmd = input("\n  Opción: ").strip()
        c = cmd.lower()

        if c == "s":
            return None
        elif c in ("a", ".."):
            if pila:
                pila.pop()
            elif drive is not None:
                drive = None
        elif c == "l" and sel:
            return list(sel.values())
        elif c.startswith("d"):
            nums = c[1:].strip()
            nuevas = []
            if not nums:
                if drive is None:
                    print("\n  Aquí estás viendo las bibliotecas. Escribe D y el número de la biblioteca (ej: D 1).")
                    continue
                nuevas.append({"drive_id": drive["id"], "drive_nombre": drive["name"],
                               "item_id": pila[-1][0] if pila else "root",
                               "ruta": "/".join(p[1] for p in pila), "tamano": pila[-1][2] if pila else None})
            else:
                indices = parse_indices(nums, len(entradas))
                if not indices:
                    print("\n  Número no válido.")
                    continue
                for i in indices:
                    e = entradas[i - 1]
                    nuevas.append({k: e[k] for k in ("drive_id", "drive_nombre", "item_id", "ruta", "tamano")})
            for s in nuevas:
                sel[(s["drive_id"], s["item_id"])] = s

            print("\n" + "-" * 78)
            print("  Se van a descargar:")
            for s in sel.values():
                print(f"    - {_ruta_legible(s)}" + (f"   ({fmt_bytes(s['tamano'])})" if s.get("tamano") else ""))
            print("-" * 78)
            print("    1) Descargar ahora (te pediré la ruta donde guardar)")
            print("    2) Elegir otra carpeta más")
            print("    3) Quitar todo y volver a elegir")
            op = input("\n  Opción [1]: ").strip() or "1"
            if op == "1":
                return list(sel.values())
            if op == "3":
                sel.clear()
        elif c.isdigit():
            i = int(c)
            if not 1 <= i <= len(entradas):
                print("\n  Número fuera de rango.")
                continue
            e = entradas[i - 1]
            if drive is None:
                drive = next(d for d in drives if d["id"] == e["drive_id"])
                pila = []
            else:
                pila.append((e["item_id"], e["nombre"], e["tamano"]))
        elif cmd:
            print("\n  No entendí esa opción. Usa un número, D + número, A, o S.")


def depurar_seleccion(sel):
    """Quita selecciones que ya están incluidas dentro de otra carpeta seleccionada."""
    res = []
    for s in sel:
        incluida = any(o is not s and o["drive_id"] == s["drive_id"] and o["ruta"] != s["ruta"]
                       and (o["ruta"] == "" or s["ruta"].startswith(o["ruta"] + "/")) for o in sel)
        if not incluida:
            res.append(s)
    return res


def pedir_destino(total):
    print("\n" + "=" * 78)
    print(" ¿DÓNDE QUIERES GUARDAR LO DESCARGADO?")
    print("=" * 78)
    print("  Escribe la ruta completa de la carpeta destino. Ejemplos:")
    print("     \\\\192.168.1.50\\Respaldos\\SharePoint      (carpeta compartida del NAS)")
    print("     Z:\\Migracion                              (unidad de red)")
    print("     D:\\Respaldo SharePoint                     (disco local)")
    print("  Si la carpeta no existe, se crea.")
    while True:
        ruta = input("\n  Ruta destino: ").strip().strip('"')
        if not ruta:
            continue
        ruta = os.path.abspath(ruta)
        try:
            os.makedirs(lp(ruta), exist_ok=True)
            prueba = os.path.join(ruta, f".spmigrar_prueba_{os.getpid()}")
            with open(lp(prueba), "w") as f:
                f.write("ok")
            os.remove(lp(prueba))
            libre = shutil.disk_usage(ruta).free
        except OSError as e:
            print(f" No se puede escribir en {ruta}: {e}")
            continue
        print(f" Espacio libre en destino: {fmt_bytes(libre)}" + (f"  |  Tamaño a copiar: {fmt_bytes(total)}" if total else ""))
        if total and libre < total * 1.05:
            if not si_no(" ¡El espacio libre parece insuficiente! ¿Usar este destino de todas formas?", False):
                continue
        return ruta


# --------------------------------------------------------------------------- trabajos
def buscar_carpetas_filtro(g, sel, filtro):
    """Recorre lo seleccionado y devuelve las carpetas que coinciden con el filtro [(selección, ruta, tamaño)].
    No entra dentro de las que coinciden (su tamaño ya incluye todo su contenido)."""
    encontradas = []
    revisadas = 0
    for s in sel:
        if coincide_filtro(s["ruta"], filtro):
            encontradas.append((s, s["ruta"], s.get("tamano") or 0))
            continue
        pila = [(s["item_id"], s["ruta"])]
        while pila:
            iid, ruta = pila.pop()
            revisadas += 1
            if revisadas % 25 == 0:
                print(f"\r  Buscando... {revisadas} carpetas revisadas, {len(encontradas)} encontradas", end="", flush=True)
            try:
                hijos = list(g.listar(s["drive_id"], iid))
            except (ErrorPermanente, AgotadoReintentos) as e:
                print(f"\n  No se pudo revisar '{ruta}': {e}")
                continue
            for h in hijos:
                if "folder" not in h and "package" not in h:
                    continue
                sub = f"{ruta}/{h['name']}" if ruta else h["name"]
                if coincide_filtro(h["name"], filtro):
                    encontradas.append((s, sub, h.get("size") or 0))
                else:
                    pila.append((h["id"], sub))
    print(f"\r  Búsqueda terminada: {revisadas} carpetas revisadas.                         ")
    return sorted(encontradas, key=lambda e: (e[0]["drive_nombre"], e[1].lower()))


def pedir_filtro(g, sel):
    """Pregunta si solo se quieren ciertas subcarpetas (p. ej. '2026'). Devuelve (filtro, tamaño, n_encontradas)."""
    print("\n" + "=" * 78)
    print(" ¿DESCARGAR TODO O SOLO CIERTAS SUBCARPETAS?")
    print("=" * 78)
    print("  Si dentro de lo elegido solo quieres las carpetas con un nombre concreto, escríbelo.")
    print("  Se buscan en TODOS los niveles y se descarga cada una completa, respetando su ruta.")
    print("  Ejemplo: eliges 'Clientes' y escribes 2026  ->  baja  Clientes/ACME/Camión/2026/...")
    print("                                                         Clientes/ACME/Contenedor/2026/...")
    print("  Varias separadas por coma (2025, 2026). Admite * (ej: 2026*).")
    while True:
        texto = input("\n  Nombre de carpeta (Enter = descargar TODO): ").strip()
        if not texto:
            return [], 0, 0
        filtro = [p.strip() for p in texto.split(",") if p.strip()]
        print(f"\n  Buscando carpetas {', '.join(repr(f) for f in filtro)}...")
        encontradas = buscar_carpetas_filtro(g, sel, filtro)
        if not encontradas:
            print(f"  No se encontró ninguna carpeta llamada {', '.join(repr(f) for f in filtro)}. Revisa el nombre.")
            continue
        total = sum(e[2] for e in encontradas)
        print(f"\n  Se encontraron {len(encontradas)} carpeta(s), {fmt_bytes(total)} en total:")
        for s, ruta, tam in encontradas[:15]:
            print(f"    - {s['drive_nombre']}/{ruta}   ({fmt_bytes(tam)})")
        if len(encontradas) > 15:
            print(f"    ... y {len(encontradas) - 15} más")
        if si_no("\n  ¿Descargar solo estas carpetas?"):
            return filtro, total, len(encontradas)


def ruta_trabajo(nombre):
    return os.path.join(DIR_TRABAJOS, f"{nombre}.json")


def listar_trabajos():
    return sorted(os.path.splitext(os.path.basename(p))[0] for p in glob.glob(os.path.join(DIR_TRABAJOS, "*.json")))


def cargar_trabajo(nombre_o_ruta):
    ruta = nombre_o_ruta if nombre_o_ruta.lower().endswith(".json") else ruta_trabajo(nombre_o_ruta)
    if not os.path.exists(ruta):
        raise SystemExit(f"No existe el trabajo: {ruta}")
    with open(ruta, encoding="utf-8") as f:
        return json.load(f)


def elegir_trabajo():
    nombres = listar_trabajos()
    if not nombres:
        print(" No hay trabajos guardados.")
        return None
    for i, n in enumerate(nombres, 1):
        t = cargar_trabajo(n)
        print(f" {i:>3}. {n}   ->  {t['destino']}   ({len(t['selecciones'])} selección(es))")
    r = preguntar(" Número de trabajo (Enter para volver)")
    if r.isdigit() and 1 <= int(r) <= len(nombres):
        return cargar_trabajo(nombres[int(r) - 1])
    return None


def nuevo_trabajo(cfg):
    stop = threading.Event()
    auth = Auth(cfg, stop)
    g = Graph(auth, cfg, stop)
    auth.token()
    print(f"\nConectado como: {auth.usuario()}")
    while True:
        url = preguntar("\nURL del sitio de SharePoint (ej. https://empresa.sharepoint.com/sites/Ventas)")
        if not url:
            return None
        try:
            sitio = g.resolver_sitio(url)
            break
        except (ErrorPermanente, AgotadoReintentos) as e:
            print(f" No se pudo abrir el sitio: {e}")
    print(f" Sitio: {sitio['displayName']}  ({sitio.get('webUrl')})")

    sel = navegar(g, sitio)
    if not sel:
        print("Cancelado.")
        return None
    sel = depurar_seleccion(sel)
    print("\nCalculando tamaño de la selección...")
    total = 0
    for s in sel:
        if s.get("tamano") is None:
            try:
                s["tamano"] = g.item(s["drive_id"], s["item_id"]).get("size")
            except Exception:
                s["tamano"] = None
        total += s.get("tamano") or 0
        print(f"  - {s['drive_nombre']}/{s['ruta'] or '(biblioteca completa)'}   {fmt_bytes(s.get('tamano'))}")
    print(f" Total aproximado: {fmt_bytes(total)}")

    filtro, total_filtrado, encontradas = pedir_filtro(g, sel)
    if filtro:
        total = total_filtrado

    destino = pedir_destino(total)

    comprimir = False
    if buscar_ffmpeg(cfg):
        print("\n" + "=" * 78)
        print(" COMPRESIÓN DE IMÁGENES")
        print("=" * 78)
        print(f"  Las imágenes JPG, PNG y WebP se pueden comprimir con FFmpeg al {cfg['calidad_imagenes']}% de calidad")
        print("  al descargarlas (mantienen su nombre y extensión; ocupan bastante menos en el NAS).")
        comprimir = si_no("\n  ¿Comprimir las imágenes?", bool(cfg["comprimir_imagenes"]))
    else:
        print("\n  (No se encontró ffmpeg.exe: las imágenes se copiarán sin comprimir.)")

    defecto = f"{limpiar(sitio.get('name') or sitio['displayName'])}_{dt.datetime.now():%Y%m%d_%H%M}"
    nombre = re.sub(r"[^\w\-]", "_", defecto)
    trabajo = {
        "nombre": nombre,
        "sitio_url": sitio.get("webUrl", url),
        "sitio_id": sitio["id"],
        "sitio_nombre": sitio["displayName"],
        "destino": destino,
        "selecciones": [{k: s[k] for k in ("drive_id", "drive_nombre", "item_id", "ruta")} for s in sel],
        "tamano_estimado": total,
        "comprimir_imagenes": comprimir,
        "filtro_carpetas": filtro,
        "creado": ahora(),
    }
    os.makedirs(DIR_TRABAJOS, exist_ok=True)
    with open(ruta_trabajo(nombre), "w", encoding="utf-8") as f:
        json.dump(trabajo, f, indent=2, ensure_ascii=False)
    carpeta_final = os.path.join(destino, limpiar(sitio["displayName"])) if cfg["incluir_carpeta_sitio"] else destino
    print("\n" + "=" * 78)
    print(" RESUMEN")
    print("=" * 78)
    for s in sel:
        print(f"   Carpeta:  {_ruta_legible(s)}")
    if filtro:
        print(f"   Filtro:   solo carpetas {', '.join(repr(f) for f in filtro)} ({encontradas} encontradas)")
    print(f"   Destino:  {carpeta_final}")
    print(f"   Tamaño:   {fmt_bytes(total)}")
    print(f"   Imágenes: {'se comprimen al ' + str(cfg['calidad_imagenes']) + '% con FFmpeg' if comprimir else 'sin comprimir'}")
    print(f"\n   Si se interrumpe, la continúas desde el menú, opción 2 ('{nombre}').")
    return trabajo


def crear_compresor(trabajo, cfg):
    if not trabajo.get("comprimir_imagenes", cfg["comprimir_imagenes"]):
        log.info("Compresión de imágenes: desactivada para este trabajo.")
        return None
    ffmpeg = buscar_ffmpeg(cfg)
    if not ffmpeg:
        log.warning("Compresión de imágenes: NO se encontró ffmpeg.exe (ponlo junto al programa o en el PATH). "
                    "Las imágenes se copiarán sin comprimir.")
        return None
    log.info(f"Compresión de imágenes: ACTIVADA, calidad {cfg['calidad_imagenes']}% (JPG/PNG/WebP) con {ffmpeg}")
    return Compresor(ffmpeg, cfg["calidad_imagenes"])


def ejecutar_trabajo(trabajo, cfg, solo_fallidos=False):
    ruta_log = configurar_log(trabajo["nombre"])
    stop = threading.Event()
    auth = Auth(cfg, stop)
    g = Graph(auth, cfg, stop)

    pulsaciones = {"n": 0}

    def al_ctrl_c(signum, frame):
        pulsaciones["n"] += 1
        if pulsaciones["n"] == 1:
            log.warning("Ctrl+C recibido: deteniendo de forma segura (pulsa otra vez para forzar).")
            stop.set()
        else:
            os._exit(130)

    anterior = signal.signal(signal.SIGINT, al_ctrl_c)
    if cfg["evitar_suspension"]:
        evitar_suspension(True)
    try:
        log.info(f"Trabajo: {trabajo['nombre']}  |  Sitio: {trabajo['sitio_nombre']}  |  Destino: {trabajo['destino']}")
        log.info(f"Log detallado: {ruta_log}")
        while True:
            try:
                auth.token()
                break
            except ERRORES_RED as e:
                g.esperar_red(e) or g.dormir(30)
        if cfg.get("hilos_auto", True):
            log.info(f"Conectado como {auth.usuario()}. Hilos automáticos: empieza con {cfg['hilos']} "
                     f"y se ajusta solo entre {cfg.get('hilos_min', 2)} y {cfg.get('hilos_max', 16)}.")
        else:
            log.info(f"Conectado como {auth.usuario()}. Hilos de descarga fijos: {cfg['hilos']}")
        compresor = crear_compresor(trabajo, cfg)
        os.makedirs(DIR_TRABAJOS, exist_ok=True)
        db = Estado(os.path.join(DIR_TRABAJOS, f"{trabajo['nombre']}.db"))
        try:
            detenido = Migrador(trabajo, g, db, cfg, stop, compresor).ejecutar(solo_fallidos)
            pendientes = len(db.fallidos()) + len(db.carpetas_fallidas())
        finally:
            db.cerrar()
        if detenido:
            log.warning(f"Detenido. Para reanudar:  python sp_migrar.py --trabajo {trabajo['nombre']}")
            return 130
        return 2 if pendientes else 0
    except Detenido:
        log.warning("Detenido por el usuario.")
        return 130
    finally:
        evitar_suspension(False)
        signal.signal(signal.SIGINT, anterior)


def comprimir_existentes(trabajo, cfg):
    """Comprime las imágenes que un trabajo ya descargó sin comprimir (no se conecta a SharePoint)."""
    ffmpeg = buscar_ffmpeg(cfg)
    if not ffmpeg:
        print(" No se encontró ffmpeg.exe. Ponlo junto al programa (o en la carpeta 'ffmpeg') e intenta de nuevo.")
        return
    ruta_db = os.path.join(DIR_TRABAJOS, f"{trabajo['nombre']}.db")
    if not os.path.exists(ruta_db):
        print(" Este trabajo aún no ha descargado nada.")
        return
    db = Estado(ruta_db)
    try:
        filas = db.imagenes_sin_comprimir()
        if not filas:
            print(" No hay imágenes pendientes de comprimir en este trabajo.")
            return
        total = sum(f["tamano_local"] or 0 for f in filas)
        print(f"\n Imágenes sin comprimir: {len(filas)}  ({fmt_bytes(total)})  ->  calidad {cfg['calidad_imagenes']}%")
        print(" Se reemplazan en el destino por su versión comprimida (solo si queda más pequeña).")
        if not si_no(" ¿Continuar?"):
            return
        configurar_log(f"{trabajo['nombre']}_compresion")
        compresor = Compresor(ffmpeg, cfg["calidad_imagenes"])
        stats = Stats()

        def una(f):
            local = f["ruta_local"]
            if not os.path.exists(lp(local)):
                log.warning(f"No existe en el destino, se omite: {local}")
                return
            antes = os.path.getsize(lp(local))
            try:
                despues = compresor.comprimir(local)
            except Exception as e:
                log.warning(f"No se pudo comprimir {f['ruta_remota']}: {e}")
                db.marcar_comprimida(f["drive_id"], f["item_id"], antes)
                return
            mt = a_timestamp(f["mtime_remoto"])
            if mt:
                try:
                    os.utime(lp(local), (mt, mt))
                except OSError:
                    pass
            db.marcar_comprimida(f["drive_id"], f["item_id"], despues)
            stats.sumar(ok=1, comprimidas=int(despues < antes), ahorro=antes - despues)
            if despues < antes:
                log.info(f"OK  {f['ruta_remota']}  {fmt_bytes(antes)} -> {fmt_bytes(despues)} "
                         f"(-{100 - despues * 100 // max(antes, 1)}%)")

        pool = ThreadPoolExecutor(max(1, int(cfg["hilos"])), thread_name_prefix="comp")
        futuros = [pool.submit(una, f) for f in filas]
        try:
            for i, fut in enumerate(futuros, 1):
                fut.result()
                if i % 100 == 0:
                    log.info(f"[ESTADO] {i}/{len(filas)} revisadas | ahorro {fmt_bytes(stats.ahorro)}")
        except KeyboardInterrupt:
            log.warning("Detenido. Lo que falta se puede comprimir después con la misma opción.")
            pool.shutdown(wait=True, cancel_futures=True)
            return
        pool.shutdown(wait=True)
        log.info(f"Listo: {stats.comprimidas} imágenes comprimidas, ahorro total {fmt_bytes(stats.ahorro)} "
                 f"({stats.ahorro * 100 // max(total, 1)}% de {fmt_bytes(total)}).")
    finally:
        db.cerrar()


def ver_estado(trabajo):
    ruta = os.path.join(DIR_TRABAJOS, f"{trabajo['nombre']}.db")
    if not os.path.exists(ruta):
        print(" Este trabajo aún no se ha ejecutado.")
        return
    db = Estado(ruta)
    try:
        print(f"\n Trabajo {trabajo['nombre']} -> {trabajo['destino']}")
        for est, (n, b) in sorted(db.resumen().items()):
            print(f"   {est:<10} {n:>8} archivos   {fmt_bytes(b)}")
        fallidos = db.fallidos()
        carpetas = db.carpetas_fallidas()
        for c in carpetas[:20]:
            print(f"   [carpeta] {c['drive_nombre']}/{c['ruta']}: {c['error']}")
        for a in fallidos[:20]:
            print(f"   [archivo] {a['drive_nombre']}/{a['ruta_remota']}: {a['error']}")
        if len(fallidos) > 20:
            print(f"   ... y {len(fallidos) - 20} más")
    finally:
        db.cerrar()


def menu(cfg):
    while True:
        print("\n" + "=" * 50)
        print("  SP-MIGRAR  ·  SharePoint -> NAS")
        print("=" * 50)
        print("  1) DESCARGAR carpetas de SharePoint")
        print("  2) Continuar una descarga que se interrumpió")
        print("  3) Reintentar solo los fallidos de un trabajo")
        print("  4) Ver estado de un trabajo")
        print("  5) Comprimir imágenes ya descargadas (FFmpeg)")
        print("  6) Cerrar sesión de Microsoft")
        print("  0) Salir")
        op = input(" > ").strip()
        try:
            if op == "1":
                t = nuevo_trabajo(cfg)
                if t and si_no("\n   ¿Comenzar la descarga ahora?"):
                    ejecutar_trabajo(t, cfg)
            elif op in ("2", "3"):
                t = elegir_trabajo()
                if t:
                    ejecutar_trabajo(t, cfg, solo_fallidos=(op == "3"))
            elif op == "4":
                t = elegir_trabajo()
                if t:
                    ver_estado(t)
            elif op == "5":
                t = elegir_trabajo()
                if t:
                    comprimir_existentes(t, cfg)
            elif op == "6":
                if os.path.exists(RUTA_TOKENS):
                    os.remove(RUTA_TOKENS)
                print(" Sesión cerrada.")
            elif op == "0":
                return
        except KeyboardInterrupt:
            print("\n Operación cancelada.")
        except (ErrorPermanente, AgotadoReintentos) as e:
            print(f"\n Error: {e}")


def main():
    preparar_consola()
    p = argparse.ArgumentParser(description="Copia carpetas de SharePoint Online a disco local / NAS.")
    p.add_argument("--trabajo", help="nombre del trabajo guardado (o ruta a su .json) para ejecutarlo sin menú")
    p.add_argument("--solo-fallidos", action="store_true", help="solo reintentar lo que falló")
    p.add_argument("--hilos", type=int, help="descargas simultáneas al empezar (por defecto el de config.json)")
    p.add_argument("--hilos-max", type=int, help="máximo de descargas simultáneas en modo automático")
    a = p.parse_args()
    cfg = cargar_config()
    if a.hilos:
        cfg["hilos"] = a.hilos
    if a.hilos_max:
        cfg["hilos_max"] = a.hilos_max
    configurar_log()
    if a.trabajo:
        sys.exit(ejecutar_trabajo(cargar_trabajo(a.trabajo), cfg, a.solo_fallidos))
    try:
        menu(cfg)
    except (KeyboardInterrupt, EOFError):
        print()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        # Con doble clic en el .exe la ventana se cerraría sin dejar leer el error.
        import traceback
        traceback.print_exc()
        if getattr(sys, "frozen", False):
            input("\nOcurrió un error inesperado. Presiona Enter para cerrar...")
        sys.exit(1)
