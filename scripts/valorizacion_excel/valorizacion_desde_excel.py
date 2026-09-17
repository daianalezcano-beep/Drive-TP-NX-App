# ============================================================
# TOURPLAN NX — CARGA DE TARIFA POR PERÍODO EN UN OPTION
# Copia adaptada para tp-nx-app (app local). El original para
# Google Colab vive sin cambios en el repo
# Valorizacion-EX-TF-dsd-servicio-madre. Mismos cambios que el
# resto de los scripts vendorizados: config por variables de
# entorno, Chrome delegado a common/chrome_bootstrap.py, sin
# --headless/puerto fijo, perfil de Chrome aislado, log en
# carpeta temporal del OS, fix del import de importlib.util. El
# resto (selectores, login, lógica de negocio) es igual.
# ------------------------------------------------------------
#   VERSION      : 1.2
#   VERSION_FECHA: 2026-08-29
#   DESCRIPCIÓN  : Busca un option/servicio (location + supplier +
#                  service type + code), entra a RATES, ubica el
#                  período por rango de fechas y price code (o lo
#                  crea copiando el último período existente) y
#                  carga un único valor en las 4 columnas de la
#                  primera fila de la grilla (Group Cost, Group
#                  Sell, FIT Cost, FIT Sell).
#   BASE         : Reutiliza el bootstrap, login/logout,
#                  buscar_producto() y la lógica de períodos
#                  (COPY DATE RANGE / INSERT / selección de price
#                  code) de valorizacion_desde_madre.py (rama
#                  claude/peaceful-goldberg-nks5yi), ya confirmados
#                  en producción. La escritura de tarifas se
#                  simplifica: un valor único en la primera fila en
#                  lugar de un valor por rango de pax derivado de
#                  costos de PCM.
#   CAMBIOS v1.1 : - STYPE_SIDEBAR ampliado al glosario completo
#                  confirmado por la usuaria (skill
#                  buscando-productos-en-tourplan/references/
#                  service-types.md): 12 siglas en vez de solo EX/TF.
#   CAMBIOS v1.2 : - Fix _copiar_ultimo_period (COPY DATE RANGE): elegía
#                  el período fuente a copiar por posición (el más
#                  reciente de la lista), ignorando su price code. Si la
#                  lista tenía períodos de varios price codes mezclados,
#                  podía extender un período con el price code equivocado
#                  sin ningún aviso (confirmado por la usuaria en corrida
#                  real). Ahora se lee el price code real de cada fila
#                  (td.tpcol-pricecodecode) y se elige la más reciente
#                  que coincida con el price code pedido.
#   CAMBIOS v1.3 : - Fix en buscar_producto(): al completar Location, si
#                  la fila `tr.selectedRow` no aparecía a tiempo, el
#                  fallback clickeaba siempre la PRIMERA fila de la tabla
#                  de sugerencias sin verificar que fuera el Location
#                  exacto (mismo bug ya arreglado en notas_srv.py
#                  v1.13/v1.14). Ahora busca entre las filas sugeridas la
#                  que tiene una celda con texto EXACTO (case-insensitive)
#                  igual al Location pedido; si ninguna coincide, no
#                  clickea nada.
#
# Excel — hoja "TARIFAS" (ver _HEADERS más abajo):
#   LOCATION, SUPPLIER, SERVICE TYPE, SERVICE CODE, RATE FROM,
#   RATE TO, PRICE CODE, VALOR, MOSTRAR CAPTURAS, ESTADO,
#   OBSERVACIONES, TIMESTAMP.
# ============================================================

VERSION       = "1.3"
VERSION_FECHA = "2026-09-01"

# ── Config — via variables de entorno (con default = valor original) ──
import os

USERNAME   = os.environ.get("TOURPLAN_USERNAME", "poner minusculas")
PASSWORD   = os.environ.get("TOURPLAN_PASSWORD", "password")
BASE_URL   = os.environ.get("TOURPLAN_BASE_URL", "https://tourplannx.eurotur.com.ar/TourplanNX_Test")
EXCEL_PATH = os.environ.get("TOURPLAN_EXCEL_PATH", "/content/cargar_tarifa_periodo.xlsx")
SHEET      = os.environ.get("TOURPLAN_HOJA", "TARIFAS")
SS_DIR     = os.environ.get("TOURPLAN_SS_DIR", "/content/screenshots")
os.makedirs(SS_DIR, exist_ok=True)

# ── Modo de ejecución ─────────────────────────────────────────
# "lectura" → no escribe nada: busca el option, ubica (o detecta que
#             falta) el período y reporta el valor actual vs. el que
#             se cargaría. Dry-run.
# cualquier otro valor ("completo"/"aplicar") → ejecuta la carga real
#             (crea el período si falta y escribe/guarda el valor).
MODO = os.environ.get("TOURPLAN_MODO", "lectura")

# Multiplicador de tiempos de espera. Producción por default —
# Tourplan responde más lento ahí que en Test.
VELOCIDAD = 1.5

# Price Code por defecto si la columna del Excel viene vacía
PRICE_CODE_DEFAULT = "TR"

# Mapping service type → número de opción en el sidebar de Product Search.
# Glosario confirmado por la usuaria (2026-08), ver skill
# buscando-productos-en-tourplan/references/service-types.md. TA y PR no
# están confirmados: si aparecen, buscar_producto() cae al fallback por
# texto/regex — no asumir un número por patrón/orden alfabético.
STYPE_SIDEBAR = {
    "HT": "01", "HX": "02", "TF": "03", "EX": "04", "ML": "05",
    "RT": "06", "CR": "07", "FT": "08", "OC": "09", "LN": "10",
    "LP": "11", "MS": "12",
}

# Siglas que Tourplan agrupa en el sidebar bajo el prefijo "Z*NO USAR*"
# (organización interna de la empresa, confirmado por la usuaria 2026-08-18 —
# no significa que estén deprecadas). Varias comparten una palabra con una
# categoría real (ej. TN="Transfer Non-Accom" vs TF="Transfer") — usadas en
# buscar_producto() para no matchear la sigla real dentro de un ítem "no usar"
# ni viceversa.
Z_NO_USAR_SIGLAS = {"GA", "TK", "TN", "CH", "CM", "CO", "EN", "GU", "PJ", "ST", "TR", "TI"}

# Columnas Excel TARIFAS (base 1)
C = {
    "location":         1,   # A - opcional, código de location (BUE). Vacío = sin filtro
    "supplier":         2,   # B - obligatorio
    "service_type":     3,   # C - obligatorio (EX, TF, etc.)
    "service_code":     4,   # D - obligatorio: código del option a abrir
    "rate_from":        5,   # E - inicio período dd/Mon/yyyy
    "rate_to":          6,   # F - fin período dd/Mon/yyyy
    "price_code":       7,   # G - TR, 34, etc. o ALL. Default: TR
    "valor":            8,   # H - valor único a cargar en Group Cost/Sell y FIT Cost/Sell
    "mostrar_capturas": 9,   # I - SI/NO
    "estado":          10,   # J - PENDIENTE para procesar
    "observaciones":   11,   # K - SALIDA: detalle del resultado
    "timestamp":       12,   # L - SALIDA: fecha/hora de procesamiento
}

MOSTRAR_CAPTURAS = False
_ss_n = [0]

# ── PASO 0: Entorno ──────────────────────────────────────────
import sys, subprocess, importlib.util, shutil, time, re
from datetime import datetime

_t_inicio = time.time()

print("🔧 Verificando entorno...\n")

# 0.1 Paquetes Python
_PIPS_NEEDED = {
    "selenium":          "selenium",
    "openpyxl":          "openpyxl",
    "webdriver_manager": "webdriver-manager",
    "IPython":           "ipython",
}
_pips_faltantes = [
    pkg for mod, pkg in _PIPS_NEEDED.items()
    if importlib.util.find_spec(mod) is None
]
if _pips_faltantes:
    print(f"  ⏳ pip install {' '.join(_pips_faltantes)} ...")
    subprocess.run(
        [sys.executable, "-m", "pip", "install", "-q"] + _pips_faltantes,
        check=True
    )
    print("  ✅ Paquetes Python OK")
else:
    print("  ✅ Paquetes Python ya instalados")

# 0.2 Google Chrome (multiplataforma — ver common/chrome_bootstrap.py)
from common.chrome_bootstrap import find_or_prepare_chrome
# 0.2b Botón Abortar de la app (ver common/abort.py)
from common.abort import chequear_abort, AbortadoPorUsuario, ABORT_EXIT_CODE

CHROMIUM_BIN, ver_chrome = find_or_prepare_chrome()

# 0.3 Imports
from webdriver_manager.chrome import ChromeDriverManager

from IPython.display import display, Image as IPyImage
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ── Helpers base ──────────────────────────────────────────────
def ss(driver, nombre):
    _ss_n[0] += 1
    p = f"{SS_DIR}/{_ss_n[0]:03d}_{nombre[:40]}_{int(time.time())}.png"
    driver.save_screenshot(p)
    print(f"  📸 {os.path.basename(p)}")
    if MOSTRAR_CAPTURAS:
        display(IPyImage(p, width=900))

def dump(driver, nombre):
    p = f"{SS_DIR}/{nombre}_{int(time.time())}.html"
    with open(p, "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    print(f"  💾 HTML: {p}")

def jc(driver, el):
    driver.execute_script("arguments[0].click();", el)

def set_val(driver, el, value):
    driver.execute_script("""
        var inp = arguments[0], val = arguments[1];
        var setter = Object.getOwnPropertyDescriptor(
            window.HTMLInputElement.prototype, 'value').set;
        setter.call(inp, val);
        inp.dispatchEvent(new Event('input',  {bubbles:true}));
        inp.dispatchEvent(new Event('change', {bubbles:true}));
    """, el, value)

def wait(driver, css, t=12):
    return WebDriverWait(driver, t).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, css)))

def waitx(driver, xpath, t=12):
    return WebDriverWait(driver, t).until(
        EC.presence_of_element_located((By.XPATH, xpath)))

# ── Driver ────────────────────────────────────────────────────
def crear_driver():
    import tempfile

    opts = Options()
    # Sin --headless: esto corre en la PC de la persona (con pantalla), no en
    # el contenedor sin pantalla de Colab. Ademas, varias empresas bloquean
    # el modo headless de Chrome por politica de seguridad.
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1366,911")
    opts.add_argument("--disable-gpu")
    # Sin --remote-debugging-port fijo: chromedriver ya maneja su propio
    # puerto; un puerto fijo puede chocar con otro Chrome ya abierto.

    # Perfil temporal y aislado para esta corrida — evita que Chrome
    # "rebote" hacia una ventana ya abierta con los perfiles reales de la
    # persona (instancia unica de Windows compartiendo el mismo
    # user-data-dir), que Selenium no controla.
    _profile_dir = tempfile.mkdtemp(prefix="tourplan_chrome_profile_")
    opts.add_argument(f"--user-data-dir={_profile_dir}")
    opts.add_argument("--no-first-run")
    opts.add_argument("--no-default-browser-check")

    if CHROMIUM_BIN:
        opts.binary_location = CHROMIUM_BIN
        print(f"Chrome binary: {CHROMIUM_BIN}  ({ver_chrome})")
    else:
        raise RuntimeError("No se encontró ningún Chrome funcional.")
    log_path = os.path.join(tempfile.gettempdir(), "chromedriver.log")
    try:
        drv_path = ChromeDriverManager().install()
        print(f"Chromedriver: {drv_path}")
        svc = Service(executable_path=drv_path, log_output=log_path)
        d = webdriver.Chrome(service=svc, options=opts)
        print("✅ Driver iniciado")
        return d
    except Exception as e:
        if os.path.exists(log_path):
            with open(log_path) as f:
                print(f"\n--- ChromeDriver log ---\n{f.read()[-3000:]}\n---")
        raise RuntimeError(f"No se pudo iniciar Chrome.\nError: {e}")

# ── Login ─────────────────────────────────────────────────────
def login(driver):
    print("🔐 Login...")
    driver.get(f"{BASE_URL}/#/login")
    time.sleep(6 * VELOCIDAD)
    ss(driver, "login_page")

    def _campos_visibles():
        return driver.execute_script("""
            function vis(e){return !!(e && (e.offsetWidth||e.offsetHeight
                                      ||e.getClientRects().length)
                                      && !e.disabled);}
            var txt = Array.from(document.querySelectorAll(
                "input[type='text'], input:not([type])")).filter(vis);
            var pwd = Array.from(document.querySelectorAll(
                "input[type='password']")).filter(vis);
            return [txt[0]||null, pwd[0]||null];
        """)

    u_el = p_el = None
    for intento in range(15):
        u_el, p_el = _campos_visibles()
        if u_el and p_el:
            break
        time.sleep(2 * VELOCIDAD)
    if not (u_el and p_el):
        ss(driver, "login_sin_campos")
        raise Exception("No aparecieron los campos de login")

    def _set(el, valor):
        try: el.clear()
        except Exception: pass
        try: el.click()
        except Exception: pass
        try:
            el.send_keys(valor)
        except Exception:
            driver.execute_script("""
                var el=arguments[0], v=arguments[1];
                var s=Object.getOwnPropertyDescriptor(
                    window.HTMLInputElement.prototype,'value').set;
                s.call(el,v);
                el.dispatchEvent(new Event('input',{bubbles:true}));
                el.dispatchEvent(new Event('change',{bubbles:true}));
            """, el, valor)

    _set(u_el, USERNAME)
    _set(p_el, PASSWORD)
    time.sleep(0.5)

    clic = driver.execute_script("""
        function vis(e){return !!(e && (e.offsetWidth||e.offsetHeight
                                  ||e.getClientRects().length) && !e.disabled);}
        var b = Array.from(document.querySelectorAll(
            "button.login, button[type='submit'], button")).filter(vis)
            .find(function(x){return /log\\s*in|ingresar|entrar|sign\\s*in/i
                                     .test((x.innerText||'')) ||
                                     x.classList.contains('login');});
        if (b){ b.click(); return (b.innerText||'button.login').trim(); }
        return null;
    """)
    if not clic:
        try: p_el.send_keys(Keys.ENTER)
        except Exception: pass
    time.sleep(8 * VELOCIDAD)
    assert "login" not in driver.current_url.lower(), "Login falló"
    ss(driver, "post_login")
    print("✅ Login OK")

# ── Logout ────────────────────────────────────────────────────
def logout(driver):
    print("\n🔒 Cerrando ventanas y haciendo logout...")
    try:
        principal = driver.window_handles[0]
        for h in driver.window_handles[1:]:
            try: driver.switch_to.window(h); driver.close()
            except Exception: pass
        driver.switch_to.window(principal)
    except Exception: pass

    def _click_item(regex):
        return driver.execute_script("""
            var rx = new RegExp(arguments[0], 'i');
            var els = Array.from(document.querySelectorAll(
                'li, label, a, button, span, div'));
            var best = null;
            for (var el of els){
                if (!el.offsetParent) continue;
                var t = (el.innerText || '').trim();
                if (!t || t.length > 40 || !rx.test(t)) continue;
                if (!best || t.length <= (best.innerText||'').trim().length)
                    best = el;
            }
            if (!best) return null;
            best.click();
            return (best.innerText || '').trim().slice(0, 40);
        """, regex)

    def _en_login():
        return driver.execute_script("""
            var pwd = document.querySelector("input[type='password']");
            return (pwd && pwd.offsetParent !== null) ||
                   /login/i.test(window.location.href);
        """)

    try:
        driver.get(f"{BASE_URL}/#/home")
        time.sleep(4 * VELOCIDAD)
        ss(driver, "logout_home")
        clicked = None
        try:
            btn_panel = WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#openUserPanel")))
            jc(driver, btn_panel)
            time.sleep(2 * VELOCIDAD)
            ss(driver, "logout_usermenu")
            btn_logout = WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.CSS_SELECTOR,
                    "div.panelHeader tp-button button, div.panelHeader button")))
            jc(driver, btn_logout)
            clicked = (btn_logout.text or "Logout").strip()
        except Exception as _e:
            print(f"    ⚠ Flujo #openUserPanel falló ({_e}) — fallback por texto")
        rx_logout = r"^(log\s?out|sign\s?out|cerrar sesi)"
        if not clicked:
            clicked = _click_item(rx_logout)
        if not clicked:
            padre = _click_item(r"logged in as")
            time.sleep(2 * VELOCIDAD)
            if padre:
                clicked = _click_item(rx_logout)
        time.sleep(5 * VELOCIDAD)
        ss(driver, "logout_done")
        if clicked and _en_login():
            print(f"  🔓 Logout OK (click en '{clicked}')")
        elif clicked:
            print(f"  ⚠ Click en '{clicked}' pero NO volvió al login")
        else:
            print("  ⚠ No encontré la opción LOG OUT")
    except Exception as e:
        print(f"  ⚠ Error en logout: {e}")

# ── Navegación ────────────────────────────────────────────────
def hamburger(driver):
    img = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "nav img")))
    jc(driver, img)
    time.sleep(2.5)

def menu_item(driver, n_or_text):
    if isinstance(n_or_text, int):
        xpath = f"(//nav//ul/li)[{n_or_text}]/div/div"
    else:
        texto = n_or_text.upper()
        xpath = (f"//nav//ul/li[.//*[contains("
                 f"translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),"
                 f"'{texto}')]]/div/div")
    el = WebDriverWait(driver, 8).until(
        EC.presence_of_element_located((By.XPATH, xpath)))
    jc(driver, el)
    time.sleep(2 * VELOCIDAD)

# ── Fechas ────────────────────────────────────────────────────
MESES_ES = {"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
            "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}

def _try_float(s):
    if s is None: return None
    try:
        return float(str(s).replace(",", "").replace("$", "").replace(" ", "").strip())
    except: return None

def parsear_fecha(txt):
    if not txt: return None
    if isinstance(txt, datetime): return txt
    txt = str(txt).strip()
    if len(txt) >= 10 and txt[4] == "-" and txt[7] == "-":
        try: return datetime.fromisoformat(txt[:10])
        except: pass
    p = txt.split("/")
    if len(p) == 3:
        try:
            dia = int(p[0])
            mes_raw = p[1]
            mes = MESES_ES.get(mes_raw[:3].capitalize(), None) or int(mes_raw)
            anio_raw = int(p[2])
            anio = anio_raw + 2000 if anio_raw < 100 else anio_raw
            return datetime(anio, mes, dia)
        except: pass
    return None

def fmt_tp(dt):
    meses = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
             7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
    return f"{dt.day:02d}/{meses[dt.month]}/{dt.year}"


# ── Buscar producto por código (idéntico a valorizacion_desde_madre.py) ──
class ProductoNoEncontrado(Exception):
    pass

def buscar_producto(driver, location, supplier, codigo, service_type=None, handle_origen=None):
    """
    Abre un producto/option específico en Product Setup.
    Reutilizado sin cambios de valorizacion_desde_madre.py (a su vez
    reutilizado de v3 del repo hermano) — ya confirmado en producción.
    """
    st_upper = (service_type or "").strip().upper()
    st_str   = f"/{st_upper}" if st_upper else ""
    print(f"\n  📦 Abriendo: {location or '—'}/{supplier}/{codigo}{st_str}")

    if handle_origen:
        driver.switch_to.window(handle_origen)

    driver.get(f"{BASE_URL}/#/home")
    time.sleep(2 * VELOCIDAD)
    driver.get(f"{BASE_URL}/#/product")
    time.sleep(5 * VELOCIDAD)
    ss(driver, f"ps_inicial_{codigo[:10]}")

    lupa = wait(driver, "#searchWrapper li:nth-of-type(2) button")
    jc(driver, lupa)
    time.sleep(3 * VELOCIDAD)
    ss(driver, f"ps_modal_{codigo[:10]}")

    try:
        WebDriverWait(driver, 4).until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "div.parameters1 input")))
    except Exception:
        tab_sel = driver.execute_script("""
            var els = Array.from(document.querySelectorAll('li,button,a,div,span'));
            for (var el of els){
                if (!el.offsetParent) continue;
                var t = (el.innerText || '').trim().toUpperCase();
                if (t === 'SELECTION'){ el.click(); return true; }
            }
            return false;
        """)
        print(f"    Tab SELECTION re-activado: {tab_sel}")
        time.sleep(2 * VELOCIDAD)

    if st_upper:
        try:
            # Traduce la sigla al prefijo numérico confirmado (STYPE_SIDEBAR) si
            # existe; si no, busca la sigla cruda. Siempre por palabra completa
            # (boundary regex) — nunca substring suelto, que puede matchear por
            # accidente dentro de otra categoría (confirmado en Flag as Deleted:
            # 'TR' tildaba "Hotel Extras" porque "Extras" contiene "tr"). Además
            # excluye los ítems "Z*NO USAR*" cuando la sigla buscada no es una de
            # esas 12 — evita que, p.ej., 'TF' (Transfer) matchee dentro de
            # "Z*NO USAR* - TN - Transfer Non-Accom" por compartir la palabra
            # "Transfer".
            st_buscar = STYPE_SIDEBAR.get(st_upper, st_upper)
            excluir_no_usar = st_upper not in Z_NO_USAR_SIGLAS
            patron_js = (r"(^|[\s\-–—/(])" + re.escape(st_buscar)
                         + r"([\s\-–—/)]|$)")
            resultado = None
            for intento in range(3):
                resultado = driver.execute_script("""
                    var rx = new RegExp(arguments[0]);
                    var excluirNoUsar = arguments[1];
                    var textos = [], elegido = null;
                    for (var li of document.querySelectorAll('li')){
                        if (!li.offsetParent) continue;
                        var t = (li.textContent || '').trim();
                        if (!t || t.length > 80) continue;
                        textos.push(t);
                        if (elegido) continue;
                        if (excluirNoUsar && t.toLowerCase().indexOf('no usar') !== -1) continue;
                        if (rx.test(t.toUpperCase())) elegido = li;
                    }
                    if (elegido){
                        elegido.click();
                        return {ok: true, texto: (elegido.textContent||'').trim().slice(0,40)};
                    }
                    return {ok: false, textos: textos.slice(0,40)};
                """, patron_js, excluir_no_usar)
                if resultado and resultado.get("ok"):
                    break
                time.sleep(1.5)
            if resultado and resultado.get("ok"):
                time.sleep(1.5)
                print(f"    Service type '{st_upper}' seleccionado: {resultado.get('texto')}")
            else:
                print(f"    ⚠ Service type '{st_upper}' no encontrado en sidebar")
                print(f"      Ítems visibles del sidebar: {(resultado or {}).get('textos', [])}")
        except Exception as e:
            print(f"    ⚠ Error seleccionando service type: {e}")

    if location:
        try:
            inp_loc = wait(driver, "div.parameters1 li:nth-of-type(1) input")
            inp_loc.click(); time.sleep(0.3)
            set_val(driver, inp_loc, location)
            time.sleep(1.5)
            try:
                row = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "tr.selectedRow > td.description")))
                jc(driver, row); time.sleep(1)
            except:
                try:
                    fila_exacta = driver.execute_script("""
                        var loc = arguments[0].toUpperCase();
                        var trs = Array.from(document.querySelectorAll(
                            'div.parameters1 li:nth-of-type(1) table tbody tr'));
                        for (var tr of trs) {
                            var tds = Array.from(tr.querySelectorAll('td'));
                            var exacto = tds.some(function(td){
                                return td.children.length === 0 &&
                                       td.innerText.trim().toUpperCase() === loc;
                            });
                            if (exacto) return tr;
                        }
                        return null;
                    """, location)
                    if fila_exacta is not None:
                        jc(driver, fila_exacta); time.sleep(1)
                    else:
                        print(f"    ⚠ Location '{location}': ninguna fila sugerida coincide "
                              f"EXACTO — no se clickeó ninguna a ciegas")
                except: pass
            print(f"    Location: {location} OK")
        except: pass

    if supplier:
        try:
            inp_sup = wait(driver, "div.parameters1 li:nth-of-type(2) input")
            inp_sup.click(); time.sleep(0.3)
            set_val(driver, inp_sup, supplier)
            time.sleep(1.5)
            try:
                row_sup = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located(
                        (By.XPATH,
                         f"//div[contains(@class,'parameters1')]//td[normalize-space(text())='{supplier.upper()}']")))
                jc(driver, row_sup); time.sleep(1)
            except:
                try:
                    inp_sup2 = driver.find_element(By.CSS_SELECTOR,
                        "div.parameters1 li:nth-of-type(2) input")
                    inp_sup2.send_keys(Keys.TAB); time.sleep(1)
                except: pass
            print(f"    Supplier: {supplier} OK")
        except: pass

    inp_cod = wait(driver, "div.parameters1 li:nth-of-type(3) input")
    inp_cod.click(); time.sleep(0.3)
    set_val(driver, inp_cod, codigo)
    time.sleep(0.5)
    print(f"    Code: {codigo}")
    ss(driver, f"ps_modal_lleno_{codigo[:10]}")

    btn_search = wait(driver, "#productSearchFilter li:nth-of-type(4) button")
    jc(driver, btn_search)
    time.sleep(6 * VELOCIDAD)
    ss(driver, f"ps_resultados_{codigo[:10]}")

    def _en_contexto_producto():
        try:
            img = WebDriverWait(driver, 4).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "nav img")))
            driver.execute_script("arguments[0].click();", img)
            time.sleep(2 * VELOCIDAD)
            items = driver.execute_script("""
                return Array.from(document.querySelectorAll('nav ul > li')).map(function(li){
                    return (li.querySelector('div div')||li).innerText.trim().split('\\n')[0].toUpperCase();
                });
            """) or []
            print(f"    Menu: {items}")
            PROD_ITEMS = {'UTILITIES','RATES','SEASONALITY','OPERATION','CONTENT','PRODUCT DETAILS'}
            ok = bool(set(items) & PROD_ITEMS)
            driver.execute_script("arguments[0].click();", img)
            time.sleep(1)
            return ok, items
        except Exception:
            return False, []

    clicked = driver.execute_script(f"""
        var cod = '{codigo}'.toUpperCase();
        var st  = '{st_upper}';
        var rows = Array.from(document.querySelectorAll('table tbody tr'));

        function matchRow(tr){{
            var tds = Array.from(tr.querySelectorAll('td'));
            var hasCod = tds.some(function(td){{
                return td.children.length===0 && td.innerText.trim().toUpperCase()===cod;
            }});
            if(!hasCod) return false;
            if(!st) return true;
            return tds.some(function(td){{
                return td.children.length===0 && td.innerText.trim().toUpperCase()===st;
            }});
        }}

        var match = null;
        for(var tr of rows){{ if(matchRow(tr)){{ match=tr; break; }} }}
        if(!match){{
            for(var tr of rows){{
                if((tr.innerText||'').toUpperCase().includes(cod)){{ match=tr; break; }}
            }}
        }}
        if(!match) return null;

        var tds = match.querySelectorAll('td');
        var target = tds.length > 4 ? tds[4] : (tds.length > 0 ? tds[tds.length-1] : match);
        target.click();
        return target.innerText.trim().slice(0,50) || 'clicked';
    """)

    if clicked:
        time.sleep(6 * VELOCIDAD)
        ss(driver, f"ps_cargado_{codigo[:10]}")
        ok_ctx, menu_items = _en_contexto_producto()
        if ok_ctx:
            print(f"  ✅ Producto {codigo} en contexto correcto (menú: {menu_items})")
            return
        print(f"    ⚠ Click ok ({clicked}) pero menú={menu_items}")
        return

    page_info = driver.execute_script("""
        return {
            url: window.location.href,
            tables: Array.from(document.querySelectorAll('table tbody tr')).slice(0,5).map(r=>r.innerText.trim().slice(0,60))
        };
    """)
    print(f"    Sin resultado: {page_info}")
    dump(driver, f"ps_sin_resultado_{codigo[:10]}")
    ss(driver, f"ps_sin_resultado_final_{codigo[:10]}")
    raise ProductoNoEncontrado(f"Producto '{codigo}' no encontrado en resultados")


# ── Rates: helpers de período (idénticos a valorizacion_desde_madre.py) ──
def _parse_rate_period(texto):
    m = re.search(r'(\d+/\w+/\d+)\s*[-–]\s*(\d+/\w+/\d+)', texto)
    if m:
        return parsear_fecha(m.group(1)), parsear_fecha(m.group(2))
    return None, None

def _seleccionar_price_code(driver, code, etiqueta=""):
    code = (code or "").strip().upper()
    if not code:
        return False
    try:
        driver.execute_script("""
            var lbl = document.querySelector('label[for="priceCodeModeSelected"]');
            var inp = document.getElementById('priceCodeModeSelected');
            if (lbl) lbl.click();
            if (inp) inp.click();
        """)
        time.sleep(1.5)
        driver.execute_script("""
            var dd = document.getElementById('priceCode');
            if (!dd) return;
            var inp = dd.querySelector('input[type="text"], input');
            if (inp){ inp.focus(); inp.click(); } else { dd.click(); }
        """)
        time.sleep(1.0)
        driver.execute_script("""
            var dd = document.getElementById('priceCode');
            if (!dd) return;
            var inp = dd.querySelector('input[type="text"], input');
            if (!inp) return;
            var setter = Object.getOwnPropertyDescriptor(
                window.HTMLInputElement.prototype, 'value').set;
            setter.call(inp, arguments[0]);
            inp.dispatchEvent(new Event('input', {bubbles:true}));
            inp.dispatchEvent(new Event('keyup', {bubbles:true}));
        """, code)
        time.sleep(1.2)
        elegido = driver.execute_script("""
            var pc = arguments[0].toUpperCase();
            var dd = document.getElementById('priceCode');
            if (!dd) return null;
            function vis(e){return !!(e.offsetWidth||e.offsetHeight||e.getClientRects().length);}
            function match(tr){
                var cod = tr.querySelector('td.code');
                var des = tr.querySelector('td.description');
                var ct = cod ? (cod.innerText||cod.textContent||'').trim().toUpperCase() : '';
                var dt = des ? (des.innerText||des.textContent||'').trim().toUpperCase() : '';
                if (ct===pc || dt===pc || dt.indexOf(pc+' ')===0 ||
                    dt.indexOf(pc+'-')===0 || dt.indexOf(pc+' -')===0)
                    return {tr:tr, txt:ct + ' | ' + dt};
                return null;
            }
            var rows = Array.from(dd.querySelectorAll('table tr'));
            var cand = null;
            for (var tr of rows){ if (vis(tr)){ var m=match(tr); if(m){cand=m;break;} } }
            if (!cand){ for (var tr2 of rows){ var m2=match(tr2); if(m2){cand=m2;break;} } }
            if (!cand) return null;
            try{ cand.tr.scrollIntoView({block:'center'}); }catch(e){}
            var des = cand.tr.querySelector('td.description');
            var cod = cand.tr.querySelector('td.code');
            (des||cod||cand.tr).click();
            cand.tr.click();
            return cand.txt;
        """, code)
        time.sleep(2.5)
        print(f"    Price Code '{code}' seleccionado ({etiqueta}): {elegido}")
        if not elegido:
            ss(driver, f"pricecode_fail_{(etiqueta or code)[:12]}")
        return bool(elegido)
    except Exception as e:
        print(f"    ⚠ No pude seleccionar Price Code '{code}' ({etiqueta}): {e}")
        return False


def _insertar_rate_period(driver, fecha_desde, fecha_hasta, cod, price_code=""):
    try:
        btn_insert = waitx(driver,
            "//button[normalize-space(text())='INSERT' or normalize-space(text())='Insert']",
            t=8)
        jc(driver, btn_insert)
        time.sleep(3 * VELOCIDAD)
        ss(driver, f"rates_insert_dlg_{cod[:10]}")

        info = driver.execute_script("""
            var dlg = document.querySelector(
                'tp-dialog, tp-modal, [role="dialog"], .modal-dialog, .dialog');
            var scope = dlg || document;
            function vis(e){ return !!(e.offsetWidth || e.offsetHeight
                                       || e.getClientRects().length); }
            var inputs = Array.from(scope.querySelectorAll('input'))
                .filter(vis)
                .filter(function(i){
                    var t = (i.getAttribute('type') || '').toLowerCase();
                    return t === '' || t === 'text' || t === 'date';
                });
            window.__tp_ins_inputs = inputs;
            return {
                inputs: inputs.map(function(i, idx){ return {
                    idx: idx, id: i.id || '', val: i.value || ''
                }; })
            };
        """)
        inputs_info = info.get("inputs", [])
        if not inputs_info:
            dump(driver, f"rates_insert_sin_inputs_{cod[:10]}")
            raise Exception("El diálogo INSERT no tiene inputs visibles")

        pat_fecha = re.compile(r'\d{1,2}/[A-Za-z]{3}/\d{2,4}')
        idx_con_fecha = [i["idx"] for i in inputs_info if pat_fecha.search(i["val"])]
        idx_vacios    = [i["idx"] for i in inputs_info if not (i["val"] or "").strip()]

        def _inp(idx):
            return driver.execute_script(
                "return window.__tp_ins_inputs[arguments[0]];", idx)

        def _fill(idx, fecha):
            el = _inp(idx)
            set_val(driver, el, fmt_tp(fecha))
            try: el.send_keys(Keys.TAB)
            except: driver.execute_script(
                "arguments[0].dispatchEvent(new Event('blur'))", el)
            time.sleep(0.5)

        if len(idx_con_fecha) >= 2:
            _fill(idx_con_fecha[0], fecha_desde)
            _fill(idx_con_fecha[1], fecha_hasta)
        elif len(idx_con_fecha) == 1 and idx_vacios:
            idx_from = idx_con_fecha[0]
            idx_to   = next((i for i in idx_vacios if i > idx_from), idx_vacios[0])
            val_from = next(i["val"] for i in inputs_info if i["idx"] == idx_from)
            if val_from != fmt_tp(fecha_desde):
                _fill(idx_from, fecha_desde)
            _fill(idx_to, fecha_hasta)
        elif len(idx_con_fecha) == 1:
            _fill(idx_con_fecha[0], fecha_hasta)
        elif len(idx_vacios) >= 2:
            idx_from, idx_to = idx_vacios[0], idx_vacios[1]
            _fill(idx_from, fecha_desde)
            _fill(idx_to, fecha_hasta)
        else:
            dump(driver, f"rates_insert_sin_fechas_{cod[:10]}")
            raise Exception(f"No identifiqué inputs de fecha en el diálogo INSERT (inputs: {inputs_info})")

        pc_warning = ""
        if price_code:
            elegido_dlg = _seleccionar_price_code(
                driver, price_code, etiqueta=f"dialog {cod}") \
                if driver.execute_script(
                    "var d=document.querySelector("
                    "'tp-dialog #priceCode, tp-modal #priceCode, "
                    "[role=\"dialog\"] #priceCode');return !!d;") else False
            if not elegido_dlg:
                pc_warning = f"Price Code '{price_code}' no se pudo asignar al período (dialog INSERT)"
                print(f"    🚩 {pc_warning}")
        else:
            pc_warning = "Sin Price Code: el período pudo quedar 'Unassigned'"
            print(f"    🚩 {pc_warning}")

        ss(driver, f"rates_insert_lleno_{cod[:10]}")

        guardado = driver.execute_script("""
            var dlg = document.querySelector(
                'tp-dialog, tp-modal, [role="dialog"], .modal-dialog, .dialog');
            var scope = dlg || document;
            function vis(e){ return !!(e.offsetWidth || e.offsetHeight
                                       || e.getClientRects().length); }
            var btn = scope.querySelector(
                'tp-button.save > button, tp-button.ok > button, tp-button.yes > button');
            if (btn && vis(btn)){ btn.click(); return btn.innerText.trim() || 'tp-button.save'; }
            var bs = Array.from(scope.querySelectorAll('button')).filter(vis);
            for (var b of bs){
                var t = (b.innerText || '').trim().toUpperCase();
                if (t === 'SAVE' || t === 'OK' || t === 'INSERT' || t === 'ACCEPT'){
                    b.click(); return t;
                }
            }
            return null;
        """)
        if not guardado:
            dump(driver, f"rates_insert_sin_save_{cod[:10]}")
            raise Exception(f"No encontré botón SAVE en el diálogo INSERT")
        print(f"    SAVE del INSERT clickeado ('{guardado}')")
        time.sleep(4 * VELOCIDAD)
        ss(driver, f"rates_insert_save_{cod[:10]}")
        print(f"    ✅ Período insertado {fmt_tp(fecha_desde)} – {fmt_tp(fecha_hasta)}")
        return pc_warning
    except Exception as e:
        try: ss(driver, f"rates_insert_error_{cod[:10]}")
        except: pass
        raise Exception(f"Error al insertar período: {e}")


def _copiar_ultimo_period(driver, fecha_desde, fecha_hasta, cod, price_code=""):
    """
    Copia el último período via COPY DATE RANGE y ajusta las fechas.
    Preserva la configuración de tarifas del período anterior (markup,
    márgenes, etc.) en lugar de insertar un período en blanco.
    Retorna (pc_warning, ya_en_detalle).

    El período fuente a copiar se elige por PRICE CODE, no simplemente
    "el más reciente de la lista": la lista puede tener períodos de varios
    price codes mezclados (el filtro `_seleccionar_price_code` del caller
    no garantiza por sí solo qué fila queda en el índice 0), y copiar a
    ciegas el primero podía extender un período con el price code
    equivocado sin ningún aviso (confirmado por la usuaria en corrida
    real de este script). Se lee el price code real de cada fila
    (td.tpcol-pricecodecode, mismo elemento que usa _leer_periodos) y se
    elige la primera (= más reciente) que coincida con `price_code` — o
    con Unassigned/ALL si `price_code` viene vacío.
    """
    filas_pc = driver.execute_script("""
        var out = [];
        document.querySelectorAll('td.tpcol-rateperiod').forEach(function(d, i){
            var tr = d.closest('tr');
            var p  = tr ? tr.querySelector('td.tpcol-pricecodecode') : null;
            out.push({idx: i, pc: (p ? (p.innerText || '').trim() : '')});
        });
        return out;
    """)
    if not filas_pc:
        print(f"    ⚠ No hay períodos existentes → INSERT normal")
        return _insertar_rate_period(driver, fecha_desde, fecha_hasta, cod, price_code), False

    PC = (price_code or "").strip().upper()
    PC_VACIOS = ("", "—", "-", "UNASSIGNED", "ALL")

    def _pc_coincide(pc_row):
        pc_row = (pc_row or "").strip().upper()
        return pc_row in PC_VACIOS if not PC else pc_row == PC

    idx_fuente = next((f["idx"] for f in filas_pc if _pc_coincide(f["pc"])), None)
    if idx_fuente is None:
        etq_pc = PC or "ALL/Unassigned"
        print(f"    ⚠ Ningún período existente tiene price code {etq_pc} → INSERT normal")
        return _insertar_rate_period(driver, fecha_desde, fecha_hasta, cod, price_code), False

    periodos_rows = driver.find_elements(By.CSS_SELECTOR, "td.tpcol-rateperiod")
    ultimo_td  = periodos_rows[idx_fuente]
    ultimo_txt = (ultimo_td.text or "").strip()
    print(f"    Período más reciente con price code {PC or 'ALL/Unassigned'}: "
          f"'{ultimo_txt}' (idx {idx_fuente}) → abriendo para COPY DATE RANGE")
    ss(driver, f"rates_copy_before_{cod[:10]}")

    jc(driver, ultimo_td)
    time.sleep(4 * VELOCIDAD)
    ss(driver, f"rates_copy_detalle_{cod[:10]}")

    try:
        btn_cdr = wait(driver, "tp-button.copydaterange > button", t=8)
    except Exception:
        print(f"    ⚠ No encontré tp-button.copydaterange → fallback a INSERT")
        ss(driver, f"rates_copy_nobtn_{cod[:10]}")
        hamburger(driver)
        menu_item(driver, "RATES")
        time.sleep(3 * VELOCIDAD)
        return _insertar_rate_period(driver, fecha_desde, fecha_hasta, cod, price_code), False

    print(f"    📋 COPY DATE RANGE encontrado → abriendo dialog")
    jc(driver, btn_cdr)
    time.sleep(3 * VELOCIDAD)
    ss(driver, f"rates_copy_dlg_{cod[:10]}")

    fecha_desde_str = fecha_desde.strftime("%d/%m/%y")
    fecha_hasta_str = fecha_hasta.strftime("%d/%m/%y")

    def _fill_dialog_date(sel, fecha_str, label):
        """Escribe la fecha y VERIFICA que haya quedado (confirmado en
        corrida real de valorizacion_desde_madre.py: el setter sintético +
        evento 'change' a veces no "pega" en este campo — TO quedaba vacío
        en pantalla sin que set_val() tirara ninguna excepción, y el caller
        seguía de largo con el dialog incompleto). Si no coincide, reintenta
        con foco + Ctrl+A + Delete + send_keys (eventos de teclado nativos).
        Si tras el reintento sigue sin coincidir, lanza excepción para que
        el caller haga el fallback a INSERT en vez de clickear OK con el
        dialog a medio llenar.

        La comparación es por FECHA, no por texto exacto: confirmado en
        corrida real (AEEZMV) que el campo TO a veces re-formatea el valor
        que se le escribió ('31/03/27') a un formato distinto pero
        equivalente ('31/Mar/2027') apenas lo acepta — comparar como texto
        literal tomaba eso como una falla y cancelaba un COPY DATE RANGE
        que en realidad había funcionado bien, terminando en un INSERT
        innecesario que además creaba el período con la fecha mal
        (01/Nov/2026–01/Nov/2026 en vez del rango pedido)."""
        def _valor_coincide(val):
            if val == fecha_str:
                return True
            f_val = parsear_fecha(val)
            f_obj = parsear_fecha(fecha_str)
            return f_val is not None and f_obj is not None and f_val == f_obj

        inp = wait(driver, sel, t=6)
        jc(driver, inp)
        time.sleep(0.2)
        set_val(driver, inp, fecha_str)
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));", inp)
        time.sleep(0.3)
        val_leido = (inp.get_attribute("value") or "").strip()

        if not _valor_coincide(val_leido):
            print(f"    ⚠ {label} quedó '{val_leido}' tras set_val — reintentando con send_keys")
            inp.click()
            time.sleep(0.2)
            inp.send_keys(Keys.CONTROL + "a")
            inp.send_keys(Keys.DELETE)
            inp.send_keys(fecha_str)
            inp.send_keys(Keys.TAB)
            time.sleep(0.3)
            val_leido = (inp.get_attribute("value") or "").strip()

        print(f"    {label}: {val_leido}")
        if not _valor_coincide(val_leido):
            raise Exception(
                f"{label} quedó '{val_leido}' en vez de '{fecha_str}' tras 2 intentos")

    def _cancelar_dialog_y_hacer_insert():
        # Botones reales del dialog COPY DATE RANGE, confirmados por
        # captura de pantalla en corrida real (valorizacion_desde_madre.py):
        # SPLIT DATE RANGE, COPY DATE RANGE, DELETE DATE RANGE, EXIT, SAVE,
        # INSERT RATE SET, DELETE RATE SET, SET CHILD RATES. El botón para
        # cerrarlo es "EXIT", NO "cancel"/"close" — por eso el click de
        # abajo no encontraba nada y el dialog quedaba abierto en silencio.
        # HTML confirmado por la usuaria (inspeccionado en Tourplan real):
        #   <button class="tpbutton tpcancel tpsecondarysystembutton">Exit</button>
        # — es un <button> plano con clase "tpcancel" (sin el wrapper
        # <tp-button> que usan otros diálogos como SAVE/OK).
        try:
            driver.execute_script("""
                function vis(e){ return !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length); }
                var b = document.querySelector(
                    'button.tpcancel, tp-button.cancel > button, tp-button.close > button, tp-button.exit > button');
                if (b && vis(b)) { b.click(); return; }
                var dlg = document.querySelector('tp-dialog:nth-of-type(2)') || document;
                var btns = Array.from(dlg.querySelectorAll('button')).filter(vis);
                for (var btn of btns) {
                    var t = (btn.innerText || '').trim().toUpperCase();
                    if (t === 'EXIT' || t === 'CANCEL' || t === 'CLOSE') { btn.click(); return; }
                }
            """)
            time.sleep(1)
        except Exception: pass

        # Verificar que el dialog realmente se cerró antes de seguir. Si
        # sigue abierto, NO seguir a INSERT — confirmado en corrida real
        # (captura de pantalla): con el dialog COPY DATE RANGE todavía
        # abierto, el "INSERT" que se clickeaba a continuación terminaba
        # siendo el botón "INSERT RATE SET" DEL PROPIO DIALOG viejo (abre
        # otro dialog "Insert Rate" apilado encima, con sus propios TRAVEL
        # DATE FROM/TO) — y _insertar_rate_period() leía los inputs
        # numéricos de la grilla de rates de atrás, no fechas. Mejor
        # abortar la fila con un error claro que arrastrar ese estado a un
        # INSERT que nunca iba a funcionar.
        for _ in range(4):
            sigue_abierto = driver.execute_script("""
                var dlg = document.querySelector('tp-dialog:nth-of-type(2)');
                return !!(dlg && dlg.offsetParent);
            """)
            if not sigue_abierto:
                break
            time.sleep(1)
        else:
            ss(driver, f"rates_copy_sigue_abierto_{cod[:10]}")
            raise Exception(
                "El dialog COPY DATE RANGE no se cerró tras cancelar — se aborta "
                "en vez de intentar INSERT con el dialog viejo todavía abierto "
                "(podía terminar clickeando 'INSERT RATE SET' del propio dialog "
                "en vez de un INSERT de período nuevo)."
            )

        hamburger(driver); menu_item(driver, "RATES"); time.sleep(3 * VELOCIDAD)
        return _insertar_rate_period(driver, fecha_desde, fecha_hasta, cod, price_code), False

    try:
        _fill_dialog_date(
            "tp-dialog:nth-of-type(2) li:nth-of-type(2) input[type='text']",
            fecha_desde_str, "FROM")
    except Exception as _e:
        print(f"    ⚠ No encontré campo FROM en dialog COPY: {_e} → cancela y hace INSERT")
        ss(driver, f"rates_copy_nofrom_{cod[:10]}")
        return _cancelar_dialog_y_hacer_insert()

    try:
        _fill_dialog_date(
            "tp-dialog:nth-of-type(2) li:nth-of-type(3) input[type='text']",
            fecha_hasta_str, "TO")
    except Exception as _e:
        print(f"    ⚠ No encontré campo TO en dialog COPY: {_e} → cancela y hace INSERT")
        ss(driver, f"rates_copy_noto_{cod[:10]}")
        return _cancelar_dialog_y_hacer_insert()
    ss(driver, f"rates_copy_lleno_{cod[:10]}")

    jc(driver, wait(driver, "tp-button.ok > button"))
    time.sleep(5 * VELOCIDAD)
    ss(driver, f"rates_copy_ok_{cod[:10]}")
    print(f"    ✅ COPY DATE RANGE OK → nuevo período {fecha_desde_str} – {fecha_hasta_str}")
    return "", True


def _leer_tabla_rates(driver):
    return driver.execute_script("""
        for(var t of document.querySelectorAll('table')){
            var ths = Array.from(t.querySelectorAll('th'));
            if(ths.some(h => h.innerText.includes('GROUP COST') ||
                              h.innerText.includes('COST'))){
                return {
                    headers: ths.map(h => h.innerText.trim()),
                    rows: Array.from(t.querySelectorAll('tbody tr')).map(tr => ({
                        celdas: Array.from(tr.querySelectorAll('td')).map(td => td.innerText.trim()),
                        inputs: Array.from(tr.querySelectorAll('input')).map(i => i.value.trim())
                    }))
                };
            }
        }
        return null;
    """)


def _leer_primera_fila_valor(driver):
    """Lee etiqueta y valor actual (1ra columna editable) de la primera
    fila con datos de la grilla de rates del período abierto."""
    tabla = _leer_tabla_rates(driver)
    if not tabla or not tabla.get("rows"):
        return None, None
    primera = next((r for r in tabla["rows"] if r["celdas"] and r["celdas"][0].strip()), None)
    if not primera:
        return None, None
    etiqueta = primera["celdas"][0]
    val_raw  = primera["inputs"][0] if (primera["inputs"] and primera["inputs"][0]) \
               else (primera["celdas"][1] if len(primera["celdas"]) > 1 else "")
    return etiqueta, (_try_float(val_raw) or 0.0)


def _escribir_valor_unico(driver, valor, cod):
    """
    Escribe `valor` en las 4 columnas de costo/venta (Group Cost, Group
    Sell, FIT Cost, FIT Sell — mismos índices ci=(0,3,4,7) que usa
    valorizacion_desde_madre.py) de la PRIMERA fila de la grilla de rates
    del período abierto, y hace SAVE. A diferencia del script base (que
    calcula un valor distinto por cada rango de pax a partir de un PCM),
    acá el mismo valor único se carga una sola vez en la primera fila.
    """
    for sel in ["[class*='tab-item']", "div.tab", "a.tab", "button.tab"]:
        for el in driver.find_elements(By.CSS_SELECTOR, sel):
            if el.text.strip().upper() == "RATES" and el.is_displayed():
                jc(driver, el); time.sleep(3 * VELOCIDAD); break

    raiz = driver
    for sel in ["tp-dialog", "tp-modal", "[role='dialog']"]:
        cand = [e for e in driver.find_elements(By.CSS_SELECTOR, sel) if e.is_displayed()]
        if cand:
            raiz = cand[-1]; break

    filas = [f for f in raiz.find_elements(By.XPATH, ".//table//tbody//tr")
             if f.is_displayed() and f.text.strip()]
    if not filas:
        ss(driver, f"valor_sin_filas_{cod[:10]}")
        dump(driver, f"valor_sin_filas_{cod[:10]}")
        raise Exception(f"No encontré ninguna fila en la grilla de rates de {cod}")

    fila = filas[0]
    etiqueta_fila = fila.text.split("\n")[0]

    edits = [inp for inp in fila.find_elements(By.TAG_NAME, "input")
             if inp.is_displayed() and not inp.get_attribute("readonly")]
    if not edits:
        ss(driver, f"valor_sin_inputs_{cod[:10]}")
        raise Exception(f"Sin inputs editables en la primera fila ('{etiqueta_fila}') de {cod}")

    valor_viejo = _try_float(edits[0].get_attribute("value")) or 0.0
    valor_txt = f"{valor:.2f}"

    for ci in (0, 3, 4, 7):
        if ci >= len(edits): continue
        inp = edits[ci]
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});", inp)
        driver.execute_script("""
            var inp = arguments[0], val = arguments[1];
            var setter = Object.getOwnPropertyDescriptor(
                window.HTMLInputElement.prototype, 'value').set;
            inp.focus();
            inp.dispatchEvent(new Event('focus', {bubbles:true}));
            setter.call(inp, val);
            inp.dispatchEvent(new Event('input',  {bubbles:true}));
            inp.dispatchEvent(new Event('change', {bubbles:true}));
            inp.dispatchEvent(new KeyboardEvent('keyup', {bubbles:true}));
            inp.blur();
            inp.dispatchEvent(new Event('blur',     {bubbles:true}));
            inp.dispatchEvent(new Event('focusout', {bubbles:true}));
        """, inp, valor_txt)
        time.sleep(0.1)

    print("    📸 PANTALLA: DESPUÉS DE ESCRIBIR — ANTES DEL SAVE")
    ss(driver, f"A_despues_escribir_antes_save_{cod[:8]}")

    # Esperar a que Angular habilite el SAVE tras procesar los eventos de input.
    for _intento in range(3):
        saves_chk = driver.find_elements(By.CSS_SELECTOR, "tp-button.save > button")
        if any(b.is_displayed() and b.is_enabled() for b in saves_chk):
            break
        print(f"    ⏳ SAVE aún disabled, esperando... (intento {_intento+1}/3)")
        time.sleep(1.5 * VELOCIDAD)

    guardado = None
    try:
        botones = driver.find_elements(By.CSS_SELECTOR, "tp-button.save > button")
        def _rank(b):
            try:
                en = b.is_enabled() and b.is_displayed()
                in_dlg = driver.execute_script(
                    "return !!arguments[0].closest('tp-dialog, tp-modal, [role=\"dialog\"]');", b)
            except Exception:
                en, in_dlg = False, False
            return (2 if (en and in_dlg) else 1 if en else 0)
        botones = sorted(botones, key=_rank, reverse=True)
        for b in botones:
            if b.is_displayed() and b.is_enabled():
                jc(driver, b)
                guardado = "tp-button.save"
                break
    except Exception:
        pass

    if not guardado:
        # Sin SAVE puede significar: (a) la fila ya tenía este valor, o
        # (b) período recién creado por COPY DATE RANGE que ya quedó
        # confirmado sin necesitar un SAVE aparte — ver v2.6 del script
        # base (valorizacion_desde_madre.py) donde se confirmó este caso
        # en producción.
        if abs(valor_viejo - valor) <= 0.011:
            print(f"    ✅ Sin cambios: la fila ya tenía este valor")
            return valor_viejo, valor, etiqueta_fila
        _, valor_actual = _leer_primera_fila_valor(driver)
        if valor_actual is not None and abs(valor_actual - valor) <= 0.30:
            print(f"    ✅ Sin botón SAVE pero el valor ya está aplicado en la grilla")
            return valor_viejo, valor, etiqueta_fila
        ss(driver, f"valor_save_error_{cod[:10]}")
        dump(driver, f"valor_save_error_{cod[:10]}")
        raise Exception(f"No encontré el botón SAVE de la grilla de rates en {cod}")

    print(f"    SAVE clickeado ('{guardado}')")
    time.sleep(4 * VELOCIDAD)
    print("    📸 PANTALLA: DESPUÉS DEL SAVE")
    ss(driver, f"B_despues_save_{cod[:8]}")
    print(f"    ✅ SAVE OK — '{etiqueta_fila}': {valor_viejo} → {valor}")
    return valor_viejo, valor, etiqueta_fila


def cargar_tarifa_option(driver, service_code, supplier, location, service_type,
                          rate_from_str, rate_to_str, price_code, valor, modo="aplicar"):
    """
    Busca un option (location+supplier+service_type+code), entra a RATES,
    ubica el período por rango de fechas + price code (o lo crea copiando
    el último existente) y carga `valor` en las 4 columnas de la primera
    fila de la grilla.

    modo="lectura": no escribe nada — solo reporta si el período existe y
    cuál es el valor actual.
    Devuelve un dict con el detalle del resultado (para el Excel).
    """
    print(f"\n📦 {service_code}  {location or '—'}/{supplier}/{service_type}  "
          f"período {rate_from_str}–{rate_to_str}  PC={price_code or PRICE_CODE_DEFAULT}  valor={valor}")

    buscar_producto(driver, location, supplier, service_code, service_type=service_type)

    hamburger(driver)
    menu_item(driver, "RATES")
    time.sleep(4 * VELOCIDAD)
    ss(driver, f"rates_lista_{service_code}")

    PC     = (price_code or PRICE_CODE_DEFAULT or "").strip().upper()
    IS_ALL = PC in ("", "ALL", "TODOS", "*", "UNASSIGNED")

    def _sel_pc(etq):
        if not IS_ALL:
            _seleccionar_price_code(driver, PC, etiqueta=etq)

    _sel_pc(f"lista {service_code}")
    time.sleep(1.5)

    def _leer_periodos():
        return driver.execute_script("""
            var out = [];
            document.querySelectorAll('td.tpcol-rateperiod').forEach(function(d){
                var tr = d.closest('tr');
                var p = tr ? tr.querySelector('td.tpcol-pricecodecode') : null;
                out.push({date:(d.innerText||'').trim(),
                          pc:(p?(p.innerText||'').trim():'')});
            });
            return out;
        """)

    rf, rt = parsear_fecha(rate_from_str), parsear_fecha(rate_to_str)
    if not (rf and rt):
        raise Exception(f"Fechas de período inválidas: '{rate_from_str}' / '{rate_to_str}'")
    PC_VACIOS = ("", "—", "-", "UNASSIGNED", "ALL")

    def _idx_periodo(rows):
        for i, r in enumerate(rows):
            pf, pt = _parse_rate_period(r["date"])
            if not (pf and pt and abs((pf - rf).days) <= 2 and abs((pt - rt).days) <= 2):
                continue
            pc_row = (r["pc"] or "").strip().upper()
            if IS_ALL:
                if pc_row in PC_VACIOS: return i
            elif pc_row == PC:
                return i
        return None

    periodos    = _leer_periodos()
    idx_periodo = _idx_periodo(periodos)
    etq_pc      = "ALL/Unassigned" if IS_ALL else PC
    print(f"    Períodos: {len(periodos)}  → {'encontrado' if idx_periodo is not None else 'NO encontrado'} "
          f"{etq_pc} {fmt_tp(rf)}–{fmt_tp(rt)}")

    if idx_periodo is None and modo == "lectura":
        print(f"    (lectura) No existe el período → se crearía copiando el último")
        return {"periodo_existia": False, "valor_actual": None, "etiqueta_fila": None}

    pc_warning    = ""
    ya_en_detalle = False
    periodo_encontrado = f"{fmt_tp(rf)}–{fmt_tp(rt)}"
    if idx_periodo is None:
        print(f"    No hay período {etq_pc} para {fmt_tp(rf)}–{fmt_tp(rt)} → COPY último período")
        pc_warning, ya_en_detalle = _copiar_ultimo_period(
            driver, rf, rt, service_code, "" if IS_ALL else PC)
        ss(driver, f"rates_creado_{service_code}")
        if not ya_en_detalle:
            _sel_pc(f"post-insert {service_code}")
            time.sleep(1.5)
            periodos    = _leer_periodos()
            idx_periodo = _idx_periodo(periodos)
            if idx_periodo is None:
                raise Exception(
                    f"El período {etq_pc} {fmt_tp(rf)}–{fmt_tp(rt)} no aparece "
                    f"en la lista de rates de {service_code} tras crearlo.")

    if not ya_en_detalle:
        periodo_encontrado = periodos[idx_periodo]["date"]
        filas_td = driver.find_elements(By.CSS_SELECTOR, "td.tpcol-rateperiod")
        jc(driver, filas_td[idx_periodo])
        time.sleep(5 * VELOCIDAD)
        ss(driver, f"rates_detalle_{service_code}")
    _sel_pc(f"escribir {service_code}")

    if modo == "lectura":
        etiqueta, valor_actual = _leer_primera_fila_valor(driver)
        print(f"    (lectura) Valor actual en '{etiqueta}': {valor_actual}  (se cargaría {valor})")
        return {"periodo_existia": True, "valor_actual": valor_actual, "etiqueta_fila": etiqueta}

    valor_viejo, valor_nuevo, etiqueta_fila = _escribir_valor_unico(driver, valor, service_code)

    # ── Verificación post-SAVE ─────────────────────────────────
    error_post = ""
    try:
        hamburger(driver)
        menu_item(driver, "RATES")
        time.sleep(4 * VELOCIDAD)
        periodos_v = _leer_periodos()
        idx_v      = _idx_periodo(periodos_v)
        if idx_v is None:
            ss(driver, f"verif_sin_periodo_{service_code[:10]}")
            error_post = f"Verificación post-SAVE: el período {periodo_encontrado} no aparece al recargar RATES"
        else:
            filas_td = driver.find_elements(By.CSS_SELECTOR, "td.tpcol-rateperiod")
            jc(driver, filas_td[idx_v])
            time.sleep(5 * VELOCIDAD)
            _sel_pc(f"verif {service_code}")

            leido = None
            for intento in range(1, 4):
                _, leido = _leer_primera_fila_valor(driver)
                if leido is not None and abs(leido - valor_nuevo) <= 0.011:
                    break
                if intento < 3:
                    print(f"    ⚠ Verificación post-SAVE: leído {leido}, esperado {valor_nuevo} "
                          f"(intento {intento}/3) — reintentando...")
                    time.sleep(2 * VELOCIDAD)

            if leido is None or abs(leido - valor_nuevo) > 0.011:
                ss(driver, f"verif_error_{service_code[:10]}")
                dump(driver, f"verif_error_{service_code[:10]}")
                error_post = f"El valor NO quedó guardado en {service_code}: esperado {valor_nuevo}, leído {leido}"
            else:
                print(f"    ✔ Verificación post-SAVE OK: {leido}")
                ss(driver, f"verif_ok_{service_code[:10]}")
    except Exception as _ve:
        error_post = f"Verificación post-SAVE falló: {_ve}"

    if error_post:
        raise Exception(error_post)

    return {"periodo_existia": periodo_encontrado != f"{fmt_tp(rf)}–{fmt_tp(rt)}" or not ya_en_detalle,
            "valor_viejo": valor_viejo, "valor_nuevo": valor_nuevo,
            "etiqueta_fila": etiqueta_fila, "pc_warning": pc_warning}


# ── Excel: cola de trabajo (ver skill armando-excel-como-cola-de-trabajo) ──
_HEADERS = ["LOCATION", "SUPPLIER", "SERVICE TYPE", "SERVICE CODE",
            "RATE FROM", "RATE TO", "PRICE CODE", "VALOR",
            "MOSTRAR CAPTURAS", "ESTADO", "OBSERVACIONES", "TIMESTAMP"]

def crear_excel_si_no_existe():
    if os.path.exists(EXCEL_PATH):
        return False
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET

    fill_h = PatternFill("solid", start_color="1F4E79", fgColor="1F4E79")
    font_h = Font(name="Arial", bold=True, color="FFFFFF", size=9)
    aln_c  = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for col, h in enumerate(_HEADERS, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = font_h; c.fill = fill_h; c.alignment = aln_c

    ws.cell(row=2, column=C["location"],         value="BUE")
    ws.cell(row=2, column=C["supplier"],         value="1EURO1")
    ws.cell(row=2, column=C["service_type"],     value="EX")
    ws.cell(row=2, column=C["service_code"],     value="CSUS01")
    ws.cell(row=2, column=C["rate_from"],        value="01/Sep/2026")
    ws.cell(row=2, column=C["rate_to"],          value="31/Dec/2026")
    ws.cell(row=2, column=C["price_code"],       value="TR")
    ws.cell(row=2, column=C["valor"],            value=100.0)
    ws.cell(row=2, column=C["mostrar_capturas"], value="NO")
    ws.cell(row=2, column=C["estado"],           value="EJEMPLO")

    for col_letter, w in (("A",10),("B",10),("C",13),("D",14),("E",13),
                          ("F",13),("G",11),("H",12),("I",16),("J",14),
                          ("K",40),("L",18)):
        ws.column_dimensions[col_letter].width = w

    wb.save(EXCEL_PATH)
    print(f"✅ Excel creado: {EXCEL_PATH}")
    return True

def leer_pendientes():
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb[SHEET]
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    col_idx = {h: i + 1 for i, h in enumerate(headers) if h}
    rows = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(row):
            continue
        d = dict(zip(headers, row))
        if str(d.get("ESTADO") or "").strip().upper() != "PENDIENTE":
            continue
        d["__row_idx__"] = row_idx
        rows.append(d)
    wb.close()
    return rows

def marcar_procesando(row_idx):
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb[SHEET]
    ws.cell(row=row_idx, column=C["estado"], value="PROCESANDO")
    wb.save(EXCEL_PATH); wb.close()

def escribir_resultado(row_idx, estado, observaciones=""):
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb[SHEET]
    ws.cell(row=row_idx, column=C["estado"], value=estado)
    ws.cell(row=row_idx, column=C["observaciones"], value=observaciones[:500] if observaciones else "")
    ws.cell(row=row_idx, column=C["timestamp"], value=datetime.now().strftime("%Y-%m-%d %H:%M"))
    wb.save(EXCEL_PATH); wb.close()


# ── MAIN ──────────────────────────────────────────────────────
def main():
    global MOSTRAR_CAPTURAS
    print("=" * 60)
    print(f"  CARGA DE TARIFA POR PERÍODO  v{VERSION} ({VERSION_FECHA})  MODO={MODO}")
    print("=" * 60)

    t_inicio = time.time()
    crear_excel_si_no_existe()
    pendientes = leer_pendientes()
    print(f"Filas PENDIENTE: {len(pendientes)}")
    if not pendientes:
        print("\n⛔ Sin filas PENDIENTE. Verificá que la columna ESTADO tenga 'PENDIENTE'.")
        return

    driver = crear_driver()
    _abortado = False
    try:
        login(driver)

        for fila in pendientes:
            chequear_abort()
            row_idx = fila["__row_idx__"]
            print(f"\n{'─' * 60}")
            print(f"Fila {row_idx}: {fila.get('SERVICE CODE')}")

            MOSTRAR_CAPTURAS = str(fila.get("MOSTRAR CAPTURAS") or "").strip().upper() == "SI"
            marcar_procesando(row_idx)

            try:
                valor = _try_float(fila.get("VALOR"))
                if valor is None:
                    raise Exception("Columna VALOR vacía o no numérica")

                resultado = cargar_tarifa_option(
                    driver,
                    service_code=str(fila.get("SERVICE CODE") or "").strip(),
                    supplier=str(fila.get("SUPPLIER") or "").strip(),
                    location=str(fila.get("LOCATION") or "").strip(),
                    service_type=str(fila.get("SERVICE TYPE") or "").strip(),
                    rate_from_str=str(fila.get("RATE FROM") or "").strip(),
                    rate_to_str=str(fila.get("RATE TO") or "").strip(),
                    price_code=str(fila.get("PRICE CODE") or "").strip(),
                    valor=valor,
                    modo=MODO,
                )

                if MODO == "lectura":
                    if resultado["periodo_existia"]:
                        obs = (f"Período existente. Valor actual en '{resultado['etiqueta_fila']}': "
                               f"{resultado['valor_actual']} → se cargaría {valor}")
                    else:
                        obs = "El período NO existe: se crearía copiando el último período existente"
                    escribir_resultado(row_idx, "OK", obs)
                else:
                    obs = f"'{resultado['etiqueta_fila']}': {resultado['valor_viejo']} → {resultado['valor_nuevo']}"
                    if resultado.get("pc_warning"):
                        obs += f" | ⚠ {resultado['pc_warning']}"
                    escribir_resultado(row_idx, "OK", obs)
                print(f"  ✅ {obs}")

            except Exception as e:
                print(f"  ❌ ERROR: {e}")
                escribir_resultado(row_idx, f"ERROR: {str(e)[:280]}", "")

    except AbortadoPorUsuario:
        _abortado = True
        print("\n⏸️  Corrida abortada por el usuario — las filas que no llegaron a "
              "procesarse quedan en PENDIENTE para retomar en otra corrida.")
    finally:
        logout(driver)
        driver.quit()
        dur = int(time.time() - t_inicio)
        m, s = divmod(dur, 60)
        print(f"\n🏁 Fin. Duración: {m}m {s:02d}s")
        print(f"📄 Excel: {EXCEL_PATH}")
        try:
            from google.colab import files
            files.download(EXCEL_PATH)
        except Exception:
            pass

    if _abortado:
        sys.exit(ABORT_EXIT_CODE)


if __name__ == "__main__":
    main()
