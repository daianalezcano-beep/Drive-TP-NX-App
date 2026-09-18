# ============================================================
# TOURPLAN NX — VALORIZACIÓN DESDE SERVICIO MADRE (CÓDIGOS NUMÉRICOS)  v1.0
# Google Colab — celda única
# ------------------------------------------------------------
#   VERSION      : 1.0
#   VERSION_FECHA: 2026-09-01
#   DESCRIPCIÓN  : Variante de valorizacion_desde_madre.py (script
#                  original, carpeta scripts/valorizacion_madre/) que
#                  invierte el filtro de códigos: el script original
#                  DESCARTA los servicios madre cuyo código empieza con
#                  un número; este script procesa EXCLUSIVAMENTE esos
#                  códigos (los que el original descarta) y descarta
#                  a su vez los que empiezan con letra. Fuera de ese
#                  filtro invertido, el resto del flujo es idéntico al
#                  original: busca todos los servicios madre (supplier +
#                  service type + location opcional), clasifica Package
#                  vs Non Accommodation, abre cada Package en USED IN
#                  buscando PCMs tipo "Package Header" con nombre PKG-*,
#                  lee costos y aplica rates.
#   HISTORIAL v1.2 a v2.11: idéntico al script original hasta esa
#                  versión (ver scripts/valorizacion_madre/valorizacion_desde_madre.py
#                  para el detalle de cada cambio heredado).
#   CAMBIOS v1.2 : - Filtro códigos no alfabéticos en búsqueda
#                  - Manejo ventana tarifas vencidas (SAVE ALL)
#                    → marca "COMPONENTES VENCIDOS" en salida
#                  - Wait 10s (antes 6s) tras cambio de fecha
#                  - Fecha base date en formato DD/MM/YY
#                  - STYPE_SIDEBAR para selección en sidebar
#   CAMBIOS v1.3 : - Nueva col H "SERVICE CODE" (opcional):
#                    si tiene valor → procesa solo ese servicio
#                    si está vacía  → procesa todos (comportamiento anterior)
#                  - Columnas de salida desplazadas a I-M
#   CAMBIOS v1.4 : - Fix cambiar_fecha_pcm(): click directo en
#                    li[1] (OLD TRAVEL DATE) en lugar de esperar
#                    sugerencia de datepicker (que nunca aparece)
#                  - Fix diálogo recálculo: flujo exacto de
#                    grabación Chrome DevTools (click directo en
#                    #calculatereplaceall via jc, labels exchange)
#                  - Ventana tarifas vencidas movida a DESPUÉS
#                    del diálogo de recálculo (orden correcto)
#   CAMBIOS v1.5 : - Nueva lógica de mapeo PCM → rangos madre:
#                    rangos 9999 con pmin≥42 → 0.0; con pmin<42
#                    → 999.0 (alerta); rangos normales → max()
#                    de todos los valores PCM en [pmin, pmax];
#                    sin match → error en pendientes + 0.0
#   CAMBIOS v1.6 : - Fix parseo de valores con separador de
#                    miles: replace(",","") en lugar de (",",".")
#                    en todos los puntos de parseo (_try_float,
#                    leer_markup_commission, _escribir_rates)
#   CAMBIOS v1.7 : - (reemplazado por v1.8)
#   CAMBIOS v1.8 : - Período nuevo: COPY DATE RANGE basado en
#                    grabación exacta Chrome DevTools.
#                    tp-button.copydaterange > button; dialog
#                    tp-dialog:nth-of-type(2) li:nth-of-type(3);
#                    formato TO DD/MM/YY; tp-button.ok.
#                    Tras OK ya estamos en detalle → no reabre
#                    desde lista (ya_en_detalle=True).
#                    Fallback a INSERT si no hay copydaterange.
#   CAMBIOS v1.9 : - COPY DATE RANGE: set FROM (li:nth-of-type(2))
#                    desde Excel además del TO (li:nth-of-type(3)).
#                    Ambas fechas siempre tomadas del Excel input.
#   CAMBIOS v2.0 : - Fix COPY DATE RANGE: si campo FROM o TO no
#                    encontrado en dialog → cancela dialog y hace
#                    fallback a INSERT normal (antes quedaba en
#                    blanco y Tourplan rechazaba con missing info).
#   CAMBIOS v2.1 : - Fix COPY DATE RANGE: tomaba periodos_rows[-1]
#                    (el más viejo, último en el DOM) en lugar de
#                    periodos_rows[0] (el más reciente, primero
#                    visualmente en la lista de arriba hacia abajo).
#   CAMBIOS v2.2 : - Nueva constante VELOCIDAD (1.0=test, 1.5=prod):
#                    multiplica los 35 sleeps de carga de página,
#                    dialogs y saves (≥3s). Microsleeps de eventos
#                    Angular (≤0.5s) permanecen fijos.
#   CAMBIOS v2.5 : - Nueva col N "FECHA RECALCULATE PCM" (opcional):
#                    si tiene valor, se usa como base del Change Base
#                    Date del PCM en vez de RATE FROM (manteniendo la
#                    regla de "siguiente día hábil" si no opera esa
#                    fecha exacta). Vacía → usa RATE FROM, igual que
#                    antes. RATE FROM sigue siendo la base para crear/
#                    buscar el período de rates a valorizar — sin
#                    cambios ahí.
#   CAMBIOS v2.6 : - Fix _copiar_ultimo_period (COPY DATE RANGE): elegía
#                    el período fuente a copiar por posición (el más
#                    reciente de la lista), ignorando su price code. Si
#                    la lista tenía períodos de varios price codes
#                    mezclados, podía extender un período con el price
#                    code equivocado sin ningún aviso (confirmado por la
#                    usuaria en corrida real de valorizacion_desde_excel.py).
#                    Ahora se lee el price code real de cada fila
#                    (td.tpcol-pricecodecode) y se elige la más reciente
#                    que coincida con el price code pedido.
#   CAMBIOS v2.7 : - STYPE_SIDEBAR ampliado al glosario completo (12
#                    siglas en vez de solo EX/TF), sincronizado con
#                    valorizacion_desde_excel.py.
#   CAMBIOS v2.8 : - La grilla de resultados de Product Search NO trae
#                    una columna Category poblada (confirmado por la
#                    usuaria) — _resolver_fila_fase1() asumía "Package"
#                    para TODO, sin distinguir Non Accommodation, con
#                    riesgo real de procesar un servicio que no
#                    correspondía. Ahora la clasificación real ocurre en
#                    _buscar_servicio_fase1(), leyendo el dropdown
#                    "Service Category" del producto ya abierto
#                    (leer_service_category(), sin navegar a ningún lado:
#                    Tourplan ya deja parada esa pantalla al seleccionar
#                    el producto desde Product Search) ANTES de ir a Used
#                    In — si es Non Accommodation,
#                    corta ahí y evita el escaneo caro de Used In
#                    (scroll virtual de hasta ~1900 filas) para nada.
#   CAMBIOS v2.9 : - Fix en buscar_producto() y buscar_servicios_madre():
#                    al completar Location, si la fila `tr.selectedRow`
#                    no aparecía a tiempo, el fallback clickeaba siempre
#                    la PRIMERA fila de la tabla de sugerencias sin
#                    verificar que fuera el Location exacto (mismo bug ya
#                    arreglado en notas_srv.py v1.13/v1.14). Ahora busca
#                    entre las filas sugeridas la que tiene una celda con
#                    texto EXACTO (case-insensitive) igual al Location
#                    pedido; si ninguna coincide, no clickea nada.
#   CAMBIOS v2.10: - Fix diálogo de recálculo del PCM: el checkbox
#                    "Update Exchange Rates" quedaba destildado pese al
#                    click (confirmado por la usuaria en corrida real,
#                    "como si no pegara"). Clickeaba su <label>
#                    (#exchange-rate-panel > tp-label > label) en vez del
#                    <input> — ahora se clickea #updateexchangerates
#                    directo (mismo patrón que #calculatereplaceall) y se
#                    VERIFICA que quede tildado, reintentando si hace
#                    falta.
#   CAMBIOS v2.11: - Corrección al fix de v2.10: quedaba un segundo click
#                    sin sacar, "tp-group.tpgroup-servicelinepricesgroup
#                    tp-checkbox > label", que una grabación nueva de
#                    Chrome DevTools confirma (por xpath) que es hijo del
#                    MISMO #exchange-rate-panel — es decir, el MISMO
#                    checkbox "Update Exchange Rates" clickeado por su
#                    widget en vez de su texto. Ese segundo click volvía a
#                    destildarlo justo después de que el fix de v2.10 lo
#                    dejara tildado. Se saca ese click; con el click
#                    directo sobre #updateexchangerates + verificación de
#                    v2.10 alcanza.
# ============================================================

VERSION       = "1.0"
VERSION_FECHA = "2026-09-01"

# ── PASO 0: Entorno ──────────────────────────────────────────
import os, sys, subprocess, importlib.util, shutil, time, re, json
from datetime import datetime, timedelta

_t_inicio = time.time()

print("🔧 Verificando entorno...\n")

_PIPS_NEEDED = {
    "selenium":            "selenium",
    "webdriver_manager":   "webdriver-manager",
    "gspread":             "gspread",
    "google_auth_oauthlib": "google-auth-oauthlib",
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
# 0.3 Botón Abortar de la app (ver common/abort.py)
from common.abort import chequear_abort, AbortadoPorUsuario, ABORT_EXIT_CODE
# 0.4 Google Sheets como cola de trabajo (ver common/sheets_client.py)
from common.sheets_client import conectar_sheets, cargar_sheet, actualizar_fila_sheet, asegurar_columnas, agregar_fila_sheet
from common.user_config import (
    CREDENTIALS_PATH as _CREDENTIALS_PATH_DEFAULT,
    TOKEN_PATH as _TOKEN_PATH_DEFAULT,
)

CHROMIUM_BIN, ver_chrome = find_or_prepare_chrome()

from webdriver_manager.chrome import ChromeDriverManager

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys

# ── Config — via variables de entorno (con default = valor original) ──
USERNAME         = os.environ.get("TOURPLAN_USERNAME", "poner minusculas")
PASSWORD         = os.environ.get("TOURPLAN_PASSWORD", "password")
BASE_URL         = os.environ.get("TOURPLAN_BASE_URL", "https://tourplannx.eurotur.com.ar/TourplanNX_Test")
SHEET_URL        = os.environ.get("TOURPLAN_SHEET_URL", "")
SHEET            = os.environ.get("TOURPLAN_HOJA", "PRODUCTOS")
CREDENTIALS_PATH = os.environ.get("TOURPLAN_CREDENTIALS_PATH", _CREDENTIALS_PATH_DEFAULT)
TOKEN_PATH       = os.environ.get("TOURPLAN_TOKEN_PATH", _TOKEN_PATH_DEFAULT)
SS_DIR           = os.environ.get("TOURPLAN_SS_DIR", "/content/screenshots")
HEADLESS         = os.environ.get("TOURPLAN_HEADLESS", "0").strip() in ("1", "true", "True")
os.makedirs(SS_DIR, exist_ok=True)

# ── Modo de ejecución ─────────────────────────────────────────
# "LEER"    → solo Fase 1: busca servicios madre, lee PCMs, guarda en Excel
# "APLICAR" → solo Fase 2: lee PCM_Detail y aplica rates en servicios madre
# "COMPLETO"→ Fase 1 + Fase 2 en la misma sesión
MODO = os.environ.get("TOURPLAN_MODO", "COMPLETO")

# Multiplicador de tiempos de espera.
# 1.0 = entorno de prueba  |  1.5 = producción (respuestas más lentas)
VELOCIDAD = float(os.environ.get("TOURPLAN_VELOCIDAD", "1.5"))

# Límite de servicios a procesar en Fase 2 (0 = sin límite)
FASE2_LIMIT = 0

# Si True, Fase 2 reprocesa TODAS las filas con SERVICE CODE sin importar ESTADO F2
FASE2_REAPLICAR_TODO = False

# Price Code por defecto para el período de rates del servicio madre
PRICE_CODE_DEFAULT = "TR"

# Mapping service type → número de opción en el sidebar de Product Search.
# Usado para seleccionar el tipo correcto de forma más robusta. Glosario
# confirmado por la usuaria (2026-08), ver skill
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
# categoría real (ej. TN="Transfer Non-Accom" vs TF="Transfer") — usadas para
# no matchear la sigla real dentro de un ítem "no usar" ni viceversa.
Z_NO_USAR_SIGLAS = {"GA", "TK", "TN", "CH", "CM", "CO", "EN", "GU", "PJ", "ST", "TR", "TI"}

# Columnas Excel PRODUCTOS (base 1)
C = {
    "location":         1,   # A - opcional, 3 letras (BUE). Vacío = sin filtro
    "supplier":         2,   # B - obligatorio
    "service_type":     3,   # C - obligatorio (EX, TF, etc.)
    "rate_from":        4,   # D - inicio período dd/Mon/yyyy
    "rate_to":          5,   # E - fin período dd/Mon/yyyy
    "price_code":       6,   # F - TR, 34, etc. Vacío/ALL = Unassigned
    "service_code":     8,   # H - opcional: si se especifica, procesa solo ese servicio madre
    "servicios_pkg":    9,   # I - SALIDA: códigos Package procesados
    "servicios_skip":  10,   # J - SALIDA: códigos saltados
    "estado":          11,   # K - PENDIENTE para procesar
    "error_msg":       12,   # L - SALIDA: error
    "timestamp":       13,   # M - SALIDA: timestamp
    "fecha_recalculate_pcm": 14,  # N - opcional: fecha base para el
                                  # recalculate/Change Base Date del PCM.
                                  # Vacía → usa RATE FROM (comportamiento
                                  # de siempre). RATE FROM sigue usándose
                                  # igual para crear/buscar el período de
                                  # rates a valorizar.
}

RATES_START     = 13   # columnas de rates dinámicos empiezan en M

_ss_n = [0]

# ── Helpers base ──────────────────────────────────────────────
def ss(driver, nombre):
    _ss_n[0] += 1
    p = f"{SS_DIR}/{_ss_n[0]:03d}_{nombre[:40]}_{int(time.time())}.png"
    driver.save_screenshot(p)
    print(f"  📸 {os.path.basename(p)}")

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
    # Sin ventana visible solo si se pide explícitamente (TOURPLAN_HEADLESS,
    # checkbox en Configuración) — por default corre con ventana real en la
    # PC de la persona, a diferencia del contenedor sin pantalla de Colab.
    if HEADLESS:
        opts.add_argument("--headless=new")
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
    try:
        items = driver.execute_script("""
            return Array.from(document.querySelectorAll('nav ul > li')).map(function(li,i){
                var txt = (li.querySelector('div div') || li).innerText.trim().split('\\n')[0];
                return (i+1) + ': ' + txt.slice(0,30);
            });
        """)
        print(f"    Menu items: {items}")
    except: pass

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

def submenu_text(driver, text, timeout=8):
    xpath = (f"//li[contains(@class,'tpnavmenuopen')]"
             f"//label[contains(translate(normalize-space(text()),"
             f"'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),"
             f"'{text.upper()}')]")
    el = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH, xpath)))
    jc(driver, el)
    time.sleep(2 * VELOCIDAD)

# ── Fechas ────────────────────────────────────────────────────
MESES_ES = {"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
            "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}

def _try_float(s):
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

def fmt_xl(dt):
    meses = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
             7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
    return f"{dt.day:02d}/{meses[dt.month]}/{dt.year}"

DIAS_SEMANA = {0:"MON",1:"TUE",2:"WED",3:"THU",4:"FRI",5:"SAT",6:"SUN"}
DIAS_IDX    = {v:k for k,v in DIAS_SEMANA.items()}

def siguiente_dia_operacion(dias_activos, desde=None):
    if not dias_activos:
        dias_activos = set(DIAS_SEMANA.values())
    base = desde or datetime.now()
    for delta in range(8):
        candidato = base + timedelta(days=delta)
        if DIAS_SEMANA[candidato.weekday()] in dias_activos:
            return candidato
    return base


# ── Buscar producto por código (para abrir servicio madre específico) ──
class ProductoNoEncontrado(Exception):
    pass

def buscar_producto(driver, location, supplier, codigo, service_type=None, handle_origen=None):
    """
    Abre un producto específico en Product Setup.
    Reutilizado de v3 sin cambios.
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
        except Exception as e:
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


# ── Buscar TODOS los servicios madre (sin código, por supplier+type+location) ──
def buscar_servicios_madre(driver, supplier, service_type, location=""):
    """
    Busca en Product Search TODOS los productos del supplier + service type
    (+ location si se especifica). Lee código y service category de cada fila.
    Retorna lista de dicts: {code, service_category, description}.
    """
    st_upper = service_type.strip().upper()
    print(f"\n  🔍 Buscando servicios madre: supplier={supplier} type={st_upper} "
          f"loc={location or '(todos)'}")

    driver.get(f"{BASE_URL}/#/home")
    time.sleep(2 * VELOCIDAD)
    driver.get(f"{BASE_URL}/#/product")
    time.sleep(5 * VELOCIDAD)
    ss(driver, f"sm_inicial_{supplier[:10]}")

    lupa = wait(driver, "#searchWrapper li:nth-of-type(2) button")
    jc(driver, lupa)
    time.sleep(3 * VELOCIDAD)

    # Asegurar tab SELECTION activo
    try:
        WebDriverWait(driver, 4).until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "div.parameters1 input")))
    except Exception:
        driver.execute_script("""
            for (var el of document.querySelectorAll('li,button,a,div,span')){
                if (!el.offsetParent) continue;
                if ((el.innerText||'').trim().toUpperCase() === 'SELECTION'){
                    el.click(); break;
                }
            }
        """)
        time.sleep(2 * VELOCIDAD)

    # Seleccionar service type en sidebar. Traduce la sigla al prefijo
    # numérico confirmado (STYPE_SIDEBAR) si existe; si no, busca la sigla
    # cruda — siempre por palabra completa (boundary regex), nunca substring
    # suelto (ver comentario largo en buscar_producto()). Excluye los ítems
    # "Z*NO USAR*" cuando la sigla buscada no es una de esas 12, para no
    # matchear p.ej. 'TF' (Transfer) dentro de "Z*NO USAR* - TN - Transfer
    # Non-Accom" por compartir la palabra "Transfer".
    if st_upper:
        st_buscar = STYPE_SIDEBAR.get(st_upper, st_upper)
        excluir_no_usar = st_upper not in Z_NO_USAR_SIGLAS
        patron_js = (r"(^|[\s\-–—/(])" + re.escape(st_buscar) + r"([\s\-–—/)]|$)")
        for intento in range(3):
            ok = driver.execute_script("""
                var rx = new RegExp(arguments[0]);
                var excluirNoUsar = arguments[1];
                for (var li of document.querySelectorAll('li')){
                    if (!li.offsetParent) continue;
                    var t = (li.textContent || '').trim();
                    if (!t || t.length > 80) continue;
                    if (excluirNoUsar && t.toLowerCase().indexOf('no usar') !== -1) continue;
                    if (rx.test(t.toUpperCase())){ li.click(); return t.slice(0,40); }
                }
                return null;
            """, patron_js, excluir_no_usar)
            if ok:
                print(f"    Service type '{st_upper}' → {ok}")
                time.sleep(1.5)
                break
            time.sleep(1.5)
        else:
            print(f"    ⚠ Service type '{st_upper}' no encontrado en sidebar")

    # Location (opcional)
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
            print(f"    Location: {location}")
        except Exception as e:
            print(f"    ⚠ Location no pudo llenarse: {e}")

    # Supplier
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
                    inp_sup.send_keys(Keys.TAB); time.sleep(1)
                except: pass
            print(f"    Supplier: {supplier}")
        except Exception as e:
            print(f"    ⚠ Supplier no pudo llenarse: {e}")

    # NO llenar código → búsqueda devuelve todos los del supplier+type

    ss(driver, f"sm_modal_lleno_{supplier[:10]}")
    btn_search = wait(driver, "#productSearchFilter li:nth-of-type(4) button")
    jc(driver, btn_search)
    time.sleep(6 * VELOCIDAD)
    ss(driver, f"sm_resultados_{supplier[:10]}")

    # ── Leer todos los resultados ─────────────────────────────
    # La grilla de resultados tiene columnas: Location | Type | Supplier | Code | Description | Category (y otras)
    # Leemos los headers de la tabla para identificar las columnas correctamente.
    servicios = []
    vistos    = set()

    def _leer_batch():
        return driver.execute_script("""
            var result = [];
            for (var tbl of document.querySelectorAll('table')) {
                var ths = Array.from(tbl.querySelectorAll('thead th, thead td'));
                var headers = ths.map(function(h){
                    return (h.innerText || h.textContent || '').trim().toUpperCase();
                });
                // tabla de resultados de búsqueda tiene al menos Code y Description
                var hasCode = headers.some(function(h){ return /^CODE$/.test(h); });
                if (!hasCode && headers.length < 3) continue;

                // índices de las columnas que nos importan
                var idxCode = -1, idxCat = -1, idxDesc = -1;
                headers.forEach(function(h, i){
                    if (/^CODE$/.test(h)) idxCode = i;
                    if (/CATEGORY|CATEGOR/i.test(h) && idxCat < 0) idxCat = i;
                    if (/^DESCRIPTION$/.test(h) && idxDesc < 0) idxDesc = i;
                });

                Array.from(tbl.querySelectorAll('tbody tr')).forEach(function(tr){
                    if (!tr.offsetParent) return;
                    var cells = Array.from(tr.querySelectorAll('td')).map(function(td){
                        return (td.innerText || td.textContent || '').trim();
                    });
                    if (!cells.some(function(c){ return c; })) return;

                    var code = idxCode >= 0 ? (cells[idxCode] || '') : '';
                    var cat  = idxCat  >= 0 ? (cells[idxCat]  || '') : '';
                    var desc = idxDesc >= 0 ? (cells[idxDesc] || '') : '';

                    // fallback si no hay header Code: buscar celda corta uppercase alfanumérica
                    if (!code) {
                        for (var c of cells) {
                            if (/^[A-Z0-9]{2,10}$/.test(c)) { code = c; break; }
                        }
                    }
                    // fallback category: buscar celda con texto reconocible
                    if (!cat) {
                        for (var c2 of cells) {
                            if (/package|non.accom|accommodation/i.test(c2)) {
                                cat = c2; break;
                            }
                        }
                    }

                    if (code) result.push({code: code, cat: cat, desc: desc, cells: cells});
                });
                if (result.length) break;
            }
            return result;
        """) or []

    def _get_scroll():
        return driver.execute_script("""
            var cdk = document.querySelector('cdk-virtual-scroll-viewport');
            if(cdk) return cdk;
            var tbl = document.querySelector('table');
            if(tbl){
                var el = tbl.parentElement;
                while(el && el !== document.body){
                    var st = getComputedStyle(el);
                    if((st.overflow==='auto'||st.overflow==='scroll'||
                        st.overflowY==='auto'||st.overflowY==='scroll')
                       && el.scrollHeight > el.clientHeight+10) return el;
                    el = el.parentElement;
                }
            }
            return document.scrollingElement || document.body;
        """)

    scroll_cont = _get_scroll()
    scroll_pos  = 0
    sin_cambio  = 0

    for _ in range(800):
        batch = _leer_batch()
        nuevas = 0
        for item in batch:
            code = item.get("code", "").strip().upper()
            if not code or code in vistos:
                continue
            vistos.add(code)
            nuevas += 1
            servicios.append({
                "code":             code,
                "service_category": item.get("cat", "").strip(),
                "description":      item.get("desc", "").strip(),
            })
        if nuevas == 0:
            sin_cambio += 1
            if sin_cambio >= 8:
                break
        else:
            sin_cambio = 0
        scroll_pos += 150
        if scroll_cont:
            driver.execute_script("arguments[0].scrollTop = arguments[1];",
                                  scroll_cont, scroll_pos)
        driver.execute_script("window.scrollTo(0, arguments[0]);", scroll_pos)
        time.sleep(0.25)

    # diagnóstico antes del filtro
    if servicios:
        cats = {}
        for s in servicios:
            cats[s["service_category"]] = cats.get(s["service_category"], 0) + 1
        print(f"    Total servicios encontrados: {len(servicios)}  categorías: {cats}")
        print(f"    Ejemplo: {servicios[0]}")
    else:
        print(f"    ⚠ Ningún servicio encontrado para supplier={supplier} type={st_upper}")
        dump(driver, f"sm_sin_resultados_{supplier[:10]}")

    # Filtro INVERTIDO respecto al script original: acá nos quedamos
    # exclusivamente con los códigos que empiezan con un número (los que
    # valorizacion_desde_madre.py descarta), y descartamos los que
    # empiezan con letra.
    servicios_todos = servicios[:]
    servicios = [s for s in servicios if not s["code"][:1].isalpha()]
    if len(servicios) < len(servicios_todos):
        filtrados = [s["code"] for s in servicios_todos if s["code"][:1].isalpha()]
        print(f"    Filtrados {len(filtrados)} código(s) alfabético(s): {filtrados[:20]}")

    return servicios


# ── Leer días de operación ────────────────────────────────────
def leer_dias_operacion(driver):
    DIAS_VALIDOS = {"MON","TUE","WED","THU","FRI","SAT","SUN"}
    try:
        dias = driver.execute_script("""
            var result = {};
            var DIAS = ['MON','TUE','WED','THU','FRI','SAT','SUN'];
            var labels = document.querySelectorAll('label.tplabel-startson');
            if(!labels.length){
                labels = Array.from(document.querySelectorAll('label')).filter(function(l){
                    var t = l.innerText.trim().toUpperCase().slice(0,3);
                    return ['MON','TUE','WED','THU','FRI','SAT','SUN'].includes(t);
                });
            }
            labels.forEach(function(lbl){
                var dia = lbl.innerText.trim().toUpperCase().slice(0,3);
                if(!['MON','TUE','WED','THU','FRI','SAT','SUN'].includes(dia)) return;
                var parent = lbl.closest('li, div, span, td') || lbl.parentElement;
                var inp = parent ? parent.querySelector('input[type=checkbox]') : null;
                if(inp){ result[dia] = inp.checked; return; }
                var ico = parent ? parent.querySelector('i') : null;
                if(ico){
                    var cls = ico.className || '';
                    result[dia] = cls.includes('check') && !cls.includes('uncheck') && !cls.includes('times');
                    return;
                }
                var tpchk = parent ? parent.querySelector('[class*="checkbox"]') : null;
                if(tpchk){
                    var cls2 = (tpchk.className || '') + ' ' + (tpchk.getAttribute('ng-reflect-model') || '');
                    result[dia] = cls2.toLowerCase().includes('true') || cls2.includes('checked');
                    return;
                }
                result[dia] = true;
            });
            return result;
        """)
        activos = {k for k, v in (dias or {}).items()
                   if v and k.upper() in DIAS_VALIDOS}
        if not activos:
            print("    ⚠ No se leyeron días operación → asumiendo todos")
            activos = set(DIAS_SEMANA.values())
        print(f"    Días operación: {sorted(activos)}")
        return activos
    except Exception as e:
        print(f"    ⚠ No pude leer días de operación: {e} → asumiendo todos")
        return set(DIAS_SEMANA.values())


# ── USED IN ───────────────────────────────────────────────────
def ir_a_used_in(driver):
    """
    Product hamburger → Used In. El menú de Tourplan es CONTEXTUAL —
    confirmado en corrida real comparando los 'Menu items' impresos por
    hamburger() en dos llamados seguidos para el mismo producto: desde la
    vista neutra recién abierto el producto, 'USED IN' aparece ANIDADO
    dentro de UTILITIES (como siempre asumió este código); pero si ya se
    había navegado a Used In antes en la misma sesión de producto (esto
    pasa ahora que _procesar_pcm_grupo_fase1() vuelve a llamar a
    ir_a_used_in() para reabrir el PCM, tras ya haber estado ahí en
    _buscar_servicio_fase1()), Tourplan lo muestra como ÍTEM DE PRIMER
    NIVEL directo, sin UTILITIES en el medio — buscar "Used In" adentro
    del submenú expandido de UTILITIES entonces nunca lo encontraba
    (timeout sin mensaje útil, solo el stacktrace de chromedriver).

    Fix: probar primero el ítem de primer nivel VISIBLE (sin expandir
    nada); si no aparece, caer al camino de siempre (UTILITIES → Used In
    anidado), sin cambios.
    """
    hamburger(driver)
    ss(driver, "hamburger_product")

    clickeado_directo = driver.execute_script("""
        function vis(e){ return !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length); }
        var items = Array.from(document.querySelectorAll('nav ul > li')).filter(vis);
        for (var li of items) {
            var t = (li.innerText || '').trim().toUpperCase().split('\\n')[0];
            if (t.indexOf('USED IN') === 0) {
                var target = li.querySelector(':scope > div > div') || li.querySelector('div div') || li;
                target.click();
                return true;
            }
        }
        return false;
    """)

    if clickeado_directo:
        print("    'Used In' encontrado como ítem de primer nivel — click directo")
        time.sleep(2 * VELOCIDAD)
    else:
        menu_item(driver, "UTILITIES")
        ss(driver, "utilities_expandido")
        submenu_text(driver, "Used In")

    time.sleep(5 * VELOCIDAD)
    ss(driver, "used_in_tabla")


def leer_pcm_list_package_header(driver):
    """
    Lee la tabla Used In y retorna lista de PCMs válidos.
    Filtros específicos para este script:
    - tipo debe contener "Package Header" (exacto, no solo "Package")
    - nombre debe empezar con "PKG-"
    Si hay múltiples PCMs Package Header válidos, retorna todos
    (el llamador toma el primero y loguea advertencia si hay más).

    Antes de escrollear, intenta una PASADA RÁPIDA (paso 0): ordena la
    grilla por la columna "Date" (los PCM Package Header siempre la traen
    en blanco, confirmado por la usuaria con captura real) y revisa si las
    filas con Date en blanco ya cumplen los filtros. Si es así, evita el
    escaneo completo con scroll. Es best-effort y no reemplaza nada: si el
    ordenamiento falla, no hay filas con Date en blanco, o esas filas no
    pasan los filtros, se sigue exactamente con el procedimiento de
    siempre (ordenar por Type + escrollear toda la lista).
    """
    def _leer_filas_visibles():
        return driver.execute_script("""
            var rows = [];
            var tbl = document.querySelector('#usedin table') ||
                      document.querySelector('tp-usedin table') ||
                      document.querySelector('[class*="usedin"] table') ||
                      document.querySelector('table');
            if(!tbl) return rows;
            var ths = tbl.querySelectorAll('thead th');
            var headers = Array.from(ths).map(h => h.innerText.trim());
            if(!headers.length){
                var firstRow = tbl.querySelector('tr');
                if(firstRow)
                    headers = Array.from(firstRow.querySelectorAll('th,td')).map(h=>h.innerText.trim());
            }
            tbl.querySelectorAll('tbody tr').forEach(function(tr){
                var celdas = Array.from(tr.querySelectorAll('td')).map(td => td.innerText.trim());
                var obj = {};
                headers.forEach(function(h,i){ obj[h] = celdas[i] || ''; });
                obj['_fila_txt'] = celdas.join(' | ');
                rows.push(obj);
            });
            return rows;
        """)

    def _valor_columna_date(info):
        """Valor de la columna cuyo encabezado es exactamente 'Date'
        (la misma que clickea _ordenar_used_in_por_date). None si esa
        columna no está presente en la fila leída."""
        for k, v in info.items():
            if k != "_fila_txt" and k.strip().lower() == "date":
                return v
        return None

    def _procesar_fila_package_header(info, el_fila):
        """Aplica los filtros de PCM Package Header válido a una fila ya
        leída de la grilla Used In. Retorna el dict a agregar a pcm_list,
        o None si la fila no pasa. Compartida por la pasada rápida (Date)
        y la pasada de siempre (Type + scroll) para que ambas den
        exactamente el mismo resultado."""
        tipo = ""
        for k in info:
            if k.upper() in ("TYPE","TIPO","USAGETYPE","USAGETYPELABEL","USAGE TYPE"):
                tipo = info[k]; break
        if not tipo:
            tipo = next((info[k] for k in info
                         if "type" in k.lower() and k != "_fila_txt"), "")

        nombre = ""
        for k in info:
            if any(x in k.lower() for x in ["pcmname","bookingpcm","booking pcm","pcm name","pcm ref"]):
                nombre = info[k]; break
        if not nombre:
            nombre = next((v for k,v in info.items()
                           if str(v).startswith("PKG-") and k != "_fila_txt"), "")
        if not nombre:
            nombre = next((v for k,v in info.items()
                           if re.match(r'[A-Z]{2,}-', str(v)) and k != "_fila_txt"), "")
        if not nombre:
            return None

        nombre_limpio = re.split(r'\s*--\s*', nombre)[0].strip()
        tipo_up       = tipo.strip().upper()
        nombre_up     = nombre_limpio.upper()
        fila_txt_up   = info.get("_fila_txt","").upper()

        if any(x in nombre_up for x in ["FLAG AS DELETED","FLAGGED","DELETED","CANCELLED","CANCELED","INACTIVE"]):
            return None
        if any(x in fila_txt_up for x in ["FLAG AS DELETED","FLAGGED"]):
            return None
        if any(x in tipo_up for x in ["DELETE","FLAG","CANCEL","INACTIVE"]):
            return None
        if "PACKAGE HEADER" not in tipo_up:
            return None
        if not nombre_up.startswith("PKG-"):
            return None

        fecha_desde_str = info.get("From","") or info.get("FROM","") or ""
        fecha_hasta_str = info.get("To","")   or info.get("TO","")   or ""
        f_desde = parsear_fecha(fecha_desde_str)
        f_hasta = parsear_fecha(fecha_hasta_str)
        if f_desde and f_hasta and (f_hasta - f_desde).days > 365:
            return None

        return {"nombre": nombre_limpio, "tipo": tipo, "fila_el": el_fila, "info": info}

    # ── Paso 0 (rápido): ordenar por Date y revisar solo las filas con
    # Date en blanco ─────────────────────────────────────────────────────
    if _ordenar_used_in_por_date(driver):
        filas_rapidas = _leer_filas_visibles()
        candidatos = []
        for info in filas_rapidas:
            val = _valor_columna_date(info)
            if val is None:
                break   # no reconocí la columna "Date" en esta fila → no confiar en el orden
            if val.strip() != "":
                break   # ya empezaron las filas con fecha → no quedan más en blanco
            candidatos.append(info)

        if candidatos:
            filas_dom_rapido = driver.find_elements(By.CSS_SELECTOR, "#usedin table tbody tr")
            if not filas_dom_rapido:
                filas_dom_rapido = driver.find_elements(By.XPATH, "//table//tbody//tr")
            pcm_list_rapido = []
            for i, info in enumerate(candidatos):
                el_fila = filas_dom_rapido[i] if i < len(filas_dom_rapido) else None
                resultado = _procesar_fila_package_header(info, el_fila)
                if resultado:
                    pcm_list_rapido.append(resultado)

            if pcm_list_rapido:
                print(f"    ✅ Pasada rápida (orden por Date): "
                      f"{len(pcm_list_rapido)} PCM Package Header válido(s) "
                      f"sin necesidad de escrollear toda la lista: "
                      f"{[p['nombre'] for p in pcm_list_rapido]}")
                return pcm_list_rapido

        print("    ⚠ Pasada rápida (orden por Date) no encontró un Package "
              "Header válido en las filas con Date en blanco — sigo con el "
              "procedimiento completo (Type + scroll)")

    # ── Paso 1 (fallback, sin cambios): ordenar por Type y escrollear toda
    # la lista ────────────────────────────────────────────────────────────
    try:
        th = wait(driver, "th.tpcol-UsageTypeLabel > span > span", t=5)
        jc(driver, th); time.sleep(2 * VELOCIDAD)
        ss(driver, "used_in_sorted")
    except: pass

    filas_vistas = {}

    def _fila_key(f):
        tipo_v = next((f[k] for k in f if "type" in k.lower() and k != "_fila_txt"), "")
        name_v = next((f[k] for k in f if any(x in k.lower()
                       for x in ["pcmname","bookingpcm","pcm name","pcm ref"]) and k != "_fila_txt"), "")
        if not name_v:
            name_v = next((v for k, v in f.items()
                           if re.match(r'[A-Z]{2,}[-]', str(v)) and k != "_fila_txt"), "")
        if not name_v:
            name_v = f.get("_fila_txt", "")
        return f"{tipo_v}|{name_v}"

    def _get_scroll_container():
        return driver.execute_script("""
            var cdk = document.querySelector('cdk-virtual-scroll-viewport');
            if(cdk) return cdk;
            var candidates = ['#usedin','tp-usedin','[class*="usedin"]','[class*="used-in"]'];
            for(var s of candidates){
                var el = document.querySelector(s);
                if(el && el.scrollHeight > el.clientHeight) return el;
            }
            var tbl = document.querySelector('table');
            if(tbl){
                var el = tbl.parentElement;
                while(el && el !== document.body){
                    var st = getComputedStyle(el);
                    var canScroll = (st.overflow==='auto'||st.overflow==='scroll'||
                                    st.overflowY==='auto'||st.overflowY==='scroll');
                    if(canScroll && el.scrollHeight > el.clientHeight + 10) return el;
                    el = el.parentElement;
                }
            }
            return document.scrollingElement || document.body;
        """)

    def _scroll_to(pos, container):
        if container:
            driver.execute_script("arguments[0].scrollTop = arguments[1];", container, pos)
        driver.execute_script("window.scrollTo(0, arguments[0]);", pos)

    scroll_container = _get_scroll_container()
    scroll_info = driver.execute_script("""
        var el = arguments[0];
        if(!el) return {scrollH:0, clientH:0};
        return {scrollH: el.scrollHeight, clientH: el.clientHeight};
    """, scroll_container)
    print(f"    Scroll container: {scroll_info}")

    scroll_step = 150
    scroll_pos  = 0
    sin_cambio  = 0

    for _ in range(1200):
        batch = _leer_filas_visibles()
        nuevas = 0
        for f in batch:
            k = _fila_key(f)
            if k and k not in filas_vistas:
                filas_vistas[k] = f; nuevas += 1
        if nuevas == 0:
            sin_cambio += 1
            if sin_cambio >= 8: break
        else:
            sin_cambio = 0
        scroll_pos += scroll_step
        _scroll_to(scroll_pos, scroll_container)
        time.sleep(0.25)

    print(f"    Total filas únicas tras scroll: {len(filas_vistas)}")
    _scroll_to(0, scroll_container)
    time.sleep(0.5)

    filas_info = list(filas_vistas.values())
    print(f"    Filas en Used In: {len(filas_info)}")
    if not filas_info:
        dump(driver, "usedin_sin_filas")

    filas_dom = driver.find_elements(By.CSS_SELECTOR, "#usedin table tbody tr")
    if not filas_dom:
        filas_dom = driver.find_elements(By.XPATH, "//table//tbody//tr")

    pcm_list = []
    for i, info in enumerate(filas_info):
        print(f"      [{i}] {info.get('_fila_txt','')[:100]}")
        resultado = _procesar_fila_package_header(info, filas_dom[i] if i < len(filas_dom) else None)
        if resultado:
            pcm_list.append(resultado)
        else:
            print(f"      → SKIP (no pasa filtros de Package Header)")

    print(f"  📋 PCMs Package Header válidos: {[p['nombre'] for p in pcm_list]}")
    return pcm_list


# ── Abrir PCM desde Used In ───────────────────────────────────
def _ordenar_used_in_por_date(driver):
    """Clickea el encabezado 'Date' de la grilla Used In para ordenar
    ascendente por esa columna — en los PCM Package Header ese campo
    viene en blanco (confirmado por la usuaria con captura de pantalla
    real: el Package Header quedó primero tras ordenar así), así que tras
    ordenar quedan primeros en la lista y casi no hace falta scrollear
    para encontrarlos.

    Es un PRIMER PASO best-effort y verificado (compara la primera fila
    visible antes/después del click) — si no encuentra el encabezado, o
    el click no cambia nada, esta función no rompe nada ni bloquea:
    _scan_from() en abrir_pcm() sigue siendo el procedimiento de scroll
    de siempre, sin cambios, como fallback.

    IMPORTANTE (confirmado en corrida real, PSUS): esta función se llama
    dos veces para el mismo producto — una en leer_pcm_list_package_header()
    (arma la lista) y otra en abrir_pcm() (para reabrir el PCM). Si la
    grilla Used In ya estaba ordenada ascendente por Date desde el primer
    llamado (Angular puede retener el estado del componente entre
    navegaciones, no siempre lo resetea), clickear el encabezado de nuevo
    es un SEGUNDO click sobre la MISMA columna ya activa — eso la invierte
    a descendente (mismo comportamiento ya documentado para el header
    Type en otros scripts: 1er click ordena, 2do invierte), y el Package
    Header (Date en blanco) deja de ser la primera fila. Se vio en corrida
    real: la segunda vez mostró un PCM totalmente distinto con la fecha
    más lejana primero — y el fallback a scroll completo SÍ encontró el
    Package Header más abajo, confirmando que la tabla era la correcta,
    solo el orden estaba invertido. Por eso ahora se verifica el estado
    ANTES de clickear: si la primera fila YA tiene la columna Date en
    blanco, no hace falta (ni conviene) clickear de nuevo."""
    def _primera_fila():
        return driver.execute_script("""
            var tr = document.querySelector('#usedin table tbody tr') ||
                     document.querySelector('table tbody tr');
            return tr ? tr.innerText.trim().slice(0, 80) : '';
        """)

    def _primera_fila_ya_con_date_vacia():
        return driver.execute_script("""
            var tbl = document.querySelector('#usedin table') || document.querySelector('table');
            if (!tbl) return null;
            var ths = Array.from(tbl.querySelectorAll('thead th'));
            var idx = ths.findIndex(function(h){
                return (h.innerText || '').trim().toLowerCase() === 'date';
            });
            if (idx < 0) return null;
            var tr = tbl.querySelector('tbody tr');
            if (!tr) return null;
            var tds = tr.querySelectorAll('td');
            if (idx >= tds.length) return null;
            return (tds[idx].innerText || '').trim() === '';
        """)

    try:
        if _primera_fila_ya_con_date_vacia():
            print(f"    ✅ Used In ya estaba ordenado por Date (primera fila con Date en "
                  f"blanco: '{_primera_fila()}') — no clickeo de nuevo para no invertirlo")
            return True
    except Exception:
        pass  # si el chequeo falla, seguir con el click de siempre (comportamiento previo)

    antes = _primera_fila()
    try:
        clickeado = driver.execute_script("""
            var spans = document.querySelectorAll('th span');
            for (var s of spans) {
                if ((s.textContent || '').trim() === 'Date') {
                    s.click();
                    return true;
                }
            }
            return false;
        """)
    except Exception as e:
        print(f"    ⚠ Error clickeando encabezado 'Date' de Used In: {e}")
        return False

    if not clickeado:
        print("    ⚠ No encontré el encabezado 'Date' de Used In para ordenar — sigo con scroll normal")
        return False

    time.sleep(2 * VELOCIDAD)
    despues = _primera_fila()
    if despues and despues != antes:
        print(f"    ✅ Used In ordenado por Date (primera fila ahora: '{despues}')")
        return True

    print("    ⚠ Click en encabezado 'Date' no cambió la primera fila visible — sigo con scroll normal")
    return False

def abrir_pcm(driver, pcm_info, start_scroll=0):
    nombre = pcm_info["nombre"]
    handles_antes = set(driver.window_handles)

    ordenado_por_date = _ordenar_used_in_por_date(driver)

    scroll_container = driver.execute_script("""
        var cdk = document.querySelector('cdk-virtual-scroll-viewport');
        if(cdk) return cdk;
        var candidates = ['#usedin','tp-usedin','[class*="usedin"]','[class*="used-in"]'];
        for(var s of candidates){
            var el = document.querySelector(s);
            if(el && el.scrollHeight > el.clientHeight) return el;
        }
        var tbl = document.querySelector('table');
        if(tbl){
            var el = tbl.parentElement;
            while(el && el !== document.body){
                var st = getComputedStyle(el);
                var canScroll = (st.overflow==='auto'||st.overflow==='scroll'||
                                st.overflowY==='auto'||st.overflowY==='scroll');
                if(canScroll && el.scrollHeight > el.clientHeight + 10) return el;
                el = el.parentElement;
            }
        }
        return document.scrollingElement || document.body;
    """)

    def _scroll_to(pos):
        if scroll_container:
            driver.execute_script("arguments[0].scrollTop = arguments[1];", scroll_container, pos)
        driver.execute_script("window.scrollTo(0, arguments[0]);", pos)

    def _find_and_click_pcm():
        return driver.execute_script(f"""
            var nombre = {json.dumps(nombre)};
            for (var td of document.querySelectorAll('td')){{
                if(td.innerText.trim() === nombre){{ td.click(); return true; }}
            }}
            return false;
        """)

    # Mismo step que ya usa la función que arma la lista inicial de PCMs
    # (ver ~línea 1106, scroll_step=150) — un step de 1500px acá (10x más
    # grande) podía saltear filas en listas de Used In grandes (virtual
    # scroll: solo lo cercano al viewport está en el DOM en cada momento),
    # confirmado en corrida real: un PCM con "cientos" de filas en Used In
    # se encontraba bien en el escaneo inicial pero después "No encontré
    # fila para PCM" acá, con el mismo nombre exacto. max_steps sube en la
    # misma proporción para cubrir la misma distancia total que antes.
    step = 150

    def _scan_from(start_pos, max_steps=2000):
        pos = max(0, start_pos)
        _scroll_to(pos)
        time.sleep(0.4)
        for _ in range(max_steps):
            if _find_and_click_pcm():
                return pos
            pos += step
            _scroll_to(pos)
            time.sleep(0.3)
        return -1

    found_pos = -1

    # Este script solo procesa el PCM tipo "Package Header": leer_pcm_list_
    # package_header() ya filtra por ese tipo, y el llamador siempre usa el
    # primero de la lista (logueando advertencia si hubiera más de uno).
    # Confirmado por la usuaria con captura real: una vez que el orden por
    # Date se verificó (Package Header con Date en blanco, siempre primero),
    # escrollear para "buscarlo" es trabajo de más — ya está en la primera
    # fila, a scroll 0. Se clickea directo ahí, sin pasar por _scan_from().
    # Si por algún motivo no está ahí pese al orden confirmado (caso no
    # esperado), cae al escaneo con scroll de siempre, sin cambios.
    if ordenado_por_date:
        _scroll_to(0)
        time.sleep(0.4)
        if _find_and_click_pcm():
            found_pos = 0
        else:
            print(f"    ⚠ Orden por Date confirmado pero no encontré {nombre} "
                  f"en la primera fila — reintentando con scroll completo")

    if found_pos < 0:
        if start_scroll > step:
            found_pos = _scan_from(start_scroll - step)
        if found_pos < 0:
            print(f"    ↩ Reintentando desde pos=0 para {nombre}")
            found_pos = _scan_from(0)

    if found_pos < 0:
        raise Exception(f"No encontré fila para PCM {nombre}")

    time.sleep(6 * VELOCIDAD)
    nuevas = set(driver.window_handles) - handles_antes
    if not nuevas:
        try:
            ico = wait(driver, "div.tpcol3-3 i, [class*='traveldate'] i", t=3)
            jc(driver, ico); time.sleep(4 * VELOCIDAD)
            nuevas = set(driver.window_handles) - handles_antes
        except: pass

    if not nuevas:
        raise Exception(f"PCM {nombre} no abrió en nueva ventana")

    handle_pcm = list(nuevas)[0]
    driver.switch_to.window(handle_pcm)
    time.sleep(5 * VELOCIDAD)
    ss(driver, f"pcm_abierto_{nombre[:15]}")
    print(f"  ✅ PCM {nombre} abierto (scroll pos: {found_pos})")
    return handle_pcm, found_pos


# ── Cerrar ventana PCM y volver ────────────────────────────────
def cerrar_pcm(driver, handle_prod):
    for h in set(driver.window_handles) - {handle_prod}:
        try: driver.switch_to.window(h); driver.close()
        except: pass
    driver.switch_to.window(handle_prod)
    time.sleep(1)


# ── Manejar ventana de tarifas vencidas ───────────────────────
def manejar_tarifas_vencidas(driver, nombre_pcm=""):
    """
    Detecta la ventana "Extension of expired rate" (tp-button.saveall).
    Si está visible, hace SAVE ALL y retorna True.
    """
    try:
        saveall = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "tp-button.saveall > button")))
        if saveall and saveall.is_displayed():
            print(f"    ⚠ Tarifas vencidas detectadas ({nombre_pcm}) → SAVE ALL")
            ss(driver, f"tarifas_vencidas_{nombre_pcm[:12]}")
            jc(driver, saveall)
            time.sleep(4 * VELOCIDAD)
            ss(driver, f"tarifas_vencidas_ok_{nombre_pcm[:12]}")
            print(f"    ✅ SAVE ALL ejecutado (tarifas vencidas)")
            return True
    except Exception:
        pass
    return False


# ── Change Base Date ──────────────────────────────────────────
def cambiar_fecha_pcm(driver, dias_operacion, nombre_pcm, rate_from_str=""):
    """
    Cambia la fecha base del PCM.
    `rate_from_str` es la fecha base para calcular el próximo día de
    operación — normalmente RATE FROM del Excel, pero el caller puede pasar
    FECHA RECALCULATE PCM en su lugar si la fila la trae cargada.
    """
    print(f"  📅 Change Base Date PCM: {nombre_pcm}")

    fecha_actual = None
    for sel in ["div.tpcol3-3 input", "[class*='traveldate'] input", "input[class*='date']"]:
        try:
            for inp in driver.find_elements(By.CSS_SELECTOR, sel):
                v = inp.get_attribute("value") or inp.text
                f = parsear_fecha(v)
                if f and 2020 < f.year < 2035:
                    fecha_actual = f; break
        except: pass
        if fecha_actual: break

    if not fecha_actual:
        for el in driver.find_elements(By.XPATH, "//*[contains(text(),'/')]"):
            try:
                f = parsear_fecha(el.text.strip())
                if f and 2020 < f.year < 2035:
                    fecha_actual = f; break
            except: pass

    print(f"    Fecha actual: {fmt_xl(fecha_actual) if fecha_actual else 'desconocida'}")

    # RATE FROM del Excel como base (no la fecha actual del PCM)
    if rate_from_str:
        base = parsear_fecha(rate_from_str) or fecha_actual or datetime.now()
    else:
        base = fecha_actual or datetime.now()
    nueva_fecha = siguiente_dia_operacion(dias_operacion, base)
    if fecha_actual and nueva_fecha.date() == fecha_actual.date():
        nueva_fecha = siguiente_dia_operacion(
            dias_operacion, nueva_fecha + timedelta(days=1))
    nueva_fecha_str = nueva_fecha.strftime("%d/%m/%y")   # DD/MM/YY (Angular datepicker)
    print(f"    Nueva fecha: {nueva_fecha_str} ({DIAS_SEMANA[nueva_fecha.weekday()]})")

    hamburger(driver)
    ss(driver, f"pcm_hamburger_{nombre_pcm[:12]}")

    itin_xpath = ("//nav//ul/li[.//*[contains("
                  "translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),"
                  "'ITINERARY')]]//div[contains(@class,'click-area')]")
    cbd_xpath = ("//nav//li//ul//li//label[contains("
                 "translate(normalize-space(text()),"
                 "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),"
                 "'CHANGE BASE DATE')]")

    _opened = False
    for _attempt in range(3):
        try:
            itin_el = WebDriverWait(driver, 6).until(
                EC.presence_of_element_located((By.XPATH, itin_xpath)))
            jc(driver, itin_el)
            time.sleep(0.5)
            hamburger(driver)
            time.sleep(0.5)
            cbd_el = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, cbd_xpath)))
            jc(driver, cbd_el)
            print(f"    📌 Change Base Date clickeado (intento {_attempt+1})")
            _opened = True
            break
        except Exception as _e:
            print(f"    ⚠ Intento {_attempt+1} Change Base Date: {_e}")
            time.sleep(1)
            hamburger(driver)

    if not _opened:
        ss(driver, f"change_date_fail_{nombre_pcm[:12]}")
        raise Exception("No pude abrir Change Base Date en el menú ITINERARY")

    time.sleep(4 * VELOCIDAD)
    ss(driver, f"pcm_change_date_dlg_{nombre_pcm[:12]}")

    inp_fecha = None
    for sel in ["tp-dialog li:nth-of-type(2) input[type='text']",
                "tp-dialog input[type='text']",
                "tp-modal input[type='text']",
                "input[type='text'][placeholder*='date' i]",
                "input[type='text'][placeholder*='fecha' i]"]:
        try:
            el = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
            inp_fecha = el; break
        except: pass

    if inp_fecha is None:
        raise Exception("No encontré input de fecha en el dialog Change Base Date")

    set_val(driver, inp_fecha, nueva_fecha_str)
    time.sleep(0.3)
    # Click en li[1] (OLD TRAVEL DATE) para que Angular procese el cambio — grabación exacta
    try:
        jc(driver, driver.find_element(By.CSS_SELECTOR, "tp-dialog li:nth-of-type(1)"))
    except:
        pass
    time.sleep(0.3)
    ss(driver, f"pcm_fecha_llena_{nombre_pcm[:12]}")

    btn_save = wait(driver, "tp-button.save > button")
    jc(driver, btn_save)
    time.sleep(4 * VELOCIDAD)
    ss(driver, f"pcm_save_clicked_{nombre_pcm[:12]}")

    # Diálogo de recálculo — flujo exacto de la grabación (Chrome DevTools recorder):
    # 1. label del primer item del grupo de precios de servicio
    # 2. checkbox #calculatereplaceall (REPLACE ALL) — click directo via jc
    # 3. label "Update Exchange" + checkbox del exchange-rate-panel
    # 4. tp-button.yes
    try:
        WebDriverWait(driver, 6).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#calculatereplaceall")))
        print("    Diálogo recálculo → siguiendo grabación exacta")
        ss(driver, f"pcm_recalc_dlg_{nombre_pcm[:12]}")
        try:
            jc(driver, driver.find_element(By.CSS_SELECTOR,
               "tp-group.tpgroup-servicelinepricesgroup li:nth-of-type(1) > tp-label > label"))
            time.sleep(0.3)
        except:
            pass
        jc(driver, wait(driver, "#calculatereplaceall"))
        time.sleep(0.5)
        # "Update Exchange Rates" (#updateexchangerates): la grabación
        # original clickeaba DOS elementos acá — "#exchange-rate-panel >
        # tp-label > label" (el texto "Update Exchange Rates") y
        # "tp-group.tpgroup-servicelinepricesgroup tp-checkbox > label"
        # (el widget del checkbox) — confirmado por una grabación nueva de
        # Chrome DevTools (xpath del segundo click:
        # //*[@id="exchange-rate-panel"]/tp-checkbox/label) que AMBOS son
        # hijos del MISMO #exchange-rate-panel, es decir el MISMO
        # checkbox: un click lo tilda, el segundo lo destilda de nuevo —
        # exactamente el "como si no pegara" que confirmó la usuaria en
        # corrida real. Se reemplazan los dos clicks por uno solo directo
        # sobre el <input> (mismo patrón ya confiable que
        # #calculatereplaceall arriba, que tampoco pasa por su label),
        # con verificación y reintentos. Mismo fix en
        # tourplan_valorizacion_pkg_v3.py.
        for _intento_uex in range(3):
            _uex_tildado = driver.execute_script(
                "var el = document.getElementById('updateexchangerates');"
                "return el ? !!el.checked : null;")
            if _uex_tildado is None or _uex_tildado:
                break
            try:
                jc(driver, driver.find_element(By.CSS_SELECTOR, "#updateexchangerates"))
            except Exception:
                break
            time.sleep(0.4)
        if _uex_tildado is False:
            print("    ⚠ 'Update Exchange Rates' no quedó tildado tras 3 intentos")
        jc(driver, wait(driver, "tp-button.yes > button"))
        time.sleep(5 * VELOCIDAD)
        ss(driver, f"pcm_recalc_done_{nombre_pcm[:12]}")
    except Exception as _re:
        print(f"    ⚠ No apareció el diálogo de recálculo: {_re}")

    # Manejar vista "Extension of expired rate" si aparece tras el recálculo
    _tiene_vencidos = manejar_tarifas_vencidas(driver, nombre_pcm)
    if _tiene_vencidos:
        time.sleep(3 * VELOCIDAD)

    print(f"    ✅ Fecha actualizada a {nueva_fecha_str}")
    return nueva_fecha, fmt_xl(nueva_fecha), fmt_xl(fecha_actual) if fecha_actual else "", _tiene_vencidos


# ── Leer costos del PCM (VOUCHER COST en DASHBOARD) ──────────
def leer_markup_commission(driver, nombre_pcm):
    print("  💰 Leyendo costos PCM desde DASHBOARD...")
    hamburger(driver)
    menu_item(driver, "DASHBOARD")
    time.sleep(10 * VELOCIDAD)   # 10s para que el dashboard refleje el cambio de fecha
    ss(driver, f"pcm_dashboard_{nombre_pcm[:12]}")

    raw_valores = driver.execute_script("""
        function celda(td) {
            if (!td) return '';
            var inp = td.querySelector('input');
            if (inp) return inp.value.trim();
            return (td.innerText || td.textContent || '').trim();
        }
        var fixedTbl = document.getElementById('paxrangesfixedcolumns');
        var rightTbl = document.getElementById('paxrangetable');
        if (!fixedTbl || !rightTbl) return {error: 'tables not found'};

        var leftRows = Array.from(fixedTbl.querySelectorAll('tbody tr'));
        var vcRowIdx = -1;
        for (var i = 0; i < leftRows.length; i++) {
            var txt = (leftRows[i].innerText || leftRows[i].textContent || '')
                      .replace(/\\s+/g,'').toUpperCase();
            if (txt.indexOf('VOUCHERCOST') >= 0) { vcRowIdx = i; break; }
        }
        if (vcRowIdx < 0) return {error: 'no VOUCHER COST row',
            rows: leftRows.map(function(r){ return (r.innerText||r.textContent||'').trim(); })};

        var rightHdrs = Array.from(rightTbl.querySelectorAll('thead th, thead td')).map(function(th){
            return (th.innerText || th.textContent || '').trim().replace(/\\s+/g,'');
        });

        var rightRows = Array.from(rightTbl.querySelectorAll('tbody tr'));
        if (vcRowIdx >= rightRows.length) return {error: 'row index out of range', vcRowIdx: vcRowIdx};
        var vcCells = Array.from(rightRows[vcRowIdx].querySelectorAll('td, th'));

        var result = {};
        for (var j = 0; j < rightHdrs.length; j++) {
            if (/^\\d+\\+\\d+$/.test(rightHdrs[j]) && j < vcCells.length) {
                result[rightHdrs[j]] = celda(vcCells[j]);
            }
        }
        return result;
    """) or {}

    if isinstance(raw_valores, dict) and 'error' in raw_valores:
        print(f"    ℹ️  VOUCHER COST no encontrado: {raw_valores} → valores = {{}}")
        raw_valores = {}

    valores = {}
    for pax, raw in raw_valores.items():
        try:
            v = float(str(raw).replace(",","").replace("$","").replace(" ","").strip())
            if v > 0:
                valores[pax] = v
        except: pass

    print(f"    Valores ({len(valores)} rangos): {valores}")
    return valores


# ── Leer Price Code del PCM ────────────────────────────────────
def leer_price_code_pcm(driver, nombre_pcm=""):
    try:
        hamburger(driver)
        menu_item(driver, "PCM DETAILS")
        time.sleep(3 * VELOCIDAD)
        info = driver.execute_script("""
            function vis(e){ return !!(e.offsetWidth || e.offsetHeight
                                       || e.getClientRects().length); }
            var out = {valor: '', candidatos: []};
            var campos = Array.from(
                document.querySelectorAll('input, select')).filter(vis);
            for (var el of campos){
                var cont = el.closest('li,tr,tp-input,tp-select,div');
                var label = cont ? (cont.innerText || '').trim()
                                        .split('\\n')[0].slice(0, 60) : '';
                var val = (el.value || '').trim();
                if (label) out.candidatos.push(label + ' = ' + val);
                if (!out.valor && /price\\s*code/i.test(label) && val){
                    out.valor = val;
                }
            }
            return out;
        """)
        pc = (info.get("valor") or "").strip()
        pc = re.split(r"\s*[-–]\s*", pc)[0].strip().upper() if pc else ""
        if pc:
            print(f"    Price Code del PCM: {pc}")
        else:
            print(f"    ⚠ Price Code no encontrado en PCM DETAILS")
            ss(driver, f"pcm_details_pc_{nombre_pcm[:10]}")
        return pc
    except Exception as e:
        print(f"    ⚠ No pude leer el Price Code del PCM: {e}")
        return ""


def leer_service_category(driver):
    """
    Lee el valor actual del dropdown "Service Category" del producto YA
    ABIERTO. No hace falta navegar a ningún lado: al seleccionar el
    producto desde la grilla de resultados de Product Search, Tourplan ya
    deja parada la pantalla donde está este campo (confirmado por la
    usuaria) — mismo momento en que buscar_producto() termina. Usado para
    saltear servicios "Non Accommodation" ANTES de ir a Used In: la
    grilla de resultados NO trae una columna Category poblada (confirmado
    por la usuaria — el fallback anterior en _resolver_fila_fase1()
    siempre caía a "categoría desconocida"), así que este dropdown es la
    única fuente confiable. No lo modifica, solo lee lo ya seleccionado.
    """
    try:
        info = driver.execute_script("""
            function vis(e){ return !!(e.offsetWidth || e.offsetHeight
                                       || e.getClientRects().length); }
            var out = {valor: '', candidatos: []};
            var campos = Array.from(
                document.querySelectorAll('input, select')).filter(vis);
            for (var el of campos){
                var cont = el.closest('li,tr,tp-input,tp-select,div');
                var label = cont ? (cont.innerText || '').trim()
                                        .split('\\n')[0].slice(0, 60) : '';
                var val = (el.value || '').trim();
                if (label) out.candidatos.push(label + ' = ' + val);
                if (!out.valor && /service\\s*category/i.test(label) && val){
                    out.valor = val;
                }
            }
            return out;
        """)
        cat = (info.get("valor") or "").strip()
        if cat:
            print(f"    Service Category: {cat}")
        else:
            print(f"    ⚠ Service Category no encontrada en la pantalla del "
                  f"producto (candidatos: {info.get('candidatos')})")
            ss(driver, "producto_sin_service_category")
        return cat
    except Exception as e:
        print(f"    ⚠ No pude leer Service Category: {e}")
        return ""


# ── Rates helpers (idénticos a v3) ────────────────────────────
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
            function labelOf(inp){
                if (inp.id){
                    var l = scope.querySelector('label[for="' + inp.id + '"]');
                    if (l) return l.innerText.trim();
                }
                var c = inp.closest('li,tr,tp-input,tp-select,div');
                return c ? c.innerText.trim().split('\\n')[0].slice(0, 60) : '';
            }
            window.__tp_ins_inputs = inputs;
            return {
                dialog: dlg ? (dlg.tagName + ' ' + (dlg.className || '')).trim() : null,
                inputs: inputs.map(function(i, idx){ return {
                    idx: idx, id: i.id || '', name: i.name || '',
                    ph: i.placeholder || '', val: i.value || '',
                    label: labelOf(i)
                }; }),
                botones: Array.from(scope.querySelectorAll('button'))
                    .filter(vis)
                    .map(function(b){ return (b.innerText || '').trim(); })
                    .filter(function(t){ return t; })
            };
        """)
        print(f"    Dialog INSERT: {info.get('dialog')}")
        for i in info.get("inputs", []):
            print(f"      input[{i['idx']}] id='{i['id']}' ph='{i['ph']}' "
                  f"val='{i['val']}' label='{i['label']}'")
        print(f"    Botones en dialog: {info.get('botones')}")

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
            print(f"    From=input[{idx_from}] To=input[{idx_to}]←{fmt_tp(fecha_hasta)}")
        elif len(idx_con_fecha) == 1:
            _fill(idx_con_fecha[0], fecha_hasta)
        elif len(idx_vacios) >= 2:
            idx_from, idx_to = idx_vacios[0], idx_vacios[1]
            _fill(idx_from, fecha_desde)
            _fill(idx_to, fecha_hasta)
            print(f"    (sin precarga) From←{fmt_tp(fecha_desde)} To←{fmt_tp(fecha_hasta)}")
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
            tipeado = "SKIP"
            opcion  = None
            if not elegido_dlg:
                tipeado = driver.execute_script("""
                    var dlg = document.querySelector(
                        'tp-dialog, tp-modal, [role="dialog"], .modal-dialog, .dialog');
                    var scope = dlg || document;
                    function vis(e){ return !!(e.offsetWidth || e.offsetHeight
                                               || e.getClientRects().length); }
                    var inputs = Array.from(scope.querySelectorAll('input')).filter(vis);
                    var inp = inputs.find(function(i){
                        return (i.value || '').trim().toLowerCase() === 'unassigned';
                    });
                    if (!inp) return 'NO_INPUT';
                    inp.click(); inp.focus();
                    var setter = Object.getOwnPropertyDescriptor(
                        window.HTMLInputElement.prototype, 'value').set;
                    setter.call(inp, arguments[0]);
                    inp.dispatchEvent(new Event('input', {bubbles: true}));
                    inp.dispatchEvent(new Event('keyup', {bubbles: true}));
                    return 'TYPED';
                """, price_code)
                if tipeado == "TYPED":
                    time.sleep(1.5)
                    opcion = driver.execute_script("""
                        var pc = arguments[0].toUpperCase();
                        function vis(e){ return !!(e.offsetWidth || e.offsetHeight
                                                   || e.getClientRects().length); }
                        var items = Array.from(document.querySelectorAll(
                            'table tr, li, tp-option, .option, [role="option"]')).filter(vis);
                        for (var it of items){
                            var cod = it.querySelector ? it.querySelector('td.code') : null;
                            var des = it.querySelector ? it.querySelector('td.description') : null;
                            var t;
                            if (cod || des){
                                t = ((cod?cod.innerText:'')+' '+(des?des.innerText:'')).trim();
                            } else { t = (it.innerText || '').trim(); }
                            if (!t || t.length > 60) continue;
                            var up = t.toUpperCase();
                            if (up === pc || up.indexOf(pc + ' ') === 0 ||
                                up.indexOf(pc + '-') === 0 || up.indexOf(pc + ' -') === 0){
                                (des||cod||it).click(); it.click(); return t;
                            }
                        }
                        return null;
                    """, price_code)
                    time.sleep(1)
            valores_dlg = driver.execute_script("""
                var dlg = document.querySelector(
                    'tp-dialog, tp-modal, [role="dialog"], .modal-dialog, .dialog');
                var scope = dlg || document;
                function vis(e){ return !!(e.offsetWidth || e.offsetHeight
                                           || e.getClientRects().length); }
                return Array.from(scope.querySelectorAll('input')).filter(vis)
                    .map(function(i){ return i.value || ''; });
            """) or []
            asignado  = bool(elegido_dlg) or any(
                price_code.upper() in (v or "").upper() for v in valores_dlg)
            sin_campo = (not elegido_dlg) and (tipeado == "NO_INPUT")
            print(f"    Price Code '{price_code}' (dialog): "
                  f"dropdown={elegido_dlg} tipeo={tipeado} opción={opcion}")
            if not asignado and not sin_campo:
                pc_warning = (f"Price Code '{price_code}' no se pudo asignar "
                              f"al período (el campo quedó: {valores_dlg})")
                print(f"    🚩 {pc_warning}")
        else:
            pc_warning = "PCM sin Price Code leído: el período pudo quedar 'Unassigned'"
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


def _cerrar_ultimo_tp_dialog(driver, max_espera=4):
    """
    Clickea EXIT/CANCEL/CLOSE en el ÚLTIMO <tp-dialog> visible del documento
    y espera a que se cierre. Angular nunca saca del DOM un tp-dialog ya
    cerrado (solo deja de mostrarlo) — confirmado con dumps de HTML reales
    en _cancelar_dialog_y_hacer_insert más abajo — así que el diálogo activo
    SIEMPRE es el último del documento, nunca el primero.
    Devuelve True si se cerró, False si seguía abierto tras `max_espera` segundos.
    """
    try:
        driver.execute_script("""
            function vis(e){ return !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length); }
            var dlgs = document.querySelectorAll('tp-dialog');
            var dlg = dlgs.length ? dlgs[dlgs.length - 1] : document;
            var b = dlg.querySelector(
                'button.tpcancel, tp-button.cancel > button, tp-button.close > button, tp-button.exit > button');
            if (b && vis(b)) { b.click(); return; }
            var btns = Array.from(dlg.querySelectorAll('button')).filter(vis);
            for (var btn of btns) {
                var t = (btn.innerText || '').trim().toUpperCase();
                if (t === 'EXIT' || t === 'CANCEL' || t === 'CLOSE') { btn.click(); return; }
            }
        """)
        time.sleep(1)
    except Exception:
        pass

    for _ in range(max_espera):
        sigue_abierto = driver.execute_script("""
            var dlgs = document.querySelectorAll('tp-dialog');
            var dlg = dlgs.length ? dlgs[dlgs.length - 1] : null;
            return !!(dlg && dlg.offsetParent);
        """)
        if not sigue_abierto:
            return True
        time.sleep(1)
    return False


def _copiar_ultimo_period(driver, fecha_desde, fecha_hasta, cod, price_code=""):
    """
    Copia el último período via COPY DATE RANGE y ajusta la fecha TO.
    Flujo exacto de grabación Chrome DevTools:
      1. Click en td.tpcol-rateperiod del último período (abre su detalle)
      2. Click en tp-button.copydaterange > button
      3. En tp-dialog:last-of-type li:nth-of-type(3) input[type='text']:
         escribir fecha TO en formato DD/MM/YY
      4. Click en tp-button.ok > button
      → Tourplan abre el NUEVO período ya con las tarifas copiadas.
         Ya estamos en el detalle: no hace falta volver a la lista.

    El dialog "Copy Rate" que abre acá es, confirmado por dump real de
    HTML, el MISMO componente tp-insert-rate / #insertrateview que usa
    también el fallback a INSERT — Tourplan reutiliza el mismo modal con
    un título distinto ("Copy Rate" vs "Insert Rate") según el contexto.
    Angular nunca saca del DOM un <tp-dialog> ya cerrado (se van
    acumulando, uno o más por período procesado) — por eso el dialog
    correcto SIEMPRE es el ÚLTIMO tp-dialog del documento
    (tp-dialog:last-of-type), nunca el segundo — ver detalle en
    _fill_dialog_date() y _cancelar_dialog_y_hacer_insert() más abajo.

    Retorna (pc_warning, ya_en_detalle).
    ya_en_detalle=True → el caller OMITE el click en la lista de períodos.
    ya_en_detalle=False → el caller debe buscar/abrir el período normalmente
                          (ocurre en el fallback a INSERT).

    El período fuente a copiar se elige por PRICE CODE, no simplemente
    "el más reciente de la lista": la lista puede tener períodos de varios
    price codes mezclados (el filtro `_seleccionar_price_code` del caller
    no garantiza por sí solo qué fila queda en el índice 0), y copiar a
    ciegas el primero podía extender un período con el price code
    equivocado sin ningún aviso (confirmado por la usuaria en corrida
    real de valorizacion_desde_excel.py). Se lee el price code real de
    cada fila (td.tpcol-pricecodecode, mismo elemento que usa
    _leer_periodos) y se elige la primera (= más reciente) que coincida
    con `price_code` — o con Unassigned/ALL si `price_code` viene vacío.
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

    # Paso 1: abrir el detalle del último período
    jc(driver, ultimo_td)
    time.sleep(4 * VELOCIDAD)
    ss(driver, f"rates_copy_detalle_{cod[:10]}")

    # Paso 2: click en COPY DATE RANGE
    try:
        btn_cdr = wait(driver, "tp-button.copydaterange > button", t=8)
    except Exception:
        print(f"    ⚠ No encontré tp-button.copydaterange → fallback a INSERT")
        ss(driver, f"rates_copy_nobtn_{cod[:10]}")
        # Volver a la lista antes de hacer INSERT
        hamburger(driver)
        menu_item(driver, "RATES")
        time.sleep(3 * VELOCIDAD)
        return _insertar_rate_period(driver, fecha_desde, fecha_hasta, cod, price_code), False

    print(f"    📋 COPY DATE RANGE encontrado → abriendo dialog")
    jc(driver, btn_cdr)
    time.sleep(3 * VELOCIDAD)
    ss(driver, f"rates_copy_dlg_{cod[:10]}")

    # Paso 3: llenar fechas FROM y TO en el dialog.
    # Estructura grabada: li:nth-of-type(2) = FROM, li:nth-of-type(3) = TO.
    # El dialog es tp-dialog:last-of-type (el último del documento — ver
    # nota en el docstring de la función sobre por qué NO es "nth-of-type(2)").
    # Ambas fechas se toman del Excel (no de la fecha auto-calculada por Tourplan).
    fecha_desde_str = fecha_desde.strftime("%d/%m/%y")
    fecha_hasta_str = fecha_hasta.strftime("%d/%m/%y")

    def _fill_dialog_date(sel, fecha_str, label):
        """Escribe la fecha y VERIFICA que haya quedado (confirmado en corrida
        real: el setter sintético + evento 'change' a veces no "pega" en este
        campo del dialog — TO quedaba vacío en pantalla sin que set_val()
        tirara ninguna excepción, y el caller seguía de largo con el dialog
        incompleto). Si no coincide, reintenta con foco + Ctrl+A + Delete +
        send_keys (eventos de teclado nativos, que Angular a veces necesita
        en vez del setter sintético). Si tras el reintento sigue sin
        coincidir, lanza excepción para que el caller haga el fallback a
        INSERT en vez de clickear OK con el dialog a medio llenar.

        La comparación es por FECHA, no por texto exacto: confirmado en
        corrida real (AEEZMV) que el campo TO a veces re-formatea el valor
        que se le escribió ('31/03/27') a un formato distinto pero
        equivalente ('31/Mar/2027') apenas lo acepta — comparar como texto
        literal tomaba eso como una falla y cancelaba un COPY DATE RANGE
        que en realidad había funcionado bien."""
        def _valor_coincide(val):
            if val == fecha_str:
                return True
            f_val = parsear_fecha(val)
            f_obj = parsear_fecha(fecha_str)
            return f_val is not None and f_obj is not None and f_val == f_obj

        inp = wait(driver, sel, t=6)   # lanza excepción si no encuentra → caller hace fallback
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
        # captura de pantalla en corrida real: SPLIT DATE RANGE, COPY DATE
        # RANGE, DELETE DATE RANGE, EXIT, SAVE, INSERT RATE SET, DELETE
        # RATE SET, SET CHILD RATES. El botón para cerrarlo es "EXIT", NO
        # "cancel"/"close" — por eso el click de abajo no encontraba nada y
        # el dialog quedaba abierto en silencio. HTML confirmado por la
        # usuaria (inspeccionado en Tourplan real):
        #   <button class="tpbutton tpcancel tpsecondarysystembutton">Exit</button>
        # — es un <button> plano con clase "tpcancel" (sin el wrapper
        # <tp-button> que usan otros diálogos como SAVE/OK).
        #
        # IMPORTANTE (confirmado con dumps de HTML reales de una corrida):
        # Angular NUNCA saca del DOM un <tp-dialog> ya "cerrado" — solo dejan
        # de mostrarse, pero el elemento sigue ahí. Cada período procesado
        # (COPY DATE RANGE que cae a este fallback) deja un <tp-dialog> más
        # acumulado (se vieron 3, 6 y hasta 8 en una misma corrida). Buscar
        # el botón EXIT con un querySelector SIN acotar a ningún diálogo
        # agarraba el del PRIMER <tp-dialog> del documento — uno viejo y
        # abandonado, no el que hay que cerrar — así que el click no hacía
        # nada sobre el diálogo real. El diálogo recién abierto SIEMPRE es
        # el ÚLTIMO en el documento (nunca el segundo, ni el primero). Ver
        # _cerrar_ultimo_tp_dialog() más arriba, que implementa este criterio
        # y se reusa acá y en la verificación post-SAVE de
        # actualizar_rates_servicio_madre.
        #
        # Si el dialog sigue abierto tras cerrar, NO seguir a INSERT —
        # confirmado en corrida real (captura de pantalla): con el dialog
        # COPY DATE RANGE todavía abierto, el "INSERT" que se clickeaba a
        # continuación terminaba siendo el botón "INSERT RATE SET" DEL
        # PROPIO DIALOG viejo (abre otro dialog "Insert Rate" apilado
        # encima, con sus propios TRAVEL DATE FROM/TO) — y
        # _insertar_rate_period() leía los inputs numéricos de la grilla de
        # rates de atrás, no fechas. Mejor abortar la fila con un error
        # claro que arrastrar ese estado a un INSERT que nunca iba a
        # funcionar.
        if not _cerrar_ultimo_tp_dialog(driver):
            print(f"    🔍 Diagnóstico tp-dialog en el documento: {_diag_tp_dialogs()}")
            ss(driver, f"rates_copy_sigue_abierto_{cod[:10]}")
            dump(driver, f"rates_copy_sigue_abierto_{cod[:10]}")
            raise Exception(
                "El dialog COPY DATE RANGE no se cerró tras cancelar — se aborta "
                "en vez de intentar INSERT con el dialog viejo todavía abierto "
                "(podía terminar clickeando 'INSERT RATE SET' del propio dialog "
                "en vez de un INSERT de período nuevo)."
            )

        hamburger(driver); menu_item(driver, "RATES"); time.sleep(3 * VELOCIDAD)
        return _insertar_rate_period(driver, fecha_desde, fecha_hasta, cod, price_code), False

    def _diag_tp_dialogs():
        """Diagnóstico: cuántos <tp-dialog> hay en el documento en este
        momento, cuáles están visibles y qué texto tienen. Confirmado con
        dumps de HTML reales (corrida PSUS): Angular nunca saca del DOM un
        <tp-dialog> ya cerrado — se acumulan uno o más por cada período
        procesado (se vieron 3, 6 y hasta 8 en una misma corrida). El
        selector posicional 'tp-dialog:nth-of-type(2)' que usaba este
        código asumía que siempre había exactamente 2 — por eso agarraba
        un diálogo viejo y abandonado en vez del recién abierto, que
        siempre es el ÚLTIMO. Ya corregido a 'tp-dialog:last-of-type' en
        _fill_dialog_date() y _cancelar_dialog_y_hacer_insert(); esta
        función queda como red de diagnóstico para la próxima falla real."""
        try:
            return driver.execute_script("""
                var dlgs = Array.from(document.querySelectorAll('tp-dialog'));
                return dlgs.map(function(d, i){
                    var vis = !!(d.offsetWidth || d.offsetHeight || d.getClientRects().length);
                    return {indice: i, visible: vis,
                            texto: (d.innerText || '').replace(/\\s+/g,' ').trim().slice(0, 80)};
                });
            """)
        except Exception as _e_diag:
            return f"(no se pudo diagnosticar: {_e_diag})"

    try:
        _fill_dialog_date(
            "tp-dialog:last-of-type li:nth-of-type(2) input[type='text']",
            fecha_desde_str, "FROM")
    except Exception as _e:
        print(f"    ⚠ No encontré campo FROM en dialog COPY: {_e} → cancela y hace INSERT")
        print(f"    🔍 Diagnóstico tp-dialog en el documento: {_diag_tp_dialogs()}")
        ss(driver, f"rates_copy_nofrom_{cod[:10]}")
        dump(driver, f"rates_copy_nofrom_{cod[:10]}")
        return _cancelar_dialog_y_hacer_insert()

    try:
        _fill_dialog_date(
            "tp-dialog:last-of-type li:nth-of-type(3) input[type='text']",
            fecha_hasta_str, "TO")
    except Exception as _e:
        print(f"    ⚠ No encontré campo TO en dialog COPY: {_e} → cancela y hace INSERT")
        print(f"    🔍 Diagnóstico tp-dialog en el documento: {_diag_tp_dialogs()}")
        ss(driver, f"rates_copy_noto_{cod[:10]}")
        dump(driver, f"rates_copy_noto_{cod[:10]}")
        return _cancelar_dialog_y_hacer_insert()
    ss(driver, f"rates_copy_lleno_{cod[:10]}")

    # Paso 4: OK
    jc(driver, wait(driver, "tp-button.ok > button"))
    time.sleep(5 * VELOCIDAD)
    ss(driver, f"rates_copy_ok_{cod[:10]}")
    print(f"    ✅ COPY DATE RANGE OK → nuevo período {fecha_desde_str} – {fecha_hasta_str}")
    # Tourplan abre directamente el detalle del nuevo período.
    return "", True   # ya_en_detalle=True


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


def _extraer_valores_ad(tabla):
    headers  = tabla["headers"]
    idx_svc  = next((i for i,h in enumerate(headers) if "SERVICE" in h.upper()), 0)
    idx_cost = next((i for i,h in enumerate(headers)
                     if "GROUP COST" in h.upper() and "FIT" not in h.upper()), 1)
    pat_rango = re.compile(r'\d+\s*[-–]\s*\d+')
    out = {}
    for row in tabla["rows"]:
        celdas = row["celdas"]
        if not celdas: continue
        svc = celdas[idx_svc] if idx_svc < len(celdas) else ""
        if not svc.strip(): continue
        if not (pat_rango.search(svc) and re.search(r'\bAD\b', svc.upper())):
            continue
        raw = ""
        if row["inputs"] and row["inputs"][0]:
            raw = row["inputs"][0]
        elif idx_cost < len(celdas):
            raw = celdas[idx_cost]
        try:
            out[svc] = float(raw.replace(",", "").replace("$", "").strip())
        except:
            out[svc] = 0.0
    return out


def _escribir_rates(driver, valores_pcm, cod):
    for sel in ["[class*='tab-item']", "div.tab", "a.tab", "button.tab"]:
        for el in driver.find_elements(By.CSS_SELECTOR, sel):
            if el.text.strip().upper() == "RATES" and el.is_displayed():
                jc(driver, el); time.sleep(3 * VELOCIDAD); break

    tabla = _leer_tabla_rates(driver)
    if not tabla:
        ss(driver, f"rates_sin_tabla_{cod[:10]}")
        dump(driver, f"rates_sin_tabla_{cod[:10]}")
        raise Exception(f"No encontré tabla de rates en {cod}")

    headers  = tabla["headers"]
    idx_svc  = next((i for i,h in enumerate(headers) if "SERVICE" in h.upper()), 0)
    idx_cost = next((i for i,h in enumerate(headers)
                     if "GROUP COST" in h.upper() and "FIT" not in h.upper()), 1)

    rangos_madre = []
    v_viejos     = {}
    pendientes   = []

    def _leer_valor(row, celdas):
        val_raw = ""
        if row["inputs"] and row["inputs"][0]:
            val_raw = row["inputs"][0]
        elif idx_cost < len(celdas):
            val_raw = celdas[idx_cost]
        try:
            return float(val_raw.replace(",", "").replace("$", "").strip())
        except:
            return 0.0

    PAT_RANGO = re.compile(r'\d+\s*[-–]\s*\d+')
    for row in tabla["rows"]:
        celdas    = row["celdas"]
        if not celdas: continue
        svc_item  = celdas[idx_svc] if idx_svc < len(celdas) else ""
        if not svc_item.strip(): continue
        valor_actual = _leer_valor(row, celdas)
        es_rango = bool(PAT_RANGO.search(svc_item))
        es_ad    = bool(re.search(r'\bAD\b', svc_item.upper()))
        if es_rango and es_ad:
            v_viejos[svc_item] = valor_actual
            rangos_madre.append(svc_item)
            continue
        if valor_actual:
            msg = f"{svc_item} tiene valor {valor_actual} (no se tocó)"
            pendientes.append(msg)
            print(f"    🚩 REVISAR: {msg}")
        else:
            print(f"    ⏭ Fila no se toca: {svc_item}")

    def pax_n(key):
        m = re.match(r'(\d+)\+', key)
        return int(m.group(1)) if m else 0

    def rango_limites(texto):
        """Devuelve (pmin, pmax) del texto de rango. Ej: '5 - 9 AD' → (5, 9)."""
        m = re.search(r'(\d+)\s*[-–]\s*(\d+)', texto)
        if m:
            return int(m.group(1)), int(m.group(2))
        m = re.search(r'(\d+)', texto)
        n = int(m.group(1)) if m else 0
        return n, n

    por_pax  = {pax_n(k): v for k, v in valores_pcm.items()}
    v_nuevos = {}
    for rango in rangos_madre:
        pmin, pmax = rango_limites(rango)
        # Regla especial para rangos abiertos (pmax = 9999)
        if pmax >= 9999:
            val = 0.0 if pmin >= 42 else 999.0
            v_nuevos[rango] = val
            print(f"    {rango}: {v_viejos.get(rango,'?')} → {val:.2f}  [regla 9999]")
            continue
        # Rango normal: tomar el MÁXIMO costo del PCM para pax en [pmin, pmax]
        vals_en_rango = [v for p, v in por_pax.items() if pmin <= p <= pmax]
        if not vals_en_rango:
            msg = (f"Inconsistencia en bases: el rango '{rango}' del servicio madre "
                   f"no tiene valores correspondientes en el PCM "
                   f"(pax disponibles en PCM: {sorted(por_pax.keys())})")
            pendientes.append(msg)
            print(f"    🚩 {msg}")
            v_nuevos[rango] = 0.0
            continue
        val = max(vals_en_rango)
        v_nuevos[rango] = val
        print(f"    {rango}: {v_viejos.get(rango,'?')} → {val:.2f}  [max de pax {pmin}-{pmax}: {vals_en_rango}]")

    # Scopear al diálogo
    raiz = driver
    for sel in ["tp-dialog", "tp-modal", "[role='dialog']"]:
        cand = [e for e in driver.find_elements(By.CSS_SELECTOR, sel) if e.is_displayed()]
        if cand:
            raiz = cand[-1]; break

    def _encontrar_fila_por_rango(rango):
        # Re-busca en vivo la fila cuyo label coincide EXACTO con `rango`.
        # IMPORTANTE: no cachear esta lista de antemano — la grilla usa CDK
        # virtual scroll, que reutiliza un pool fijo de <tr>/<input> del DOM
        # y los re-vincula a otra fila lógica a medida que se scrollea. Una
        # referencia capturada antes de escribir en filas anteriores (lo que
        # dispara scrollIntoView y por lo tanto re-render) puede terminar
        # apuntando a OTRA fila para cuando se usa — eso hacía que el valor
        # de una fila se terminara guardando en la fila vecina.
        for f in raiz.find_elements(By.XPATH, ".//table//tbody//tr"):
            if not (f.is_displayed() and f.text.strip()):
                continue
            if f.text.split("\n")[0] == rango:
                return f
        return None

    inputs_escritos = []
    for rango in rangos_madre:
        fila = _encontrar_fila_por_rango(rango)
        if fila is None:
            print(f"    ⚠ No encontré la fila {rango} en la grilla (¿scroll?)")
            continue
        edits = [inp for inp in fila.find_elements(By.TAG_NAME, "input")
                 if inp.is_displayed() and not inp.get_attribute("readonly")]
        if not edits:
            print(f"    ⚠ Sin inputs editables en fila {rango}")
            continue
        valor_txt = f"{v_nuevos.get(rango,0):.2f}"
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
            if ci == 0:
                inputs_escritos.append(inp)
            time.sleep(0.1)

    try:
        diag = driver.execute_script("""
            function vis(e){ return !!(e && (e.offsetWidth || e.offsetHeight
                                       || e.getClientRects().length)); }
            var inputs = Array.from(document.querySelectorAll('input'))
                .filter(function(i){ return vis(i) && !i.readOnly
                        && (i.className||'').indexOf('tpnumber') >= 0; });
            var vals = inputs.slice(0, 12).map(function(i){ return i.value; });
            var saves = Array.from(document.querySelectorAll('tp-button.save > button'))
                .map(function(b, k){
                    return {idx:k, disabled:!!b.disabled, visible:vis(b),
                            inDialog: !!b.closest('tp-dialog, tp-modal, [role="dialog"]')};
                });
            return {nInputsCosto: inputs.length, valsCosto: vals, saves: saves};
        """)
        print(f"    [diag] inputs de costo: {diag['nInputsCosto']} · saves: {diag['saves']}")
        leidos_ad = []
        for inp in inputs_escritos[:8]:
            try: leidos_ad.append(inp.get_attribute("value"))
            except Exception: leidos_ad.append("?")
        print(f"    [diag] valores AD escritos: {leidos_ad}")
    except Exception as _de:
        print(f"    [diag] no se pudo inspeccionar: {_de}")

    print("    📸 PANTALLA: DESPUÉS DE ESCRIBIR — ANTES DEL SAVE")
    ss(driver, f"A_DESPUES_ESCRIBIR_ANTES_SAVE_{cod[:8]}")

    # Esperar a que Angular habilite el SAVE tras procesar los eventos de input.
    # Si todos los botones están disabled, reintentar hasta 3 veces con 1.5s entre intentos.
    for _save_intento in range(3):
        _saves_chk = driver.find_elements(By.CSS_SELECTOR, "tp-button.save > button")
        if any(b.is_displayed() and b.is_enabled() for b in _saves_chk):
            break
        print(f"    ⏳ SAVE aún disabled, esperando... (intento {_save_intento+1}/3)")
        time.sleep(1.5 * VELOCIDAD)

    guardado = None
    try:
        botones = driver.find_elements(By.CSS_SELECTOR, "tp-button.save > button")
        def _rank(b):
            try:
                en    = b.is_enabled() and b.is_displayed()
                in_dlg = driver.execute_script(
                    "return !!arguments[0].closest('tp-dialog, tp-modal, [role=\"dialog\"]');", b)
            except Exception:
                en, in_dlg = False, False
            return (2 if (en and in_dlg) else 1 if en else 0)
        botones = sorted(botones, key=_rank, reverse=True)
        for b in botones:
            try:
                if b.is_displayed() and b.is_enabled():
                    in_dlg = driver.execute_script(
                        "return !!arguments[0].closest('tp-dialog, tp-modal, [role=\"dialog\"]');", b)
                    jc(driver, b)
                    guardado = f"tp-button.save (click real{', dialog' if in_dlg else ''})"
                    break
            except Exception: continue
    except Exception: pass

    if not guardado:
        guardado = driver.execute_script("""
            function vis(e){ return !!(e.offsetWidth || e.offsetHeight
                                       || e.getClientRects().length); }
            var saves = Array.from(document.querySelectorAll('tp-button.save > button'))
                .filter(function(b){ return vis(b) && !b.disabled; });
            saves.sort(function(a,b){
                var da = a.closest('tp-dialog,tp-modal,[role="dialog"]') ? 1 : 0;
                var db = b.closest('tp-dialog,tp-modal,[role="dialog"]') ? 1 : 0;
                return db - da;
            });
            if (saves.length){ saves[0].click(); return 'tp-button.save enabled (JS)'; }
            var bs = Array.from(document.querySelectorAll('button')).filter(vis);
            for (var b of bs){
                var t = (b.innerText || '').trim().toUpperCase();
                if (t === 'SAVE' && !b.disabled){ b.click(); return t + ' (JS)'; }
            }
            return null;
        """)

    if not guardado:
        # El Save puede estar DESHABILITADO por dos motivos MUY distintos:
        # (a) el período ya tenía exactamente estos valores (re-corrida) → nada
        #     que guardar, Angular deja el form 'pristine'.
        # (b) el período es NUEVO (se acaba de crear con COPY DATE RANGE) y el
        #     valor que acabamos de escribir YA se confirmó solo (sin necesitar
        #     un SAVE aparte) — visto en producción: el form vuelve a estar
        #     'pristine' apenas se escribe, mucho antes de que busquemos el
        #     botón. Comparar contra el viejo (anterior a escribir) da un
        #     falso error en este caso — hay que comparar contra lo que HAY
        #     AHORA en la grilla.
        def _f(x):
            try: return float(str(x).replace(",", "") or 0)
            except Exception: return 0.0

        def _ya_aplicado(tolerancia):
            tabla_actual = _leer_tabla_rates(driver)
            if not tabla_actual:
                return False
            actuales = _extraer_valores_ad(tabla_actual)
            return all(
                r in actuales and abs(_f(actuales[r]) - _f(v_nuevos.get(r, 0))) <= tolerancia
                for r in rangos_madre)

        sin_cambios = all(
            abs(_f(v_viejos.get(r, 0)) - _f(v_nuevos.get(r, 0))) <= 0.011
            for r in rangos_madre)
        if sin_cambios:
            print(f"    ✅ Sin cambios: el período ya tenía estos valores")
            return v_viejos, v_nuevos, rangos_madre, pendientes
        if _ya_aplicado(0.30):
            print(f"    ✅ Sin botón SAVE pero los valores ya están aplicados en "
                  f"la grilla (dentro de tolerancia) — período nuevo confirmado "
                  f"sin SAVE aparte ({len(rangos_madre)} rangos AD)")
            return v_viejos, v_nuevos, rangos_madre, pendientes
        ss(driver, f"rates_save_error_{cod[:10]}")
        dump(driver, f"rates_save_error_{cod[:10]}")
        raise Exception(f"No encontré el botón SAVE de la grilla de rates en {cod}")

    print(f"    SAVE de rates clickeado ('{guardado}')")
    time.sleep(4 * VELOCIDAD)
    print("    📸 PANTALLA: DESPUÉS DEL SAVE")
    ss(driver, f"B_DESPUES_DEL_SAVE_{cod[:8]}")
    print(f"    ✅ SAVE OK — {len(rangos_madre)} rangos AD actualizados"
          + (f"  (🚩 {len(pendientes)} pendientes)" if pendientes else ""))
    return v_viejos, v_nuevos, rangos_madre, pendientes


def _buscar_servicio_madre_fase2(driver, service_code, supplier, location,
                                  service_type=None):
    """
    Busca el servicio madre en Product Setup UNA VEZ y verifica que el
    supplier coincida. Se reutiliza para todos los PCM_Detail pendientes
    de Fase 2 que compartan el mismo servicio madre (mismo service_code/
    supplier/location/service_type), evitando repetir la búsqueda completa
    del producto por cada período a aplicar. La navegación a RATES (barata)
    NO se hace acá — se repite por período en actualizar_rates_servicio_madre,
    porque la verificación post-SAVE del período anterior puede dejar al
    driver parado en el detalle de otro período, no en la lista.
    """
    print(f"\n  🔍 Buscando servicio madre: {service_code}")
    buscar_producto(driver, location, supplier, service_code,
                    service_type=service_type)

    # Verificar supplier en el producto cargado
    supplier_actual = ""
    try:
        inp_sup = wait(driver, "input.tplabel-supplier, #searchWrapper input", t=5)
        supplier_actual = (inp_sup.get_attribute("value") or "").strip().upper()
    except: pass
    if not supplier_actual:
        try:
            el = driver.find_element(By.CSS_SELECTOR, ".tplabel-supplier, [class*='supplier'] input")
            supplier_actual = (el.get_attribute("value") or el.text or "").strip().upper()
        except: pass

    # El campo del producto muestra el supplier como "1EURO1 - EUROTUR"
    # (código + nombre): comparar solo el código.
    cod_supplier_actual = re.split(r"\s*[-–]\s*", supplier_actual)[0].strip()
    if cod_supplier_actual and supplier and cod_supplier_actual != supplier.upper():
        raise Exception(
            f"Supplier del servicio madre ({supplier_actual}) ≠ supplier esperado ({supplier}). "
            f"No se actualiza.")


def actualizar_rates_servicio_madre(driver, service_code, supplier,
                                     rate_from_str, rate_to_str,
                                     location, valores_pcm,
                                     service_type=None, price_code=""):
    """
    Fase 2: aplica UN período de rates sobre un servicio madre YA BUSCADO
    (ver _buscar_servicio_madre_fase2, llamada por el caller antes de este
    loop): navega a RATES, busca/crea el período y escribe los valores.
    """
    print(f"\n  🔄 Actualizando servicio madre: {service_code}  periodo {rate_from_str}–{rate_to_str}")

    hamburger(driver)
    ss(driver, f"madre_hamburger_{service_code}")
    menu_item(driver, "RATES")
    time.sleep(4 * VELOCIDAD)
    ss(driver, f"madre_rates_lista_{service_code}")

    PC     = (price_code or PRICE_CODE_DEFAULT or "").strip().upper()
    IS_ALL = PC in ("", "ALL", "TODOS", "*", "UNASSIGNED")

    def _sel_pc(etq):
        if not IS_ALL:
            _seleccionar_price_code(driver, PC, etiqueta=etq)

    _sel_pc(f"lista {service_code}")
    time.sleep(1.5)
    ss(driver, f"madre_rates_lista_pc_{service_code}")

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

    rf = parsear_fecha(rate_from_str)
    rt = parsear_fecha(rate_to_str)
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
    print("    Períodos: {} (ej: {})".format(
        len(periodos),
        ", ".join("{}[{}]".format(r["date"], r["pc"] or "—") for r in periodos[:4])))

    idx_periodo  = _idx_periodo(periodos)
    pc_warning   = ""
    ya_en_detalle = False
    if idx_periodo is None:
        etq_pc = "ALL/Unassigned" if IS_ALL else PC
        print(f"    No hay período {etq_pc} para {fmt_tp(rf)}–{fmt_tp(rt)} → COPY último período")
        pc_warning, ya_en_detalle = _copiar_ultimo_period(
            driver, rf, rt, service_code, "" if IS_ALL else PC)
        ss(driver, f"madre_rates_creado_{service_code}")

        if not ya_en_detalle:
            # Fallback INSERT: volver a buscar el período en la lista
            _sel_pc(f"post-insert {service_code}")
            time.sleep(1.5)
            periodos    = _leer_periodos()
            idx_periodo = _idx_periodo(periodos)
            print(f"    Períodos tras INSERT: {len(periodos)}  (idx={idx_periodo})")
            if idx_periodo is None:
                raise Exception(
                    f"El período {etq_pc} {fmt_tp(rf)}–{fmt_tp(rt)} no aparece "
                    f"en la lista de rates de {service_code} tras el INSERT.")

    # Default para cuando ya_en_detalle=True (COPY DATE RANGE nos dejó
    # directo en el detalle, sin pasar por la lista) — sin esto,
    # periodo_encontrado queda sin asignar y la verificación post-SAVE de
    # más abajo revienta con UnboundLocalError si el período no aparece al
    # recargar RATES (confirmado en corrida real, 2026-08).
    periodo_encontrado = f"{fmt_tp(rf)}–{fmt_tp(rt)}"
    if not ya_en_detalle:
        periodo_encontrado = periodos[idx_periodo]["date"]
        filas_td = driver.find_elements(By.CSS_SELECTOR, "td.tpcol-rateperiod")
        print(f"    Abriendo período {PC if not IS_ALL else 'ALL'}: "
              f"{periodo_encontrado} (idx {idx_periodo})")
        jc(driver, filas_td[idx_periodo])
        time.sleep(5 * VELOCIDAD)
        ss(driver, f"madre_rates_detalle_{service_code}")
    else:
        print(f"    Ya en detalle del período copiado → omitiendo apertura desde lista")

    _sel_pc(f"escribir {service_code}")
    ss(driver, f"madre_rates_pricecode_{service_code}")

    v_viejos, v_nuevos, rangos, pendientes = _escribir_rates(
        driver, valores_pcm, service_code)
    if pc_warning:
        pendientes.append(pc_warning)

    # ── Verificación post-SAVE ─────────────────────────────────
    # No se reabre el menú lateral (hamburger) acá: reabrirlo con la
    # pantalla de detalle del período todavía activa dejaba colgado el
    # backdrop de navegación (.tpnavbackdrop) y ese backdrop interceptaba
    # el click/lectura siguiente — el valor quedaba bien guardado en
    # Tourplan pero la verificación leía mal y reportaba "no se actualizó".
    # En su lugar: cerrar el detalle con EXIT (mismo botón que usa el
    # flujo COPY DATE RANGE, ver _cerrar_ultimo_tp_dialog) para volver a la
    # lista de períodos subyacente, y reabrir el período desde ahí — igual
    # que la apertura inicial más arriba.
    error_post = ""
    try:
        cerrado = _cerrar_ultimo_tp_dialog(driver)
        if not cerrado:
            ss(driver, f"verif_dialog_no_cerro_{service_code[:10]}")
            error_post = (f"Verificación post-SAVE: el detalle del período "
                          f"{periodo_encontrado} no se cerró (EXIT) tras el SAVE")
        else:
            time.sleep(2 * VELOCIDAD)
            periodos_v  = _leer_periodos()
            idx_v       = _idx_periodo(periodos_v)
            if idx_v is None:
                ss(driver, f"verif_sin_periodo_{service_code[:10]}")
                error_post = (f"Verificación post-SAVE: el período {periodo_encontrado} "
                              f"no aparece en la lista tras reabrir")
            else:
                filas_td = driver.find_elements(By.CSS_SELECTOR, "td.tpcol-rateperiod")
                jc(driver, filas_td[idx_v])
                time.sleep(5 * VELOCIDAD)
                _sel_pc(f"verif {service_code}")

                # Reintentos antes de reportar error: la grilla usa virtual scroll
                # (CDK) y Angular puede tardar en re-bindear los <input> de las
                # filas recicladas justo después del reload. Una lectura inmediata
                # puede traer el valor de OTRA fila (patrón típico: "leído" de una
                # fila coincide con "esperado" de la fila vecina) — no es que no
                # se haya guardado, es que se leyó antes de que termine de asentar.
                MAX_INTENTOS_VERIF = 3
                tabla_post = None
                leidos = {}
                difs = []
                for intento in range(1, MAX_INTENTOS_VERIF + 1):
                    tabla_post = _leer_tabla_rates(driver)
                    if not tabla_post:
                        break
                    leidos = _extraer_valores_ad(tabla_post)
                    difs = []
                    for rango in rangos:
                        esperado = float(v_nuevos.get(rango, 0) or 0)
                        actual   = leidos.get(rango)
                        if actual is None or abs(actual - esperado) > 0.011:
                            difs.append(f"{rango}: esperado {esperado:.2f}, leído {actual}")
                    if not difs:
                        break
                    if intento < MAX_INTENTOS_VERIF:
                        print(f"    ⚠ Verificación post-SAVE: {len(difs)}/{len(rangos)} rangos no "
                              f"coinciden (intento {intento}/{MAX_INTENTOS_VERIF}) — puede ser que "
                              f"Angular todavía no re-renderizó la grilla, reintentando...")
                        time.sleep(2 * VELOCIDAD)

                if not tabla_post:
                    ss(driver, f"verif_sin_tabla_{service_code[:10]}")
                    dump(driver, f"verif_sin_tabla_{service_code[:10]}")
                    error_post = "Verificación post-SAVE: no pude releer la grilla del período"
                elif difs:
                    ss(driver, f"verif_error_{service_code[:10]}")
                    dump(driver, f"verif_error_{service_code[:10]}")
                    error_post = (f"Los valores NO quedaron guardados en {service_code}: "
                                  f"{len(difs)}/{len(rangos)} rangos difieren tras recargar y "
                                  f"{MAX_INTENTOS_VERIF} intentos (ej: {'; '.join(difs[:3])})")
                else:
                    print(f"    ✔ Verificación post-SAVE OK: {len(rangos)} rangos AD persistidos")
                    ss(driver, f"verif_ok_{service_code[:10]}")
    except Exception as _ve:
        error_post = f"Verificación post-SAVE falló: {_ve}"

    return v_viejos, v_nuevos, rangos, pendientes, error_post


# ── Sheets ─────────────────────────────────────────────────────
SHEET_DETAIL = "PCM_Detail"

_DETAIL_HEADERS = [
    "TIMESTAMP", "SERVICE CODE", "SERVICE TYPE", "LOCATION", "SUPPLIER",
    "RATE FROM", "RATE TO", "PCM", "MARKUP JSON", "PRICE CODE",
    "ESTADO F1", "ERROR F1", "ESTADO F2", "ERROR F2",
]
_DETAIL_COL = {h: i+1 for i, h in enumerate(_DETAIL_HEADERS)}

_MARKUP_PAX          = [f"{n}+0" for n in range(1, 42)]
_DETAIL_MARKUP_START = len(_DETAIL_HEADERS) + 1
_DETAIL_RATES_START  = _DETAIL_MARKUP_START + len(_MARKUP_PAX)

# Comentarios de ayuda en los headers del Excel original — decisión
# explícita: no se portan a Sheets por ahora (la ayuda de cada columna
# sigue en el README/skills del repo).

# Worksheets conectados (ver conectar_sheets_madre()) y sus encabezados —
# se leen una vez y se reusan entre llamadas, para no re-autenticar
# contra Google en cada fila.
_ws = None
_ws_detail = None
_columnas = None
_campo_a_columna = None
_columnas_detail = None

def conectar_sheets_madre():
    global _ws, _ws_detail
    _ws = conectar_sheets(SHEET_URL, SHEET, CREDENTIALS_PATH, TOKEN_PATH)
    _ws_detail = conectar_sheets(SHEET_URL, SHEET_DETAIL, CREDENTIALS_PATH, TOKEN_PATH)
    return _ws, _ws_detail


def escribir_pcm_detail_fase1(service_code, service_type, location, supplier,
                               rate_from, rate_to, nombre_pcm,
                               markup_dict, price_code="",
                               estado_fase1="LEIDO", error_fase1=""):
    """Fase 1: agrega una fila en PCM_Detail — por nombre de columna (no
    por posición fija), para no depender de que el Sheet tenga las
    columnas en el orden exacto de _DETAIL_HEADERS/_MARKUP_PAX."""
    global _columnas_detail
    estado_f2 = "PENDIENTE_APLICAR" if estado_fase1 == "LEIDO" else "SKIP"
    markup_str = json.dumps(markup_dict, ensure_ascii=False) if markup_dict else ""

    valores = {
        "TIMESTAMP":    datetime.now().strftime("%Y-%m-%d %H:%M"),
        "SERVICE CODE": service_code,
        "SERVICE TYPE": service_type,
        "LOCATION":     location,
        "SUPPLIER":     supplier,
        "RATE FROM":    rate_from,
        "RATE TO":      rate_to,
        "PCM":          nombre_pcm,
        "MARKUP JSON":  markup_str,
        "PRICE CODE":   price_code,
        "ESTADO F1":    estado_fase1,
        "ERROR F1":     error_fase1[:300] if error_fase1 else "",
        "ESTADO F2":    estado_f2,
        "ERROR F2":     "",
    }
    valores.update({
        rango: (markup_dict.get(rango, "") if markup_dict else "")
        for rango in _MARKUP_PAX
    })

    asegurar_columnas(_ws_detail, _DETAIL_HEADERS + _MARKUP_PAX)
    _columnas_detail = _ws_detail.row_values(1)
    agregar_fila_sheet(_ws_detail, _columnas_detail, valores)
    print(f"    📝 PCM_Detail F1: {nombre_pcm} → {estado_fase1}/{estado_f2}")


def _col_detail(nombre):
    """Nombre de columna real en PCM_Detail para un campo de _DETAIL_HEADERS
    — por texto exacto, o si no por la posición fija de _DETAIL_COL (mismo
    fallback que usaba el Excel original)."""
    if nombre in _columnas_detail:
        return nombre
    pos = _DETAIL_COL.get(nombre)
    if pos and 0 <= pos - 1 < len(_columnas_detail):
        return _columnas_detail[pos - 1]
    return nombre


def leer_pcm_detail_pendientes_fase2():
    global _columnas_detail
    try:
        filas_sheet, _columnas_detail = cargar_sheet(_ws_detail)
    except Exception:
        return []

    def val(fila, nombre):
        return str(fila.get(_col_detail(nombre)) or "").strip()

    filas = []
    for fila in filas_sheet:
        if FASE2_REAPLICAR_TODO:
            if not val(fila, "SERVICE CODE"):
                continue
        elif val(fila, "ESTADO F2").upper() != "PENDIENTE_APLICAR":
            continue

        try:
            markup_dict = json.loads(val(fila, "MARKUP JSON")) if val(fila, "MARKUP JSON") else {}
        except Exception:
            markup_dict = {}

        filas.append({
            "_row":         fila["__row_idx__"],
            "service_code": val(fila, "SERVICE CODE"),
            "service_type": val(fila, "SERVICE TYPE"),
            "location":     val(fila, "LOCATION"),
            "supplier":     val(fila, "SUPPLIER"),
            "rate_from":    val(fila, "RATE FROM"),
            "rate_to":      val(fila, "RATE TO"),
            "pcm":          val(fila, "PCM"),
            "markup_dict":  markup_dict,
            "price_code":   val(fila, "PRICE CODE"),
        })
    return filas


def marcar_pcm_detail_row(row_num, estado_fase2, error="",
                           rangos=None, v_viejos=None, v_nuevos=None):
    global _columnas_detail
    valores = {
        _col_detail("ESTADO F2"): estado_fase2,
        _col_detail("ERROR F2"):  error[:300] if error else "",
    }
    if rangos and v_viejos and v_nuevos:
        encabezados_rate = []
        for rango in rangos:
            encabezados_rate.append(f"VIEJO {rango}")
            encabezados_rate.append(f"NUEVO {rango}")
        asegurar_columnas(_ws_detail, encabezados_rate)
        _columnas_detail = _ws_detail.row_values(1)
        for rango in rangos:
            valores[f"VIEJO {rango}"] = v_viejos.get(rango, "")
            valores[f"NUEVO {rango}"] = v_nuevos.get(rango, "")
    actualizar_fila_sheet(_ws_detail, row_num, _columnas_detail, valores)


HEADER_ALIASES = {
    "LOCATION":         "location",
    "SUPPLIER":         "supplier",
    "SERVICE TYPE":     "service_type",
    "SERVICETYPE":      "service_type",
    "TIPO SERVICIO":    "service_type",
    "TIPO":             "service_type",
    "RATE FROM":        "rate_from",
    "FROM":             "rate_from",
    "RATE TO":          "rate_to",
    "TO":               "rate_to",
    "PRICE CODE":       "price_code",
    "PRICECODE":        "price_code",
    "SERVICE CODE":     "service_code",
    "SERVICECODE":      "service_code",
    "PRODUCT CODE":     "service_code",
    "COD SERVICIO":     "service_code",
    "SERVICIOS_PKG":    "servicios_pkg",
    "SERVICIOS PKG":    "servicios_pkg",
    "SERVICIOS_SKIP":   "servicios_skip",
    "SERVICIOS SKIP":   "servicios_skip",
    "ESTADO":           "estado",
    "STATUS":           "estado",
    "ERROR":            "error_msg",
    "TIMESTAMP":        "timestamp",
    "FECHA RECALCULATE PCM": "fecha_recalculate_pcm",
}

def _mapear_columnas(columnas):
    """Traduce cada campo lógico al nombre de columna real que tiene en el
    Sheet — por alias de encabezado, o si no por la posición fija del
    dict C (mismo fallback que get_col() usaba con el Excel)."""
    mapa = {}
    for encabezado in columnas:
        campo = HEADER_ALIASES.get(str(encabezado).strip().upper())
        if campo and campo not in mapa:
            mapa[campo] = encabezado
    for campo, pos in C.items():
        if campo not in mapa and 0 <= pos - 1 < len(columnas):
            mapa[campo] = columnas[pos - 1]
    return mapa


def leer_pendientes():
    global _columnas, _campo_a_columna
    ESTADOS_PENDIENTE = {"PENDIENTE", "PENDING", "PEND"}
    filas_sheet, _columnas = cargar_sheet(_ws)
    _campo_a_columna = _mapear_columnas(_columnas)

    col_estado = _campo_a_columna.get("estado")
    filas = []
    for fila in filas_sheet:
        est = str(fila.get(col_estado) or "").strip().upper()
        if est in ESTADOS_PENDIENTE:
            d = {campo: fila.get(col) for campo, col in _campo_a_columna.items()}
            d["_row"] = fila["__row_idx__"]
            filas.append(d)
    return filas


def marcar_procesando(row_num):
    col_estado = _campo_a_columna["estado"]
    actualizar_fila_sheet(_ws, row_num, _columnas, {col_estado: "PROCESANDO"})


def escribir_resultado_productos(row_num, estado, servicios_pkg="",
                                  servicios_skip="", error=""):
    valores = {
        _campo_a_columna["servicios_pkg"]:  servicios_pkg,
        _campo_a_columna["servicios_skip"]: servicios_skip,
        _campo_a_columna["timestamp"]:      datetime.now().strftime("%Y-%m-%d %H:%M"),
        _campo_a_columna["error_msg"]:      error[:300] if error else "",
        _campo_a_columna["estado"]:         estado,
    }
    actualizar_fila_sheet(_ws, row_num, _columnas, valores)


# ── Fase 1: resolver, agrupar y procesar ──────────────────────
# Mismo patrón de optimización que ya tiene tourplan_valorizacion_pkg_v3.py
# para su Fase 1 (_buscar_componente_fase1 / _procesar_pcm_todos_periodos):
# agrupar por servicio madre para no repetir búsqueda + Used In + apertura
# del PCM una vez por cada período del mismo servicio. Acá con una vuelta
# extra: cada fila del Excel puede traer un SERVICE_CODE puntual o disparar
# una búsqueda masiva (supplier+type+location) que devuelve varios códigos
# — no se sabe qué servicios madre va a tocar una fila hasta resolverla, así
# que el agrupamiento se hace en dos pasos: primero se RESUELVE cada fila a
# su lista de códigos (sin abrir nada todavía), y recién ahí se agrupan
# todas las tareas resultantes por servicio madre real, sin importar si el
# mismo código se descubrió por SERVICE_CODE directo en una fila o por
# búsqueda masiva en otra.

def _resolver_fila_fase1(driver, fila):
    """
    Paso 1: resuelve qué servicios madre corresponden a esta fila del
    Excel — directo si trae SERVICE_CODE, o vía búsqueda masiva si no —
    SIN abrir ningún producto todavía y SIN clasificar por categoría acá:
    la grilla de resultados de Product Search NO trae una columna
    Category poblada (confirmado por la usuaria), así que no hay forma
    confiable de distinguir Package de Non Accommodation en este punto.
    La clasificación real ocurre más adelante, por servicio madre ya
    agrupado, en _buscar_servicio_fase1() — recién ahí se abre el
    producto y se lee el dropdown "Service Category" (leer_service_category()),
    ANTES de pagar el costo de ir a Used In.

    Devuelve (tareas, skips):
      tareas: lista de dicts {fila, code, location, supplier, service_type}
              — todos los servicios encontrados, candidatos a procesar.
      skips:  lista de strings ("CODIGO [motivo]") — hoy siempre vacía acá;
              el filtro real se aplica más adelante, por servicio madre.
    """
    loc    = str(fila.get("location")     or "").strip()
    sup    = str(fila.get("supplier")     or "").strip()
    stype  = str(fila.get("service_type") or "").strip().upper()
    sc_in  = str(fila.get("service_code") or "").strip().upper()   # filtro opcional

    if sc_in:
        # Modo código específico: no hace búsqueda masiva, procesa solo ese código
        print(f"  🎯 SERVICE CODE especificado: '{sc_in}' → procesando solo ese servicio")
        todos_servicios = [{"code": sc_in, "description": ""}]
    else:
        # Modo general: buscar todos los servicios del supplier+type+location
        todos_servicios = buscar_servicios_madre(driver, sup, stype, loc)
        if not todos_servicios:
            raise Exception(
                f"No se encontraron servicios para supplier={sup} type={stype} loc={loc}")

    tareas = [
        {"fila": fila, "code": svc["code"],
         "location": loc, "supplier": sup, "service_type": stype}
        for svc in todos_servicios
    ]
    return tareas, []


def _agrupar_tareas_fase1(tareas):
    """
    Agrupa tareas (fila+código Package a procesar, ya resueltas por
    _resolver_fila_fase1) por servicio madre real
    (location+supplier+code+service_type), preservando el orden de
    primera aparición. Tareas para el mismo servicio madre pero distinto
    período (RATE FROM/RATE TO, tomado de la fila de cada una) quedan
    juntas — así se reutiliza la búsqueda del producto, el escaneo de
    Used In y la apertura del PCM en vez de repetirlos por cada período.
    """
    grupos, orden = {}, []
    for t in tareas:
        key = (t["location"].upper(), t["supplier"].upper(),
               t["code"].upper(), t["service_type"].upper())
        if key not in grupos:
            grupos[key] = []
            orden.append(key)
        grupos[key].append(t)
    return [grupos[k] for k in orden]


def _buscar_servicio_fase1(driver, code, location, supplier, service_type):
    """
    Busca el servicio madre UNA VEZ (Product Search + Service Category +
    días de operación + USED IN + lista de PCMs Package Header). El
    resultado se reutiliza para todas las tareas que compartan el mismo
    servicio madre pero distinto período, evitando repetir la búsqueda
    completa por cada fecha a valorizar del mismo producto — mismo patrón
    que tourplan_valorizacion_pkg_v3.py (_buscar_componente_fase1).

    Corta ANTES de ir a Used In si el producto es Non Accommodation: la
    grilla de resultados de Product Search no trae esa categoría (ver
    _resolver_fila_fase1), así que acá, con el producto ya abierto (no
    hace falta abrirlo de nuevo), se lee el dropdown real "Service
    Category" — si da Non Accommodation, se evita el escaneo caro de
    Used In (grilla con scroll virtual de hasta ~1900 filas) para un
    servicio que no corresponde procesar.

    Devuelve (handle_prod, dias_op, dias_str, pcm_list, motivo_skip).
    motivo_skip != "" → el caller debe saltear el grupo entero sin usar
    dias_op/dias_str/pcm_list (quedan en None).
    """
    handle_prod = driver.current_window_handle

    buscar_producto(driver, location, supplier, code,
                     service_type=service_type, handle_origen=handle_prod)

    cat = leer_service_category(driver)
    cat_up = cat.strip().upper()
    if "NON" in cat_up and "ACCOM" in cat_up:
        print(f"  ⏭ {code}: Service Category = '{cat}' → Non Accommodation, "
              f"SKIP sin ir a Used In")
        return handle_prod, None, None, None, "Non Accommodation"

    dias_op  = leer_dias_operacion(driver)
    dias_str = ",".join(sorted(dias_op))

    ir_a_used_in(driver)
    pcm_list = leer_pcm_list_package_header(driver)

    return handle_prod, dias_op, dias_str, pcm_list, ""


def _procesar_pcm_grupo_fase1(driver, pcm_info, grupo, handle_prod,
                               dias_op, dias_str, resultados_por_fila):
    """
    Abre UN PCM Package Header una sola vez y aplica Change Base Date +
    lee costos para TODOS los períodos (tareas) del grupo sobre esa misma
    ventana ya abierta, cerrándola recién al final — evita reabrir el
    mismo PCM (con el escaneo de Used In) una vez por cada período. Mismo
    patrón que tourplan_valorizacion_pkg_v3.py
    (_procesar_pcm_todos_periodos); acá siempre hay un solo PCM por grupo
    (Package Header), no una lista de PCMs a recorrer.

    resultados_por_fila: dict {row_num: {"pkg": [...], "skip": [...]}} que
    se va completando; el caller (MAIN) lo usa para escribir el resultado
    final de cada fila del Excel.
    """
    nombre_pcm = pcm_info["nombre"]
    code  = grupo[0]["code"]
    loc   = grupo[0]["location"]
    sup   = grupo[0]["supplier"]
    stype = grupo[0]["service_type"]

    print(f"\n  ▶ PCM: {nombre_pcm}  ({len(grupo)} período(s) a aplicar)")

    # _buscar_servicio_fase1() (llamado antes, para este mismo grupo) deja
    # la página parada en Used In, scroll 0 — tanto en la pasada rápida de
    # leer_pcm_list_package_header() (sin scroll) como en su fallback con
    # scroll completo (que también resetea a 0 antes de retornar). Por eso
    # NO hace falta volver a navegar a Used In acá: abrir_pcm() ya reordena
    # por Date (con el fix que no clickea de nuevo si ya está ordenado) y
    # busca la fila ahí mismo, sobre la misma vista. Evita por completo la
    # re-navegación que disparaba el bug del menú contextual de Tourplan
    # ("Used In" pasa a ítem de primer nivel una vez que ya se navegó ahí
    # antes en la sesión — ver sección "Fix — ir_a_used_in()..." en el
    # README). Si de todos modos esto fallara (por ejemplo si algo más
    # cambió el estado de la página entre medio), se reintenta una vez
    # con la re-navegación explícita de siempre, como red de seguridad.
    def _intentar_abrir():
        return abrir_pcm(driver, pcm_info, start_scroll=0)

    try:
        _handle_pcm, _found_scroll = _intentar_abrir()
    except Exception as _e_directo:
        print(f"    ⚠ Abrir PCM sin re-navegar falló ({_e_directo}) — "
              f"reintentando con re-navegación a Used In")
        try:
            ir_a_used_in(driver)
            time.sleep(2 * VELOCIDAD)
        except Exception as _nav_e:
            print(f"    ⚠ Re-nav USED IN falló ({_nav_e}), continuando con vista actual")
        try:
            _handle_pcm, _found_scroll = _intentar_abrir()
        except Exception as e:
            # No se pudo abrir el PCM: todos los períodos del grupo fallan.
            print(f"  ❌ Error abriendo {nombre_pcm}: {e}")
            ss(driver, f"error_f1_abrir_{nombre_pcm[:12]}")
            for t in grupo:
                fila = t["fila"]; row = fila["_row"]
                rf = str(fila.get("rate_from") or "").strip()
                rt = str(fila.get("rate_to")   or "").strip()
                escribir_pcm_detail_fase1(
                    service_code  = code, service_type = stype, location = loc, supplier = sup,
                    rate_from = rf, rate_to = rt, nombre_pcm = nombre_pcm, markup_dict = None,
                    estado_fase1 = "ERROR", error_fase1 = str(e),
                )
                resultados_por_fila[row]["skip"].append(f"{code} [ERROR: {str(e)[:60]}]")
            try:
                cerrar_pcm(driver, handle_prod)
            except Exception:
                pass
            return

    # Verificar si hay tarifas vencidas al abrir el PCM (una vez, aplica a
    # todos los períodos del grupo)
    _vencidos_apertura = manejar_tarifas_vencidas(driver, nombre_pcm)

    for t in grupo:
        fila = t["fila"]
        row  = fila["_row"]
        rf   = str(fila.get("rate_from")  or "").strip()
        rt   = str(fila.get("rate_to")    or "").strip()
        pc_in = str(fila.get("price_code") or "").strip()
        # FECHA RECALCULATE PCM (columna opcional): si la fila la trae
        # cargada, se usa como base del Change Base Date del PCM en vez de
        # RATE FROM. RATE FROM sigue siendo la base para crear/buscar el
        # período de rates a valorizar (rf/rt de acá abajo) — eso no cambia.
        rf_recalc = str(fila.get("fecha_recalculate_pcm") or "").strip() or rf

        print(f"    · Período [{rf} – {rt}] (fila {row})")
        try:
            # Cambiar fecha base: usa FECHA RECALCULATE PCM si la fila la
            # trae cargada, si no cae a RATE FROM (comportamiento de siempre)
            nueva_fecha, nueva_str, ant_str, _vencidos_fecha = cambiar_fecha_pcm(
                driver, dias_op, nombre_pcm, rate_from_str=rf_recalc)
            tiene_vencidos = _vencidos_apertura or _vencidos_fecha

            # Leer costos del PCM (VOUCHER COST en DASHBOARD)
            valores_pcm = leer_markup_commission(driver, nombre_pcm)
            if not valores_pcm:
                print("    ℹ️  VOUCHER COST vacío — se guardará con markup vacío")

            # Price Code: prioridad columna Excel, luego PCM
            price_code = pc_in or leer_price_code_pcm(driver, nombre_pcm)

            escribir_pcm_detail_fase1(
                service_code  = code,
                service_type  = stype,
                location      = loc,
                supplier      = sup,
                rate_from     = rf,
                rate_to       = rt,
                nombre_pcm    = nombre_pcm,
                markup_dict   = valores_pcm,
                price_code    = price_code,
                estado_fase1  = "LEIDO",
            )
            entrada_pkg = code + (" [COMPONENTES VENCIDOS]" if tiene_vencidos else "")
            resultados_por_fila[row]["pkg"].append(entrada_pkg)

        except Exception as e:
            print(f"  ❌ Error leyendo {nombre_pcm} (fila {row}): {e}")
            ss(driver, f"error_f1_{nombre_pcm[:12]}_row{row}")
            escribir_pcm_detail_fase1(
                service_code  = code,
                service_type  = stype,
                location      = loc,
                supplier      = sup,
                rate_from     = rf,
                rate_to       = rt,
                nombre_pcm    = nombre_pcm,
                markup_dict   = None,
                estado_fase1  = "ERROR",
                error_fase1   = str(e),
            )
            resultados_por_fila[row]["skip"].append(f"{code} [ERROR: {str(e)[:60]}]")

    cerrar_pcm(driver, handle_prod)


# ── Procesar una fila de PCM_Detail — Fase 2 ──────────────────
def procesar_fila_fase2(driver, pcm_row):
    """
    Fase 2: lee una fila de PCM_Detail (PENDIENTE_APLICAR) y aplica
    las tasas en el servicio madre correspondiente.
    """
    row          = pcm_row["_row"]
    service_code = pcm_row["service_code"]
    service_type = pcm_row["service_type"]
    location     = pcm_row["location"]
    supplier     = pcm_row["supplier"]
    rate_from    = pcm_row["rate_from"]
    rate_to      = pcm_row["rate_to"]
    nombre_pcm   = pcm_row["pcm"]
    markup_dict  = pcm_row["markup_dict"]
    price_code   = (pcm_row.get("price_code") or "").strip().upper() or PRICE_CODE_DEFAULT

    print(f"\n  ▶ FASE 2 — {service_code} [{service_type}]  [{rate_from} – {rate_to}]  "
          f"PCM: {nombre_pcm}")

    if not markup_dict:
        marcar_pcm_detail_row(row, "SKIP", error="markup_dict vacío")
        return

    try:
        v_viejos, v_nuevos, rangos, pendientes, error_post = \
            actualizar_rates_servicio_madre(
                driver,
                service_code  = service_code,
                supplier      = supplier,
                rate_from_str = rate_from,
                rate_to_str   = rate_to,
                location      = location,
                valores_pcm   = markup_dict,
                service_type  = service_type,
                price_code    = price_code,
            )

        if error_post:
            print(f"    ❌ {service_code}: {error_post}")
            ss(driver, f"error_f2_{service_code[:12]}")
            err_full = error_post + (("; " + "; ".join(pendientes)) if pendientes else "")
            marcar_pcm_detail_row(
                row, "ERROR", error=err_full,
                rangos=rangos, v_viejos=v_viejos, v_nuevos=v_nuevos)
            return

        estado = "OK_REVISAR" if pendientes else "OK"
        marcar_pcm_detail_row(
            row, estado, error="; ".join(pendientes),
            rangos=rangos, v_viejos=v_viejos, v_nuevos=v_nuevos)
        if pendientes:
            print(f"    🚩 {service_code} OK_REVISAR: {len(pendientes)} filas con valor inesperado")
        else:
            print(f"    ✅ {service_code} OK")

    except ProductoNoEncontrado as e:
        print(f"    ⚠ Servicio madre {service_code} no encontrado → SVS_MADRE_NO_ENCONTRADO")
        ss(driver, f"error_f2_{service_code[:12]}")
        marcar_pcm_detail_row(row, "SVS_MADRE_NO_ENCONTRADO", error=str(e))

    except Exception as e:
        print(f"    ❌ Error fase 2 {service_code}: {e}")
        ss(driver, f"error_f2_{service_code[:12]}")
        marcar_pcm_detail_row(row, "ERROR", error=str(e))


# ── Agrupamiento para evitar re-búsquedas ──────────────────────
def _agrupar_por_servicio_madre(filas):
    """
    Agrupa filas PENDIENTE_APLICAR de Fase 2 por servicio madre
    (location+service_code+supplier+service_type), preservando el orden de
    primera aparición. Filas para el mismo servicio madre pero distinto
    período quedan juntas para reutilizar la búsqueda del producto (no la
    navegación a RATES, que se repite por período — ver
    _buscar_servicio_madre_fase2).
    """
    grupos, orden = {}, []
    for f in filas:
        key = (
            str(f.get("location") or "").strip().upper(),
            str(f.get("service_code") or "").strip().upper(),
            str(f.get("supplier") or "").strip().upper(),
            str(f.get("service_type") or "").strip().upper(),
        )
        if key not in grupos:
            grupos[key] = []
            orden.append(key)
        grupos[key].append(f)
    return [grupos[k] for k in orden]


# ── MAIN ──────────────────────────────────────────────────────
print("="*60)
print(f"  📌 TOURPLAN NX — VALORIZACIÓN DESDE SERVICIO MADRE  v{VERSION}  [MODO={MODO}]")
print(f"  📌 VERSION {VERSION}  ·  {VERSION_FECHA}")
print("="*60)

if not SHEET_URL:
    raise ValueError("No se indicó la URL del Google Sheet (TOURPLAN_SHEET_URL).")

conectar_sheets_madre()
print(f"📄 Sheet: {SHEET_URL}")

pendientes_f1 = leer_pendientes()              if MODO in ("LEER",    "COMPLETO") else []
pendientes_f2 = leer_pcm_detail_pendientes_fase2() if MODO in ("APLICAR", "COMPLETO") else []

print(f"📊 Filas PENDIENTE Fase 1: {len(pendientes_f1)}")
print(f"📊 PCMs  PENDIENTE Fase 2: {len(pendientes_f2)}")

for f in pendientes_f1:
    sc = f.get('service_code') or ''
    print(f"   · F1 fila {f.get('_row')}: location={f.get('location')!r} "
          f"supplier={f.get('supplier')!r} service_type={f.get('service_type')!r} "
          f"service_code={sc!r} "
          f"price_code={f.get('price_code')!r} "
          f"rate={f.get('rate_from')!r}–{f.get('rate_to')!r}")

if not pendientes_f1 and not pendientes_f2:
    print("\n" + "="*60)
    print("⛔ EL SHEET NO TIENE FILAS PENDIENTE.")
    print("   Poné ESTADO=PENDIENTE en las filas a procesar (hoja PRODUCTOS)")
    print("   y volvé a correr.")
    print("="*60)
    raise SystemExit("Sin filas PENDIENTE.")
else:
    driver = crear_driver()
    _abortado = False
    try:
        login(driver)

        if pendientes_f1:
            print(f"\n{'─'*60}")
            print(f"  ▶▶ FASE 1 — Leyendo {len(pendientes_f1)} filas")
            print(f"{'─'*60}")

            resultados_por_fila = {f["_row"]: {"pkg": [], "skip": []} for f in pendientes_f1}

            # Paso 1: resolver cada fila a sus servicios madre (SERVICE_CODE
            # directo o búsqueda masiva), clasificando Package vs
            # Non-Accommodation — sin abrir ningún producto todavía.
            tareas_totales = []
            for fila in pendientes_f1:
                chequear_abort()
                print(f"\n{'='*60}")
                print(f"  FASE 1 — supplier={fila.get('supplier')} "
                      f"type={fila.get('service_type')} loc={fila.get('location') or '—'}  "
                      f"[{fila.get('rate_from')} – {fila.get('rate_to')}]")
                print(f"{'='*60}")
                marcar_procesando(fila["_row"])
                try:
                    tareas, skips = _resolver_fila_fase1(driver, fila)
                    tareas_totales.extend(tareas)
                    resultados_por_fila[fila["_row"]]["skip"].extend(skips)
                except Exception as e:
                    print(f"\n❌ ERROR F1 (resolución) fila {fila.get('_row')}: {e}")
                    ss(driver, f"error_f1_resolucion_row{fila.get('_row','x')}")
                    escribir_resultado_productos(
                        fila["_row"], "ERROR_F1", error=str(e))
                    del resultados_por_fila[fila["_row"]]

            # Paso 2: agrupar por servicio madre real y procesar cada grupo
            # una sola vez (buscar producto + identificar PCM + abrirlo),
            # aplicando todos los períodos del grupo antes de cerrarlo.
            if tareas_totales:
                grupos = _agrupar_tareas_fase1(tareas_totales)
                print(f"\n  📋 {len(tareas_totales)} servicio(s) Package a procesar, "
                      f"agrupados en {len(grupos)} servicio(s) madre distinto(s)")

                for grupo in grupos:
                    chequear_abort()
                    code = grupo[0]["code"]
                    if len(grupo) > 1:
                        print(f"  ℹ️  {len(grupo)} período(s) para el mismo servicio "
                              f"madre {code} — reutilizando búsqueda y PCM")

                    try:
                        handle_prod, dias_op, dias_str, pcm_list, motivo_skip = \
                            _buscar_servicio_fase1(
                                driver, code, grupo[0]["location"], grupo[0]["supplier"],
                                grupo[0]["service_type"])
                    except Exception as e:
                        print(f"\n❌ ERROR F1 (búsqueda de servicio) {code}: {e}")
                        ss(driver, f"error_f1_busqueda_{code[:12]}")
                        for t in grupo:
                            resultados_por_fila[t["fila"]["_row"]]["skip"].append(
                                f"{code} [ERROR: {str(e)[:60]}]")
                        try:
                            cerrar_pcm(driver, driver.window_handles[0])
                        except Exception:
                            pass
                        continue

                    if motivo_skip:
                        for t in grupo:
                            resultados_por_fila[t["fila"]["_row"]]["skip"].append(
                                f"{code} [{motivo_skip}]")
                        cerrar_pcm(driver, handle_prod)
                        continue

                    if not pcm_list:
                        razon = "SIN PCM - CARGA DIRECTA"
                        print(f"    ⚠ {code}: no hay PCM Package Header → {razon}")
                        for t in grupo:
                            resultados_por_fila[t["fila"]["_row"]]["skip"].append(
                                f"{code} [{razon}]")
                        cerrar_pcm(driver, handle_prod)
                        continue

                    if len(pcm_list) > 1:
                        print(f"    ⚠ {code}: {len(pcm_list)} PCMs Package Header "
                              f"encontrados → procesando el primero: {pcm_list[0]['nombre']}")

                    pcm_info = pcm_list[0]

                    try:
                        _procesar_pcm_grupo_fase1(driver, pcm_info, grupo, handle_prod,
                                                   dias_op, dias_str, resultados_por_fila)
                    except Exception as e:
                        # _procesar_pcm_grupo_fase1 ya maneja sus propios errores
                        # por PCM/período; esto es un resguardo para no frenar el
                        # resto de los grupos ante algo inesperado.
                        print(f"\n❌ ERROR F1 inesperado procesando PCM "
                              f"{pcm_info.get('nombre','?')}: {e}")
                        ss(driver, f"error_f1_pcm_inesperado_"
                                   f"{str(pcm_info.get('nombre','x'))[:12]}")
                        for t in grupo:
                            resultados_por_fila[t["fila"]["_row"]]["skip"].append(
                                f"{code} [ERROR inesperado: {str(e)[:60]}]")
                        try:
                            cerrar_pcm(driver, handle_prod)
                        except Exception:
                            pass

            # Paso 3: escribir el resultado final de cada fila del Excel,
            # juntando lo que aportó cada servicio madre resuelto a partir
            # de ella.
            for fila in pendientes_f1:
                row = fila["_row"]
                if row not in resultados_por_fila:
                    continue  # ya se escribió como ERROR_F1 en la resolución
                r = resultados_por_fila[row]
                if r["pkg"]:
                    estado = "FASE1_OK"
                elif r["skip"]:
                    estado = "FASE1_SKIP"
                else:
                    estado = "FASE1_OK"  # sin servicios de ese tipo, sin error
                escribir_resultado_productos(
                    row_num       = row,
                    estado        = estado,
                    servicios_pkg = "\n".join(r["pkg"]),
                    servicios_skip= "\n".join(r["skip"]),
                )
                print(f"\n✅ Fase 1 fila {row} — {len(r['pkg'])} Package procesados, "
                      f"{len(r['skip'])} saltados")

        if MODO == "COMPLETO" and pendientes_f1:
            pendientes_f2 = leer_pcm_detail_pendientes_fase2()
            print(f"\n  📋 PCMs listos para Fase 2: {len(pendientes_f2)}")

        if FASE2_LIMIT and pendientes_f2:
            print(f"  ⚡ FASE2_LIMIT={FASE2_LIMIT} → proceso solo los primeros "
                  f"{FASE2_LIMIT} de {len(pendientes_f2)}")
            pendientes_f2 = pendientes_f2[:FASE2_LIMIT]

        if pendientes_f2:
            print(f"\n{'─'*60}")
            print(f"  ▶▶ FASE 2 — Aplicando rates en {len(pendientes_f2)} servicios madre")
            print(f"{'─'*60}")
            for grupo in _agrupar_por_servicio_madre(pendientes_f2):
                chequear_abort()
                ref = grupo[0]
                service_code_ref = str(ref.get("service_code") or "").strip()
                location_ref = str(ref.get("location") or "").strip()
                supplier_ref = str(ref.get("supplier") or "").strip()
                service_type_ref = str(ref.get("service_type") or "").strip()

                try:
                    _buscar_servicio_madre_fase2(
                        driver, service_code_ref, supplier_ref, location_ref,
                        service_type=service_type_ref)
                except ProductoNoEncontrado as e:
                    print(f"    ⚠ Servicio madre {service_code_ref} no existe en Tourplan "
                          f"→ SVS_MADRE_NO_ENCONTRADO")
                    ss(driver, f"error_f2_busqueda_{service_code_ref[:12]}")
                    for pcm_row in grupo:
                        marcar_pcm_detail_row(pcm_row["_row"], "SVS_MADRE_NO_ENCONTRADO", error=str(e))
                    continue
                except Exception as e:
                    print(f"\n❌ ERROR F2 (búsqueda de servicio madre) {service_code_ref}: {e}")
                    ss(driver, f"error_f2_busqueda_{service_code_ref[:12]}")
                    for pcm_row in grupo:
                        marcar_pcm_detail_row(pcm_row["_row"], "ERROR", error=str(e))
                    continue

                if len(grupo) > 1:
                    print(f"  ℹ️  {len(grupo)} período(s) para el mismo servicio madre "
                          f"{service_code_ref} — reutilizando búsqueda")

                for pcm_row in grupo:
                    chequear_abort()
                    procesar_fila_fase2(driver, pcm_row)

    except AbortadoPorUsuario:
        _abortado = True
        print("\n⏸️  Corrida abortada por el usuario — las filas que no llegaron a "
              "procesarse quedan en PENDIENTE para retomar en otra corrida.")
    finally:
        try:
            logout(driver)
        except Exception as _e:
            print(f"  ⚠ Logout falló: {_e}")
        driver.quit()
        _dur = int(time.time() - _t_inicio)
        _h, _resto = divmod(_dur, 3600)
        _m, _s = divmod(_resto, 60)
        print("\n🏁 Finalizado.")
        print(f"⏱ Duración total: {_h:d}h {_m:02d}m {_s:02d}s"
              if _h else f"⏱ Duración total: {_m:d}m {_s:02d}s")
        print(f"📂 Screenshots: {SS_DIR}")

    if _abortado:
        sys.exit(ABORT_EXIT_CODE)
