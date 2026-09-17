# ============================================================
# TOURPLAN NX — COPIAR PCM DESDE SERVICIO MADRE  v1.0
# Copia adaptada para tp-nx-app (app local). El original para
# Google Colab vive sin cambios en el repo Copy-PCM-Linkeo.
# ------------------------------------------------------------
#   VERSION      : 1.0
#   DESCRIPCIÓN  : Dado un código de servicio madre (Product Code),
#                  lo abre, va a USED IN, busca el PCM tipo
#                  "Package Header" con nombre PKG-*, lo abre y
#                  ejecuta COPY PCM. El nuevo nombre del PCM se
#                  arma tomando el nombre del PCM original y
#                  reemplazando todo lo que hay desde la 5ta
#                  posición (después del 4to guion) en adelante
#                  por el "CODIGO NUEVO" indicado en el Excel.
#                  Ej: PKG-BUE-EX-6CRIO1-1XXXXX + CODIGO NUEVO
#                  "1COCLA" → PKG-BUE-EX-6CRIO1-1COCLA
#                  Luego de OK, maneja el diálogo de recálculo
#                  (REPLACE ALL + Update Exchange + YES) y la
#                  ventana de tarifas vencidas si aparece.
#   NOTA         : Este script cubre hasta la copia del PCM
#                  inclusive. El siguiente paso (qué hacer con
#                  el PCM ya copiado) se agrega en una v1.1.
#   CAMBIOS v1.1 : Fix en buscar_producto(): al completar Location, si la
#                  fila `tr.selectedRow` no aparecía a tiempo, el fallback
#                  clickeaba siempre la PRIMERA fila de la tabla de
#                  sugerencias sin verificar que fuera el Location exacto
#                  (mismo bug ya arreglado en notas_srv.py v1.13/v1.14).
#                  Ahora busca entre las filas sugeridas la que tiene una
#                  celda con texto EXACTO (case-insensitive) igual al
#                  Location pedido; si ninguna coincide, no clickea nada.
# ============================================================

VERSION       = "1.1"
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
from common.sheets_client import conectar_sheets, cargar_sheet, actualizar_fila_sheet

CHROMIUM_BIN, ver_chrome = find_or_prepare_chrome()

from webdriver_manager.chrome import ChromeDriverManager

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys

# ── Config ────────────────────────────────────────────────────
import getpass

from common.user_config import (
    CREDENTIALS_PATH as _CREDENTIALS_PATH_DEFAULT,
    TOKEN_PATH as _TOKEN_PATH_DEFAULT,
)

USERNAME         = os.environ.get("TOURPLAN_USERNAME") or input("Usuario Tourplan (minúsculas): ").strip()
PASSWORD         = os.environ.get("TOURPLAN_PASSWORD") or getpass.getpass("Password Tourplan: ")
BASE_URL         = os.environ.get("TOURPLAN_BASE_URL", "https://tourplannx.eurotur.com.ar/TourplanNX_Test")
SHEET_URL        = os.environ.get("TOURPLAN_SHEET_URL", "")
SHEET            = os.environ.get("TOURPLAN_HOJA", "PRODUCTOS")
CREDENTIALS_PATH = os.environ.get("TOURPLAN_CREDENTIALS_PATH", _CREDENTIALS_PATH_DEFAULT)
TOKEN_PATH       = os.environ.get("TOURPLAN_TOKEN_PATH", _TOKEN_PATH_DEFAULT)
SS_DIR           = os.environ.get("TOURPLAN_SS_DIR", "/content/screenshots")
HEADLESS         = os.environ.get("TOURPLAN_HEADLESS", "0").strip() in ("1", "true", "True")
os.makedirs(SS_DIR, exist_ok=True)

# Multiplicador de tiempos de espera. 1.0 = test  |  1.5 = producción
VELOCIDAD = 1.5

# Mapping service type → número de opción en el sidebar de Product Search
STYPE_SIDEBAR = {"EX": "04", "TF": "03"}

# Siglas que Tourplan agrupa en el sidebar bajo el prefijo "Z*NO USAR*"
# (organización interna de la empresa, confirmado por la usuaria 2026-08-18 —
# no significa que estén deprecadas). Varias comparten una palabra con una
# categoría real (ej. TN="Transfer Non-Accom" vs TF="Transfer") — usadas en
# buscar_producto() para no matchear la sigla real dentro de un ítem "no usar"
# ni viceversa.
Z_NO_USAR_SIGLAS = {"GA", "TK", "TN", "CH", "CM", "CO", "EN", "GU", "PJ", "ST", "TR", "TI"}

# Columnas Excel PRODUCTOS (base 1)
C = {
    "location":         1,   # A - opcional, 3 letras (BUE)
    "supplier":         2,   # B - obligatorio
    "service_type":     3,   # C - obligatorio (EX, TF, etc.)
    "service_code":     4,   # D - obligatorio: código del servicio madre a abrir
    "codigo_nuevo":     5,   # E - obligatorio: reemplaza la 5ta posición (post 4to guion)
                              #     del nombre del PCM original
    "estado":           7,   # G - PENDIENTE para procesar
    "pcm_original":     8,   # H - SALIDA: nombre del PCM original encontrado
    "pcm_copiado":      9,   # I - SALIDA: nombre final asignado a la copia
    "error_msg":       10,   # J - SALIDA: error
    "timestamp":       11,   # K - SALIDA: timestamp
}

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

# ── Login / Logout ──────────────────────────────────────────────
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

# ── Fechas (usado por el filtro >365 días en Used In) ──────────
MESES_ES = {"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
            "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}

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

# ── Buscar producto por código (abrir servicio madre específico) ──
class ProductoNoEncontrado(Exception):
    pass

def buscar_producto(driver, location, supplier, codigo, service_type=None, handle_origen=None):
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

# ── USED IN ───────────────────────────────────────────────────
def ir_a_used_in(driver):
    hamburger(driver)
    ss(driver, "hamburger_product")
    menu_item(driver, "UTILITIES")
    ss(driver, "utilities_expandido")
    submenu_text(driver, "Used In")
    time.sleep(5 * VELOCIDAD)
    ss(driver, "used_in_tabla")

def leer_pcm_list_package_header(driver):
    """
    Lee la tabla Used In y retorna PCMs tipo Package Header
    con nombre que empieza con PKG-.
    """
    try:
        th = wait(driver, "th.tpcol-UsageTypeLabel > span > span", t=5)
        jc(driver, th); time.sleep(2 * VELOCIDAD)
        ss(driver, "used_in_sorted")
    except: pass

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
            continue

        nombre_limpio = re.split(r'\s*--\s*', nombre)[0].strip()

        tipo_up       = tipo.strip().upper()
        nombre_up     = nombre_limpio.upper()
        fila_txt_up   = info.get("_fila_txt","").upper()

        if any(x in nombre_up for x in ["FLAG AS DELETED","FLAGGED","DELETED","CANCELLED","CANCELED","INACTIVE"]):
            print(f"      → SKIP (nombre indica eliminado)")
            continue
        if any(x in fila_txt_up for x in ["FLAG AS DELETED","FLAGGED"]):
            print(f"      → SKIP (fila indica eliminado)")
            continue
        if any(x in tipo_up for x in ["DELETE","FLAG","CANCEL","INACTIVE"]):
            print(f"      → SKIP (tipo indica eliminado: {tipo})")
            continue

        if "PACKAGE HEADER" not in tipo_up:
            print(f"      → SKIP (no es Package Header: tipo='{tipo}')")
            continue

        if not nombre_up.startswith("PKG-"):
            print(f"      → SKIP (nombre no empieza con PKG-: '{nombre_limpio}')")
            continue

        fecha_desde_str = info.get("From","") or info.get("FROM","") or ""
        fecha_hasta_str = info.get("To","")   or info.get("TO","")   or ""
        f_desde = parsear_fecha(fecha_desde_str)
        f_hasta = parsear_fecha(fecha_hasta_str)
        if f_desde and f_hasta and (f_hasta - f_desde).days > 365:
            print(f"      → SKIP (rango > 365 días)")
            continue

        el_fila = filas_dom[i] if i < len(filas_dom) else None
        pcm_list.append({
            "nombre":  nombre_limpio,
            "tipo":    tipo,
            "fila_el": el_fila,
            "info":    info,
        })

    print(f"  📋 PCMs Package Header válidos: {[p['nombre'] for p in pcm_list]}")
    return pcm_list

def abrir_pcm(driver, pcm_info, start_scroll=0):
    nombre = pcm_info["nombre"]
    handles_antes = set(driver.window_handles)

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

    step = 1500

    def _scan_from(start_pos, max_steps=200):
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

def cerrar_pcm(driver, handle_prod):
    for h in set(driver.window_handles) - {handle_prod}:
        try: driver.switch_to.window(h); driver.close()
        except: pass
    driver.switch_to.window(handle_prod)
    time.sleep(1)

def manejar_tarifas_vencidas(driver, nombre_pcm=""):
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

# ── Diálogo de recálculo (REPLACE ALL + Update Exchange + YES) ─
# Reutilizado tanto por Change Base Date como por Copy PCM: es el
# mismo widget (tp-get-recalc-parameters) en ambos flujos.
def manejar_dialogo_recalculo(driver, nombre_pcm=""):
    try:
        WebDriverWait(driver, 6).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#calculatereplaceall")))
        print("    Diálogo recálculo → siguiendo grabación exacta")
        ss(driver, f"recalc_dlg_{nombre_pcm[:12]}")
        try:
            jc(driver, driver.find_element(By.CSS_SELECTOR,
               "tp-group.tpgroup-servicelinepricesgroup li:nth-of-type(1) > tp-label > label"))
            time.sleep(0.3)
        except:
            pass
        jc(driver, wait(driver, "#calculatereplaceall"))
        time.sleep(0.5)
        try:
            jc(driver, driver.find_element(By.CSS_SELECTOR,
               "#exchange-rate-panel > tp-label > label"))
            time.sleep(0.3)
            jc(driver, driver.find_element(By.CSS_SELECTOR,
               "tp-group.tpgroup-servicelinepricesgroup tp-checkbox > label"))
            time.sleep(0.3)
        except:
            pass
        jc(driver, wait(driver, "tp-button.yes > button"))
        time.sleep(5 * VELOCIDAD)
        ss(driver, f"recalc_done_{nombre_pcm[:12]}")
        return True
    except Exception as _re:
        print(f"    ⚠ No apareció el diálogo de recálculo: {_re}")
        return False

# ── Nombre del PCM copiado ──────────────────────────────────────
def calcular_nombre_pcm_nuevo(nombre_viejo, codigo_nuevo):
    """
    Reemplaza todo lo que hay desde la 5ta posición del nombre
    (después del 4to guion) por 'codigo_nuevo'.
    Ej: PKG-BUE-EX-6CRIO1-1XXXXX + '1COCLA' → PKG-BUE-EX-6CRIO1-1COCLA
    """
    partes = nombre_viejo.split("-", 4)
    if len(partes) < 5:
        raise Exception(
            f"El nombre del PCM '{nombre_viejo}' no tiene el formato esperado "
            f"PKG-LOC-TYPE-XXX-CODIGO (menos de 5 posiciones separadas por guion)")
    partes[4] = codigo_nuevo
    return "-".join(partes)

# ── COPY PCM ─────────────────────────────────────────────────
def copiar_pcm(driver, nombre_pcm_actual, codigo_nuevo):
    """
    Con el PCM ya abierto (ventana actual = ventana del PCM):
      1. Ir a PCM DETAILS
      2. Click COPY PCM
      3. Escribir el nuevo nombre en el input (calculado a partir
         del nombre actual + codigo_nuevo)
      4. Click fuera (blur) + OK
      5. Manejar diálogo de recálculo (REPLACE ALL + Update Exchange + YES)
      6. Manejar ventana de tarifas vencidas si aparece
    Retorna (nombre_pcm_nuevo, tiene_vencidos).
    """
    nombre_nuevo = calcular_nombre_pcm_nuevo(nombre_pcm_actual, codigo_nuevo)
    print(f"  📄 Copiando PCM '{nombre_pcm_actual}' → nuevo nombre: '{nombre_nuevo}'")

    hamburger(driver)
    time.sleep(1 * VELOCIDAD)

    # Click en el label del header del nav (paso de la grabación).
    # No siempre es imprescindible, pero lo replicamos tal cual.
    try:
        lbl_head = driver.find_element(By.CSS_SELECTOR, "div.nav-head > label")
        jc(driver, lbl_head)
        time.sleep(1 * VELOCIDAD)
    except Exception:
        pass

    # Click en "PCM DETAILS" para desplegar su submenú. A veces requiere
    # más de un click para que Angular agregue la clase tpnavmenuopen.
    submenu_abierto = False
    for _intento in range(3):
        menu_item(driver, "PCM DETAILS")
        time.sleep(1 * VELOCIDAD)
        if driver.find_elements(By.CSS_SELECTOR, "li.tpnavmenuopen li"):
            submenu_abierto = True
            break

    if not submenu_abierto:
        ss(driver, f"pcm_details_sin_submenu_{nombre_pcm_actual[:12]}")
        dump(driver, f"pcm_details_sin_submenu_{nombre_pcm_actual[:12]}")
        raise Exception("El submenú de PCM DETAILS no se desplegó (sin li.tpnavmenuopen)")

    ss(driver, f"pcm_details_submenu_{nombre_pcm_actual[:12]}")

    # Click en el primer ítem del submenú ya desplegado → GENERAL SETUP
    sub_el = wait(driver, "li.tpnavmenuopen li:nth-of-type(1)", t=6)
    jc(driver, sub_el)
    time.sleep(3 * VELOCIDAD)
    ss(driver, f"pcm_details_precopy_{nombre_pcm_actual[:12]}")

    try:
        btn_copy = wait(driver, "tp-button.copypcm > button", t=8)
    except Exception as _e:
        diag = driver.execute_script("""
            function vis(e){ return !!(e.offsetWidth || e.offsetHeight
                                       || e.getClientRects().length); }
            var tpbtns = Array.from(document.querySelectorAll('tp-button')).filter(vis)
                .map(function(b){
                    return {clase: b.className, texto: (b.innerText||'').trim().slice(0,30)};
                });
            var btns = Array.from(document.querySelectorAll('button')).filter(vis)
                .map(function(b){ return (b.innerText||'').trim().slice(0,30); })
                .filter(function(t){ return t; });
            return {tpbtns: tpbtns, btns: btns, url: window.location.href};
        """)
        print(f"    ⚠ No encontré 'tp-button.copypcm > button' tras 8s.")
        print(f"    tp-button visibles: {diag.get('tpbtns')}")
        print(f"    <button> visibles : {diag.get('btns')}")
        print(f"    URL: {diag.get('url')}")
        ss(driver, f"pcm_copy_sin_boton_{nombre_pcm_actual[:12]}")
        dump(driver, f"pcm_copy_sin_boton_{nombre_pcm_actual[:12]}")
        raise Exception(
            f"No encontré el botón COPY PCM en PCM DETAILS "
            f"(tp-button visibles: {diag.get('tpbtns')})") from _e
    jc(driver, btn_copy)
    time.sleep(3 * VELOCIDAD)
    ss(driver, f"pcm_copy_dlg_{nombre_pcm_actual[:12]}")

    inp_nombre = None
    for sel in ["div.tpcol1 li:nth-of-type(1) input",
                "#pcmDetailName input",
                "tp-dialog input[type='text']"]:
        try:
            inp_nombre = wait(driver, sel, t=5)
            break
        except Exception:
            pass
    if inp_nombre is None:
        ss(driver, f"pcm_copy_sin_input_{nombre_pcm_actual[:12]}")
        dump(driver, f"pcm_copy_sin_input_{nombre_pcm_actual[:12]}")
        raise Exception("No encontré el campo de nombre en el diálogo COPY PCM")

    set_val(driver, inp_nombre, nombre_nuevo)
    time.sleep(0.3)

    # Click fuera del input para disparar blur/validación (igual que grabación)
    try:
        jc(driver, driver.find_element(By.CSS_SELECTOR, "div.withmargin"))
    except Exception:
        pass
    time.sleep(0.3)
    ss(driver, f"pcm_copy_lleno_{nombre_pcm_actual[:12]}")

    btn_ok = wait(driver, "tp-button.ok > button", t=6)
    jc(driver, btn_ok)
    time.sleep(4 * VELOCIDAD)
    ss(driver, f"pcm_copy_ok_click_{nombre_pcm_actual[:12]}")

    manejar_dialogo_recalculo(driver, nombre_pcm_actual)

    tiene_vencidos = manejar_tarifas_vencidas(driver, nombre_pcm_actual)
    if tiene_vencidos:
        time.sleep(3 * VELOCIDAD)

    ss(driver, f"pcm_copy_final_{nombre_pcm_actual[:12]}")
    print(f"    ✅ PCM copiado como '{nombre_nuevo}'")
    return nombre_nuevo, tiene_vencidos

# ── Linkeo del PCM copiado a su servicio madre ────────────────
# Checkboxes de handling/documentation que se marcan siempre igual
# (según grabación Chrome DevTools) tras seleccionar el producto.
_HANDLING_CHECKBOX_SELECTORS = [
    "tp-group.tpgroup-packagehandling li:nth-of-type(2) i",
    "tp-group.tpgroup-packagehandling li:nth-of-type(6) i",
    "tp-group.servicehandling li:nth-of-type(2) i",
    "#handling > div li:nth-of-type(3) i",
    "tp-group.tpgroup-documentation li:nth-of-type(2) i",
    "tp-group.tpgroup-documentation li:nth-of-type(5) i",
    "tp-group.tpgroup-documentation li:nth-of-type(6) i",
]

def _marcar_handling_checkboxes(driver, etiqueta=""):
    for sel in _HANDLING_CHECKBOX_SELECTORS:
        try:
            el = driver.find_element(By.CSS_SELECTOR, sel)
            jc(driver, el)
            time.sleep(0.4)
        except Exception as e:
            print(f"    ⚠ No até checkbox '{sel}' ({etiqueta}): {e}")


def linkear_pcm_a_servicio(driver, location, supplier, codigo_nuevo):
    """
    Con el PCM (copiado) ya abierto en la ventana actual:
      1. OPERATION → submenú "Package Setup"
      2. PRODUCT SEARCH → completa Location/Supplier/Código
      3. FIND PRODUCTS → SELECT en el resultado
      4. Tilda los checkboxes fijos de handling/documentation
      5. SAVE
    """
    print(f"\n  🔗 Linkeando PCM copiado a servicio madre: "
          f"{location or '—'}/{supplier}/{codigo_nuevo}")

    hamburger(driver)
    time.sleep(1 * VELOCIDAD)

    submenu_abierto = False
    for _intento in range(3):
        menu_item(driver, "OPERATION")
        time.sleep(1 * VELOCIDAD)
        if driver.find_elements(By.CSS_SELECTOR, "li.tpnavmenuopen li"):
            submenu_abierto = True
            break

    if not submenu_abierto:
        ss(driver, f"link_sin_submenu_{codigo_nuevo[:12]}")
        dump(driver, f"link_sin_submenu_{codigo_nuevo[:12]}")
        raise Exception("El submenú de OPERATION no se desplegó (sin li.tpnavmenuopen)")

    ss(driver, f"link_operation_submenu_{codigo_nuevo[:12]}")

    try:
        submenu_text(driver, "Package Setup", timeout=6)
        print("    Submenú 'Package Setup' seleccionado")
    except Exception as _e:
        print(f"    ⚠ No até el submenú 'Package Setup' por texto ({_e}); "
              f"pruebo el primer ítem del submenú")
        sub_el = wait(driver, "li.tpnavmenuopen li:nth-of-type(1)", t=6)
        jc(driver, sub_el)
    time.sleep(3 * VELOCIDAD)
    ss(driver, f"link_package_setup_{codigo_nuevo[:12]}")

    btn_search = wait(driver, "tp-button.findproduct > button", t=8)
    jc(driver, btn_search)
    time.sleep(3 * VELOCIDAD)
    ss(driver, f"link_search_dlg_{codigo_nuevo[:12]}")

    if location:
        try:
            inp_loc = wait(driver, "tp-dialog div > div > ul > li:nth-of-type(1) input", t=6)
            inp_loc.click(); time.sleep(0.3)
            set_val(driver, inp_loc, location)
            time.sleep(1)
            inp_loc.send_keys(Keys.ENTER)
            time.sleep(0.4)
            inp_loc.send_keys(Keys.TAB)
            time.sleep(1)
            print(f"    Location: {location}")
        except Exception as e:
            print(f"    ⚠ No pude llenar Location: {e}")

    inp_sup = wait(driver, "tp-dialog div > div > ul > li:nth-of-type(2) input", t=6)
    inp_sup.click(); time.sleep(0.3)
    set_val(driver, inp_sup, supplier)
    time.sleep(1.2)
    try:
        inp_sup.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.3)
        inp_sup.send_keys(Keys.ENTER)
    except Exception:
        pass
    time.sleep(0.3)
    try:
        inp_sup.send_keys(Keys.TAB)
    except Exception:
        pass
    time.sleep(1)
    print(f"    Supplier: {supplier}")

    inp_cod = None
    for sel in ["#optionCode input", "tp-dialog div > div > ul > li:nth-of-type(3) input"]:
        try:
            inp_cod = wait(driver, sel, t=5)
            break
        except Exception:
            pass
    if inp_cod is None:
        ss(driver, f"link_sin_input_codigo_{codigo_nuevo[:12]}")
        dump(driver, f"link_sin_input_codigo_{codigo_nuevo[:12]}")
        raise Exception("No encontré el input de código (optionCode) en PRODUCT SEARCH")
    inp_cod.click(); time.sleep(0.3)
    set_val(driver, inp_cod, codigo_nuevo)
    time.sleep(0.5)
    print(f"    Código: {codigo_nuevo}")
    ss(driver, f"link_search_lleno_{codigo_nuevo[:12]}")

    btn_find = wait(driver, "#tabSelection li:nth-of-type(3) button", t=8)
    jc(driver, btn_find)
    time.sleep(4 * VELOCIDAD)
    ss(driver, f"link_resultados_{codigo_nuevo[:12]}")

    try:
        btn_select = wait(driver, "#tabResults tbody button", t=8)
    except Exception as e:
        ss(driver, f"link_sin_resultados_{codigo_nuevo[:12]}")
        dump(driver, f"link_sin_resultados_{codigo_nuevo[:12]}")
        raise Exception(
            f"No encontré resultados para {location or '—'}/{supplier}/{codigo_nuevo} "
            f"en PRODUCT SEARCH") from e
    jc(driver, btn_select)
    time.sleep(4 * VELOCIDAD)
    ss(driver, f"link_seleccionado_{codigo_nuevo[:12]}")

    _marcar_handling_checkboxes(driver, etiqueta=codigo_nuevo)
    ss(driver, f"link_checkboxes_{codigo_nuevo[:12]}")

    btn_save = wait(driver, "tp-button.save > button", t=8)
    jc(driver, btn_save)
    time.sleep(4 * VELOCIDAD)
    ss(driver, f"link_guardado_{codigo_nuevo[:12]}")
    print(f"    ✅ PCM linkeado a {location or '—'}/{supplier}/{codigo_nuevo}")

# ── Excel ──────────────────────────────────────────────────────
HEADER_ALIASES = {
    "LOCATION":         "location",
    "SUPPLIER":         "supplier",
    "SERVICE TYPE":     "service_type",
    "SERVICE CODE":     "service_code",
    "CODIGO NUEVO":     "codigo_nuevo",
    "CÓDIGO NUEVO":     "codigo_nuevo",
    "ESTADO":           "estado",
    "PCM ORIGINAL":     "pcm_original",
    "PCM COPIADO":      "pcm_copiado",
    "ERROR":            "error_msg",
    "TIMESTAMP":        "timestamp",
}

_ws = None            # worksheet ya conectado (ver conectar_sheet())
_columnas = None       # encabezados tal cual figuran en el Sheet
_campo_a_columna = None  # "estado" -> nombre real de columna en el Sheet

def _mapear_columnas(columnas):
    """Traduce cada campo lógico (location/supplier/.../estado/...) al
    nombre de columna real que tiene en el Sheet — por alias de
    encabezado si lo encuentra, o si no por la posición original del
    Excel (dict C), igual que hacía get_col() con el Excel."""
    mapa = {}
    for encabezado in columnas:
        campo = HEADER_ALIASES.get(str(encabezado).strip().upper())
        if campo and campo not in mapa:
            mapa[campo] = encabezado
    for campo, pos in C.items():
        if campo not in mapa and 0 <= pos - 1 < len(columnas):
            mapa[campo] = columnas[pos - 1]
    return mapa

def conectar_sheet():
    global _ws
    _ws = conectar_sheets(SHEET_URL, SHEET, CREDENTIALS_PATH, TOKEN_PATH)
    return _ws

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

def escribir_resultado(row_num, estado, pcm_original="", pcm_copiado="", error=""):
    valores = {
        _campo_a_columna["pcm_original"]: pcm_original,
        _campo_a_columna["pcm_copiado"]:  pcm_copiado,
        _campo_a_columna["timestamp"]:    datetime.now().strftime("%Y-%m-%d %H:%M"),
        _campo_a_columna["error_msg"]:    error[:300] if error else "",
        _campo_a_columna["estado"]:       estado,
    }
    actualizar_fila_sheet(_ws, row_num, _columnas, valores)

# ── Procesar una fila ────────────────────────────────────────
def procesar_fila(driver, fila):
    loc         = str(fila.get("location")     or "").strip()
    sup         = str(fila.get("supplier")     or "").strip()
    stype       = str(fila.get("service_type") or "").strip().upper()
    code        = str(fila.get("service_code") or "").strip().upper()
    codigo_nuevo = str(fila.get("codigo_nuevo") or "").strip()
    row         = fila["_row"]

    print(f"\n{'='*60}")
    print(f"  COPIAR PCM — servicio madre {code}  [{sup}/{stype}/{loc or '—'}]  "
          f"codigo_nuevo={codigo_nuevo}")
    print(f"{'='*60}")

    if not code:
        escribir_resultado(row, "ERROR", error="SERVICE CODE vacío")
        return
    if not codigo_nuevo:
        escribir_resultado(row, "ERROR", error="CODIGO NUEVO vacío")
        return

    marcar_procesando(row)
    handle_prod = driver.current_window_handle
    nombre_pcm   = ""   # se completan a medida que se conocen, para no perder
    nombre_nuevo = ""   # el rastro si algo falla después de copiar/renombrar

    try:
        buscar_producto(driver, loc, sup, code, service_type=stype,
                        handle_origen=handle_prod)

        ir_a_used_in(driver)
        pcm_list = leer_pcm_list_package_header(driver)

        if not pcm_list:
            razon = "SIN PCM Package Header en USED IN"
            print(f"    ⚠ {code}: {razon}")
            escribir_resultado(row, "ERROR", error=razon)
            cerrar_pcm(driver, handle_prod)
            return

        if len(pcm_list) > 1:
            print(f"    ⚠ {code}: {len(pcm_list)} PCMs Package Header encontrados "
                  f"→ procesando el primero: {pcm_list[0]['nombre']}")

        pcm_info   = pcm_list[0]
        nombre_pcm = pcm_info["nombre"]
        print(f"    PCM encontrado: {nombre_pcm}")

        handle_pcm, _ = abrir_pcm(driver, pcm_info, start_scroll=0)

        # Chequeo de tarifas vencidas al abrir, por las dudas
        manejar_tarifas_vencidas(driver, nombre_pcm)

        nombre_nuevo, tiene_vencidos = copiar_pcm(driver, nombre_pcm, codigo_nuevo)

        # El PCM copiado se linkea al servicio madre ORIGINAL (buscado por
        # su SERVICE CODE, `code`) -- no por CODIGO NUEVO. CODIGO NUEVO solo
        # identifica al PCM copiado en sí (va en su nombre, ver
        # calcular_nombre_pcm_nuevo); no es un product code buscable en
        # Tourplan, por eso buscarlo en PRODUCT SEARCH siempre fallaba.
        linkear_pcm_a_servicio(driver, loc, sup, code)

        cerrar_pcm(driver, handle_prod)

        estado = "OK" if not tiene_vencidos else "OK"
        pcm_copiado_out = nombre_nuevo + (" [COMPONENTES VENCIDOS]" if tiene_vencidos else "")
        escribir_resultado(row, estado,
                            pcm_original=nombre_pcm,
                            pcm_copiado=pcm_copiado_out)
        print(f"\n✅ {code}: PCM copiado OK → {pcm_copiado_out}")

    except Exception as e:
        print(f"  ❌ Error procesando {code}: {e}")
        ss(driver, f"error_{code[:12]}")
        # Si ya se conocía el PCM original y/o ya se había copiado/renombrado
        # antes de que falle un paso posterior (ej. el linkeo), se deja ese
        # rastro en el Excel -- sin esto, un PCM ya creado en Tourplan podía
        # quedar sin registrar, dificultando la revisión manual.
        escribir_resultado(row, "ERROR", error=str(e),
                            pcm_original=nombre_pcm, pcm_copiado=nombre_nuevo)
        try:
            cerrar_pcm(driver, handle_prod)
        except Exception:
            pass

# ── MAIN ──────────────────────────────────────────────────────
print("="*60)
print(f"  📌 TOURPLAN NX — COPIAR PCM DESDE SERVICIO MADRE  v{VERSION}")
print(f"  📌 {VERSION_FECHA}")
print("="*60)

if not SHEET_URL:
    raise ValueError("No se indicó la URL del Google Sheet (TOURPLAN_SHEET_URL).")

conectar_sheet()
print(f"📄 Sheet: {SHEET_URL}")

pendientes = leer_pendientes()
print(f"📊 Filas PENDIENTE: {len(pendientes)}")
for f in pendientes:
    print(f"   · fila {f.get('_row')}: location={f.get('location')!r} "
          f"supplier={f.get('supplier')!r} service_type={f.get('service_type')!r} "
          f"service_code={f.get('service_code')!r} codigo_nuevo={f.get('codigo_nuevo')!r}")

if not pendientes:
    print("\n" + "="*60)
    print("⛔ EL SHEET NO TIENE FILAS PENDIENTE.")
    print("   Poné ESTADO=PENDIENTE en las filas a procesar y volvé a correr.")
    print("="*60)
    raise SystemExit("Sin filas PENDIENTE.")

driver = crear_driver()
_abortado = False
try:
    login(driver)
    for fila in pendientes:
        chequear_abort()
        try:
            procesar_fila(driver, fila)
        except Exception as e:
            print(f"\n❌ ERROR fila {fila.get('_row')}: {e}")
            ss(driver, f"error_row{fila.get('_row','x')}")
            escribir_resultado(fila["_row"], "ERROR", error=str(e))
            try:
                cerrar_pcm(driver, driver.window_handles[0])
            except Exception:
                pass
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
