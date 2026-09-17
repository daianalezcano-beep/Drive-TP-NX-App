# ============================================================
# TOURPLAN NX — UPDATE OPTION DESCRIPTION  v1.1
# Copia adaptada para tp-nx-app (app local). El original para
# Google Colab vive sin cambios en el repo Copy-products.
# Mismos cambios que el resto de los scripts vendorizados: config
# por variables de entorno, Chrome delegado a
# common/chrome_bootstrap.py, sin --headless/puerto fijo, perfil
# de Chrome aislado, log en carpeta temporal del OS, fix del
# import de importlib.util. El resto (selectores, login, lógica
# de negocio) es igual.
# ------------------------------------------------------------
# Modifica los campos "Description" y/o "Comment" (Option
# Description / Option Comment) de un option existente en
# Product Setup, procesando un Excel como cola de trabajo (una
# fila por option a modificar).
#
# NEW_DESCRIPTION y NEW_COMMENT son independientes entre sí: si
# alguna de las dos viene vacía en una fila, ESE campo se deja
# tal cual está en Tourplan — nunca se borra ni se pisa con
# vacío. Si ambas vienen vacías, la fila se marca OK sin hacer
# nada (no hay ningún campo para actualizar).
#
# Columnas del Excel:
#   LOCATION          — (opcional) filtro de location
#   SUPPLIER          — código del supplier que tiene el option
#   PRODUCT_CODE      — código del option a buscar
#   NEW_DESCRIPTION   — (opcional) nuevo texto para "Option Description"
#   NEW_COMMENT       — (opcional) nuevo texto para "Option Comment"
#   SERVICE_TYPE      — (opcional) código de tipo de servicio, ej: EX
#   ESTADO            — PENDIENTE → OK / ERROR: detalle
#   OBSERVACIONES     — (opcional) contexto adicional del resultado
# ============================================================

VERSION       = "1.2"
VERSION_FECHA = "2026-09-01"

# ── CAMBIOS ──
# v1.0: primera versión (solo Option Description).
# v1.1: agregado soporte para Option Comment (columna NEW_COMMENT,
#       independiente de NEW_DESCRIPTION — cualquiera de las dos puede
#       venir vacía sin afectar a la otra).
# v1.2: Fix en select_from_dropdown() (Supplier del panel de búsqueda):
#       tomaba siempre el PRIMER ítem del dropdown (items[0]) sin
#       verificar que coincidiera con el valor buscado — el dropdown
#       puede devolver más de un ítem parecido, sin garantía de que el
#       buscado quede primero (mismo tipo de bug ya confirmado con
#       Location en buscar_producto(), ver notas_srv.py v1.13/v1.14).
#       Ahora busca el ítem que contiene el valor como PALABRA COMPLETA;
#       si ninguno coincide, no clickea nada a ciegas.

# ── CONFIGURACIÓN — via variables de entorno (con default = valor original) ──
import os

from common.user_config import (
    CREDENTIALS_PATH as _CREDENTIALS_PATH_DEFAULT,
    TOKEN_PATH as _TOKEN_PATH_DEFAULT,
)

MODO             = os.environ.get("TOURPLAN_MODO", "lectura")     # "lectura" (dry-run) | cualquier otro valor = aplicar
SHEET_URL        = os.environ.get("TOURPLAN_SHEET_URL", "")
HOJA             = os.environ.get("TOURPLAN_HOJA", "PRODUCTOS")
CREDENTIALS_PATH = os.environ.get("TOURPLAN_CREDENTIALS_PATH", _CREDENTIALS_PATH_DEFAULT)
TOKEN_PATH       = os.environ.get("TOURPLAN_TOKEN_PATH", _TOKEN_PATH_DEFAULT)

USERNAME   = os.environ.get("TOURPLAN_USERNAME", "poner minusculas")
PASSWORD   = os.environ.get("TOURPLAN_PASSWORD", "password")
BASE_URL   = os.environ.get("TOURPLAN_BASE_URL", "https://tourplannx.eurotur.com.ar/TourplanNX_Test")

# Resguardo: si de todos modos el login terminara en una ruta distinta a
# BASE_URL (ej. un ambiente que redirige internamente), toda navegación
# posterior al login usa APP_BASE_URL, detectada de la URL real ya
# logueado (ver login()) en vez de BASE_URL a secas. No editar a mano —
# se sobreescribe en cada corrida.
APP_BASE_URL = BASE_URL

SS_DIR     = os.environ.get("TOURPLAN_SS_DIR", "/content/screenshots")
HEADLESS   = os.environ.get("TOURPLAN_HEADLESS", "0").strip() in ("1", "true", "True")

# Multiplicador de tiempos de espera. Producción por default —
# Tourplan responde más lento ahí que en Test.
VELOCIDAD = 1.5

# ── PASO 0: Entorno ──────────────────────────────────────────
import sys, subprocess, importlib.util, shutil, re, time, traceback
from datetime import datetime

_t_inicio = time.time()

print("🔧 Verificando entorno...\n")

# 0.1 Paquetes Python
_PIPS_NEEDED = {
    "selenium":            "selenium",
    "webdriver_manager":   "webdriver-manager",
    "gspread":             "gspread",
    "google_auth_oauthlib": "google-auth-oauthlib",
}
_faltantes = [pkg for mod, pkg in _PIPS_NEEDED.items()
              if importlib.util.find_spec(mod) is None]
if _faltantes:
    print(f"  ⏳ pip install {' '.join(_faltantes)} ...")
    subprocess.run([sys.executable, "-m", "pip", "install", "-q"] + _faltantes, check=True)
    print("  ✅ Paquetes OK")
else:
    print("  ✅ Paquetes ya instalados")

# 0.2 Google Chrome (multiplataforma — ver common/chrome_bootstrap.py)
from common.chrome_bootstrap import find_or_prepare_chrome
from common.abort import chequear_abort, AbortadoPorUsuario, ABORT_EXIT_CODE
# 0.4 Google Sheets como cola de trabajo (ver common/sheets_client.py)
from common.sheets_client import conectar_sheets, cargar_sheet, actualizar_fila_sheet

CHROMIUM_BIN, ver_chrome = find_or_prepare_chrome()

# 0.3 Imports
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

os.makedirs(SS_DIR, exist_ok=True)

WAIT_SHORT  = 5
WAIT_MEDIUM = 15
WAIT_LONG   = 30

SEL = {
    # Búsqueda de options — mismos selectores confirmados/probados en
    # producción por copy_products.py / rename_products.py.
    "SEARCH_BTN"       : "#searchWrapper li:nth-of-type(2) button",
    "FIELD_LOCATION"   : "div.parameters1 li:nth-of-type(1) input",
    "FIELD_SUPPLIER"   : "div.parameters1 li:nth-of-type(2) input",
    "FIELD_CODE"       : "#option tp-validator input",
    "DROPDOWN_ITEMS"   : "ul.dropdown-menu li",
    "TAB_RESULTS"      : "#tptablabel-productSearchResults",
    "RESULTS_ROWS"     : "td.tpcol-optioncode",
    "EXIT_PRODUCT_BTN" : "tp-button.cancelproduct > button",
    # Edición de Option Description / Option Comment dentro del option
    # ya abierto (ambos en el mismo tab de Product Details).
    "DESCRIPTION_INPUT"          : "#optionDescription input",
    "DESCRIPTION_INPUT_FALLBACK" : "tp-group.tpgroup-productdetailsgroup ul:nth-of-type(1) > li:nth-of-type(1) input",
    "COMMENT_INPUT"              : "#optionComment input",
    "COMMENT_INPUT_FALLBACK"     : "#tabs-product ul:nth-of-type(1) > li:nth-of-type(3) input",
    "SAVE_BTN"         : "tp-button.save > button",
}

COL_LOCATION        = "LOCATION"
COL_SUPPLIER        = "SUPPLIER"
COL_PRODUCT_CODE    = "PRODUCT_CODE"
COL_NEW_DESCRIPTION = "NEW_DESCRIPTION"
COL_NEW_COMMENT     = "NEW_COMMENT"
COL_SERVICE_TYPE    = "SERVICE_TYPE"
COL_ESTADO          = "ESTADO"
COL_OBSERVACIONES   = "OBSERVACIONES"

SAVE_POLL_INTERVAL = 0.4
SAVE_POLL_TIMEOUT  = 15  # seg — con polling ~400ms, no un solo intento

# Mapping sigla → prefijo numérico del sidebar de Service Type (confirmado
# por la usuaria, ver skill buscando-productos-en-tourplan/references/
# service-types.md). Las categorías "reales" (no agrupadas bajo "Z*NO
# USAR*") se muestran en el panel como "<número> -<Descripción>" (ej.
# "12 -Miscelaneos") — la sigla cruda ('MS') NUNCA aparece como texto ahí,
# así que buscarla sin traducir nunca matchea. TA y PR no tienen número
# confirmado: si aparecen, select_service_type() cae al fallback por
# sigla cruda — no asumir un número por patrón/orden alfabético.
STYPE_SIDEBAR = {
    "HT": "01", "HX": "02", "TF": "03", "EX": "04", "ML": "05",
    "RT": "06", "CR": "07", "FT": "08", "OC": "09", "LN": "10",
    "LP": "11", "MS": "12",
}

# Siglas que Tourplan agrupa en el sidebar de Service Type bajo el prefijo
# "Z*NO USAR*" (organización interna de la empresa, confirmado por la
# usuaria 2026-08-18 — no significa que estén deprecadas). Estas SÍ muestran
# la sigla cruda como texto (ej. "Z*NO USAR* - GU - GUIDE SERVICES"), por
# eso no están en STYPE_SIDEBAR. select_service_type() excluye estos ítems
# cuando la sigla buscada NO es una de estas 12, para no matchear por
# accidente una categoría real dentro de un ítem "no usar" que comparte una
# palabra (ej. TF="Transfer" vs TN="Transfer Non-Accom").
Z_NO_USAR_SIGLAS = {"GA", "TK", "TN", "CH", "CM", "CO", "EN", "GU", "PJ", "ST", "TR", "TI"}

# ── Helpers de DOM ────────────────────────────────────────────
_ss_n = [0]

def ss(driver, nombre):
    _ss_n[0] += 1
    nombre_seguro = re.sub(r"[^A-Za-z0-9_-]+", "_", nombre)[:60]
    p = f"{SS_DIR}/{_ss_n[0]:03d}_{nombre_seguro}_{int(time.time())}.png"
    try:
        ok = driver.save_screenshot(p)
        if not ok:
            print(f"  ⚠️ No se pudo guardar captura: {os.path.basename(p)}")
            return None
        print(f"  📸 {os.path.basename(p)}")
        return p
    except Exception as e:
        print(f"  ⚠️ Error guardando captura ({e})")
        return None

def dump(driver, nombre):
    p = f"{SS_DIR}/{nombre}_{int(time.time())}.html"
    try:
        with open(p, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print(f"  💾 HTML: {p}")
    except Exception as e:
        print(f"  ⚠️ Error volcando HTML ({e})")

def jc(driver, el):
    driver.execute_script("arguments[0].click();", el)

def wait(driver, css, t=WAIT_MEDIUM):
    return WebDriverWait(driver, t).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, css)))

def wait_click(driver, css, t=WAIT_MEDIUM):
    return WebDriverWait(driver, t).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, css)))

def esperar_fin_carga(driver, timeout=15):
    """Tourplan muestra un <dialog> nativo con 'PLEASE WAIT...' mientras
    termina de procesar algo en curso (ej. terminar de guardar/navegar
    tras la fila anterior). Confirmado en corrida real: si se clickea
    algo mientras ese dialog sigue abierto, Selenium tira
    ElementClickIntercepted contra el propio <dialog> en vez de llegar
    al botón real. Poll corto en vez de un sleep fijo — no todas las
    esperas duran lo mismo."""
    fin = time.time() + timeout * VELOCIDAD
    while time.time() < fin:
        if not driver.find_elements(By.CSS_SELECTOR, "dialog[open]"):
            return
        time.sleep(0.3)
    print("    ⚠ El dialog de carga ('PLEASE WAIT...') seguía abierto tras esperar")

def set_val(driver, el, value):
    """Versión corta — inputs de filtro/búsqueda (no confirman al perder foco)."""
    driver.execute_script("""
        var inp = arguments[0], val = arguments[1];
        var setter = Object.getOwnPropertyDescriptor(
            window.HTMLInputElement.prototype, 'value').set;
        setter.call(inp, val);
        inp.dispatchEvent(new Event('input',  {bubbles:true}));
        inp.dispatchEvent(new Event('change', {bubbles:true}));
    """, el, value)

def set_val_con_blur(driver, el, value):
    """Versión completa — foco → setter nativo → eventos → blur. Necesaria
    para que Angular confirme el valor de Option Description (ver skill
    escribiendo-en-formularios-angular-de-tourplan)."""
    driver.execute_script("""
        var inp = arguments[0], val = arguments[1];
        inp.focus();
        inp.dispatchEvent(new Event('focus', {bubbles:true}));
        var setter = Object.getOwnPropertyDescriptor(
            window.HTMLInputElement.prototype, 'value').set;
        setter.call(inp, val);
        inp.dispatchEvent(new Event('input',  {bubbles:true}));
        inp.dispatchEvent(new Event('change', {bubbles:true}));
        inp.dispatchEvent(new KeyboardEvent('keyup', {bubbles:true}));
        inp.blur();
        inp.dispatchEvent(new Event('blur',     {bubbles:true}));
        inp.dispatchEvent(new Event('focusout', {bubbles:true}));
    """, el, value)

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
    opts.add_argument("--window-size=1704,1012")
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
        raise RuntimeError("No se encontró Chrome funcional.")
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

# ── Sesión: login / logout ────────────────────────────────────
# REGLA DE NEGOCIO (licencias/sesiones concurrentes limitadas en Tourplan):
# el flujo principal SIEMPRE va envuelto en try/finally con logout(driver)
# + driver.quit() — ver skill arrancando-un-script-de-tourplan.

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
    for _ in range(15):
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
    time.sleep(0.5 * VELOCIDAD)

    clic = driver.execute_script("""
        function vis(e){return !!(e && (e.offsetWidth||e.offsetHeight
                                  ||e.getClientRects().length) && !e.disabled);}
        var b = Array.from(document.querySelectorAll(
            "button.login, button[type='submit'], button")).filter(vis)
            .find(function(x){return /log\\s*in|ingresar|entrar|sign\\s*in/i
                                     .test((x.innerText||'')) ||
                                     x.classList.contains('login');});
        if (b){ b.click(); return (b.innerText||'button').trim(); }
        return null;
    """)
    if not clic:
        try: p_el.send_keys(Keys.ENTER)
        except Exception: pass

    time.sleep(8 * VELOCIDAD)
    if "login" in driver.current_url.lower():
        # No asumir "usuario/password mal" a ciegas — Tourplan tiene
        # licencias/sesiones concurrentes limitadas (ver skill
        # arrancando-un-script-de-tourplan/references/licencias-y-sesiones.md):
        # una sesión colgada de una corrida anterior puede bloquear el login
        # nuevo con un mensaje que este flujo no reconoce. Volcar ese mensaje
        # en vez de fallar en blanco.
        mensaje_pantalla = driver.execute_script("""
            function vis(e){ return !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length); }
            var candidatos = Array.from(document.querySelectorAll(
                '.error, .alert, [class*="error"], [class*="alert"], [class*="message"]'))
                .filter(vis).map(function(e){ return (e.innerText || '').trim(); })
                .filter(Boolean);
            return {
                candidatos: candidatos.slice(0, 5),
                textoVisible: (document.body.innerText || '').replace(/\\s+/g, ' ').trim().slice(0, 300),
            };
        """)
        ss(driver, "login_fallido")
        dump(driver, "login_fallido")
        raise AssertionError(
            "Login falló — sigue en #/login tras enviar usuario/password. "
            f"Mensajes en pantalla: {mensaje_pantalla.get('candidatos')}. "
            f"Texto visible: {mensaje_pantalla.get('textoVisible')!r}. "
            "Si corriste el script varias veces seguidas, puede ser una sesión "
            "colgada ocupando la licencia — esperá unos minutos y reintentá. "
            "Revisá screenshots/login_fallido_*.png y el .html volcado."
        )

    # Resguardo: usar la base real donde quedó la sesión ya logueada (en
    # el caso normal, coincide con BASE_URL) para toda navegación
    # posterior — si por algún motivo no coincidieran, forzar BASE_URL a
    # secas haría que el router de Angular vuelva silenciosamente a
    # #/home en vez de navegar a #/product.
    global APP_BASE_URL
    APP_BASE_URL = driver.current_url.split("#")[0].rstrip("/")
    if APP_BASE_URL != BASE_URL:
        print(f"    Entorno real detectado tras login: {APP_BASE_URL}")

    ss(driver, "post_login")
    print("✅ Login OK")

def logout(driver):
    """Cierra ventanas secundarias y hace logout real. El ícono de usuario
    es contextual (dentro de Product Setup puede decir "Close Setup" en
    vez de "Log Out") — por eso siempre se fuerza navegación a #/home
    antes de intentar el logout (ver skill arrancando-un-script-de-tourplan
    / references/licencias-y-sesiones.md)."""
    print("\n🔒 Finalización: cerrando ventanas y haciendo logout...")
    try:
        principal = driver.window_handles[0]
        for h in driver.window_handles[1:]:
            try:
                driver.switch_to.window(h); driver.close()
            except Exception:
                pass
        driver.switch_to.window(principal)
    except Exception:
        pass

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
        driver.get(f"{APP_BASE_URL}/#/home")
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
            texto_btn = (btn_logout.text or "Logout").strip()
            jc(driver, btn_logout)
            clicked = texto_btn or "Logout"
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
            print(f"  🔓 Logout OK (click en '{clicked}', volvió al login)")
        elif clicked:
            print(f"  ⚠ Click en '{clicked}' pero NO volvió al login — logout no confirmado")
        else:
            print("  ⚠ No encontré la opción LOG OUT en el menú — logout no realizado")
    except Exception as e:
        print(f"  ⚠ Error en logout: {e}")

# ── Búsqueda de producto ──────────────────────────────────────
# open_product_setup() / select_from_dropdown() / select_service_type() /
# search_options() son el mismo proceso que copy_products.py líneas
# 292-398 (probado en producción) — únicas diferencias no funcionales:
# el multiplicador VELOCIDAD en los sleeps y un guard por si SERVICE_TYPE
# viene vacío (acá es opcional; en copy_products.py es obligatorio y
# nunca llega vacío a select_service_type()). A partir de acá (equivalente
# a la línea 400 de copy_products.py) la lógica de negocio cambia: en vez
# de copiar el option, se abre y se edita su Option Description.

def open_product_setup(driver):
    """Navega a Product Setup una sola vez, al principio de la corrida.
    Tourplan a veces abre Product Setup en una pestaña/ventana nueva — si
    no se detecta y se cambia el foco de Selenium a esa ventana, todo lo
    que sigue (lupa, filtros, resultados) se busca en la ventana vieja y
    nunca aparece."""
    print("Navegando a Product Setup...")
    handles_before = set(driver.window_handles)

    driver.get(f"{APP_BASE_URL}/#/home")
    time.sleep(2 * VELOCIDAD)
    driver.get(f"{APP_BASE_URL}/#/product")
    time.sleep(5 * VELOCIDAD)

    handles_after = set(driver.window_handles)
    new_handles = handles_after - handles_before
    if new_handles:
        driver.switch_to.window(new_handles.pop())
        print("  Nueva pestaña de Product Setup detectada — foco cambiado")

    # Hasta acá el proceso es idéntico a copy_products.py líneas 292-398
    # (confirmado en corrida real). Si #searchWrapper igual no aparece,
    # ya no es una diferencia de código — volcar diagnóstico en vez de
    # dejar propagar un TimeoutException sin contexto.
    try:
        wait(driver, "#searchWrapper", t=WAIT_LONG)
    except TimeoutException:
        diagnostico = driver.execute_script("""
            function vis(e){ return !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length); }
            return {
                url: window.location.href,
                title: document.title,
                texto: (document.body.innerText || '').replace(/\\s+/g, ' ').trim().slice(0, 400),
                idsVisibles: Array.from(document.querySelectorAll('[id]'))
                    .filter(vis).map(function(e){ return e.id; }).slice(0, 50),
                dialogsAbiertos: document.querySelectorAll('tp-dialog, .modal, [class*="dialog"]').length,
            };
        """)
        print(f"    ⚠ #searchWrapper no apareció tras {WAIT_LONG}s. Diagnóstico: {diagnostico}")
        ss(driver, "product_setup_timeout")
        dump(driver, "product_setup_timeout")
        raise RuntimeError(
            "Product Setup no cargó: #searchWrapper no apareció tras el mismo proceso "
            "que copy_products.py. "
            f"URL={diagnostico.get('url')!r} título={diagnostico.get('title')!r} "
            f"texto visible={diagnostico.get('texto')!r} "
            f"ids visibles={diagnostico.get('idsVisibles')} "
            f"dialogs abiertos={diagnostico.get('dialogsAbiertos')}. "
            "Revisá screenshots/product_setup_timeout_*.png y el .html volcado en la misma carpeta."
        ) from None

    print("✅ Product Setup listo")

def select_from_dropdown(driver, input_el, value):
    """Busca, entre los ítems visibles del dropdown, el que contiene
    `value` como PALABRA COMPLETA (delimitada por inicio/fin de texto o
    cualquier caracter no alfanumérico) en vez de asumir que el primer
    ítem (`items[0]`) es siempre el correcto — el dropdown puede devolver
    más de un ítem con texto parecido y no hay garantía de que el buscado
    quede primero (mismo tipo de bug ya confirmado con Location en
    buscar_producto(), ver notas_srv.py v1.13/v1.14). Mismo criterio de
    palabra completa ya usado en select_service_type() de
    flag_as_deleted.py para el mismo tipo de problema."""
    input_el.click()
    time.sleep(0.3 * VELOCIDAD)
    set_val(driver, input_el, value)
    time.sleep(1.2 * VELOCIDAD)
    try:
        items = WebDriverWait(driver, WAIT_SHORT).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, SEL["DROPDOWN_ITEMS"])))
        if items:
            valor_upper = value.strip().upper()
            rx = re.compile(r"(^|[^A-Z0-9])" + re.escape(valor_upper) + r"($|[^A-Z0-9])")
            exacto = next((it for it in items if rx.search(it.text.strip().upper())), None)
            if exacto is not None:
                exacto.click()
                time.sleep(0.5 * VELOCIDAD)
                return True
            print(f"    ⚠ '{value}': ningún ítem del dropdown coincide como palabra "
                  f"completa — no se clickeó ninguno a ciegas")
    except TimeoutException:
        pass
    try:
        input_el.send_keys(Keys.RETURN)
    except Exception:
        pass
    time.sleep(0.5 * VELOCIDAD)
    return False

def select_service_type(driver, service_type_code):
    """Guard inicial: un código vacío matchea cualquier texto por substring
    y terminaría clickeando el primer ítem de la lista (no necesariamente
    'All Services'). En copy_products.py SERVICE_TYPE es obligatorio y
    nunca llega vacío acá; en este script es opcional, así que si no vino
    no se toca el panel.

    Escanea TODOS los <li> visibles de la página (igual que
    copy_products.py, ver comentario largo ahí) en vez de restringirse a un
    selector con un tag específico ("tp-list-selector ul li") que nunca se
    confirmó contra Tourplan real — en una corrida real de copy_products.py
    (2026-08) ese selector no encontró NINGÚN <li> en 98 filas seguidas.
    Como acá SERVICE_TYPE es solo un filtro opcional (PRODUCT_CODE ya
    identifica el option exacto), que no se pueda seleccionar NO aborta la
    fila — sólo se avisa y se sigue sin ese filtro.

    Traduce la sigla al prefijo numérico confirmado (STYPE_SIDEBAR) antes de
    buscar — las categorías reales se muestran como "<número> -<Descripción>"
    (ej. "12 -Miscelaneos"), la sigla cruda ('MS') no aparece ahí como texto.
    Si la sigla no tiene número confirmado, busca la sigla cruda (fallback
    para TA/PR, o para cualquier categoría "Z*NO USAR*", que sí muestra la
    sigla como texto — ver Z_NO_USAR_SIGLAS).

    El match es por PALABRA COMPLETA (boundary regex, nunca substring
    suelto — evita el bug confirmado en Flag as Deleted donde 'TR' tildaba
    "Hotel Extras" porque "Extras" contiene "tr"), y descarta los ítems
    "Z*NO USAR*" cuando la sigla buscada no es una de las 12 agrupadas ahí
    (Z_NO_USAR_SIGLAS), para no matchear una categoría real dentro de un
    ítem "no usar" que comparte una palabra."""
    if not service_type_code:
        return
    code_upper = service_type_code.upper()
    excluir_no_usar = code_upper not in Z_NO_USAR_SIGLAS
    st_buscar = STYPE_SIDEBAR.get(code_upper, code_upper)
    patron_js = r"(^|[\s\-–—/(])" + re.escape(st_buscar) + r"([\s\-–—/)]|$)"

    resultado = None
    for intento in range(3):
        resultado = driver.execute_script("""
            var rx = new RegExp(arguments[0]);
            var excluirNoUsar = arguments[1];
            var textos = [], elegido = null;
            for (var li of document.querySelectorAll('li')) {
                if (!li.offsetParent) continue;
                var t = (li.textContent || '').trim();
                if (!t || t.length > 80) continue;
                textos.push(t);
                if (elegido) continue;
                if (excluirNoUsar && t.toLowerCase().indexOf('no usar') !== -1) continue;
                if (rx.test(t.toUpperCase())) elegido = li;
            }
            if (elegido) {
                elegido.click();
                return {ok: true, texto: (elegido.textContent || '').trim().slice(0, 60)};
            }
            return {ok: false, textos: textos.slice(0, 60)};
        """, patron_js, excluir_no_usar)
        if resultado.get("ok"):
            print(f"    Service Type: {resultado.get('texto')}")
            time.sleep(0.8 * VELOCIDAD)
            return
        time.sleep(1.5 * VELOCIDAD)

    print(f"    ⚠ Service Type '{service_type_code}' no encontrado")

def search_options(driver, supplier, service_type, location="", option_code=""):
    """Busca options por supplier (+ location/service_type/code si se
    dan) y devuelve la lista de filas de resultado, cada una con su
    'code' y el elemento 'row_el' para poder clickearla después."""
    print(f"\n  🔍 Buscando: supplier={supplier} st={service_type or '-'} "
          f"loc={location or '-'} code={option_code or 'todos'}")

    # Re-navegar para limpiar el estado de una búsqueda anterior.
    driver.get(f"{APP_BASE_URL}/#/home")
    time.sleep(2 * VELOCIDAD)
    driver.get(f"{APP_BASE_URL}/#/product")
    time.sleep(5 * VELOCIDAD)

    # Confirmado en corrida real: al procesar la 2ª fila en adelante,
    # Tourplan a veces todavía muestra el dialog "PLEASE WAIT..." (sigue
    # terminando de procesar la fila anterior) cuando llegamos acá — sin
    # esperar a que cierre, el click de la lupa rebota contra ese dialog.
    esperar_fin_carga(driver)

    wait_click(driver, SEL["SEARCH_BTN"]).click()
    time.sleep(1 * VELOCIDAD)

    if location:
        loc_inp = wait(driver, SEL["FIELD_LOCATION"])
        set_val(driver, loc_inp, location)
        loc_inp.send_keys(Keys.RETURN)
        time.sleep(0.8 * VELOCIDAD)

    sup_inp = wait(driver, SEL["FIELD_SUPPLIER"])
    selected = select_from_dropdown(driver, sup_inp, supplier)
    if not selected:
        print(f"    ⚠ Supplier '{supplier}' — sin dropdown, se usó Enter")
    time.sleep(0.5 * VELOCIDAD)

    if option_code:
        code_inp = wait(driver, SEL["FIELD_CODE"])
        set_val(driver, code_inp, option_code)
        code_inp.send_keys(Keys.RETURN)
        time.sleep(0.5 * VELOCIDAD)

    select_service_type(driver, service_type)

    tab = wait_click(driver, SEL["TAB_RESULTS"])
    jc(driver, tab)
    time.sleep(2 * VELOCIDAD)

    try:
        row_els = WebDriverWait(driver, WAIT_MEDIUM).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, SEL["RESULTS_ROWS"])))
    except TimeoutException:
        print("    Sin resultados")
        return []

    options = [{"code": el.text.strip(), "row_el": el} for el in row_els]
    print(f"    ✓ {len(options)} option(s): {[o['code'] for o in options]}")
    return options

def exit_to_results(driver):
    """Cierra el option abierto y vuelve a la grilla de resultados —
    siempre llamarla al terminar de procesar un option, haya salido bien
    o mal, para no arrastrar estado a la próxima fila."""
    try:
        jc(driver, wait_click(driver, SEL["EXIT_PRODUCT_BTN"]))
        time.sleep(1.5 * VELOCIDAD)
    except Exception:
        print("    ⚠ No se pudo hacer Exit Product")

def abrir_option(driver, opt):
    """Clickea la fila de resultado para abrir el option (Product
    Details, Rates, Utilities, etc. quedan disponibles) — igual que
    copy_products.py/rename_products.py: clickear la fila y seguir, sin
    verificación extra vía el menú hamburguesa (no hace falta abrirlo
    para nada acá, y abrirlo/cerrarlo podía dejar un `.tpnavbackdrop`
    colgado tapando la pantalla)."""
    jc(driver, opt["row_el"])
    time.sleep(2 * VELOCIDAD)
    ss(driver, f"option_abierto_{opt['code'][:10]}")

# ── Option Description / Option Comment ─────────────────────────
# Ambos campos viven en el mismo formulario (tab Product Details) y se
# manejan con el mismo patrón — foco → setter nativo → eventos → blur →
# click en el label — solo cambia el selector del input. Se generaliza
# acá para no duplicar la lógica, con un wrapper por campo abajo para
# que el resto del script (process_row) siga llamando funciones con
# nombre explícito.

def _obtener_input(driver, sel_primario, sel_fallback):
    """Prioriza el id estable (#optionDescription / #optionComment) sobre
    el selector posicional del recording, que puede romperse si cambia
    el orden de los grupos del formulario."""
    try:
        return driver.find_element(By.CSS_SELECTOR, sel_primario)
    except NoSuchElementException:
        return wait(driver, sel_fallback)

def _leer_campo(driver, sel_primario, sel_fallback):
    return (_obtener_input(driver, sel_primario, sel_fallback).get_attribute("value") or "").strip()

def _click_label_campo(driver, sel_primario, sel_fallback):
    """Click en el <label> del mismo campo (recording manual) para forzar
    blur/confirmación antes de ir a buscar el SAVE."""
    return bool(driver.execute_script("""
        var inp = document.querySelector(arguments[0]) || document.querySelector(arguments[1]);
        if (!inp) return false;
        var li = inp.closest('li');
        var label = li ? li.querySelector('label') : null;
        if (!label) return false;
        label.click();
        return true;
    """, sel_primario, sel_fallback))

def _escribir_campo(driver, sel_primario, sel_fallback, valor, nombre_campo):
    """Escribe valor en el input con el patrón Angular completo (foco →
    setter nativo → eventos → blur) y clickea el label del campo, tal
    como el recording manual."""
    inp = _obtener_input(driver, sel_primario, sel_fallback)
    inp.click()
    time.sleep(0.3 * VELOCIDAD)
    set_val_con_blur(driver, inp, valor)
    time.sleep(0.5 * VELOCIDAD)
    if not _click_label_campo(driver, sel_primario, sel_fallback):
        print(f"    ⚠ No se pudo clickear el label de {nombre_campo} (no crítico)")
    time.sleep(0.5 * VELOCIDAD)

def leer_option_description(driver):
    return _leer_campo(driver, SEL["DESCRIPTION_INPUT"], SEL["DESCRIPTION_INPUT_FALLBACK"])

def actualizar_option_description(driver, nueva_descripcion):
    _escribir_campo(driver, SEL["DESCRIPTION_INPUT"], SEL["DESCRIPTION_INPUT_FALLBACK"],
                     nueva_descripcion, "Option Description")

def leer_option_comment(driver):
    return _leer_campo(driver, SEL["COMMENT_INPUT"], SEL["COMMENT_INPUT_FALLBACK"])

def actualizar_option_comment(driver, nuevo_comentario):
    _escribir_campo(driver, SEL["COMMENT_INPUT"], SEL["COMMENT_INPUT_FALLBACK"],
                     nuevo_comentario, "Option Comment")

def wait_save_habilitado(driver, save_selector, timeout=SAVE_POLL_TIMEOUT, interval=SAVE_POLL_INTERVAL):
    """Polling activo del botón SAVE — no un sleep fijo. Puede haber más
    de un tp-button.save en la vista; se toma el primero visible+enabled
    (ver skill escribiendo-en-formularios-angular-de-tourplan /
    references/boton-save-ambiguo.md)."""
    fin = time.time() + timeout * VELOCIDAD
    encontrado_alguna_vez = False
    while time.time() < fin:
        elems = driver.find_elements(By.CSS_SELECTOR, save_selector)
        if elems:
            encontrado_alguna_vez = True
            for el in elems:
                try:
                    if el.is_displayed() and el.is_enabled():
                        return el, None
                except Exception:
                    pass  # elemento stale por un re-render; se reintenta en la próxima vuelta
        time.sleep(interval)
    if not encontrado_alguna_vez:
        return None, "el botón SAVE no se encontró en la página"
    return None, "el botón SAVE nunca se habilitó (timeout)"

def dump_save_candidates(driver):
    return driver.execute_script("""
        var all = Array.from(document.querySelectorAll('button, tp-button'));
        var out = [];
        for (var i = 0; i < all.length && out.length < 30; i++) {
            var el = all[i];
            var cls = (el.className && el.className.toString) ? el.className.toString() : '';
            var txt = (el.innerText || el.textContent || '').trim().replace(/\\s+/g, ' ');
            var esCandidato = cls.toLowerCase().indexOf('save') !== -1 ||
                               txt.toLowerCase().indexOf('save') !== -1;
            if (!esCandidato) continue;
            var vis = !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);
            out.push(el.tagName.toLowerCase() + '.' + cls + '=' + JSON.stringify(txt.slice(0, 40)) +
                      ' visible=' + vis + ' disabled=' + !!el.disabled);
        }
        return out;
    """)

# ── Sheet como cola de trabajo ────────────────────────────────
def load_sheet():
    ws = conectar_sheets(SHEET_URL, HOJA, CREDENTIALS_PATH, TOKEN_PATH)
    rows, columnas = cargar_sheet(ws)
    print(f"Sheet cargado: {len(rows)} fila(s) en '{HOJA}'")
    return ws, rows, columnas

def update_row(ws, row_idx, columnas, estado=None, observaciones=None):
    """Escribe ESTADO/OBSERVACIONES de UNA fila con un batch_update — así
    una corrida cortada a mitad de camino deja registro de lo ya
    procesado."""
    valores = {}
    if estado is not None:
        valores[COL_ESTADO] = estado
    if observaciones is not None and COL_OBSERVACIONES in columnas:
        valores[COL_OBSERVACIONES] = observaciones
    actualizar_fila_sheet(ws, row_idx, columnas, valores)

# Campos editables — nombre visible, columna del Excel, lector y editor.
# Agregar un campo nuevo (ej. otro input del mismo tab) es sumar una
# entrada acá; process_row() no necesita tocarse.
CAMPOS_EDITABLES = [
    ("Description", COL_NEW_DESCRIPTION, leer_option_description, actualizar_option_description),
    ("Comment",     COL_NEW_COMMENT,     leer_option_comment,     actualizar_option_comment),
]

# ── Lógica de negocio por fila ────────────────────────────────
def process_row(driver, row):
    location     = str(row.get(COL_LOCATION) or "").strip()
    supplier     = str(row.get(COL_SUPPLIER) or "").strip()
    codigo       = str(row.get(COL_PRODUCT_CODE) or "").strip()
    service_type = str(row.get(COL_SERVICE_TYPE) or "").strip()

    # objetivos[nombre] = valor deseado, solo para los campos que vinieron
    # con algo en el Excel — un campo vacío significa "no tocar", nunca
    # "borrar lo que ya está en Tourplan".
    objetivos = {nombre: str(row.get(col) or "").strip()
                 for nombre, col, _lector, _editor in CAMPOS_EDITABLES}
    objetivos = {nombre: valor for nombre, valor in objetivos.items() if valor}

    if not supplier or not codigo:
        return "ERROR", "Faltan campos obligatorios (SUPPLIER, PRODUCT_CODE)"
    if not objetivos:
        return "OK", "NEW_DESCRIPTION y NEW_COMMENT vacíos — no había nada para actualizar"

    try:
        options = search_options(driver, supplier, service_type, location, codigo)
    except Exception as e:
        ss(driver, f"search_error_{codigo}")
        return "ERROR", f"Falla en búsqueda: {e}"

    if not options:
        return "ERROR", f"No se encontró ningún option con los filtros dados (supplier={supplier}, code={codigo})"

    opt = next((o for o in options if o["code"].upper() == codigo.upper()), None)
    if opt is None:
        dump(driver, f"sin_match_{codigo}")
        return "ERROR", f"'{codigo}' no está entre los resultados: {[o['code'] for o in options]}"

    try:
        abrir_option(driver, opt)
    except Exception as e:
        ss(driver, f"abrir_option_error_{codigo}")
        exit_to_results(driver)
        return "ERROR", f"Error abriendo el option: {e}"

    # Leer el estado actual SOLO de los campos con objetivo — no hace
    # falta tocar el campo que el Excel dejó vacío.
    actuales = {}
    try:
        for nombre, _col, lector, _editor in CAMPOS_EDITABLES:
            if nombre in objetivos:
                actuales[nombre] = lector(driver)
                print(f"    {nombre} actual: '{actuales[nombre]}'  →  nueva: '{objetivos[nombre]}'")
    except Exception as e:
        ss(driver, f"no_campo_{codigo}")
        exit_to_results(driver)
        return "ERROR", f"No se encontró algún campo a editar: {e}"

    if MODO == "lectura":
        exit_to_results(driver)
        detalle = " | ".join(f"{n}: actual='{actuales[n]}' → nueva='{v}'" for n, v in objetivos.items())
        return "OK", f"LECTURA: {detalle} (no se guardó)"

    # Solo hace falta editar (y guardar) los campos donde el objetivo
    # difiere del valor actual.
    a_editar = [(nombre, objetivos[nombre], editor)
                for nombre, _col, _lector, editor in CAMPOS_EDITABLES
                if nombre in objetivos and objetivos[nombre] != actuales[nombre]]

    if not a_editar:
        exit_to_results(driver)
        coinciden = " y ".join(f"{n} ya coincidía" for n in objetivos)
        return "OK", f"{coinciden} — no hizo falta editar"

    try:
        for nombre, valor, editor in a_editar:
            editor(driver, valor)
    except Exception as e:
        ss(driver, f"edit_error_{codigo}")
        exit_to_results(driver)
        return "ERROR", f"Error escribiendo campo(s): {e}"

    def _verificar(lectores_por_nombre):
        """Relee cada campo editado y compara contra su objetivo.
        Devuelve (todo_ok, detalle_legible)."""
        resultados, todo_ok = [], True
        for nombre, valor_objetivo, _editor in a_editar:
            final = lectores_por_nombre[nombre](driver)
            ok = (final == valor_objetivo)
            todo_ok = todo_ok and ok
            resultados.append(f"{nombre}: OK" if ok else f"{nombre}: quedó '{final}', esperaba '{valor_objetivo}'")
        return todo_ok, "; ".join(resultados)

    lectores_por_nombre = {nombre: lector for nombre, _col, lector, _editor in CAMPOS_EDITABLES}

    ss(driver, f"pre_save_{codigo}")
    btn, err = wait_save_habilitado(driver, SEL["SAVE_BTN"])

    if btn:
        try:
            jc(driver, btn)
        except Exception as e:
            exit_to_results(driver)
            return "ERROR", f"Error clickeando SAVE: {e}"
        time.sleep(2 * VELOCIDAD)
        ss(driver, f"post_save_{codigo}")
        todo_ok, detalle = _verificar(lectores_por_nombre)
        exit_to_results(driver)
        if todo_ok:
            return "OK", None
        return "ERROR", f"SAVE clickeado pero algún valor no coincide — {detalle}"

    # No apareció ningún SAVE habilitado — no asumir error automáticamente:
    # releer los campos editados y comparar contra el objetivo (ver skill
    # escribiendo-en-formularios-angular-de-tourplan / boton-save-ambiguo.md)
    ss(driver, f"sin_save_{codigo}")
    candidatos = dump_save_candidates(driver)
    todo_ok, detalle = _verificar(lectores_por_nombre)
    exit_to_results(driver)
    if todo_ok:
        return "OK", f"SAVE no se habilitó ({err}) pero los valores ya quedaron confirmados — {detalle}"
    return "ERROR", f"SAVE no disponible ({err}); {detalle}. Candidatos SAVE: {candidatos}"

# ── MAIN ──────────────────────────────────────────────────────
print("=" * 60)
print(f"  UPDATE OPTION DESCRIPTION  v{VERSION}  [{VERSION_FECHA}]  MODO={MODO}")
print("=" * 60)

if not SHEET_URL:
    raise ValueError("No se indicó la URL del Google Sheet (TOURPLAN_SHEET_URL).")

ws, rows, columnas = load_sheet()

_columnas_obligatorias = [COL_SUPPLIER, COL_PRODUCT_CODE, COL_NEW_DESCRIPTION, COL_NEW_COMMENT, COL_ESTADO]
_faltantes_col = [c for c in _columnas_obligatorias if c not in columnas]
if _faltantes_col:
    raise ValueError(
        f"Columna(s) {_faltantes_col} no encontrada(s) en el Sheet.\n"
        f"Columnas detectadas: {columnas}"
    )

pendientes = [r for r in rows
              if str(r.get(COL_ESTADO) or "").strip().upper() == "PENDIENTE"]

print(f"Filas PENDIENTE: {len(pendientes)}")
for f in pendientes:
    print(f"  · fila {f['__row_idx__']}: {f.get(COL_SUPPLIER)}/{f.get(COL_PRODUCT_CODE)}  "
          f"→ desc='{f.get(COL_NEW_DESCRIPTION) or '-'}' comment='{f.get(COL_NEW_COMMENT) or '-'}'")

if not pendientes:
    print("\n⛔ Sin filas PENDIENTE. Verificá que la columna ESTADO tenga 'PENDIENTE'.")
    raise SystemExit

driver = crear_driver()
_abortado = False
try:
    login(driver)
    open_product_setup(driver)

    for row in pendientes:
        chequear_abort()
        row_idx = row["__row_idx__"]
        print(f"\n{'─'*60}")
        print(f"Fila {row_idx}: {row.get(COL_SUPPLIER)}/{row.get(COL_PRODUCT_CODE)}")

        estado, observaciones = "ERROR", "Error desconocido"
        try:
            estado, observaciones = process_row(driver, row)
        except Exception:
            estado = "ERROR"
            observaciones = traceback.format_exc(limit=3)
            ss(driver, f"fatal_row{row_idx}")

        print(f"  Estado: {estado}" + (f" — {observaciones}" if observaciones else ""))
        update_row(ws, row_idx, columnas, estado=estado, observaciones=observaciones)

except AbortadoPorUsuario:
    _abortado = True
    print("\n⏸️  Corrida abortada por el usuario — las filas que no llegaron a "
          "procesarse quedan en PENDIENTE para retomar en otra corrida.")
finally:
    logout(driver)
    driver.quit()
    dur = int(time.time() - _t_inicio)
    m, s = divmod(dur, 60)
    print(f"\n🏁 Fin. Duración: {m}m {s:02d}s")
    print(f"📄 Sheet: {SHEET_URL}")
    print(f"📂 Screenshots: {SS_DIR}")

if _abortado:
    sys.exit(ABORT_EXIT_CODE)
