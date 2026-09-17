# ============================================================
# TOURPLAN NX — NEW OPTION (Copy Product)  v1.4
# Copia adaptada para tp-nx-app (app local). El original para
# Google Colab vive sin cambios en el repo Copy-products.
# Mismos cambios que el resto de los scripts vendorizados: config
# por variables de entorno, Chrome delegado a
# common/chrome_bootstrap.py, sin --headless/puerto fijo, perfil
# de Chrome aislado, log en carpeta temporal del OS, fix del
# import de importlib.util. El resto (selectores, login, lógica
# de negocio) es igual.
# ------------------------------------------------------------
# Copia un option existente y lo guarda como una opción nueva,
# permitiendo además cambiar en el mismo paso: location, service
# type, supplier, description y comment de la copia.
#
# Todos los campos "destino" (salvo OPTION_CODE_DESTINO) son
# opcionales: si la celda del Excel está vacía, ESE CAMPO NO SE
# TOCA en el diálogo Copy Product y la opción nueva conserva el
# valor que Tourplan precarga al abrir el diálogo (que es el del
# option de origen).
#
# Columnas del Excel — ubicación del option de origen:
#   SUPPLIER_ORIGEN      — supplier que tiene el option original
#   SERVICE_TYPE         — service type del option original (ej: GU).
#                          Es un filtro DETERMINANTE de la búsqueda, no
#                          cosmético: puede haber dos options con el
#                          mismo OPTION_CODE pero distinto service type
#                          — si no se puede seleccionar en el panel de
#                          la búsqueda, la fila se marca ERROR en vez de
#                          arriesgar traer el option equivocado.
#   OPTION_CODE          — código del option a buscar en origen
#   LOCATION             — (opcional) filtro de location del origen
#
# Columnas del Excel — valores de la opción nueva (destino):
#   OPTION_CODE_DESTINO  — nuevo código (obligatorio)
#   LOCATION_DESTINO     — (opcional) nueva location
#   SERVICE_TYPE_DESTINO — (opcional) nuevo service type
#   SUPPLIER_DESTINO     — (opcional) nuevo supplier
#   DESCRIPTION          — (opcional) nueva descripción
#   COMMENT              — (opcional) nuevo comentario
#
#   ESTADO               — PENDIENTE → OK / ERROR: detalle
#
# VELOCIDAD (más abajo) multiplica todos los tiempos de espera del
# script — subirlo a 1.5 si se corre contra producción y se ven
# timeouts, en vez de tocar sleeps individuales a mano.
#
# CAMBIOS v1.5 : Fix en select_from_dropdown() (Supplier del panel de
#                búsqueda): tomaba siempre el PRIMER ítem del dropdown
#                (items[0]) sin verificar que coincidiera con el valor
#                buscado — el dropdown puede devolver más de un ítem
#                parecido, sin garantía de que el buscado quede primero
#                (mismo tipo de bug ya confirmado con Location en
#                buscar_producto(), ver notas_srv.py v1.13/v1.14). Ahora
#                busca el ítem que contiene el valor como PALABRA
#                COMPLETA; si ninguno coincide, no clickea nada a ciegas.
# ============================================================

VERSION       = "1.5"
VERSION_FECHA = "2026-09-01"

# ── CONFIGURACIÓN — via variables de entorno (con default = valor original) ──
import os

MODO       = os.environ.get("TOURPLAN_MODO", "completo")    # "lectura" | "escritura" | "completo"
EXCEL_PATH = os.environ.get("TOURPLAN_EXCEL_PATH", "/content/NEW_OPTION_input.xlsx")
HOJA       = os.environ.get("TOURPLAN_HOJA", "PRODUCTOS")

USERNAME   = os.environ.get("TOURPLAN_USERNAME", "poner minusculas")
PASSWORD   = os.environ.get("TOURPLAN_PASSWORD", "password")
BASE_URL   = os.environ.get("TOURPLAN_BASE_URL", "https://tourplannx.eurotur.com.ar/tourplannx")

SS_DIR     = os.environ.get("TOURPLAN_SS_DIR", "/content/screenshots")

# Multiplicador de tiempos de espera (sleeps y timeouts de WebDriverWait).
# Producción por default — Tourplan responde más lento ahí que en Test.
VELOCIDAD = 1.5

# ── PASO 0: Entorno ──────────────────────────────────────────
import re, sys, subprocess, importlib.util, shutil, time, traceback
from datetime import datetime

_t_inicio = time.time()

print("🔧 Verificando entorno...\n")

# 0.1 Paquetes Python
_PIPS_NEEDED = {
    "selenium":          "selenium",
    "openpyxl":          "openpyxl",
    "webdriver_manager": "webdriver-manager",
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
# 0.3 Botón Abortar de la app (ver common/abort.py)
from common.abort import chequear_abort, AbortadoPorUsuario, ABORT_EXIT_CODE

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
from selenium.common.exceptions import TimeoutException
import openpyxl

os.makedirs(SS_DIR, exist_ok=True)

WAIT_SHORT  = int(5 * VELOCIDAD)
WAIT_MEDIUM = int(15 * VELOCIDAD)
WAIT_LONG   = int(30 * VELOCIDAD)

SEL = {
    "SEARCH_BTN"          : "#searchWrapper li:nth-of-type(2) button",
    "FIELD_LOCATION"       : "div.parameters1 li:nth-of-type(1) input",
    "FIELD_SUPPLIER"       : "div.parameters1 li:nth-of-type(2) input",
    "FIELD_CODE"           : "#option tp-validator input",
    "DROPDOWN_ITEMS"       : "ul.dropdown-menu li",
    "TAB_RESULTS"          : "#tptablabel-productSearchResults",
    "RESULTS_ROWS"         : "td.tpcol-optioncode",
    "COPY_PRODUCT_BTN"     : "tp-button.copyproduct > button",
    "EXIT_PRODUCT_BTN"     : "tp-button.cancelproduct > button",
    "DIALOG"               : "tp-dialog",
    # Campos del diálogo "Copy Product" — orden real del diálogo:
    # (1) Location  (2) Service Type  (3) Supplier  (4) Option Code
    # (5) Description  (7) Comment
    "DIALOG_LOCATION"      : "tp-dialog li:nth-of-type(1) input",
    "DIALOG_SERVICE_TYPE"  : "tp-dialog li:nth-of-type(2) input",
    "DIALOG_SUPPLIER"      : "tp-dialog li:nth-of-type(3) input",
    "DIALOG_CODE"          : "#productOptionCode tp-validator input",
    "DIALOG_DESCRIPTION"   : "#productDescription tp-validator input",
    "DIALOG_COMMENT"       : "#productComment tp-validator input",
}

SCROLL_STEP_PCT = 0.8   # fracción del alto visible que se avanza por paso de scroll
SCROLL_MAX_PASOS = 40   # tope de pasos de scroll al buscar una fila en el diálogo

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

# ── Columnas Excel — origen ───────────────────────────────────
COL_SUPPLIER_ORIGEN      = "SUPPLIER_ORIGEN"
COL_SERVICE_TYPE         = "SERVICE_TYPE"
COL_OPTION_CODE          = "OPTION_CODE"
COL_LOCATION             = "LOCATION"

# ── Columnas Excel — destino (opción nueva) ───────────────────
COL_OPTION_CODE_DEST     = "OPTION_CODE_DESTINO"
COL_LOCATION_DEST        = "LOCATION_DESTINO"
COL_SERVICE_TYPE_DEST    = "SERVICE_TYPE_DESTINO"
COL_SUPPLIER_DEST        = "SUPPLIER_DESTINO"
COL_DESCRIPTION          = "DESCRIPTION"
COL_COMMENT              = "COMMENT"

COL_ESTADO               = "ESTADO"

_ss_n = [0]

def ss(driver, nombre):
    _ss_n[0] += 1
    p = f"{SS_DIR}/{_ss_n[0]:03d}_{nombre[:40]}_{int(time.time())}.png"
    try:
        driver.save_screenshot(p)
        print(f"  📸 {os.path.basename(p)}")
    except Exception:
        pass

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

def wait(driver, css, t=WAIT_MEDIUM):
    return WebDriverWait(driver, t).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, css)))

def wait_click(driver, css, t=WAIT_MEDIUM):
    return WebDriverWait(driver, t).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, css)))

def wait_overlay_gone(driver, t=WAIT_LONG):
    """
    Espera a que desaparezca cualquier <dialog open> ('PLEASE WAIT...' u
    otro overlay de carga) que tape la página tras una navegación. Sin
    esto, un click puede fallar con "element click intercepted" porque el
    dialog nativo sigue arriba del botón aunque éste ya sea "clickeable"
    para Selenium (visible + enabled no detecta overlays superpuestos).
    """
    try:
        WebDriverWait(driver, t).until(lambda d: not d.execute_script("""
            var dlg = document.querySelector('dialog[open]');
            return !!(dlg && (dlg.offsetWidth || dlg.offsetHeight || dlg.getClientRects().length));
        """))
    except TimeoutException:
        pass

def crear_driver():
    import tempfile

    opts = Options()
    # Sin --headless: esto corre en la PC de la persona (con pantalla), no en
    # el contenedor sin pantalla de Colab. Ademas, varias empresas bloquean
    # el modo headless de Chrome por politica de seguridad.
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
    assert "login" not in driver.current_url.lower(), "Login falló — verificá usuario/password"
    ss(driver, "post_login")
    print("✅ Login OK")

def open_product_setup(driver):
    print("Navegando a Product Setup...")
    handles_before = set(driver.window_handles)
    driver.get(f"{BASE_URL}/#/home")
    time.sleep(2 * VELOCIDAD)
    driver.get(f"{BASE_URL}/#/product")
    time.sleep(5 * VELOCIDAD)
    handles_after = set(driver.window_handles)
    new_handles = handles_after - handles_before
    if new_handles:
        driver.switch_to.window(new_handles.pop())
        print("  Nueva pestaña de Product Setup detectada")
    wait(driver, "#searchWrapper", t=WAIT_LONG)
    wait_overlay_gone(driver)
    print("✅ Product Setup listo")

def select_from_dropdown(driver, input_el, value):
    """Usado para los campos del panel de búsqueda (fuera del diálogo).

    Busca, entre los ítems visibles del dropdown, el que contiene `value`
    como PALABRA COMPLETA (delimitada por inicio/fin de texto o cualquier
    caracter no alfanumérico) en vez de asumir que el primer ítem
    (`items[0]`) es siempre el correcto — el dropdown puede devolver más
    de un ítem con texto parecido y no hay garantía de que el buscado
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
    """
    Selecciona el service type en el panel/sidebar de la búsqueda.

    Es un filtro DETERMINANTE, no cosmético: puede haber dos options con
    el mismo OPTION_CODE pero distinto service type, y sin este filtro la
    búsqueda podría traer el equivocado. Por eso, si el panel no carga o
    el código no aparece en él, esta función NO sigue adelante en
    silencio — devuelve (False, motivo) para que search_options aborte la
    fila con ERROR en vez de arriesgar una búsqueda ambigua.

    Escanea TODOS los <li> visibles de la página (filtrando por longitud de
    texto para descartar contenedores anidados) en vez de restringirse a un
    selector con un tag específico ("tp-list-selector ul li") — ese
    selector nunca se confirmó contra Tourplan real y, en una corrida real
    (2026-08), hizo que el panel "nunca cargara" (0 <li> encontrados,
    timeout en los 2 intentos) para el 100% de las filas de un Excel real.
    Este escaneo amplio es el mismo mecanismo ya validado en producción
    para este mismo modal en valorizacion_pkg.py / valorizacion_madre.py /
    copy_pcm_linkeo.py (buscar_producto()).

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
    ítem "no usar" que comparte una palabra.

    Reintenta unas veces porque el panel puede tardar en poblarse.
    Devuelve (True, "") si lo seleccionó, o (False, motivo) si no.
    """
    code_upper = service_type_code.upper()
    excluir_no_usar = code_upper not in Z_NO_USAR_SIGLAS
    st_buscar = STYPE_SIDEBAR.get(code_upper, code_upper)
    patron_js = r"(^|[\s\-–—/(])" + re.escape(st_buscar) + r"([\s\-–—/)]|$)"

    resultado = None
    for intento in range(3):
        wait_overlay_gone(driver)
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
            print(f"  Service Type: {resultado.get('texto')}")
            time.sleep(0.8 * VELOCIDAD)
            return True, ""
        time.sleep(2 * VELOCIDAD)

    textos = (resultado or {}).get("textos", [])
    print(f"  🩺 Service Type '{service_type_code}' (buscado como '{st_buscar}'): no "
          f"encontrado. Ítems <li> visibles en la página (máx 60): {textos}")
    return False, (f"Service Type '{service_type_code}' no aparece en el panel de la "
                    f"búsqueda ({len(textos)} ítems <li> visibles en la página)")

def search_options(driver, supplier, service_type, location="", option_code=""):
    print(f"\n  🔍 Buscando: supplier={supplier} st={service_type} "
          f"loc={location or '-'} code={option_code or 'todos'}")
    driver.get(f"{BASE_URL}/#/home")
    time.sleep(2 * VELOCIDAD)
    driver.get(f"{BASE_URL}/#/product")
    time.sleep(5 * VELOCIDAD)
    wait_overlay_gone(driver)

    jc(driver, wait_click(driver, SEL["SEARCH_BTN"]))
    time.sleep(1 * VELOCIDAD)

    if location:
        loc_inp = wait(driver, SEL["FIELD_LOCATION"])
        set_val(driver, loc_inp, location)
        loc_inp.send_keys(Keys.RETURN)
        time.sleep(0.8 * VELOCIDAD)

    sup_inp = wait(driver, SEL["FIELD_SUPPLIER"])
    selected = select_from_dropdown(driver, sup_inp, supplier)
    if not selected:
        print(f"  ⚠ Supplier '{supplier}' — sin dropdown, se usó Enter")
    time.sleep(0.5 * VELOCIDAD)

    # Service Type va ANTES del código: es el filtro que distingue dos
    # options con el mismo OPTION_CODE pero distinto service type, y
    # completar el código antes puede dejar el filtro sin aplicar sobre
    # el resultado. Si no se puede seleccionar, se aborta la búsqueda en
    # vez de arriesgar traer el option equivocado.
    ok, motivo = select_service_type(driver, service_type)
    if not ok:
        raise Exception(
            f"No se pudo filtrar por Service Type '{service_type}': {motivo}. "
            f"Se aborta la búsqueda para no arriesgar traer un option con el "
            f"código correcto pero service type equivocado."
        )

    if option_code:
        code_inp = wait(driver, SEL["FIELD_CODE"])
        set_val(driver, code_inp, option_code)
        code_inp.send_keys(Keys.RETURN)
        time.sleep(0.5 * VELOCIDAD)

    tab = wait_click(driver, SEL["TAB_RESULTS"])
    jc(driver, tab)
    time.sleep(2 * VELOCIDAD)

    try:
        row_els = WebDriverWait(driver, WAIT_MEDIUM).until(
            EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, SEL["RESULTS_ROWS"])))
    except TimeoutException:
        print("  Sin resultados")
        return []

    options = [{"code": el.text.strip(), "row_el": el} for el in row_els]
    print(f"  ✓ {len(options)} option(s): {[o['code'] for o in options]}")
    return options

# ── Campos autocompletables del diálogo (tabla de resultados) ─
def _buscar_fila_exacta_en_li(driver, li_css, valor_upper):
    """
    Busca, entre las filas visibles de la tabla dentro de `li_css`, la que
    tiene alguna celda cuyo texto coincide EXACTO (sin distinguir
    mayúsculas) con `valor_upper`. Nunca asume que la primera fila es la
    correcta: estas tablas pueden mostrar resultados sin filtrar por lo
    tipeado (grillas de Tourplan — ver skill de grillas virtuales).
    Devuelve el WebElement <tr> encontrado, o None.
    """
    return driver.execute_script("""
        var liCss = arguments[0], valor = arguments[1];
        var li = document.querySelector(liCss);
        if (!li) return null;
        var rows = li.querySelectorAll('table tbody tr');
        for (var tr of rows) {
            var tds = tr.querySelectorAll('td');
            for (var td of tds) {
                if ((td.innerText || '').trim().toUpperCase() === valor) {
                    tr.scrollIntoView({block: 'center'});
                    return tr;
                }
            }
        }
        return null;
    """, li_css, valor_upper)

def _contenedor_scroll_de_li(driver, li_css):
    """Contenedor scrolleable de la tabla dentro de `li_css` (heurística
    de detección de scroll interno — CDK, o el primer ancestro con
    overflow real subiendo desde la tabla)."""
    return driver.execute_script("""
        var li = document.querySelector(arguments[0]);
        if (!li) return null;
        var cdk = li.querySelector('cdk-virtual-scroll-viewport');
        if (cdk) return cdk;
        var tbl = li.querySelector('table');
        if (!tbl) return null;
        var el = tbl.parentElement;
        while (el && el !== li) {
            var st = getComputedStyle(el);
            var canScroll = (st.overflow === 'auto' || st.overflow === 'scroll' ||
                             st.overflowY === 'auto' || st.overflowY === 'scroll');
            if (canScroll && el.scrollHeight > el.clientHeight + 10) return el;
            el = el.parentElement;
        }
        return tbl.parentElement;
    """, li_css)

def select_from_dialog_table(driver, li_index, value, campo_nombre):
    """
    Escribe `value` en el input del li:nth-of-type(li_index) del diálogo
    Copy Product y busca, entre las filas de la tabla de resultados que
    aparece debajo, la que coincide EXACTO (código o descripción) con
    `value`. Si no aparece entre las filas visibles, scrollea el
    contenedor interno en pasos buscando de nuevo tras cada scroll — igual
    que exige cualquier grilla de Tourplan con posible virtual scroll.
    """
    input_css = f"tp-dialog li:nth-of-type({li_index}) input"
    li_css    = f"tp-dialog li:nth-of-type({li_index})"
    valor_upper = value.strip().upper()

    try:
        inp = wait(driver, input_css)
        jc(driver, inp)
        time.sleep(0.3 * VELOCIDAD)
        set_val(driver, inp, value)
        time.sleep(1.2 * VELOCIDAD)
        wait(driver, f"{li_css} table tbody tr")
    except TimeoutException:
        return False, f"No aparecieron resultados de {campo_nombre}='{value}' en el diálogo"
    except Exception as e:
        return False, f"Error escribiendo {campo_nombre}='{value}' en el diálogo: {e}"

    fila = _buscar_fila_exacta_en_li(driver, li_css, valor_upper)

    if not fila:
        contenedor = _contenedor_scroll_de_li(driver, li_css)
        if contenedor:
            for _ in range(SCROLL_MAX_PASOS):
                driver.execute_script(
                    "arguments[0].scrollTop += arguments[0].clientHeight * arguments[1];",
                    contenedor, SCROLL_STEP_PCT)
                time.sleep(0.25 * VELOCIDAD)
                fila = _buscar_fila_exacta_en_li(driver, li_css, valor_upper)
                if fila:
                    break
                al_final = driver.execute_script(
                    "return arguments[0].scrollTop + arguments[0].clientHeight "
                    ">= arguments[0].scrollHeight - 2;", contenedor)
                if al_final:
                    break

    if not fila:
        return False, (f"No encontré ninguna fila que coincida EXACTO con "
                        f"{campo_nombre}='{value}' — revisá que el código/descripción "
                        f"sea exactamente como aparece en Tourplan")

    jc(driver, fila)
    time.sleep(0.6 * VELOCIDAD)
    return True, "OK"

# ── Campos de texto validados del diálogo (option code, description, comment)
def set_validator_field(driver, css, value, campo_nombre):
    try:
        inp = wait(driver, css)
        inp.click()
        time.sleep(0.3 * VELOCIDAD)
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
        """, inp, value)
        time.sleep(0.5 * VELOCIDAD)
        val_leido = inp.get_attribute("value")
        return True, val_leido
    except Exception as e:
        return False, f"Error completando {campo_nombre}: {e}"

# ── Copia como opción nueva ────────────────────────────────────
def _abrir_dialogo_copy_product(driver, code_origen):
    """Clickea el botón "Copy Product" (asume que ya se está posicionado
    dentro del option de origen, sea porque se acaba de clickear su fila
    en resultados, o porque ya se estaba ahí de una copia anterior — ver
    copy_option_encadenado) y espera a que el diálogo aparezca."""
    try:
        copy_btn = wait_click(driver, SEL["COPY_PRODUCT_BTN"])
        jc(driver, copy_btn)
        time.sleep(1.5 * VELOCIDAD)
    except TimeoutException:
        ss(driver, f"no_copy_btn_{code_origen}")
        return False, "Botón Copy Product no encontrado"

    try:
        wait(driver, SEL["DIALOG"])
        time.sleep(1.5 * VELOCIDAD)
    except TimeoutException:
        ss(driver, f"no_dialog_{code_origen}")
        return False, "tp-dialog no apareció"

    ss(driver, f"dialog_abierto_{code_origen}")
    return True, "OK"

def _completar_y_guardar_dialogo(driver, code_origen, destino):
    """Completa los campos del diálogo "Copy Product" ya abierto y
    guarda. `destino` con claves:
      option_code_destino  (obligatorio)
      location_destino, service_type_destino, supplier_destino,
      description, comment  (opcionales — "" = no tocar ese campo,
      la opción nueva conserva el valor precargado por Tourplan)
    """
    # Location destino (opcional — li 1)
    if destino["location_destino"]:
        ok, msg = select_from_dialog_table(driver, 1, destino["location_destino"], "LOCATION_DESTINO")
        if not ok:
            ss(driver, f"dialog_location_error_{code_origen}")
            return False, msg
        print(f"    Location destino seteada: '{destino['location_destino']}'")

    # Service type destino (opcional — li 2)
    if destino["service_type_destino"]:
        ok, msg = select_from_dialog_table(driver, 2, destino["service_type_destino"], "SERVICE_TYPE_DESTINO")
        if not ok:
            ss(driver, f"dialog_service_type_error_{code_origen}")
            return False, msg
        print(f"    Service Type destino seteado: '{destino['service_type_destino']}'")

    # Supplier destino (opcional — li 3)
    if destino["supplier_destino"]:
        ok, msg = select_from_dialog_table(driver, 3, destino["supplier_destino"], "SUPPLIER_DESTINO")
        if not ok:
            ss(driver, f"dialog_supplier_error_{code_origen}")
            return False, msg
        print(f"    Supplier destino seteado: '{destino['supplier_destino']}'")

    # Option code destino (obligatorio)
    ok, val_leido = set_validator_field(
        driver, SEL["DIALOG_CODE"], destino["option_code_destino"], "OPTION_CODE_DESTINO")
    if not ok:
        ss(driver, f"dialog_code_error_{code_origen}")
        return False, val_leido
    print(f"    Option code destino seteado: '{val_leido}'")
    if val_leido != destino["option_code_destino"]:
        print(f"    ⚠ Valor leído '{val_leido}' difiere del esperado '{destino['option_code_destino']}'")

    # Description (opcional)
    if destino["description"]:
        ok, val_leido = set_validator_field(
            driver, SEL["DIALOG_DESCRIPTION"], destino["description"], "DESCRIPTION")
        if not ok:
            ss(driver, f"dialog_description_error_{code_origen}")
            return False, val_leido
        print(f"    Description seteada: '{val_leido}'")

    # Comment (opcional)
    if destino["comment"]:
        ok, val_leido = set_validator_field(
            driver, SEL["DIALOG_COMMENT"], destino["comment"], "COMMENT")
        if not ok:
            ss(driver, f"dialog_comment_error_{code_origen}")
            return False, val_leido
        print(f"    Comment seteado: '{val_leido}'")

    ss(driver, f"dialog_lleno_{code_origen}")

    # SAVE — buscar el botón enabled dentro del tp-dialog
    guardado = driver.execute_script("""
        var dlg = document.querySelector('tp-dialog');
        if (!dlg) return null;
        function vis(e){ return !!(e.offsetWidth||e.offsetHeight||e.getClientRects().length); }
        var tpSave = dlg.querySelector('tp-button.save > button');
        if (tpSave && vis(tpSave) && !tpSave.disabled){ tpSave.click(); return 'tp-button.save'; }
        var btns = Array.from(dlg.querySelectorAll('button')).filter(vis);
        for (var b of btns){
            var t = (b.innerText||'').trim().toUpperCase();
            if ((t==='SAVE'||t==='OK'||t==='ACCEPT') && !b.disabled){ b.click(); return t; }
        }
        for (var b2 of btns){
            if (!b2.disabled){ b2.click(); return 'primer-btn:' + (b2.innerText||'').trim(); }
        }
        return null;
    """)

    if not guardado:
        ss(driver, f"dialog_save_error_{code_origen}")
        return False, "No encontré botón SAVE enabled en el dialog"

    print(f"    SAVE clickeado: '{guardado}'")
    time.sleep(3 * VELOCIDAD)

    dialog_cerrado = driver.execute_script("""
        var dlg = document.querySelector('tp-dialog');
        return !dlg || !dlg.offsetParent;
    """)
    ss(driver, f"post_save_{code_origen}")

    if dialog_cerrado:
        print(f"    ✓ Dialog cerrado → guardado OK")
        return True, "OK"
    else:
        print(f"    ⚠ Dialog sigue abierto tras SAVE — posible error de validación")
        error_msg = driver.execute_script("""
            var dlg = document.querySelector('tp-dialog');
            if (!dlg) return '';
            var err = dlg.querySelector('.error, .alert, [class*="error"], [class*="alert"]');
            return err ? err.innerText.trim() : '';
        """) or ""
        return False, f"Dialog no se cerró tras SAVE. {error_msg}".strip()

def _copiar_option_ya_abierto(driver, code_origen, destino):
    ok, msg = _abrir_dialogo_copy_product(driver, code_origen)
    if not ok:
        return False, msg
    return _completar_y_guardar_dialogo(driver, code_origen, destino)

def copy_option(driver, option, destino):
    """Clickea la fila de origen en la grilla de resultados y copia como
    opción nueva. Usar cuando todavía no se está posicionado dentro de un
    option con la identidad correcta (primera fila de un grupo) — ver
    copy_option_encadenado() para el caso en que ya se está (fila
    siguiente de un grupo que copia el mismo option a muchos suppliers)."""
    code_origen = option["code"]
    print(f"  → Copiando '{code_origen}' → '{destino['option_code_destino']}' (opción nueva)")

    try:
        jc(driver, option["row_el"])
        time.sleep(1.5 * VELOCIDAD)
    except Exception as e:
        return False, f"No se pudo clickear la fila: {e}"

    return _copiar_option_ya_abierto(driver, code_origen, destino)

def copy_option_encadenado(driver, code_origen, destino):
    """Copia como opción nueva SIN buscar ni clickear ninguna fila — asume
    que el driver ya está posicionado dentro de un option con la misma
    location + service type + código que el origen: la copia recién
    creada por la fila anterior del mismo grupo (ver
    _copia_preserva_identidad y _agrupar_por_origen). Ahorra la búsqueda
    completa (home→product→modal→4 filtros→resultados→click fila) para
    cada fila de un lote que copia el mismo option a muchos suppliers
    distintos, manteniendo todos los demás campos iguales — el caso que
    reportó la usuaria (160 filas, mismo OPTION_CODE/SERVICE_TYPE/
    LOCATION, solo cambia SUPPLIER_DESTINO)."""
    print(f"  → Copiando '{code_origen}' → '{destino['option_code_destino']}' "
          f"(encadenado, sin re-buscar el origen)")
    return _copiar_option_ya_abierto(driver, code_origen, destino)

def exit_to_results(driver):
    try:
        jc(driver, wait_click(driver, SEL["EXIT_PRODUCT_BTN"]))
        time.sleep(1.5 * VELOCIDAD)
    except Exception:
        print("    ⚠ No se pudo hacer Exit Product")

def refresh_results_rows(driver):
    try:
        return WebDriverWait(driver, WAIT_MEDIUM).until(
            EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, SEL["RESULTS_ROWS"])))
    except TimeoutException:
        return []

def load_excel(path):
    wb = openpyxl.load_workbook(path)
    ws = wb[HOJA]
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    col_idx = {h: i + 1 for i, h in enumerate(headers) if h}
    rows = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(row):
            continue
        d = dict(zip(headers, row))
        d["__row_idx__"] = row_idx
        rows.append(d)
    print(f"Excel cargado: {len(rows)} fila(s) en '{HOJA}'")
    return wb, rows, col_idx

def update_estado(wb, row_idx, col_idx, estado, path):
    ws = wb[HOJA]
    ws.cell(row=row_idx, column=col_idx[COL_ESTADO]).value = estado
    wb.save(path)

def _clave_origen(row):
    """Identidad del origen de una fila: mismo SUPPLIER_ORIGEN + SERVICE_TYPE
    + OPTION_CODE + LOCATION → search_options() encontraría exactamente el
    mismo option. Usada para agrupar filas que copian desde el mismo origen
    (ver _agrupar_por_origen)."""
    return (
        str(row.get(COL_SUPPLIER_ORIGEN) or "").strip().upper(),
        str(row.get(COL_SERVICE_TYPE) or "").strip().upper(),
        str(row.get(COL_OPTION_CODE) or "").strip().upper(),
        str(row.get(COL_LOCATION) or "").strip().upper(),
    )

def _copia_preserva_identidad(row):
    """True si esta fila copia manteniendo la misma location + service type
    + código que el origen (solo cambia supplier y/o description/comment).
    Es la condición para poder encadenar la fila SIGUIENTE del mismo grupo
    sin re-buscar: tras guardar, Tourplan queda posicionado en la copia
    recién creada, que por tener la misma location+service type+código que
    el origen es un punto de partida válido para la próxima copia del
    grupo. Si esta fila cambiara location/service type, o renombrara el
    código, la copia resultante ya NO tendría la identidad del grupo y no
    se puede reusar."""
    option_code = str(row.get(COL_OPTION_CODE) or "").strip().upper()
    option_code_destino = str(row.get(COL_OPTION_CODE_DEST) or "").strip().upper()
    return (
        not str(row.get(COL_LOCATION_DEST) or "").strip()
        and not str(row.get(COL_SERVICE_TYPE_DEST) or "").strip()
        and option_code_destino == option_code
    )

def _agrupar_por_origen(pendientes):
    """Agrupa las filas PENDIENTE que copian desde el mismo origen (mismo
    SUPPLIER_ORIGEN + SERVICE_TYPE + OPTION_CODE + LOCATION), preservando
    el orden relativo dentro de cada grupo y el orden de aparición de los
    grupos. Devuelve una lista de grupos (cada grupo es una lista de
    filas). Ahorra la búsqueda completa (home→product→modal→4 filtros→
    resultados→click fila) para todas las filas de un grupo salvo la
    primera — ver copy_option_encadenado()."""
    grupos = {}
    orden = []
    for row in pendientes:
        clave = _clave_origen(row)
        if clave not in grupos:
            grupos[clave] = []
            orden.append(clave)
        grupos[clave].append(row)
    return [grupos[clave] for clave in orden]

def process_row(driver, row, ya_posicionado=False):
    """Procesa una fila. Si `ya_posicionado` es True, el driver ya está
    posicionado dentro de un option con la misma location+service type+
    código que esta fila necesita (dejado ahí por la copia de la fila
    anterior del mismo grupo — ver _agrupar_por_origen) y se salta la
    búsqueda completa, yendo directo a Copy Product.

    Devuelve (nuevo_estado, se_puede_encadenar): `se_puede_encadenar` es
    True solo si la copia salió OK y esta fila no cambió location/service
    type/código en el destino (ver _copia_preserva_identidad) — así la
    fila siguiente del grupo puede reusar esta misma posición."""
    supplier_origen      = str(row.get(COL_SUPPLIER_ORIGEN)   or "").strip()
    service_type         = str(row.get(COL_SERVICE_TYPE)      or "").strip()
    option_code          = str(row.get(COL_OPTION_CODE)       or "").strip()
    location             = str(row.get(COL_LOCATION)          or "").strip()

    option_code_destino  = str(row.get(COL_OPTION_CODE_DEST)  or "").strip()
    location_destino     = str(row.get(COL_LOCATION_DEST)     or "").strip()
    service_type_destino = str(row.get(COL_SERVICE_TYPE_DEST) or "").strip()
    supplier_destino     = str(row.get(COL_SUPPLIER_DEST)     or "").strip()
    description          = str(row.get(COL_DESCRIPTION)       or "").strip()
    comment              = str(row.get(COL_COMMENT)           or "").strip()
    row_idx              = row["__row_idx__"]

    if not supplier_origen or not service_type:
        return "ERROR: Faltan campos obligatorios (SUPPLIER_ORIGEN, SERVICE_TYPE)", False
    if not option_code:
        return "ERROR: OPTION_CODE es obligatorio (se necesita saber qué option buscar en el origen)", False
    if not option_code_destino:
        return "ERROR: OPTION_CODE_DESTINO es obligatorio (es el código de la opción nueva)", False

    destino = {
        "option_code_destino":  option_code_destino,
        "location_destino":     location_destino,
        "service_type_destino": service_type_destino,
        "supplier_destino":     supplier_destino,
        "description":          description,
        "comment":              comment,
    }

    if MODO == "lectura":
        # En lectura nunca se abre el diálogo de copia (no hay adónde
        # "quedar posicionado"), así que siempre busca desde cero.
        try:
            options = search_options(driver, supplier_origen, service_type, location, option_code)
        except Exception as e:
            ss(driver, f"search_error_row{row_idx}")
            return f"ERROR: Falla en búsqueda — {e}", False
        if not options:
            return f"ERROR: No se encontró el option '{option_code}' con los filtros dados", False
        opt = next((o for o in options if o["code"] == option_code), options[0])
        cambios = ", ".join(f"{k}='{v}'" for k, v in destino.items() if v)
        print(f"  [LECTURA] Origen: '{opt['code']}' → {cambios}")
        return f"LECTURA: '{opt['code']}' → {cambios}", False

    puede_encadenar = _copia_preserva_identidad(row)

    try:
        if ya_posicionado:
            ok, msg = copy_option_encadenado(driver, option_code, destino)
        else:
            try:
                options = search_options(driver, supplier_origen, service_type, location, option_code)
            except Exception as e:
                ss(driver, f"search_error_row{row_idx}")
                return f"ERROR: Falla en búsqueda — {e}", False
            if not options:
                return f"ERROR: No se encontró el option '{option_code}' con los filtros dados", False
            # Tomamos solo el primero que coincida exactamente con el código buscado
            opt = next((o for o in options if o["code"] == option_code), options[0])
            if opt["code"] != option_code:
                print(f"  ⚠ No hay coincidencia exacta; usando el primero: '{opt['code']}'")
            ok, msg = copy_option(driver, opt, destino)

        # Si no se puede encadenar la fila siguiente (falló, o esta fila
        # cambió location/service type/código), salir del option para no
        # arrastrar un estado a medias a la próxima búsqueda desde cero.
        if not (ok and puede_encadenar):
            exit_to_results(driver)
    except Exception as e:
        ss(driver, f"copy_error_row{row_idx}")
        try: exit_to_results(driver)
        except Exception: pass
        return f"ERROR: {e}", False

    if not ok:
        return f"ERROR: {msg}", False
    return "OK", puede_encadenar

# ── MAIN ──────────────────────────────────────────────────────
print("=" * 60)
print(f"  NEW OPTION (Copy Product)  v{VERSION}  [{VERSION_FECHA}]  MODO={MODO}")
print("=" * 60)

if not os.path.exists(EXCEL_PATH):
    raise FileNotFoundError(f"No encontré el Excel en {EXCEL_PATH} — subilo a /content/")

wb, rows, col_idx = load_excel(EXCEL_PATH)

if COL_OPTION_CODE_DEST not in col_idx:
    raise ValueError(
        f"Columna '{COL_OPTION_CODE_DEST}' no encontrada en el Excel.\n"
        f"Columnas detectadas: {list(col_idx.keys())}"
    )

pendientes = [r for r in rows
              if str(r.get(COL_ESTADO) or "").strip().upper() == "PENDIENTE"]

print(f"Filas PENDIENTE: {len(pendientes)}")
for f in pendientes:
    print(f"  · fila {f['__row_idx__']}: {f.get(COL_SUPPLIER_ORIGEN)}  "
          f"code={f.get(COL_OPTION_CODE)} → {f.get(COL_OPTION_CODE_DEST)}  "
          f"supplier_destino={f.get(COL_SUPPLIER_DEST) or '(mismo)'}")

if not pendientes:
    print("\n⛔ Sin filas PENDIENTE. Verificá que la columna ESTADO tenga 'PENDIENTE'.")
    raise SystemExit

grupos = _agrupar_por_origen(pendientes)
n_grupos_multi = sum(1 for g in grupos if len(g) > 1)
if n_grupos_multi:
    print(f"Agrupadas por mismo origen (SUPPLIER_ORIGEN+SERVICE_TYPE+OPTION_CODE+"
          f"LOCATION): {n_grupos_multi} grupo(s) de 2+ filas — para esas, la fila "
          f"siguiente reusa la copia recién creada en vez de re-buscar el origen "
          f"desde cero (solo cuando no cambia location/service type/código).")

driver = crear_driver()
_abortado = False
try:
    login(driver)
    open_product_setup(driver)

    for grupo in grupos:
        ya_posicionado = False
        for row in grupo:
            chequear_abort()
            print(f"\n{'─'*60}")
            print(f"Fila {row['__row_idx__']}: "
                  f"{row.get(COL_SUPPLIER_ORIGEN)}  '{row.get(COL_OPTION_CODE)}'"
                  f" → '{row.get(COL_OPTION_CODE_DEST)}'"
                  + ("  (encadenado, sin re-buscar)" if ya_posicionado else ""))

            nuevo_estado = "PENDIENTE"
            puede_encadenar = False
            try:
                nuevo_estado, puede_encadenar = process_row(driver, row, ya_posicionado)
            except Exception as e:
                nuevo_estado = f"ERROR: {traceback.format_exc(limit=3)}"
                ss(driver, f"fatal_row{row['__row_idx__']}")

            print(f"  Estado: {nuevo_estado}")
            update_estado(wb, row["__row_idx__"], col_idx, nuevo_estado, EXCEL_PATH)
            ya_posicionado = puede_encadenar

except AbortadoPorUsuario:
    _abortado = True
    print("\n⏸️  Corrida abortada por el usuario — las filas que no llegaron a "
          "procesarse quedan en PENDIENTE para retomar en otra corrida.")
finally:
    driver.quit()
    dur = int(time.time() - _t_inicio)
    m, s = divmod(dur, 60)
    print(f"\n🏁 Fin. Duración: {m}m {s:02d}s")
    print(f"📄 Excel: {EXCEL_PATH}")
    print(f"📂 Screenshots: {SS_DIR}")

if _abortado:
    sys.exit(ABORT_EXIT_CODE)

try:
    from google.colab import files
    files.download(EXCEL_PATH)
    print("📥 Excel descargado.")
except Exception:
    pass
