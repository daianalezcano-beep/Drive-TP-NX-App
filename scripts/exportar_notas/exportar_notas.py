# ============================================================
# TOURPLAN NX — EXPORTAR NOTAS (Product Notes)
# Lee el contenido de una nota puntual de cada producto y lo vuelca al
# Excel — script hermano de notas_srv.py (comparte toda la navegación),
# pero de SOLO LECTURA: nunca inserta, edita ni guarda nada en Tourplan.
# Copia adaptada para tp-nx-app (app local). El original para Google Colab
# vive sin cambios en el repo copy-products, rama notas-SRV. Mismos
# cambios que el resto de los scripts vendorizados: config por variables
# de entorno, Chrome delegado a common/chrome_bootstrap.py, sin
# --headless (corre con pantalla, no en el contenedor de Colab), perfil
# de Chrome aislado, log en carpeta temporal del OS. El resto (selectores,
# login, navegación a Product Notes, lectura de notas) es igual — ver
# ESTADO.md en el repo origen para el historial completo de bugs y fixes
# ya resueltos antes de tocar esa parte.
# ------------------------------------------------------------
#   VERSION : 1.4
#   FECHA   : 2026-09-01
# ============================================================

VERSION       = "1.4"
VERSION_FECHA = "2026-09-01"

# ── CAMBIOS ──
# v1.0: primera versión — basada en notas_srv.py v1.9 (bootstrap, login,
#       buscar_producto, abrir_product_notes y lectura de la tabla de
#       notas ya confirmados funcionando en corrida real). El código de
#       nota a exportar no está restringido a una lista fija — se busca
#       lo que venga en la columna Codigo_Nota de cada fila.
# v1.1: fix — la tabla de Product Notes usa Angular CDK virtual scroll
#       cuando el producto tiene varias notas cargadas (ej. los 5
#       Descriptivo por idioma). parsear_notas() leía solo las filas
#       actualmente renderizadas en el DOM, así que códigos que existían
#       pero no estaban en el viewport inicial se reportaban como
#       "NO EXISTE". Reemplazado por buscar_nota_por_codigo(): scrollea
#       progresivamente el contenedor scrolleable (detectado por
#       cascada: viewport CDK → clase confirmada .productnoteslistview →
#       ancestro con overflow → fallback documento), releyendo
#       tbody tr en cada paso (nunca cachea WebElements) y comparando la
#       columna CAT contra el código buscado, con salida temprana tras
#       8 pasos sin códigos nuevos vistos. Basado en el skill
#       recorriendo-grillas-virtuales-de-tourplan.
# v1.2: el fix de v1.1 seguía dando "NO EXISTE" para un código real
#       (confirmado en corrida real, código UES sobre IGR/1RIOTU/PARBR1).
#       Causa: _detectar_contenedor_scroll() buscaba el primer
#       `cdk-virtual-scroll-viewport` con `document.querySelector` en TODA
#       la página — si existe cualquier otro viewport CDK en pantalla
#       (otro panel, otro dropdown) que no tiene nada que ver con esta
#       tabla, se scrollea ese elemento equivocado, no aparece ninguna
#       fila nueva nunca, y el corte por "8 pasos sin cambios" da el
#       mismo falso "NO EXISTE" casi de inmediato — el fix de v1.1 nunca
#       llegaba a scrollear la tabla real. Corregido: la búsqueda del
#       viewport CDK y de `.productnoteslistview` ahora está limitada a
#       los ANCESTROS de la tabla (subiendo con `parentElement`/
#       `closest`), nunca al documento completo. Se agrega además
#       diagnóstico: se loguea qué contenedor se detectó (tag, clase,
#       scrollHeight/clientHeight) y, si el código buscado no aparece, la
#       lista completa de códigos CAT realmente vistos durante el
#       recorrido — vuelca esa lista en DETALLE_PROCESO y toma
#       screenshot + dump HTML, para poder confirmar en el próximo
#       "NO EXISTE" si es un problema de scroll o si el código
#       simplemente no existe en ese producto puntual.
# v1.3: el diagnóstico agregado en v1.2 mostró la causa REAL, distinta de
#       las dos hipótesis anteriores: "contenedor detectado:
#       HTML.tp-viewport-zoom scrollHeight=869 clientHeight=869" y
#       "Códigos vistos: []" — o sea, cero filas encontradas, incluso sin
#       necesidad de scroll. La usuaria compartió el dump completo del
#       DOM real (producto IGR/1RIOTU/PARBR1, código UES): confirma que
#       Product Notes NO usa Angular CDK virtual scroll en absoluto — no
#       hay ningún `cdk-virtual-scroll-viewport` en la página, y las ~20
#       notas del producto (incluyendo UES) están TODAS presentes en el
#       DOM a la vez, sin scrollear nada. El diagnóstico original de
#       "virtual scroll" (v1.0/v1.1/v1.2) era incorrecto. Los dos bugs
#       reales, confirmados por el dump: (1) encontrar_tabla_notas()
#       busca el <table> cuyo header dice "CAT"+"DESCRIPTION" y lo
#       devuelve tal cual — pero esa es la tabla del HEADER, que no tiene
#       ningún <tbody> con filas; los datos están en tablas SEPARADAS,
#       una por nota, envueltas cada una en <div class="tprow-even"/
#       "tprow-odd">, todas hermanas dentro de un contenedor común
#       (.productnoteslistview) — por eso "tbody tr" nunca encontraba
#       nada. (2) _fila_a_nota() leía columnas por índice posicional de
#       <td> asumiendo 6 columnas en orden fijo, pero cada fila real
#       tiene más <td> (celdas de expansor + ícono de lápiz intercaladas)
#       — aun si se hubiese encontrado la fila, el mapeo de columnas
#       habría sido incorrecto. Corregido: encontrar_tabla_notas() ahora
#       devuelve el contenedor común (.productnoteslistview), no la tabla
#       de header, para que "tbody tr" vea todas las tablas hermanas;
#       _fila_a_nota() lee cada campo por su clase CSS estable
#       (td.tpcol-category, td.tpcol-description, etc.), inmune al orden
#       o cantidad real de columnas. Se elimina toda la lógica de
#       detección de contenedor scrolleable y scroll progresivo
#       (_detectar_contenedor_scroll, _scroll_a) — no resolvía nada real
#       y agregaba complejidad innecesaria una vez confirmado que no hay
#       virtualización en esta grilla.
# v1.4: bug encontrado al revisar el resto de los scripts vendorizados
#       (mismo patrón ya arreglado en notas_srv.py v1.13/v1.14): en
#       buscar_producto(), al completar Location, si la fila que Tourplan
#       marca como `tr.selectedRow` no aparece a tiempo, el fallback
#       clickeaba SIEMPRE la primera fila de la tabla de sugerencias sin
#       verificar que fuera el Location exacto buscado. Corregido: busca,
#       entre las filas sugeridas, la que tiene una celda con texto EXACTO
#       (case-insensitive) igual al Location pedido; si ninguna coincide,
#       no se clickea nada a ciegas.

# ── CONFIGURACIÓN — via variables de entorno (con default = valor original) ──
import os

from common.user_config import (
    CREDENTIALS_PATH as _CREDENTIALS_PATH_DEFAULT,
    TOKEN_PATH as _TOKEN_PATH_DEFAULT,
)

SHEET_URL        = os.environ.get("TOURPLAN_SHEET_URL", "")
HOJA             = os.environ.get("TOURPLAN_HOJA", "EXPORTAR_NOTAS")
CREDENTIALS_PATH = os.environ.get("TOURPLAN_CREDENTIALS_PATH", _CREDENTIALS_PATH_DEFAULT)
TOKEN_PATH       = os.environ.get("TOURPLAN_TOKEN_PATH", _TOKEN_PATH_DEFAULT)

USERNAME   = os.environ.get("TOURPLAN_USERNAME", "poner minusculas")
PASSWORD   = os.environ.get("TOURPLAN_PASSWORD", "password")
# Test y Producción son rutas BASE distintas (no un flag de ambiente) —
# confirmar cuál corresponde antes de correr. Ver skill
# arrancando-un-script-de-tourplan/references/convenciones-arranque.md
BASE_URL   = os.environ.get("TOURPLAN_BASE_URL", "https://tourplannx.eurotur.com.ar/TourplanNX_Test")

SS_DIR     = os.environ.get("TOURPLAN_SS_DIR", "/content/screenshots")
HEADLESS   = os.environ.get("TOURPLAN_HEADLESS", "0").strip() in ("1", "true", "True")

# Multiplicador de tiempos de espera. Producción por default — Tourplan
# responde más lento ahí que en Test (mismo criterio que el resto de los
# scripts vendorizados; no ajustar sleeps a mano).
VELOCIDAD = 1.5

# ── PASO 0: Entorno ───────────────────────────────────────────
import html
import importlib.util
import os
import re
import subprocess
import sys
import time
import traceback

print("🔧 Verificando entorno...\n")

_PIPS_NEEDED = {
    "selenium": "selenium",
    "webdriver_manager": "webdriver-manager",
    "gspread": "gspread",
    "google_auth_oauthlib": "google-auth-oauthlib",
}
_faltantes = [pkg for mod, pkg in _PIPS_NEEDED.items()
              if importlib.util.find_spec(mod) is None]
if _faltantes:
    print(f"  ⏳ pip install {' '.join(_faltantes)} ...")
    subprocess.run([sys.executable, "-m", "pip", "install", "-q"] + _faltantes, check=True)
    print("  ✅ Paquetes Python OK")
else:
    print("  ✅ Paquetes Python ya instalados")


# Detección/instalación de Chrome, multiplataforma (ver common/chrome_bootstrap.py).
# En Windows/Mac nunca instala solo — la lógica original (apt-get/.deb) solo
# tenía sentido en el contenedor Linux de Colab.
from common.chrome_bootstrap import find_or_prepare_chrome
# Botón Abortar de la app (ver common/abort.py)
from common.abort import chequear_abort, AbortadoPorUsuario, ABORT_EXIT_CODE
# Google Sheets como cola de trabajo (ver common/sheets_client.py)
from common.sheets_client import conectar_sheets, cargar_sheet, actualizar_fila_sheet

CHROMIUM_BIN, ver_chrome = find_or_prepare_chrome()

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

os.makedirs(SS_DIR, exist_ok=True)

WAIT_SHORT = 5
WAIT_MEDIUM = 15
WAIT_LONG = 30

# ── Sigla de Service Type → prefijo numérico del sidebar de Product
# Search (ver skill buscando-productos-en-tourplan/references/service-types.md).
# Si aparece una sigla nueva sin número confirmado, NO asumir uno acá — dejar
# que buscar_producto() use su fallback por texto/regex y confirmarlo contra
# Test antes de sumarlo.
STYPE_SIDEBAR = {
    "HT": "01", "HX": "02", "TF": "03", "EX": "04", "ML": "05",
    "RT": "06", "CR": "07", "FT": "08", "OC": "09", "LN": "10",
    "LP": "11", "MS": "12",
}

SEL = {
    "NAV_ICON":    "nav img",
    "NOTE_EDITOR": "#noteeditorview",
    "EDIT_ICON":   "i.fa-pencil-square-o.note-icon",
}


class ProductoNoEncontrado(Exception):
    """El código de producto buscado no apareció en los resultados de
    Product Search. Se levanta explícitamente para distinguir 'no existe'
    de un error real de Selenium/Tourplan."""
    pass


# ── Helpers de DOM ─────────────────────────────────────────────
def jc(driver, el):
    """Click vía JavaScript — único método confiable en Angular."""
    driver.execute_script("arguments[0].click();", el)


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
    """Vuelca el HTML de la página — complementa a ss() cuando el problema
    es de estructura del DOM (selector no encontrado) más que visual."""
    p = f"{SS_DIR}/{nombre}_{int(time.time())}.html"
    try:
        with open(p, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print(f"  💾 HTML: {p}")
    except Exception as e:
        print(f"  ⚠️ Error guardando HTML ({e})")


def set_val(driver, el, value):
    """set_val corto — sin focus/blur, para inputs de filtro/búsqueda."""
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


def cerrar_nav_backdrop(driver, timeout=3):
    """Limpia `.tpnavbackdrop` si re-clickear el ícono del menú hamburguesa
    no alcanza a sacarlo a tiempo (bug confirmado en corrida real — ver
    skill buscando-productos-en-tourplan)."""
    fin = time.time() + timeout * VELOCIDAD
    while time.time() < fin:
        if not driver.find_elements(By.CSS_SELECTOR, ".tpnavbackdrop"):
            return
        time.sleep(0.2)
    driver.execute_script("""
        document.querySelectorAll('.tpnavbackdrop').forEach(function(e){ e.remove(); });
    """)
    print("    ⚠ .tpnavbackdrop seguía presente — removido a mano")


def esperar_fin_carga(driver, timeout=15):
    """Espera a que desaparezca el <dialog open> nativo 'PLEASE WAIT...'
    de Tourplan antes de clickear la lupa de búsqueda entre filas."""
    fin = time.time() + timeout * VELOCIDAD
    while time.time() < fin:
        if not driver.find_elements(By.CSS_SELECTOR, "dialog[open]"):
            return
        time.sleep(0.3)
    print("    ⚠ El dialog 'PLEASE WAIT...' seguía abierto tras esperar")


# ── Driver ─────────────────────────────────────────────────────
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

    # Perfil temporal y aislado para esta corrida — evita que Chrome
    # "rebote" hacia una ventana ya abierta con los perfiles reales de la
    # persona (instancia única de Windows compartiendo el mismo
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


# ── Login / Logout ──────────────────────────────────────────────
def login(driver):
    global BASE_URL
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
        try:
            el.clear()
        except Exception:
            pass
        try:
            el.click()
        except Exception:
            pass
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
        if (b){ b.click(); return (b.innerText||'button').trim(); }
        return null;
    """)
    if not clic:
        try:
            p_el.send_keys(Keys.ENTER)
        except Exception:
            pass

    time.sleep(8 * VELOCIDAD)
    assert "login" not in driver.current_url.lower(), "Login falló — verificá usuario/password"
    ss(driver, "post_login")
    print("✅ Login OK")

    # Resguardo defensivo (ver skill arrancando-un-script-de-tourplan):
    # Test/Producción son rutas BASE distintas — detectar la base real tras
    # el login evita fallas silenciosas de navegación si BASE_URL quedó mal
    # configurado.
    base_real = driver.current_url.split("#")[0].rstrip("/")
    if base_real and base_real != BASE_URL:
        print(f"  ⚠ BASE_URL configurado ({BASE_URL}) difiere de la base real "
              f"post-login ({base_real}) — usando la real para el resto de la corrida")
        BASE_URL = base_real


def logout(driver):
    """Cierra ventanas secundarias y hace logout real. Ver regla dura de
    sesiones/licencias concurrentes en skill arrancando-un-script-de-tourplan."""
    print("\n🔒 Finalización: cerrando ventanas y haciendo logout...")
    try:
        principal = driver.window_handles[0]
        for h in driver.window_handles[1:]:
            try:
                driver.switch_to.window(h)
                driver.close()
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


# ── Búsqueda de producto (buscar_producto — entra de lleno al registro) ──
def buscar_producto(driver, location, supplier, codigo, service_type=None):
    """
    Flujo (ver skill buscando-productos-en-tourplan), idéntico al de
    notas_srv.py — ya confirmado funcionando en corrida real:
    1. #/home → #/product (siempre pasar por #/home antes).
    2. Click lupa → modal Product Search (tabs SELECTION/RESULTS).
    3. Service Type en el sidebar izquierdo (prefijo numérico, no sigla).
    4. Location → Supplier → Code.
    5. SEARCH → click en la fila de resultado que matchea código+service type.
    6. Verificar contexto abriendo el menú hamburguesa (UTILITIES/RATES/...).
    """
    st_upper = (service_type or "").strip().upper()
    st_str = f"/{st_upper}" if st_upper else ""
    print(f"\n  📦 Buscando: {location}/{supplier}/{codigo}{st_str}")

    driver.get(f"{BASE_URL}/#/home")
    time.sleep(2 * VELOCIDAD)
    driver.get(f"{BASE_URL}/#/product")
    time.sleep(5 * VELOCIDAD)
    ss(driver, f"ps_inicial_{codigo[:10]}")

    esperar_fin_carga(driver)

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
            st_buscar = STYPE_SIDEBAR.get(st_upper, st_upper)
            patron_js = (r"(^|[\s\-–—/(])" + re.escape(st_buscar) + r"([\s\-–—/)]|$)")
            resultado = None
            for _ in range(3):
                resultado = driver.execute_script("""
                    var rx = new RegExp(arguments[0]);
                    var textos = [], elegido = null;
                    for (var li of document.querySelectorAll('li')){
                        if (!li.offsetParent) continue;
                        var t = (li.textContent || '').trim();
                        if (!t || t.length > 80) continue;
                        textos.push(t);
                        if (!elegido && rx.test(t.toUpperCase())) elegido = li;
                    }
                    if (elegido){
                        elegido.click();
                        return {ok: true, texto: (elegido.textContent||'').trim().slice(0,40)};
                    }
                    return {ok: false, textos: textos.slice(0,40)};
                """, patron_js)
                if resultado and resultado.get("ok"):
                    break
                time.sleep(1.5)
            if resultado and resultado.get("ok"):
                time.sleep(1.5)
                print(f"    Service type '{st_upper}' seleccionado: {resultado.get('texto')}")
            else:
                print(f"    ⚠ Service type '{st_upper}' no encontrado en sidebar — buscando sin filtro")
        except Exception as e:
            print(f"    ⚠ Error seleccionando service type: {e}")

    if location:
        try:
            inp_loc = wait(driver, "div.parameters1 li:nth-of-type(1) input")
            inp_loc.click()
            time.sleep(0.3)
            set_val(driver, inp_loc, location)
            time.sleep(1.5)
            row = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "tr.selectedRow > td.description")))
            jc(driver, row)
            time.sleep(1)
            print(f"    Location: {location} OK")
        except Exception:
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
                    jc(driver, fila_exacta)
                    time.sleep(1)
                else:
                    print(f"    ⚠ Location '{location}': ninguna fila sugerida coincide "
                          f"EXACTO — no se clickeó ninguna a ciegas (antes se tomaba la "
                          f"primera, podía ser otro Location)")
            except Exception:
                pass

    if supplier:
        try:
            inp_sup = wait(driver, "div.parameters1 li:nth-of-type(2) input")
            inp_sup.click()
            time.sleep(0.3)
            set_val(driver, inp_sup, supplier)
            time.sleep(1.5)
            row_sup = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located(
                    (By.XPATH,
                     f"//div[contains(@class,'parameters1')]//td[normalize-space(text())='{supplier.upper()}']")))
            jc(driver, row_sup)
            time.sleep(1)
            print(f"    Supplier: {supplier} OK")
        except Exception:
            try:
                inp_sup2 = driver.find_element(By.CSS_SELECTOR,
                    "div.parameters1 li:nth-of-type(2) input")
                inp_sup2.send_keys(Keys.TAB)
                time.sleep(1)
                print(f"    Supplier: {supplier} (TAB)")
            except Exception:
                pass

    inp_cod = wait(driver, "div.parameters1 li:nth-of-type(3) input")
    inp_cod.click()
    time.sleep(0.3)
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
            PROD_ITEMS = {'UTILITIES', 'RATES', 'SEASONALITY', 'OPERATION', 'CONTENT', 'PRODUCT DETAILS'}
            ok = bool(set(items) & PROD_ITEMS)
            driver.execute_script("arguments[0].click();", img)
            time.sleep(1)
            cerrar_nav_backdrop(driver)
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

    page_info = driver.execute_script(f"""
        return {{
            url: window.location.href,
            inputs: Array.from(document.querySelectorAll('input')).map(i=>i.value).filter(Boolean).slice(0,5),
            tables: Array.from(document.querySelectorAll('table tbody tr')).slice(0,5).map(r=>r.innerText.trim().slice(0,60))
        }};
    """)
    print(f"    Sin resultado: {page_info}")
    dump(driver, f"ps_sin_resultado_{codigo[:10]}")
    ss(driver, f"ps_sin_resultado_final_{codigo[:10]}")
    raise ProductoNoEncontrado(f"Producto '{codigo}' no encontrado en resultados")


# ── Navegación a Product Notes ────────────────────────────────
def abrir_product_notes(driver, timeout=10):
    """
    Abre el submenú "Product Notes" del producto ya cargado por
    buscar_producto(). Idéntico al de notas_srv.py — ya confirmado
    funcionando en corrida real: "Product Notes" vive anidado dentro de
    "Content", ubicado siempre por TEXTO (nunca por posición, el menú de
    primer nivel varía según los módulos habilitados de cada producto).
    """
    nav_icon = wait(driver, SEL["NAV_ICON"], t=timeout)
    jc(driver, nav_icon)
    time.sleep(1.5 * VELOCIDAD)

    def _click_item_de_nivel(rx_texto):
        return driver.execute_script("""
            var rx = new RegExp(arguments[0], 'i');
            var items = Array.from(document.querySelectorAll('nav ul > li'));
            for (var li of items) {
                if (!li.offsetParent) continue;
                var heading = li.querySelector(':scope > div') || li;
                var t = (heading.innerText || '').trim().split('\\n')[0];
                if (rx.test(t)) {
                    var clickArea = heading.querySelector(':scope > div') || heading;
                    clickArea.click();
                    return t;
                }
            }
            return null;
        """, rx_texto)

    def _click_en_submenu_abierto(rx_texto):
        return driver.execute_script("""
            var rx = new RegExp(arguments[0], 'i');
            var abierto = document.querySelector('li.tpnavmenuopen');
            if (!abierto) return null;
            var labels = Array.from(abierto.querySelectorAll('label'));
            for (var el of labels) {
                if (!el.offsetParent) continue;
                var t = (el.innerText || '').trim().split('\\n')[0];
                if (rx.test(t)) { el.click(); return t; }
            }
            return null;
        """, rx_texto)

    def _content_ya_expandido():
        return bool(driver.execute_script("""
            var abierto = document.querySelector('li.tpnavmenuopen');
            if (!abierto) return false;
            var heading = abierto.querySelector(':scope > div') || abierto;
            var t = (heading.innerText || '').trim().split('\\n')[0];
            return /^\\s*content\\s*$/i.test(t);
        """))

    def _esperar_y_click_submenu(rx_texto, timeout_sub=8):
        fin = time.time() + timeout_sub * VELOCIDAD
        while time.time() < fin:
            encontrado = _click_en_submenu_abierto(rx_texto)
            if encontrado:
                return encontrado
            time.sleep(0.3)
        return None

    clicked = _click_en_submenu_abierto(r"^\s*product notes\s*$")

    if not clicked:
        if _content_ya_expandido():
            print("    'Content' ya estaba expandido — no se re-clickea (evitar colapsarlo)")
        else:
            padre = _click_item_de_nivel(r"^\s*content\s*$")
            if not padre:
                dump(driver, "product_notes_sin_content")
                ss(driver, "product_notes_sin_content")
                raise Exception("No encontré 'Content' entre los ítems de primer nivel del menú")
            print("    'Content' expandido")
        clicked = _esperar_y_click_submenu(r"^\s*product notes\s*$")

    if not clicked:
        dump(driver, "product_notes_no_encontrado")
        ss(driver, "product_notes_no_encontrado")
        raise Exception("No se encontró 'Product Notes' dentro de 'Content' — revisar selector contra Test")

    time.sleep(2 * VELOCIDAD)
    cerrar_nav_backdrop(driver)
    print(f"    Product Notes abierto (click en '{clicked}')")


# ── Lectura de la tabla de notas existentes ───────────────────
def encontrar_tabla_notas(driver, timeout=15):
    """Ubica el CONTENEDOR de la grilla de Product Notes (no una única
    <table>). CONFIRMADO POR DUMP REAL DE DOM (2026-08): esta grilla no es
    una <table> con varias <tr> en un mismo <tbody> — Tourplan arma una
    <table> aparte solo para el header (CAT/DESCRIPTION/...), sin tbody
    con filas, y una <table> DISTINTA por cada nota (envuelta en
    <div class="tprow-even"/"tprow-odd">), todas hermanas dentro de un
    contenedor común (`.productnoteslistview`). Buscar la tabla cuyo
    header dice "CAT"+"DESCRIPTION" devuelve la tabla de header — que no
    tiene ningún <tbody><tr> de datos — así que hay que subir al
    contenedor común para que `tbody tr` encuentre las filas reales."""
    fin = time.time() + timeout * VELOCIDAD
    while time.time() < fin:
        contenedor = driver.execute_script("""
            var tables = Array.from(document.querySelectorAll('table'));
            for (var t of tables) {
                var head = (t.querySelector('thead') || t);
                var txt = (head.innerText || '').toUpperCase();
                if (txt.indexOf('CAT') !== -1 && txt.indexOf('DESCRIPTION') !== -1) {
                    return t.closest('.productnoteslistview') || t.parentElement;
                }
            }
            return null;
        """)
        if contenedor is not None:
            return contenedor
        time.sleep(0.4)
    return None


def esperar_filas_estables(driver, tabla, timeout=8):
    """No decidir sobre una tabla todavía cargando: esperar a que el
    número de filas se mantenga estable un rato antes de leerla."""
    fin = time.time() + timeout * VELOCIDAD
    prev = None
    estable_desde = None
    while time.time() < fin:
        rows = tabla.find_elements(By.CSS_SELECTOR, "tbody tr")
        n = len(rows)
        if n == prev:
            if estable_desde is None:
                estable_desde = time.time()
            elif time.time() - estable_desde > 0.6:
                return rows
        else:
            estable_desde = None
        prev = n
        time.sleep(0.3)
    return tabla.find_elements(By.CSS_SELECTOR, "tbody tr")


def _fila_a_nota(r):
    """Lee cada campo por su clase CSS estable (confirmado por dump real
    de DOM), no por índice posicional de <td>. Cada fila real tiene más
    columnas que las 6 documentadas — celdas de expansor y de ícono de
    lápiz intercaladas entre CAT y DESCRIPTION — así que asumir un layout
    de "6 columnas en orden fijo" leía datos de la celda equivocada.
    Devuelve None si la fila no tiene `td.tpcol-category` (ej. una fila
    de placeholder sin datos)."""
    try:
        cat = r.find_element(By.CSS_SELECTOR, "td.tpcol-category").text.strip()
    except Exception:
        return None

    def _txt(clase):
        try:
            return r.find_element(By.CSS_SELECTOR, f"td.{clase}").text.strip()
        except Exception:
            return ""

    return {
        "cat": cat,
        "description": _txt("tpcol-description"),
        "created": _txt("tpcol-created"),
        "created_by": _txt("tpcol-createdby"),
        "updated": _txt("tpcol-updated"),
        "updated_by": _txt("tpcol-updatedby"),
        "row_el": r,
    }


def buscar_nota_por_codigo(driver, tabla, codigo_nota):
    """Busca la fila cuyo CAT coincide con codigo_nota dentro del
    contenedor de Product Notes.

    CONFIRMADO POR DUMP REAL DE DOM (2026-08): esta grilla NO usa Angular
    CDK virtual scroll — no hay ningún `cdk-virtual-scroll-viewport` en la
    página, y las ~20 notas de un producto real de prueba estaban TODAS
    presentes en el DOM al mismo tiempo (confirmado por captura), sin
    necesidad de scrollear nada. El intento de scroll era una solución
    para un problema que no existía. Alcanza con leer `tbody tr`
    directamente del contenedor devuelto por `encontrar_tabla_notas()`
    (que ya lo ubica correctamente, ver ese fix).

    Devuelve (nota_o_None, vistos_ordenados) — vistos_ordenados es la
    lista de todos los códigos CAT realmente encontrados, para poder
    diagnosticar un "NO EXISTE" sin adivinar."""
    codigo_nota = codigo_nota.strip().upper()
    vistos = set()
    encontrada = None
    for r in tabla.find_elements(By.CSS_SELECTOR, "tbody tr"):
        nota = _fila_a_nota(r)
        if nota is None:
            continue
        clave = nota["cat"].upper()
        if clave:
            vistos.add(clave)
        if clave == codigo_nota and encontrada is None:
            encontrada = nota
    return encontrada, sorted(vistos)


# ── Lectura del contenido de una nota ────────────────────────────
def esperar_editor(driver, timeout=15):
    """Espera a que el editor de la nota (#noteeditorview) y su iframe
    interno terminen de cargar antes de leer el contenido. A diferencia
    de notas_srv.py, este script no ESCRIBE nada (no llama a setData() ni
    a fire('change')), así que no hace falta esperar el status 'ready' de
    CKEDITOR — alcanza con que el iframe exista para leer su DOM."""
    wait(driver, SEL["NOTE_EDITOR"], t=timeout)
    wait(driver, f"{SEL['NOTE_EDITOR']} iframe", t=timeout)
    time.sleep(1.0 * VELOCIDAD)


def leer_iframe_directo(driver):
    """Lee el HTML actual del <body contenteditable> del iframe del
    editor. Mismo camino que notas_srv.py usa para escribir (la API de
    CKEDITOR resultó poco confiable ahí) — para lectura, ir directo al
    DOM es igual de simple y evita cualquier dependencia de esa API."""
    iframe = wait(driver, f"{SEL['NOTE_EDITOR']} iframe")
    driver.switch_to.frame(iframe)
    try:
        return driver.execute_script("return document.body.innerHTML;")
    finally:
        driver.switch_to.default_content()


def cerrar_nota(driver, timeout=6):
    """
    Cierra la vista de lectura de una nota (#noteeditorview) haciendo click
    en su botón "Exit". A diferencia de Notas SRV, este script nunca hace
    SAVE, así que sin este click la nota queda abierta — la usuaria
    reportó en corrida real que, sin cerrarla, terminaban abriéndose
    varias notas juntas al pasar a la nota siguiente del mismo producto
    (el ícono de lápiz de la fila siguiente abre otra vista de lectura
    encima de la que ya estaba abierta). HTML real confirmado por la
    usuaria (inspeccionado en Tourplan):
      <button class="tpbutton tpcancel tpsecondarysystembutton">Exit</button>
    Selector por clase ("tpcancel") + texto ("EXIT") como respaldo — nunca
    por atributo `_ngcontent-*`, que cambia de build a build. Mismo botón
    y mismo criterio que `_cerrar_ultimo_tp_dialog()` en
    valorizacion_desde_madre.py (confirmado ahí en otro diálogo).

    Devuelve True si logró cerrarla (o si ya estaba cerrada), False si no
    encontró el botón o seguía abierta tras `timeout` segundos.
    """
    try:
        ya_cerrada = driver.execute_script("""
            var root = document.querySelector(arguments[0]);
            return !root || !root.offsetParent;
        """, SEL["NOTE_EDITOR"])
    except Exception:
        ya_cerrada = True
    if ya_cerrada:
        return True

    try:
        clickeado = driver.execute_script("""
            var root = document.querySelector(arguments[0]);
            if (!root) return false;
            function vis(e){ return !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length); }
            var b = root.querySelector('button.tpcancel');
            if (b && vis(b)) { b.click(); return true; }
            var btns = Array.from(root.querySelectorAll('button')).filter(vis);
            for (var btn of btns) {
                if ((btn.innerText || btn.textContent || '').trim().toUpperCase() === 'EXIT') {
                    btn.click(); return true;
                }
            }
            return false;
        """, SEL["NOTE_EDITOR"])
    except Exception:
        clickeado = False

    if not clickeado:
        return False

    for _ in range(timeout):
        time.sleep(1)
        try:
            sigue_abierto = driver.execute_script("""
                var root = document.querySelector(arguments[0]);
                return !!(root && root.offsetParent);
            """, SEL["NOTE_EDITOR"])
        except Exception:
            sigue_abierto = False
        if not sigue_abierto:
            return True
    return False


def html_a_texto_plano(contenido_html):
    """Convierte el HTML del editor (un <pre> con <br> entre líneas, para
    notas en formato Plain Text) a texto legible para el Excel: los <br>
    y </p> se vuelven saltos de línea, se sacan el resto de las etiquetas,
    y se decodifican entidades HTML (&aacute;, &amp;, etc.)."""
    if contenido_html is None:
        return None
    texto = re.sub(r"(?i)<br\s*/?>", "\n", contenido_html)
    texto = re.sub(r"(?i)</p\s*>", "\n", texto)
    texto = re.sub(r"<[^>]+>", "", texto)
    texto = html.unescape(texto)
    return texto.strip()


# ── Sheet — I/O ───────────────────────────────────────────────
def load_sheet(hoja):
    ws = conectar_sheets(SHEET_URL, hoja, CREDENTIALS_PATH, TOKEN_PATH)
    rows, columnas = cargar_sheet(ws)
    print(f"Sheet cargado: {len(rows)} fila(s) en '{hoja}'")
    return ws, rows, columnas


def update_row(ws, row_idx, columnas, estado=None, detalle=None, texto_exportado=None):
    """Escribe INMEDIATAMENTE tras cada fila — no acumular para el final."""
    valores = {}
    if estado is not None:
        valores["ESTADO"] = estado
    if detalle is not None and "DETALLE_PROCESO" in columnas:
        valores["DETALLE_PROCESO"] = detalle
    if texto_exportado is not None and "Texto_Exportado" in columnas:
        valores["Texto_Exportado"] = texto_exportado
    actualizar_fila_sheet(ws, row_idx, columnas, valores)


# ── Agrupar filas por producto (encadenamiento) ──────────────────
def _agrupar_por_producto(filas):
    """
    Agrupa filas PENDIENTE por producto (Location+Supplier+Code+
    Service_Type), preservando el orden de primera aparición. Filas del
    mismo producto pero distinto Codigo_Nota quedan juntas para reutilizar
    la búsqueda del producto y la apertura de Product Notes — no hace
    falta repetirlas por cada nota del mismo producto. Mismo patrón que
    notas_srv.py (que a su vez sigue el de los scripts de Valorización).
    """
    grupos, orden = {}, []
    for f in filas:
        key = (
            str(f.get("Location")     or "").strip().upper(),
            str(f.get("Supplier")     or "").strip().upper(),
            str(f.get("Code")         or "").strip().upper(),
            str(f.get("Service_Type") or "").strip().upper(),
        )
        if key not in grupos:
            grupos[key] = []
            orden.append(key)
        grupos[key].append(f)
    return [grupos[k] for k in orden]


def _buscar_y_abrir_producto(driver, location, supplier, code, service_type):
    """
    Busca el producto y abre Product Notes UNA VEZ por grupo. El caller
    (MAIN) reutiliza esta apertura para todas las filas del Excel que
    compartan el mismo producto, sin repetir buscar_producto() por cada
    Codigo_Nota — sí se vuelve a llamar abrir_product_notes() (sola, sin
    buscar_producto()) entre notas del mismo grupo, ver "Re-sincronización"
    en process_nota()/MAIN más abajo.
    """
    buscar_producto(driver, location, supplier, code, service_type=service_type)
    abrir_product_notes(driver)


# ── Procesamiento de una nota, sobre un producto ya abierto ──────
def process_nota(driver, row):
    """Procesa UN Codigo_Nota sobre un producto YA ABIERTO en Product
    Notes (ver _buscar_y_abrir_producto(), llamada por el caller una sola
    vez por grupo de filas del mismo producto — no repite la búsqueda del
    producto ni la apertura de Product Notes).

    Devuelve (estado, detalle, texto_exportado).
    estado: "EXPORTADA" | "NO EXISTE" | "ERROR: <detalle>"
    """
    location     = str(row.get("Location") or "").strip()
    supplier     = str(row.get("Supplier") or "").strip()
    service_type = str(row.get("Service_Type") or "").strip()
    code         = str(row.get("Code") or "").strip()
    codigo_nota  = str(row.get("Codigo_Nota") or "").strip().upper()
    row_idx      = row["__row_idx__"]

    if not (location and supplier and service_type and code and codigo_nota):
        return ("ERROR: faltan campos obligatorios "
                "(Location/Supplier/Service_Type/Code/Codigo_Nota)", None, None)

    try:
        tabla = encontrar_tabla_notas(driver)
        if tabla is None:
            ss(driver, f"tabla_notas_no_encontrada_f{row_idx}")
            return "ERROR: no cargó la tabla de Product Notes", None, None
        esperar_filas_estables(driver, tabla)
        existente, vistos = buscar_nota_por_codigo(driver, tabla, codigo_nota)
    except Exception as e:
        ss(driver, f"leer_tabla_error_f{row_idx}")
        return f"ERROR: falla leyendo tabla de notas — {e}", None, None

    if not existente:
        ss(driver, f"no_existe_{codigo_nota}_f{row_idx}")
        dump(driver, f"no_existe_{codigo_nota}_f{row_idx}")
        detalle = (f"Código '{codigo_nota}' no existe para este producto — "
                    f"códigos CAT vistos durante la búsqueda: {vistos}")
        return "NO EXISTE", detalle, None

    try:
        icono = existente["row_el"].find_element(By.CSS_SELECTOR, SEL["EDIT_ICON"])
        jc(driver, icono)
        try:
            esperar_editor(driver)
            contenido_html = leer_iframe_directo(driver)
        finally:
            # No se guarda nada — nunca se modifica el contenido, así que no
            # hay nada que perder al cerrar. Pero a diferencia de Notas SRV
            # (donde el SAVE cierra el editor solo), acá hay que cerrar la
            # nota a mano con "Exit" — si no, al pasar a la nota siguiente
            # del mismo producto se abre otra encima de esta sin haberla
            # cerrado. Se intenta cerrar tanto si la lectura salió bien como
            # si tiró una excepción (finally), para no arrastrar la vista
            # abierta a la nota siguiente en ningún caso.
            if not cerrar_nota(driver):
                print(f"    ⚠ No pude cerrar la nota con Exit — puede quedar "
                      f"abierta para la próxima")
                ss(driver, f"exit_no_cerro_f{row_idx}")
        if contenido_html is None:
            ss(driver, f"lectura_vacia_f{row_idx}")
            return "ERROR: no se pudo leer el contenido de la nota", None, None
        texto = html_a_texto_plano(contenido_html)
        detalle = (f"CREATED={existente['created']} ({existente['created_by']}) "
                   f"UPDATED={existente['updated']} ({existente['updated_by']})")
        return "EXPORTADA", detalle, texto
    except Exception as e:
        ss(driver, f"export_error_f{row_idx}")
        return f"ERROR: falla exportando nota — {e}", None, None


# ── MAIN ──────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print(f"  EXPORTAR NOTAS (Product Notes)  v{VERSION}  [{VERSION_FECHA}]")
    print("=" * 60)

    t_inicio = time.time()

    if not SHEET_URL:
        raise ValueError("No se indicó la URL del Google Sheet (TOURPLAN_SHEET_URL).")

    ws, rows, columnas = load_sheet(HOJA)
    pendientes = [r for r in rows
                  if str(r.get("ESTADO") or "").strip().upper() == "PENDIENTE"]

    print(f"Filas PENDIENTE: {len(pendientes)}")
    if not pendientes:
        print("\n⛔ Sin filas PENDIENTE. Verificá que la columna ESTADO tenga 'PENDIENTE'.")
        return

    driver = crear_driver()
    _abortado = False
    try:
        login(driver)

        grupos = _agrupar_por_producto(pendientes)
        print(f"Productos a procesar: {len(grupos)} (de {len(pendientes)} fila(s) PENDIENTE)")

        for grupo in grupos:
            chequear_abort()
            ref          = grupo[0]
            location     = str(ref.get("Location")     or "").strip()
            supplier     = str(ref.get("Supplier")     or "").strip()
            service_type = str(ref.get("Service_Type") or "").strip()
            code         = str(ref.get("Code")         or "").strip()

            print(f"\n{'=' * 60}")
            print(f"Producto: {location}/{supplier}/{code} [{service_type}]"
                  + (f"  ({len(grupo)} nota(s))" if len(grupo) > 1 else ""))
            print(f"{'=' * 60}")

            try:
                _buscar_y_abrir_producto(driver, location, supplier, code, service_type)
            except ProductoNoEncontrado as e:
                estado = f"ERROR: producto no encontrado — {e}"
                for row in grupo:
                    print(f"  Estado: {estado}")
                    update_row(ws, row["__row_idx__"], columnas,
                               estado=estado, detalle=None)
                continue
            except Exception as e:
                ss(driver, f"buscar_error_grupo_{code[:12]}")
                estado = f"ERROR: falla buscando/abriendo producto — {e}"
                for row in grupo:
                    print(f"  Estado: {estado}")
                    update_row(ws, row["__row_idx__"], columnas,
                               estado=estado, detalle=None)
                continue

            if len(grupo) > 1:
                print(f"  ℹ️  {len(grupo)} nota(s) para el mismo producto — "
                      f"reutilizando búsqueda y apertura de Product Notes")

            for i, row in enumerate(grupo):
                chequear_abort()
                row_idx = row["__row_idx__"]
                print(f"\n{'─' * 60}")
                print(f"Fila {row_idx}: Codigo_Nota={row.get('Codigo_Nota')}")

                estado, detalle, texto_exportado = "ERROR", "Error desconocido", None
                try:
                    estado, detalle, texto_exportado = process_nota(driver, row)
                except Exception:
                    estado = "ERROR"
                    detalle = traceback.format_exc(limit=3)

                print(f"  Estado: {estado}" + (f" — {detalle}" if detalle else ""))
                update_row(ws, row_idx, columnas,
                           estado=estado, detalle=detalle, texto_exportado=texto_exportado)

                # Re-sincronización: process_nota() ya se encarga de cerrar
                # la nota (botón "Exit") apenas termina de leerla — ver
                # cerrar_nota(), llamada ahí mismo — así que el caso normal
                # ya vuelve limpio a la tabla de Product Notes sin hacer
                # nada más acá. Igual que en Notas SRV, el riesgo real es
                # que una nota termine en `ERROR`: puede ser justamente
                # porque el cierre con Exit falló (cerrar_nota() devolvió
                # False) y algo quedó abierto. Si quedan más notas del mismo
                # producto, se re-sincroniza: primero liviano (reabrir
                # Product Notes, sin repetir la búsqueda del producto) y, si
                # eso también falla, una re-sincronización completa
                # (`_buscar_y_abrir_producto()`, con nueva búsqueda).
                hay_mas_notas = i < len(grupo) - 1
                if estado.startswith("ERROR") and hay_mas_notas:
                    print("  ↩ Nota en ERROR — re-sincronizando Product Notes "
                          "antes de la próxima nota del mismo producto")
                    try:
                        abrir_product_notes(driver)
                    except Exception:
                        # Si ni reabrir Product Notes funciona, probar una
                        # re-sincronización completa (buscar_producto() de
                        # nuevo) antes de rendirse con este producto.
                        try:
                            _buscar_y_abrir_producto(driver, location, supplier, code, service_type)
                        except Exception as e2:
                            ss(driver, f"resync_error_grupo_{code[:12]}")
                            estado_resync = (f"ERROR: no se pudo re-sincronizar Product "
                                             f"Notes tras la nota anterior — {e2}")
                            for row_restante in grupo[i + 1:]:
                                print(f"  Estado: {estado_resync}")
                                update_row(ws, row_restante["__row_idx__"], columnas,
                                           estado=estado_resync, detalle=None)
                            break

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
        print(f"📄 Sheet: {SHEET_URL}")
        print(f"📂 Screenshots: {SS_DIR}")

    if _abortado:
        sys.exit(ABORT_EXIT_CODE)


if __name__ == "__main__":
    main()
