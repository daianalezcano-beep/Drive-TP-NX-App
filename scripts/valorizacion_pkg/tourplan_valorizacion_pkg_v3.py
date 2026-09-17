# ============================================================
# TOURPLAN NX — VALORIZACIÓN SERVICIOS MADRE PKG  v3
# Google Colab — celda única
# ------------------------------------------------------------
#   VERSION : 3.43
#   v3.35: leer_pcm_list() -- los sleeps del scroll progresivo de Used In
#          (poll cada 150px + corte tras N intentos sin filas nuevas)
#          no escalaban con VELOCIDAD, a diferencia del resto del archivo.
#          Ahora los sleeps escalan por VELOCIDAD y el corte sube de 8 a
#          20 intentos consecutivos. (Descartado como causa real -- ver
#          v3.37, el corte real era otro.)
#   v3.36: primer intento de verificar el click que ordena Used In por
#          Type -- reemplazado por v3.37 (ordenaba A-Z, no Z-A).
#   v3.37: **causa real confirmada con 2 corridas reales**: Used In sin
#          ordenar por Type se cortaba siempre en la misma fila (~1089),
#          igual en ambas corridas pese a timings distintos -- no era un
#          problema de timing/scroll, sino un límite real que el scroll
#          nunca lograba superar. Con el header Type ordenado A-Z (un
#          solo click, v3.36) las filas FITS/GROUPS (muchísimas, letra
#          baja) quedan ANTES que Package/PCM Service (letra alta) y hay
#          que atravesarlas todas para llegar a las válidas. Fix real:
#          un SEGUNDO click sobre el mismo header invierte a Z-A (dato
#          de la usuaria, confirmado en Tourplan real) -- así Package/
#          PCM Service quedan primero, sin necesidad de escrollear más
#          allá del límite real. Cada click se verifica (compara la
#          primera fila antes/después) y se reintenta si no registró.
#   v3.38: Verificación post-SAVE de actualizar_rates_servicio_madre ya no
#          reabre el menú lateral (hamburger) con la pantalla de detalle
#          del período todavía activa -- eso dejaba colgado el backdrop
#          de navegación (.tpnavbackdrop), que interceptaba la lectura
#          siguiente: el valor quedaba bien guardado en Tourplan pero la
#          verificación reportaba "no se actualizó". Ahora cierra el
#          detalle con EXIT (_cerrar_ultimo_tp_dialog) y reabre el período
#          desde la lista subyacente antes de comparar. Mismo bug y mismo
#          fix ya confirmados en valorizacion_desde_madre.py. Sin cambios
#          en Fase 1 / búsqueda de componente / Used In.
#   v3.39: _copiar_ultimo_period (COPY DATE RANGE) usaba un selector
#          posicional fijo 'tp-dialog:nth-of-type(2)' para ubicar el
#          diálogo a llenar/cerrar. Angular nunca saca del DOM un
#          <tp-dialog> ya cerrado -- se van acumulando uno o más por
#          período procesado en la misma corrida -- así que tras el
#          primer período ese "segundo diálogo" deja de ser el recién
#          abierto y pasa a ser uno viejo abandonado. Cambiado a
#          'tp-dialog:last-of-type' (siempre el más reciente) y el
#          cierre por EXIT ahora reusa _cerrar_ultimo_tp_dialog(), con
#          diagnóstico _diag_tp_dialogs() agregado para la próxima
#          falla real. Mismo bug y mismo fix ya confirmados en
#          valorizacion_desde_madre.py. Sin cambios en Fase 1 /
#          búsqueda de componente / Used In.
#   v3.40: Columna reservada "RATE FROM D" (antes sin uso) pasa a ser
#          opcional de ENTRADA: si la fila la trae cargada, se usa como
#          base del Change Base Date del PCM en vez de RATE FROM
#          (manteniendo la regla de "siguiente día hábil" si no opera
#          esa fecha exacta). Vacía → usa RATE FROM, igual que antes.
#          RATE FROM sigue siendo la base para crear/buscar el período
#          de rates a valorizar — sin cambios ahí. Mismo cambio ya
#          hecho en valorizacion_desde_madre.py (columna nueva "FECHA
#          RECALCULATE PCM" ahí, ya que no tenía una reservada).
#   v3.41: Fix _copiar_ultimo_period (COPY DATE RANGE): elegía el
#          período fuente a copiar por posición (el más reciente de la
#          lista), ignorando su price code. Si la lista tenía períodos
#          de varios price codes mezclados, podía extender un período
#          con el price code equivocado sin ningún aviso (confirmado
#          por la usuaria en corrida real de valorizacion_desde_excel.py).
#          Ahora se lee el price code real de cada fila
#          (td.tpcol-pricecodecode) y se elige la más reciente que
#          coincida con el price code pedido. Mismo fix portado a
#          valorizacion_desde_madre.py y valorizacion_desde_excel.py.
#   v3.42: STYPE_SIDEBAR ampliado al glosario completo (12 siglas en
#          vez de solo EX/TF), sincronizado con valorizacion_desde_excel.py.
#   v3.43: Fix en buscar_producto(): al completar Location, si la fila
#          `tr.selectedRow` no aparecía a tiempo, el fallback clickeaba
#          siempre la PRIMERA fila de la tabla de sugerencias sin
#          verificar que fuera el Location exacto (mismo bug ya arreglado
#          en notas_srv.py v1.13/v1.14). Ahora busca entre las filas
#          sugeridas la que tiene una celda con texto EXACTO
#          (case-insensitive) igual al Location pedido; si ninguna
#          coincide, no clickea nada.
#   v3.44: Fix diálogo de recálculo del PCM: el checkbox "Update Exchange
#          Rates" quedaba destildado pese al click (confirmado por la
#          usuaria en corrida real, "como si no pegara"). Clickeaba su
#          <label> (#exchange-rate-panel > tp-label > label) en vez del
#          <input> -- ahora se clickea #updateexchangerates directo
#          (mismo patrón que #calculatereplaceall) y se VERIFICA que
#          quede tildado, reintentando si hace falta. Mismo fix portado
#          a valorizacion_desde_madre.py.
#   v3.45: Corrección al fix de v3.44: quedaba un segundo click sin sacar,
#          "tp-group.tpgroup-servicelinepricesgroup tp-checkbox > label",
#          que una grabación nueva de Chrome DevTools confirma (por
#          xpath) que es hijo del MISMO #exchange-rate-panel -- es decir,
#          el MISMO checkbox "Update Exchange Rates" clickeado por su
#          widget en vez de su texto. Ese segundo click volvía a
#          destildarlo justo después de que el fix de v3.44 lo dejara
#          tildado. Se saca ese click; con el click directo sobre
#          #updateexchangerates + verificación de v3.44 alcanza.
# ------------------------------------------------------------
#   ⚠ Si el log NO imprime esta misma VERSION/FECHA al arrancar,
#     estás corriendo una copia vieja: re-pegá el archivo entero.
# ============================================================

VERSION      = "3.45"
VERSION_FECHA = "2026-09-01"

# ── PASO 0: Entorno ──────────────────────────────────────────
# Este bloque verifica e instala todo lo necesario cada vez que
# se ejecuta. Colab resetea el entorno cuando pierde actividad,
# así que este chequeo siempre corre al inicio.

import os, sys, subprocess, importlib.util, shutil, time, re, json
from datetime import datetime, timedelta

_t_inicio = time.time()   # timer de la corrida completa (incluye setup)

print("🔧 Verificando entorno...\n")

# ── 0.1  Paquetes Python ─────────────────────────────────────
_PIPS_NEEDED = {
    "selenium":         "selenium",
    "openpyxl":         "openpyxl",
    "webdriver_manager":"webdriver-manager",
    "IPython":          "ipython",
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

# ── 0.2  Google Chrome (multiplataforma — ver common/chrome_bootstrap.py) ──
from common.chrome_bootstrap import find_or_prepare_chrome
# ── 0.2b  Botón Abortar de la app (ver common/abort.py) ──────
from common.abort import chequear_abort, AbortadoPorUsuario, ABORT_EXIT_CODE

CHROMIUM_BIN, ver_chrome = find_or_prepare_chrome()

# ── 0.3  ChromeDriver via webdriver-manager ──────────────────
# (se descarga automático más abajo al crear el driver)
from webdriver_manager.chrome import ChromeDriverManager

# ── 0.4  Imports del resto del script ───────────────────────
from IPython.display import display, Image as IPyImage
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

# ── Config — via variables de entorno (con default = valor original) ──
USERNAME   = os.environ.get("TOURPLAN_USERNAME", "poner minusculas")
PASSWORD   = os.environ.get("TOURPLAN_PASSWORD", "password")
BASE_URL   = os.environ.get("TOURPLAN_BASE_URL", "https://tourplannx.eurotur.com.ar/TourplanNX_Test")
EXCEL_PATH = os.environ.get("TOURPLAN_EXCEL_PATH", "/content/tourplan_valorizacion_pkg.xlsx")
SHEET      = os.environ.get("TOURPLAN_HOJA", "PRODUCTOS")
SS_DIR     = os.environ.get("TOURPLAN_SS_DIR", "/content/screenshots")
os.makedirs(SS_DIR, exist_ok=True)

# ── Modo de ejecución ─────────────────────────────────────────
# "LEER"    → solo Fase 1: navega PCMs, cambia fecha, lee markup, guarda en Excel
# "APLICAR" → solo Fase 2: lee PCM_Detail y aplica rates en servicios madre
# "COMPLETO"→ Fase 1 + Fase 2 en la misma sesión
MODO = os.environ.get("TOURPLAN_MODO", "COMPLETO")

# Multiplicador de tiempos de espera.
# 1.0 = entorno de prueba  |  1.5 = producción (respuestas más lentas)
VELOCIDAD = float(os.environ.get("TOURPLAN_VELOCIDAD", "1.5"))

# Límite de PCMs a procesar en Fase 2 (0 = sin límite). Para iterar rápido
# sobre el problema de escritura: poné MODO="APLICAR" y FASE2_LIMIT=1 y corre
# SOLO la Fase 2 del primer servicio madre (≈2-3 min en vez de ~38 min).
# Requiere que el Excel (EXCEL_PATH) ya tenga la Fase 1 cargada (subilo a Colab).
FASE2_LIMIT = 0
# Si True, la Fase 2 reprocesa TODAS las filas con COD DESTINO sin importar el
# ESTADO F2 (sirve para re-testear con un Excel cuya Fase 2 ya falló/corrió, sin
# tener que re-correr la Fase 1 ni editar el estado a mano).
FASE2_REAPLICAR_TODO = False

# Price Code por defecto para el período de rates del servicio madre.
# El script intenta leerlo del PCM (PCM DETAILS); si no lo encuentra usa
# este valor. Sin un Price Code el período queda 'Unassigned' y Tourplan
# NO persiste los costos (los valores se leen 0.0 al recargar).
PRICE_CODE_DEFAULT = "TR"

# Columnas Excel (base 1)
C = {
    "location":          1,   # A
    "supplier":          2,   # B
    "cod_origen":        3,   # C
    "service_type":      4,   # D  ← tipo de servicio (TK, EX, TF, etc.)
    "rate_from":         5,   # E
    "rate_to":           6,   # F
    "price_code":        7,   # G  ← price code a usar (TR, 34, ... o ALL)
    "mostrar_capturas":  8,   # H  ← mostrar capturas inline (SI/NO, default NO)
    "cod_destino":       9,   # I  override servicio madre
    "rate_from_d":      10,   # J
    "rate_to_d":        11,   # K
    "pcm_ref":          12,   # L
    "dias_op":          13,   # M  días operación leídos
    "estado":           14,   # N
    "error_msg":        15,   # O
    "timestamp":        16,   # P
}
# Columnas dinámicas de tarifas empiezan en Q (17)
RATES_START = 17

_ss_n = [0]

# Mostrar las capturas inline en Colab. Se lee de la columna MOSTRAR CAPTURAS
# del Excel (default NO). Aunque sea NO, las capturas SÍ se guardan en disco
# (para depurar errores); sólo se evita renderizarlas en el notebook, que es
# lo que hace lento el output con ~370 imágenes.
MOSTRAR_CAPTURAS = False

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
    """JavaScript click — único método confiable en Angular."""
    driver.execute_script("arguments[0].click();", el)

def set_val(driver, el, value):
    """Setea valor en input Angular disparando eventos nativos."""
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
        raise RuntimeError(
            "No se encontró ningún Chrome funcional. "
            "Verificá que google-chrome-stable esté instalado en Colab.")

    log_path = os.path.join(tempfile.gettempdir(), "chromedriver.log")
    try:
        drv_path = ChromeDriverManager().install()
        print(f"Chromedriver: {drv_path}")
        svc = Service(executable_path=drv_path, log_output=log_path)
        d = webdriver.Chrome(service=svc, options=opts)
        print("✅ Driver iniciado")
        return d
    except Exception as e:
        # Mostrar log del chromedriver para diagnóstico
        if os.path.exists(log_path):
            with open(log_path) as f:
                tail = f.read()[-3000:]
            print(f"\n--- ChromeDriver log ---\n{tail}\n---")
        raise RuntimeError(
            f"No se pudo iniciar Chrome.\n"
            f"Binary: {CHROMIUM_BIN} ({ver_chrome})\nError: {e}")

# ── Login ─────────────────────────────────────────────────────
def login(driver):
    print("🔐 Login...")
    driver.get(f"{BASE_URL}/#/login")
    time.sleep(6 * VELOCIDAD)
    ss(driver, "login_page")

    def _campos_visibles():
        """Devuelve (input_usuario, input_password) eligiendo los que están
        realmente visibles/interactables. La SPA Angular puede tener inputs
        ocultos o todavía no renderizados → send_keys directo tira
        ElementNotInteractableException."""
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

    # Esperar a que el formulario esté renderizado e interactable (hasta ~30s)
    u_el = p_el = None
    for intento in range(15):
        u_el, p_el = _campos_visibles()
        if u_el and p_el:
            break
        time.sleep(2 * VELOCIDAD)
    if not (u_el and p_el):
        ss(driver, "login_sin_campos")
        raise Exception("No aparecieron los campos de login (usuario/password)")

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
            # último recurso: setear por JS + disparar eventos Angular
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

    # Botón login: clickear el visible; si no, Enter en el password
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
        try:
            p_el.send_keys(Keys.ENTER)
        except Exception:
            pass
    time.sleep(8 * VELOCIDAD)
    assert "login" not in driver.current_url.lower(), "Login falló"
    ss(driver, "post_login")
    print("✅ Login OK")

# ── Logout ────────────────────────────────────────────────────
def logout(driver):
    """
    Cierra ventanas secundarias y hace logout real.
    El logout solo se considera OK si después del click aparece el
    formulario de login (input password) o la URL vuelve a #/login.
    Encontrar el texto 'logged in as' NO es prueba de logout: significa
    que la sesión sigue abierta.
    """
    print("\n🔒 Finalización: cerrando ventanas y haciendo logout...")
    try:
        principal = driver.window_handles[0]
        for h in driver.window_handles[1:]:
            try:
                driver.switch_to.window(h); driver.close()
            except Exception: pass
        driver.switch_to.window(principal)
    except Exception: pass

    def _click_item(regex):
        """Clickea el elemento visible cuyo texto matchee el regex.
        Busca en TODO el documento (la opción Log out vive en la barra
        superior junto al usuario, no en el <nav> hamburger).
        Prefiere el match más interno/corto (evita clickear contenedores)."""
        return driver.execute_script("""
            var rx = new RegExp(arguments[0], 'i');
            var els = Array.from(document.querySelectorAll(
                'li, label, a, button, span, div'));
            var best = null;
            for (var el of els){
                if (!el.offsetParent) continue;          // no visible
                var t = (el.innerText || '').trim();
                if (!t || t.length > 40 || !rx.test(t)) continue;
                if (!best || t.length <= (best.innerText||'').trim().length)
                    best = el;                            // más corto/interno
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

        # Flujo grabado (Chrome DevTools recorder):
        #   1. click #openUserPanel            (panel de usuario en el header)
        #   2. click div.panelHeader button    (botón "Logout" del panel)
        #      → navega a #/login
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

        # Fallback: búsqueda por texto en todo el documento
        rx_logout = r"^(log\s?out|sign\s?out|cerrar sesi)"
        if not clicked:
            clicked = _click_item(rx_logout)
        if not clicked:
            padre = _click_item(r"logged in as")
            time.sleep(2 * VELOCIDAD)
            ss(driver, "logout_submenu")
            if padre:
                clicked = _click_item(rx_logout)

        if not clicked:
            # Diagnóstico: qué textos visibles contienen 'log' o el usuario
            candidatos = driver.execute_script("""
                var out = [];
                for (var el of document.querySelectorAll('*')){
                    if (!el.offsetParent || el.children.length > 0) continue;
                    var t = (el.innerText || '').trim();
                    if (t && t.length < 40 && /log|user|usuario/i.test(t)) out.push(t);
                }
                return out.slice(0, 25);
            """)
            print(f"      Textos visibles con 'log/user': {candidatos}")

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

# ── Navegación (selectores reales de grabaciones) ─────────────

def hamburger(driver):
    """nav img = ícono hamburger en cualquier vista."""
    img = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "nav img")))
    jc(driver, img)
    time.sleep(2.5)
    # Diagnóstico: mostrar items visibles del menú
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
    """
    Clickea un ítem del menú lateral.
    Si n_or_text es int → usa posición (li[n] > div > div).
    Si n_or_text es str → busca por texto (case-insensitive).
    """
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
    """Clickea sub-item por texto dentro de li.tpnavmenuopen."""
    xpath = (f"//li[contains(@class,'tpnavmenuopen')]"
             f"//label[contains(translate(normalize-space(text()),"
             f"'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),"
             f"'{text.upper()}')]")
    el = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH, xpath)))
    jc(driver, el)
    time.sleep(2 * VELOCIDAD)

def submenu_n(driver, n, timeout=8):
    """Clickea sub-item N dentro de li.tpnavmenuopen."""
    xpath = f"//li[contains(@class,'tpnavmenuopen')]//ul/li[{n}]//label"
    el = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH, xpath)))
    jc(driver, el)
    time.sleep(2)

# ── Fechas ────────────────────────────────────────────────────
MESES_ES = {"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
            "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}

def _try_float(s):
    """Convierte string a float limpiando $, comas, espacios. Retorna None si falla."""
    try:
        return float(str(s).replace(",", "").replace("$", "").replace(" ", "").strip())
    except: return None

def parsear_fecha(txt):
    if not txt: return None
    if isinstance(txt, datetime): return txt
    txt = str(txt).strip()
    # ISO format: "YYYY-MM-DD" or "YYYY-MM-DD HH:MM:SS" (openpyxl output)
    if len(txt) >= 10 and txt[4] == "-" and txt[7] == "-":
        try:
            return datetime.fromisoformat(txt[:10])
        except: pass
    # dd/Mon/yyyy  o  dd/mm/yyyy  o  dd/mm/yy
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
    """dd/Mon/yyyy para inputs de Tourplan (4-digit year)."""
    meses = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
             7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
    return f"{dt.day:02d}/{meses[dt.month]}/{dt.year}"

def fmt_xl(dt):
    """dd/Mon/yyyy para Excel."""
    meses = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
             7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
    return f"{dt.day:02d}/{meses[dt.month]}/{dt.year}"

DIAS_SEMANA = {0:"MON",1:"TUE",2:"WED",3:"THU",4:"FRI",5:"SAT",6:"SUN"}
DIAS_IDX    = {v:k for k,v in DIAS_SEMANA.items()}

def siguiente_dia_operacion(dias_activos, desde=None):
    """
    Retorna el próximo datetime donde el día de la semana esté en dias_activos.
    dias_activos: set de strings como {'MON','WED','FRI'}
    desde: datetime base (default=hoy)
    """
    if not dias_activos:
        dias_activos = set(DIAS_SEMANA.values())  # todos los días si no se pudo leer
    base = desde or datetime.now()
    for delta in range(8):
        candidato = base + timedelta(days=delta)
        if DIAS_SEMANA[candidato.weekday()] in dias_activos:
            return candidato
    return base  # fallback

# ── Buscar producto (flujo real de grabaciones) ───────────────
class ProductoNoEncontrado(Exception):
    """El código buscado no apareció en los resultados de Product Search."""
    pass

def buscar_producto(driver, location, supplier, codigo, service_type=None, handle_origen=None):
    """
    Flujo correcto según capturas reales:
    1. driver.get(#/product)
    2. Click lupa (abre Product Search modal con dos tabs: SELECTION / RESULTS)
    3. En el sidebar izquierdo del modal, click en el service type (TK, EX, etc.)
    4. Llenar Location, Supplier, Code en el formulario de la derecha
    5. Click SEARCH → tab RESULTS se activa con filas de resultados
    6. Click en la fila del resultado correcto → carga el producto en detalle
    7. Verificar que el menú hamburger tiene UTILITIES/RATES (contexto correcto)
    """
    st_upper = (service_type or "").strip().upper()
    st_str   = f"/{st_upper}" if st_upper else ""
    print(f"\n  📦 Buscando: {location}/{supplier}/{codigo}{st_str}")

    if handle_origen:
        driver.switch_to.window(handle_origen)

    # Pasar por #/home antes de #/product: si la ventana ya está en
    # #/product (con un producto cargado de una búsqueda anterior),
    # driver.get() con la misma URL hash NO recarga nada y la lupa abre
    # un popover de búsqueda rápida (Supplier/Product/Description/Comment)
    # en vez del modal completo con SELECTION + sidebar de service types.
    driver.get(f"{BASE_URL}/#/home")
    time.sleep(2 * VELOCIDAD)
    driver.get(f"{BASE_URL}/#/product")
    time.sleep(5 * VELOCIDAD)
    ss(driver, f"ps_inicial_{codigo[:10]}")

    # ── Abrir modal de búsqueda ──────────────────────────────
    lupa = wait(driver, "#searchWrapper li:nth-of-type(2) button")
    jc(driver, lupa)
    time.sleep(3 * VELOCIDAD)
    ss(driver, f"ps_modal_{codigo[:10]}")

    # ── Asegurar tab SELECTION activo ─────────────────────────
    # El modal recuerda el tab usado la última vez: tras una búsqueda previa
    # puede abrirse en RESULTS y el formulario (parameters1) no existe,
    # con lo que todos los waits de Location/Supplier/Code expiran.
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
        time.sleep(2)
        ss(driver, f"ps_tab_selection_{codigo[:10]}")

    # ── Service Type en sidebar izquierdo del modal ──────────
    if st_upper:
        try:
            # El sidebar tiene ítems tipo "TK - TICKETS", "Z*NO USAR* - TK -Tickets",
            # "EX -Excursion", etc. Búsqueda y click en UNA sola llamada JS:
            # con find_elements + click separados el elemento puede quedar stale
            # entre la lectura y el click (el sidebar se re-renderiza). El código
            # se acepta como palabra completa delimitada por inicio/fin, espacios,
            # guiones (-, –, —), '/' o paréntesis — nunca substring suelto, que
            # puede matchear por accidente dentro de otra categoría (confirmado en
            # Flag as Deleted: 'TR' tildaba "Hotel Extras" porque "Extras"
            # contiene "tr").
            # Traducir sigla al prefijo numérico del sidebar. Glosario confirmado
            # por la usuaria (2026-08), ver skill buscando-productos-en-tourplan/
            # references/service-types.md. TA y PR no están confirmados: si
            # aparecen, cae al fallback por texto/regex de acá abajo — no asumir
            # un número por patrón/orden alfabético.
            STYPE_SIDEBAR = {
                "HT": "01", "HX": "02", "TF": "03", "EX": "04", "ML": "05",
                "RT": "06", "CR": "07", "FT": "08", "OC": "09", "LN": "10",
                "LP": "11", "MS": "12",
            }
            # Siglas agrupadas por Tourplan bajo "Z*NO USAR*" (organización interna
            # de la empresa, confirmado por la usuaria 2026-08-18 — no significa
            # deprecadas). Varias comparten una palabra con una categoría real (ej.
            # TN="Transfer Non-Accom" vs TF="Transfer") — si `st_upper` no es una de
            # estas 12, se excluyen esos ítems para no matchear la real dentro de
            # uno "no usar" (o viceversa).
            Z_NO_USAR_SIGLAS = {"GA", "TK", "TN", "CH", "CM", "CO", "EN", "GU", "PJ", "ST", "TR", "TI"}
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
                        if (!li.offsetParent) continue;            // no visible
                        var t = (li.textContent || '').trim();
                        // ítems del sidebar son cortos; descarta contenedores anidados
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
                time.sleep(1.5)  # el sidebar puede renderizar tarde
            if resultado and resultado.get("ok"):
                time.sleep(1.5)
                print(f"    Service type '{st_upper}' seleccionado: {resultado.get('texto')}")
            else:
                print(f"    ⚠ Service type '{st_upper}' no encontrado en sidebar — buscando sin filtro")
                print(f"      Ítems visibles del sidebar: {(resultado or {}).get('textos', [])}")
        except Exception as e:
            print(f"    ⚠ Error seleccionando service type: {e}")

    # ── Location ─────────────────────────────────────────────
    if location:
        try:
            inp_loc = wait(driver, "div.parameters1 li:nth-of-type(1) input")
            inp_loc.click(); time.sleep(0.3)
            set_val(driver, inp_loc, location)
            time.sleep(1.5)
            row = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "tr.selectedRow > td.description")))
            jc(driver, row); time.sleep(1)
            print(f"    Location: {location} OK")
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

    # ── Supplier ─────────────────────────────────────────────
    if supplier:
        try:
            inp_sup = wait(driver, "div.parameters1 li:nth-of-type(2) input")
            inp_sup.click(); time.sleep(0.3)
            set_val(driver, inp_sup, supplier)
            time.sleep(1.5)
            row_sup = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located(
                    (By.XPATH,
                     f"//div[contains(@class,'parameters1')]//td[normalize-space(text())='{supplier.upper()}']")))
            jc(driver, row_sup); time.sleep(1)
            print(f"    Supplier: {supplier} OK")
        except:
            try:
                inp_sup2 = driver.find_element(By.CSS_SELECTOR,
                    "div.parameters1 li:nth-of-type(2) input")
                inp_sup2.send_keys(Keys.TAB); time.sleep(1)
                print(f"    Supplier: {supplier} (TAB)")
            except: pass

    # ── Code ─────────────────────────────────────────────────
    inp_cod = wait(driver, "div.parameters1 li:nth-of-type(3) input")
    inp_cod.click(); time.sleep(0.3)
    set_val(driver, inp_cod, codigo)
    time.sleep(0.5)
    print(f"    Code: {codigo}")
    ss(driver, f"ps_modal_lleno_{codigo[:10]}")

    # ── SEARCH ───────────────────────────────────────────────
    btn_search = wait(driver, "#productSearchFilter li:nth-of-type(4) button")
    jc(driver, btn_search)
    time.sleep(6 * VELOCIDAD)
    ss(driver, f"ps_resultados_{codigo[:10]}")

    # ── Verificar si quedó un solo resultado y cargó directo ─
    # La URL en Tourplan NX SIEMPRE es #/product (sin ID).
    # Detectamos "en detalle" por el menú hamburger (tiene UTILITIES/RATES)
    # O por el input que muestra el código completo del producto.

    def _en_contexto_producto():
        """Abre hamburger y verifica presencia de UTILITIES o RATES en el menú."""
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
            # Cerrar hamburger
            driver.execute_script("arguments[0].click();", img)
            time.sleep(1)
            return ok, items
        except Exception as e:
            return False, []

    # ── Click en fila de resultado (todo en JS para evitar stale) ─
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

        // Buscar fila coincidente (con service type si fue indicado)
        var match = null;
        for(var tr of rows){{ if(matchRow(tr)){{ match=tr; break; }} }}
        // Fallback: solo por código (el service type puede mostrarse distinto en la grilla)
        if(!match){{
            for(var tr of rows){{
                if((tr.innerText||'').toUpperCase().includes(cod)){{ match=tr; break; }}
            }}
        }}
        if(!match) return null;

        // Click en la descripción (columna 5 = índice 4)
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
        # Asumir cargado y continuar — el error real aparecerá en UTILITIES
        return

    # Sin resultado: diagnóstico
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

# ── Leer días de operación (STAY MUST START ON) ───────────────
def leer_dias_operacion(driver):
    """
    Lee qué días opera el producto (STAY MUST START ON checkboxes MON-SUN).
    Navega al submenu 'Policies' via hamburger (NO es un tab sino un item de menú).
    Retorna set de strings: {'MON','WED','FRI'} etc.
    Si no puede leerlos retorna todos los días.
    """
    DIAS_VALIDOS = {"MON","TUE","WED","THU","FRI","SAT","SUN"}
    try:
        # POLICIES es un label nested dentro de PRODUCT DETAILS (li.tpnavmenuopen).
        # NO usar menu_item: su XPath busca div/div y matchea PRODUCT DETAILS primero
        # (colapsa el menú en lugar de navegar). Buscar el label por texto directamente.
        hamburger(driver)
        xpath_pol = ("//nav//label[contains("
                     "translate(normalize-space(text()),"
                     "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),"
                     "'POLICIES')]")
        pol_lbl = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, xpath_pol)))
        jc(driver, pol_lbl)
        time.sleep(3 * VELOCIDAD)

        dias = driver.execute_script("""
            var result = {};
            var DIAS = ['MON','TUE','WED','THU','FRI','SAT','SUN'];
            // Selector confirmado por DevTools inspector:
            // #tabs-policies .tpgroup-startson ul li tp-checkbox label i
            var container = document.querySelector('#tabs-policies .tpgroup-startson') ||
                            document.querySelector('.tpgroup-startson');
            if (!container) return result;
            container.querySelectorAll('li').forEach(function(li) {
                // El label del día está en tp-label adyacente al tp-checkbox
                var tpLbl = li.querySelector('tp-label label') ||
                            li.querySelector('label.tplabel');
                if (!tpLbl) return;
                var dia = tpLbl.innerText.trim().toUpperCase().slice(0, 3);
                if (!DIAS.includes(dia)) return;
                // Estado: primero input nativo, fallback ng-reflect-model del tp-checkbox
                var inp   = li.querySelector('input[type=checkbox]');
                var tpChk = li.querySelector('tp-checkbox');
                if (inp) {
                    result[dia] = inp.checked;
                } else if (tpChk) {
                    var m = tpChk.getAttribute('ng-reflect-model');
                    result[dia] = (m === 'true' || m === '1');
                }
            });
            return result;
        """)

        activos = {k for k, v in (dias or {}).items()
                   if v and k.upper() in DIAS_VALIDOS}
        if not activos:
            print("    ⚠ No se leyeron días operación (resultado vacío) → asumiendo todos")
            dump(driver, "policies_dias_no_leidos")
            activos = set(DIAS_SEMANA.values())
        print(f"    Días operación: {sorted(activos)}")
        return activos
    except Exception as e:
        print(f"    ⚠ No pude leer días de operación: {e} → asumiendo todos")
        return set(DIAS_SEMANA.values())

# ── USED IN ───────────────────────────────────────────────────
def ir_a_used_in(driver):
    """
    Product hamburger → UTILITIES → Used In
    Usa búsqueda por texto en lugar de posición fija (el número varía por contexto).
    """
    hamburger(driver)
    ss(driver, "hamburger_product")
    menu_item(driver, "UTILITIES")   # busca por texto, no posición
    ss(driver, "utilities_expandido")
    submenu_text(driver, "Used In")
    time.sleep(5 * VELOCIDAD)
    ss(driver, "used_in_tabla")

def leer_pcm_list(driver):
    """
    Lee la tabla Used In y retorna lista de PCMs válidos.
    Filtros:
    - tipo = "Package" o "PCM Service" (columna UsageTypeLabel)
    - nombre debe empezar con "PKG-"
    - NO deleted/flagged/cancelled/inactive
    - Travel Date con antigüedad menor a 365 días (no se procesan viajes de más de 1 año atrás)

    Antes de tocar la columna Type, intenta una PASADA RÁPIDA (paso 0):
    ordena por la columna "Date" (los PCM Package Header la traen siempre
    en blanco, confirmado por la usuaria con captura real) y revisa si las
    filas con Date en blanco ya cumplen estos filtros. Si es así, evita el
    escaneo completo con scroll. Es best-effort: si el ordenamiento falla,
    no hay filas con Date en blanco, o esas filas no pasan los filtros, se
    sigue exactamente con el procedimiento de siempre (ordenar por Type +
    escrollear toda la lista), sin cambios.
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
                if(firstRow){
                    headers = Array.from(firstRow.querySelectorAll('th,td')).map(h=>h.innerText.trim());
                }
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

    def _procesar_fila_pcm(info, el_fila):
        """Aplica los filtros de PCM válido (Package/PCM Service + PKG- +
        no eliminado + Travel Date <= 365 días) a una fila ya leída de la
        grilla Used In. Retorna el dict a agregar a pcm_list, o None si la
        fila no pasa. Compartida por la pasada rápida (Date) y la pasada
        de siempre (Type + scroll) para que ambas den el mismo resultado."""
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
        fila_txt_up = info.get("_fila_txt","").upper()
        nombre_up   = nombre.upper()

        if any(x in nombre_up for x in ["FLAG AS DELETED","FLAGGED","DELETED","CANCELLED","CANCELED","INACTIVE"]):
            return None
        if any(x in fila_txt_up for x in ["FLAG AS DELETED","FLAGGED"]):
            return None
        if any(x in tipo.upper() for x in ["DELETE","FLAG","CANCEL","INACTIVE"]):
            return None

        _TIPOS_VALIDOS = {"PACKAGE", "PCM SERVICE"}
        tipo_up = tipo.upper().strip()
        if tipo_up and not any(t in tipo_up for t in _TIPOS_VALIDOS):
            return None
        if not nombre_limpio.startswith("PKG-"):
            return None

        travel_date_str = (info.get("Travel Date","") or info.get("TRAVEL DATE","")
                           or info.get("travel date","") or "")
        f_travel = parsear_fecha(travel_date_str)
        if f_travel and (datetime.now() - f_travel).days > 365:
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
                resultado = _procesar_fila_pcm(info, el_fila)
                if resultado:
                    pcm_list_rapido.append(resultado)

            if pcm_list_rapido:
                print(f"    ✅ Pasada rápida (orden por Date): "
                      f"{len(pcm_list_rapido)} PCM válido(s) sin necesidad "
                      f"de escrollear toda la lista: "
                      f"{[p['nombre'] for p in pcm_list_rapido]}")
                return pcm_list_rapido

        print("    ⚠ Pasada rápida (orden por Date) no encontró un PCM "
              "válido en las filas con Date en blanco — sigo con el "
              "procedimiento completo (Type + scroll)")

    # ── Paso 1 (fallback, sin cambios): ordenar por tipo (columna Type)
    # DESCENDENTE (Z-A) para que las filas
    # "Package"/"PCM Service" (empiezan con letra alta) queden AL PRINCIPIO
    # de la lista, antes que los miles de filas "FITS"/"GROUPS" (letra más
    # baja) que hay en componentes reales con mucho volumen -- confirmado
    # con la usuaria que en Tourplan un PRIMER click ordena A-Z (FITS
    # primero) y un SEGUNDO click sobre el mismo encabezado invierte a Z-A.
    # Sin esto, encontrar los PCMs válidos dependía de escrollear más allá
    # de un límite real (~1089 filas, reproducido en 2 corridas reales
    # distintas) que el scroll progresivo nunca llegó a superar.
    def _primer_tipo_visible():
        return driver.execute_script("""
            var tr = document.querySelector('#usedin table tbody tr') ||
                     document.querySelector('table tbody tr');
            return tr ? tr.innerText.trim().slice(0, 60) : '';
        """)

    def _click_header_type():
        th = wait(driver, "th.tpcol-UsageTypeLabel > span > span", t=5)
        jc(driver, th)
        time.sleep(2 * VELOCIDAD)
        return _primer_tipo_visible()

    _estado_orden = _primer_tipo_visible()   # sin ordenar todavía
    _clicks_ok = 0
    for _click_num in range(1, 3):   # click 1 = A-Z, click 2 = Z-A
        _logrado = False
        for _intento in range(2):    # reintenta el MISMO click si no registró
            try:
                _nuevo = _click_header_type()
                if _nuevo and _nuevo != _estado_orden:
                    _estado_orden = _nuevo
                    _logrado = True
                    break
                print(f"    ⚠ Click {_click_num} sobre el header Type "
                      f"(intento {_intento+1}) no cambió el orden visible")
            except Exception as _e_sort:
                print(f"    ⚠ Click {_click_num} sobre el header Type "
                      f"(intento {_intento+1}) falló: {_e_sort}")
        if not _logrado:
            break
        _clicks_ok += 1

    if _clicks_ok == 2:
        print(f"    ✅ Used In ordenado Z-A por columna Type "
              f"(primera fila ahora: '{_estado_orden}')")
        ss(driver, "used_in_sorted_ok")
    else:
        print(f"    ⚠ Solo se confirmaron {_clicks_ok}/2 clicks del header Type — "
              f"el orden puede haber quedado A-Z o sin ordenar. La búsqueda de "
              f"PCMs válidos puede depender más del scroll completo")
        ss(driver, "used_in_sorted_fallo")

    # Scroll progresivo para cargar todas las filas (virtual scroll)
    # Clave de dedup: tipo|nombre (más estable que _fila_txt completo)
    filas_vistas = {}

    def _fila_key(f):
        tipo_v = next((f[k] for k in f if "type" in k.lower() and k != "_fila_txt"), "")
        name_v = next((f[k] for k in f if any(x in k.lower()
                       for x in ["pcmname","bookingpcm","pcm name","pcm ref"]) and k != "_fila_txt"), "")
        if not name_v:
            name_v = next((v for k, v in f.items()
                           if re.match(r'[A-Z]{2,}[-]', str(v)) and k != "_fila_txt"), "")
        # Fallback: usar texto completo para que filas FITS distintas no colapsen
        if not name_v:
            name_v = f.get("_fila_txt", "")
        return f"{tipo_v}|{name_v}"

    def _get_scroll_container():
        return driver.execute_script("""
            // 1. Angular CDK virtual scroll viewport (más común en Angular)
            var cdk = document.querySelector('cdk-virtual-scroll-viewport');
            if(cdk) return cdk;

            // 2. Contenedores específicos de Tourplan
            var candidates = ['#usedin','tp-usedin','[class*="usedin"]','[class*="used-in"]'];
            for(var s of candidates){
                var el = document.querySelector(s);
                if(el && el.scrollHeight > el.clientHeight) return el;
            }

            // 3. Subir desde la tabla buscando el primer ancestor scrollable con altura real
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

            // 4. Último recurso: body / scrollingElement
            return document.scrollingElement || document.body;
        """)

    def _scroll_to(pos, container):
        if container:
            driver.execute_script("arguments[0].scrollTop = arguments[1];", container, pos)
        driver.execute_script("window.scrollTo(0, arguments[0]);", pos)

    def _get_max_scroll(container):
        return driver.execute_script("""
            var el = arguments[0];
            return el ? el.scrollHeight : document.body.scrollHeight;
        """, container)

    scroll_container = _get_scroll_container()
    scroll_info = driver.execute_script("""
        var el = arguments[0];
        if(!el) return {scrollH:0, clientH:0, tag:'null'};
        return {scrollH: el.scrollHeight, clientH: el.clientHeight,
                tag: el.tagName + (el.id?'#'+el.id:'') + (el.className?'.'+el.className.split(' ')[0]:'')};
    """, scroll_container)
    print(f"    Scroll container: {scroll_info}")

    # sin_cambio y el sleep por paso escalan con VELOCIDAD: en producción,
    # con listas Used In muy largas, Angular puede tardar más en re-renderizar
    # el siguiente batch de filas del virtual scroll de lo que tarda en test
    # -- un corte fijo de "8 intentos sin filas nuevas" con sleep sin escalar
    # puede frenar el scroll antes de llegar a filas reales más abajo (visto
    # en un componente real con muchísimas líneas en Used In, sin ningún PCM
    # Package Header detectado a pesar de haberlos).
    scroll_step        = 150
    scroll_pos         = 0
    sin_cambio         = 0
    SIN_CAMBIO_LIMITE  = 20
    max_scrolls        = 1200   # 1200 × 150px = 180 000px
    for _ in range(max_scrolls):
        batch = _leer_filas_visibles()
        nuevas = 0
        for f in batch:
            k = _fila_key(f)
            if k and k not in filas_vistas:
                filas_vistas[k] = f
                nuevas += 1
        if nuevas == 0:
            sin_cambio += 1
            if sin_cambio >= SIN_CAMBIO_LIMITE:
                break
        else:
            sin_cambio = 0
        scroll_pos += scroll_step
        _scroll_to(scroll_pos, scroll_container)
        time.sleep(0.25 * VELOCIDAD)

    print(f"    Total filas únicas tras scroll: {len(filas_vistas)}")

    # Reset scroll al inicio
    _scroll_to(0, scroll_container)
    time.sleep(0.5 * VELOCIDAD)

    filas_info = list(filas_vistas.values())
    print(f"    Filas en Used In: {len(filas_info)}")
    if not filas_info:
        dump(driver, "usedin_sin_filas")

    pcm_list = []
    filas_dom = driver.find_elements(By.CSS_SELECTOR, "#usedin table tbody tr")
    if not filas_dom:
        filas_dom = driver.find_elements(By.XPATH, "//table//tbody//tr")

    for i, info in enumerate(filas_info):
        print(f"      [{i}] {info.get('_fila_txt','')[:100]}")
        resultado = _procesar_fila_pcm(info, filas_dom[i] if i < len(filas_dom) else None)
        if resultado:
            pcm_list.append(resultado)
        else:
            print(f"      → SKIP (no pasa filtros de PCM válido)")

    print(f"  📋 PCMs válidos: {[p['nombre'] for p in pcm_list]}")
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
    el click no cambia nada, esta función no rompe nada ni bloquea: el
    scroll de siempre (_scan_from() más abajo) sigue siendo el
    procedimiento de fallback, sin cambios."""
    def _primera_fila():
        return driver.execute_script("""
            var tr = document.querySelector('#usedin table tbody tr') ||
                     document.querySelector('table tbody tr');
            return tr ? tr.innerText.trim().slice(0, 80) : '';
        """)

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
    """
    Clickea la celda tpcol-bookingpcmname → abre PCM en nueva ventana.
    Retorna (handle_pcm, scroll_pos_encontrado).
    Scrollea el contenedor virtual hasta encontrar la fila por texto.
    start_scroll: posición inicial de scroll para optimizar búsqueda.
    """
    nombre = pcm_info["nombre"]
    handles_antes = set(driver.window_handles)

    _ordenar_used_in_por_date(driver)

    # Obtener contenedor scrolleable de la tabla Used In (CDK-aware)
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
        """Intenta encontrar y clickear la celda del PCM en el DOM actual."""
        return driver.execute_script(f"""
            var nombre = {json.dumps(nombre)};
            var celdas = Array.from(document.querySelectorAll('td'));
            for(var td of celdas){{
                if(td.innerText.trim() === nombre){{
                    td.click();
                    return true;
                }}
            }}
            return false;
        """)

    def _scan_from(start_pos, max_steps=2000):
        """Scrollea desde start_pos hacia adelante buscando el PCM. Retorna pos o -1."""
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

    # Mismo step que ya usa la función que arma la lista inicial de PCMs
    # (scroll_step=150, línea ~1051) — un step de 1500px acá (10x más
    # grande) podía saltear filas en listas de Used In grandes (virtual
    # scroll: solo lo cercano al viewport está en el DOM en cada momento),
    # confirmado en corrida real de valorizacion_desde_madre.py: un PCM con
    # "cientos" de filas en Used In se encontraba bien en el escaneo
    # inicial pero después "No encontré fila para PCM" acá, con el mismo
    # nombre exacto. max_steps sube en la misma proporción para cubrir la
    # misma distancia total que antes.
    step = 150
    found_pos = -1

    # Primera pasada: desde start_scroll
    if start_scroll > step:
        found_pos = _scan_from(start_scroll - step)

    # Fallback: desde el inicio si no encontró
    if found_pos < 0:
        print(f"    ↩ Reintentando desde pos=0 para {nombre}")
        found_pos = _scan_from(0)

    if found_pos < 0:
        raise Exception(f"No encontré fila para PCM {nombre}")

    time.sleep(6 * VELOCIDAD)

    nuevas = set(driver.window_handles) - handles_antes
    if not nuevas:
        # A veces no abre nueva ventana, intenta icono de calendario
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

# ── Leer supplier del PCM ──────────────────────────────────────
def parsear_pcm_nombre(nombre_pcm):
    """
    PKG-BUE-EX-1EURO1-PCSSTE → supplier=1EURO1, cod_madre=PCSSTE
    Formato: PKG-{ciudad}-{tipo}-{supplier}-{cod_madre}
    cod_madre = últimas posiciones después del último '-' (max 6 chars)
    supplier = penúltimo segmento
    """
    partes = nombre_pcm.split("-")
    if len(partes) < 3:
        return None, nombre_pcm
    cod_madre = partes[-1][:6]
    supplier  = partes[-2] if len(partes) >= 4 else ""
    return supplier, cod_madre

# ── Change Base Date ──────────────────────────────────────────
def cambiar_fecha_pcm(driver, dias_operacion, nombre_pcm, rate_from_str=""):
    """
    Cambia la fecha base del PCM para forzar recálculo de costos.
    Navegación por texto (no posición fija): hamburger → "ITINERARY" → "CHANGE BASE DATE".
    Reabre el hamburger entre pasos si es necesario (hasta 3 reintentos).
    Fecha nueva: `rate_from_str` ajustada al primer día de operación válido
    — normalmente RATE FROM del Excel, pero el caller puede pasar RATE FROM D
    en su lugar si la fila la trae cargada.
    Formato de fecha en el input: DD/MM/YY (grabación Chrome DevTools).
    Trigger Angular: click en li[1] (OLD TRAVEL DATE) tras escribir la fecha.
    Maneja la vista "Extension of expired rate" (#manualratesview / tp-button.saveall) post-save.
    Diálogo de recálculo: 5 pasos exactos de la grabación (labels + #calculatereplaceall + YES).
    """
    print(f"  📅 Change Base Date PCM: {nombre_pcm}")

    # Leer fecha actual del PCM (mostrada en el summary)
    fecha_actual = None
    for sel in ["div.tpcol3-3 input", "[class*='traveldate'] input",
                "input[class*='date']"]:
        try:
            for inp in driver.find_elements(By.CSS_SELECTOR, sel):
                v = inp.get_attribute("value") or inp.text
                f = parsear_fecha(v)
                if f and 2020 < f.year < 2035:
                    fecha_actual = f; break
        except: pass
        if fecha_actual: break

    # Fallback: leer cualquier fecha visible
    if not fecha_actual:
        for el in driver.find_elements(By.XPATH, "//*[contains(text(),'/')]"):
            try:
                f = parsear_fecha(el.text.strip())
                if f and 2020 < f.year < 2035:
                    fecha_actual = f; break
            except: pass

    print(f"    Fecha actual: {fmt_xl(fecha_actual) if fecha_actual else 'desconocida'}")

    # Usar RATE FROM del Excel como fecha objetivo del Change Base Date.
    # Si no viene rate_from_str, caer en siguiente día de operación desde hoy.
    if rate_from_str:
        base_rf = parsear_fecha(rate_from_str) or fecha_actual or datetime.now()
        nueva_fecha = siguiente_dia_operacion(dias_operacion, base_rf)
    else:
        base = fecha_actual or datetime.now()
        nueva_fecha = siguiente_dia_operacion(dias_operacion, base)
    # La nueva fecha debe ser DISTINTA de la actual del PCM: si coincide,
    # Tourplan no dispara el diálogo de recálculo y los rates/tipos de
    # cambio no se actualizan. En ese caso usar el siguiente día de operación.
    if fecha_actual and nueva_fecha.date() == fecha_actual.date():
        nueva_fecha = siguiente_dia_operacion(
            dias_operacion, nueva_fecha + timedelta(days=1))
    nueva_fecha_str = fmt_tp(nueva_fecha)
    print(f"    Nueva fecha: {nueva_fecha_str} ({DIAS_SEMANA[nueva_fecha.weekday()]})")

    # ── Abrir hamburger ──────────────────────────────────────────
    hamburger(driver)
    ss(driver, f"pcm_hamburger_{nombre_pcm[:12]}")

    # ── Click ITINERARY → (reabrir hamburger) → Change Base Date ──
    # Flujo grabado (Chrome DevTools recorder):
    #   1. hamburger
    #   2. click li[3] > div > div   (ITINERARY click-area)  → cierra hamburger + expande sub-menu
    #   3. hamburger de nuevo         (reabre con sub-menu ITINERARY ya expandido)
    #   4. click li[3] > ul > li[5] > div > label  (Change Base Date)
    #
    # XPaths explícitos del recording:
    #   ITINERARY:  //*[@id="pcmview"]/nav/tp-nav/.../ul/li[3]/div/div
    #   CBD label:  //*[@id="pcmview"]/nav/tp-nav/.../ul/li[3]/ul/li[5]/div/label
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
            # Paso 1: click en ITINERARY click-area (expande sub-menu, cierra hamburger)
            itin_el = WebDriverWait(driver, 6).until(
                EC.presence_of_element_located((By.XPATH, itin_xpath)))
            jc(driver, itin_el)
            time.sleep(0.5)
            # Paso 2: reabrir hamburger (sub-menu ITINERARY ya está expandido)
            hamburger(driver)
            time.sleep(0.5)
            # Paso 3: click en Change Base Date label
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

    # Formato de fecha que acepta el input del CBD: DD/MM/YY (grabación real)
    nueva_fecha_input = nueva_fecha.strftime("%d/%m/%y")

    # Input NEW TRAVEL DATE — id="newtraveldate", segundo input dentro del li[2]
    inp_fecha = wait(driver, "tp-dialog li:nth-of-type(2) input[type='text']")

    # Escribir la fecha: click + change event (reproduce exactamente la grabación)
    jc(driver, inp_fecha)
    time.sleep(0.2)
    set_val(driver, inp_fecha, nueva_fecha_input)
    driver.execute_script(
        "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));", inp_fecha)
    time.sleep(0.3)

    # Click en li[1] (OLD TRAVEL DATE) para que Angular procese el cambio — grabación
    try:
        jc(driver, driver.find_element(By.CSS_SELECTOR, "tp-dialog li:nth-of-type(1)"))
    except:
        pass
    time.sleep(0.3)
    ss(driver, f"pcm_fecha_llena_{nombre_pcm[:12]}")

    # SAVE del CBD dialog
    btn_save = wait(driver, "tp-button.save > button")
    jc(driver, btn_save)
    time.sleep(4 * VELOCIDAD)
    ss(driver, f"pcm_save_clicked_{nombre_pcm[:12]}")

    # Diálogo de recálculo — flujo exacto de la grabación:
    # 1. label del primer item del grupo de precios de servicio
    # 2. checkbox #calculatereplaceall (REPLACE ALL)
    # 3. label "Update Exchange"
    # 4. checkbox del exchange-rate-panel
    # 5. tp-button.yes
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
        # valorizacion_desde_madre.py.
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

    # Manejar vista "Extension of expired rate" (#manualratesview) si aparece.
    # NO es un diálogo modal — es una vista completa con tp-button.saveall.
    _tiene_vencidos = False
    try:
        WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "tp-button.saveall > button")))
        print(f"    ⚠ Vista 'Extension of expired rate' detectada → SAVE ALL")
        ss(driver, f"pcm_ext_rate_{nombre_pcm[:12]}")
        jc(driver, driver.find_element(By.CSS_SELECTOR, "tp-button.saveall > button"))
        time.sleep(5 * VELOCIDAD)
        ss(driver, f"pcm_ext_rate_done_{nombre_pcm[:12]}")
        _tiene_vencidos = True
    except Exception:
        pass  # La vista no apareció, continuar normalmente

    print(f"    ✅ Fecha actualizada a {nueva_fecha_str} (input: {nueva_fecha_input})")
    return (nueva_fecha, fmt_xl(nueva_fecha), fmt_xl(fecha_actual) if fecha_actual else "", _tiene_vencidos)

# ── Leer costos del PCM ───────────────────────────────────────
def _extraer_tabla_pax(driver, etiqueta):
    """
    Busca en el DOM la primera tabla con columnas N+N (pax ranges).
    Normaliza headers quitando espacios ('1 + 0' -> '1+0').
    Lee input.value para celdas editables.
    """
    diag = driver.execute_script("""
        function celda(td) {
            var inp = td.querySelector('input');
            if (inp) return inp.value.trim();
            return (td.innerText || td.textContent || '').trim();
        }
        function normH(s) { return s.replace(/\\s+/g,''); }
        var result = {tablas: [], encontrada: null};
        var all = document.querySelectorAll('table');
        result.total_tablas = all.length;
        for(var i=0; i<all.length; i++){
            var t = all[i];
            var ths = Array.from(t.querySelectorAll('th'));
            var tds_first = Array.from(t.querySelectorAll('tr:first-child td'));
            var rawHdrs = ths.concat(tds_first).map(function(c){
                return (c.innerText||c.textContent||'').trim();
            });
            var normHdrs = rawHdrs.map(normH);
            var tiene_pax = normHdrs.some(function(h){ return /^\\d+\\+\\d+$/.test(h); });
            var filas_labels = Array.from(t.querySelectorAll('tbody tr')).slice(0,8).map(function(tr){
                var c = tr.querySelector('th,td');
                return c ? celda(c).slice(0,30) : '';
            });
            result.tablas.push({
                idx: i,
                id: t.id || '',
                class: t.className.slice(0,40),
                tiene_pax: tiene_pax,
                headers: normHdrs.slice(0,12),
                filas_labels: filas_labels
            });
            if(tiene_pax && !result.encontrada){
                result.encontrada = Array.from(t.querySelectorAll('tr')).map(function(tr){
                    return Array.from(tr.querySelectorAll('th,td')).map(function(td){
                        return normH(celda(td));
                    });
                });
            }
        }
        return result;
    """)

    print(f"    [{etiqueta}] tablas en DOM: {diag.get('total_tablas', 0)}")
    for ti in diag.get('tablas', []):
        print(f"      tabla[{ti['idx']}] id='{ti['id']}' class='{ti['class']}'")
        print(f"        headers={ti['headers']}")
        print(f"        filas_labels={ti['filas_labels']}")

    encontrada = diag.get('encontrada') or []

    # Fallback especifico para DASHBOARD: #paxrangesfixedcolumns + #paxrangetable
    if not encontrada:
        encontrada = driver.execute_script("""
            function celda(td) {
                var inp = td.querySelector('input');
                if (inp) return inp.value.trim();
                return (td.innerText||td.textContent||'').trim();
            }
            function normH(s) { return s.replace(/\\s+/g,''); }
            var fixedTbl = document.getElementById('paxrangesfixedcolumns');
            var paxTbl   = document.getElementById('paxrangetable');
            if (!paxTbl) return null;
            // headers del paxrangetable
            var ths = Array.from(paxTbl.querySelectorAll('thead th,thead td'));
            if (!ths.length) {
                var fr = paxTbl.querySelector('tr');
                if (fr) ths = Array.from(fr.querySelectorAll('th,td'));
            }
            var headers = ths.map(function(h){ return normH(celda(h)); });
            if (!headers.some(function(h){ return /^\\d+\\+\\d+$/.test(h); })) return null;
            // etiquetas de fila del panel fijo
            var rowLabels = [];
            if (fixedTbl) {
                Array.from(fixedTbl.querySelectorAll('tbody tr')).forEach(function(tr){
                    var label = '';
                    for (var td of tr.querySelectorAll('td,th')) {
                        var t = celda(td);
                        if (t) { label = t; break; }
                    }
                    rowLabels.push(label);
                });
            }
            // filas de datos
            var rows = Array.from(paxTbl.querySelectorAll('tbody tr')).map(function(tr,idx){
                var vals = Array.from(tr.querySelectorAll('td,th')).map(celda);
                return [rowLabels[idx]||''].concat(vals);
            });
            return [['ROW_LABEL'].concat(headers)].concat(rows);
        """) or []
        if encontrada:
            print(f"    [{etiqueta}] Usando fallback #paxrangetable ({len(encontrada)-1} filas)")

    return encontrada


def leer_markup_commission(driver, nombre_pcm):
    """
    Navega al DASHBOARD y lee VOUCHER COST por rango de pax.
    Layout: pax ranges son columnas (thead de #paxrangetable);
            VOUCHER COST es una fila (label en #paxrangesfixedcolumns tbody).
    Si no existe fila VOUCHER COST retorna {} sin fallback.
    Retorna: (valores_dict, supplier_pcm, cod_madre, nombre_pcm)
    """
    print("  💰 Leyendo costos PCM desde DASHBOARD...")

    # Supplier y cod_madre se extraen del nombre (ya conocido desde USED IN)
    sup_nombre, cod_madre = parsear_pcm_nombre(nombre_pcm)
    pcm_supplier = sup_nombre or ""
    nombre_leido = nombre_pcm

    # Navegar al DASHBOARD via hamburger → li[1] click-area (grabación exacta)
    hamburger(driver)
    try:
        dash_el = driver.find_element(By.CSS_SELECTOR,
            "li:nth-of-type(1) > div > div")
        jc(driver, dash_el)
    except Exception:
        menu_item(driver, "DASHBOARD")
    time.sleep(10 * VELOCIDAD)
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

        // Panel izquierdo: encontrar índice de la fila VOUCHER COST
        var leftRows = Array.from(fixedTbl.querySelectorAll('tbody tr'));
        var vcRowIdx = -1;
        for (var i = 0; i < leftRows.length; i++) {
            var txt = (leftRows[i].innerText || leftRows[i].textContent || '')
                      .replace(/\\s+/g,'').toUpperCase();
            if (txt.indexOf('VOUCHERCOST') >= 0) { vcRowIdx = i; break; }
        }
        if (vcRowIdx < 0) return {error: 'no VOUCHER COST row',
            rows: leftRows.map(function(r){ return (r.innerText||r.textContent||'').trim(); })};

        // Panel derecho: pax ranges como columnas en el header
        var rightHdrs = Array.from(rightTbl.querySelectorAll('thead th, thead td')).map(function(th){
            return (th.innerText || th.textContent || '').trim().replace(/\\s+/g,'');
        });

        // Panel derecho: fila de VOUCHER COST (mismo índice de fila)
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
            v = float(str(raw).replace(",", "").replace("$", "").replace(" ", "").strip())
            if v > 0:
                valores[pax] = v
        except:
            pass

    print(f"    Valores ({len(valores)} rangos): {valores}")
    print(f"    Supplier PCM: {pcm_supplier}  Cod madre: {cod_madre}")
    return valores, pcm_supplier, cod_madre, nombre_leido

# ── Leer Price Code del PCM ────────────────────────────────────
def leer_price_code_pcm(driver, nombre_pcm=""):
    """
    Lee el PRICE CODE del PCM desde la vista PCM DETAILS (ej: 'TR').
    Ese mismo Price Code se asigna en Fase 2 al crear el período de
    rates del servicio madre (si no, el período queda 'Unassigned').
    Devuelve "" si no se pudo leer (y deja diagnóstico en el log).
    """
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
        # El campo suele mostrar 'TR - Descripción': quedarse con el código
        pc = re.split(r"\s*[-–]\s*", pc)[0].strip().upper() if pc else ""
        if pc:
            print(f"    Price Code del PCM: {pc}")
        else:
            print(f"    ⚠ Price Code no encontrado en PCM DETAILS — "
                  f"campos visibles:")
            for c in info.get("candidatos", [])[:25]:
                print(f"      {c}")
            ss(driver, f"pcm_details_pc_{nombre_pcm[:10]}")
        return pc
    except Exception as e:
        print(f"    ⚠ No pude leer el Price Code del PCM: {e}")
        return ""

# ── Cerrar ventana PCM y volver ────────────────────────────────
def cerrar_pcm(driver, handle_prod):
    for h in set(driver.window_handles) - {handle_prod}:
        try: driver.switch_to.window(h); driver.close()
        except: pass
    driver.switch_to.window(handle_prod)
    time.sleep(1)

# ── Buscar / actualizar rates del servicio madre ──────────────
def _parse_rate_period(texto):
    """Parsea '01/Apr/2026 - 31/Aug/2026' → (datetime, datetime)."""
    m = re.search(r'(\d+/\w+/\d+)\s*[-–]\s*(\d+/\w+/\d+)', texto)
    if m:
        return parsear_fecha(m.group(1)), parsear_fecha(m.group(2))
    return None, None

def _seleccionar_price_code(driver, code, etiqueta=""):
    """En el editor de un período de rates, pasa el modo a 'Selected Price
    Code' y elige el price code indicado (ej. 'TR').

    Los rates se editan POR price code: en modo 'All Price Codes' (default) la
    grilla es la vista agregada y lo que se escribe ahí NO persiste (al recargar
    se lee 0.0). Hay que seleccionar el price code concreto para editar/leer la
    grilla que sí persiste. (Hardcode simple a TR vía PRICE_CODE_DEFAULT.)
    """
    code = (code or "").strip().upper()
    if not code:
        return False
    try:
        # 1. Modo "Selected Price Code" (radio #priceCodeModeSelected)
        driver.execute_script("""
            var lbl = document.querySelector('label[for="priceCodeModeSelected"]');
            var inp = document.getElementById('priceCodeModeSelected');
            if (lbl) lbl.click();
            if (inp) inp.click();
        """)
        time.sleep(1.5)
        # 2. Abrir el dropdown #priceCode (tp-combo: hay un input de búsqueda)
        driver.execute_script("""
            var dd = document.getElementById('priceCode');
            if (!dd) return;
            var inp = dd.querySelector('input[type="text"], input');
            if (inp){ inp.focus(); inp.click(); } else { dd.click(); }
        """)
        time.sleep(1.0)
        # 2b. Escribir el code en el buscador para filtrar la lista
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
        # 3. Elegir la fila cuyo td.code (o descripción) matchee el code.
        #    DOM real (tp-dropdown#priceCode): table > tbody > tr[key] con
        #    <td class="code">TR</td><td class="description">TR - ...</td>
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
            // 1º preferir filas visibles (la lista pudo filtrarse al tipear)
            var cand = null;
            for (var tr of rows){ if (vis(tr)){ var m=match(tr); if(m){cand=m;break;} } }
            // 2º fallback: cualquier fila aunque esté oculta
            if (!cand){ for (var tr2 of rows){ var m2=match(tr2); if(m2){cand=m2;break;} } }
            if (!cand) return null;
            try{ cand.tr.scrollIntoView({block:'center'}); }catch(e){}
            var des = cand.tr.querySelector('td.description');
            var cod = cand.tr.querySelector('td.code');
            (des||cod||cand.tr).click();
            cand.tr.click();
            return cand.txt;
        """, code)
        time.sleep(2.5)  # la grilla se recarga según el price code elegido
        print(f"    Price Code '{code}' seleccionado ({etiqueta}): {elegido}")
        if not elegido:
            ss(driver, f"pricecode_fail_{(etiqueta or code)[:12]}")
            # dump del DOM del dropdown para diagnosticar si vuelve a fallar
            try:
                html = driver.execute_script(
                    "var d=document.getElementById('priceCode');"
                    "return d?d.outerHTML.slice(0,4000):'NO #priceCode';")
                print(f"    [debug priceCode DOM] {html[:1500]}")
            except Exception:
                pass
        return bool(elegido)
    except Exception as e:
        print(f"    ⚠ No pude seleccionar Price Code '{code}' ({etiqueta}): {e}")
        return False

def _buscar_servicio_madre_fase2(driver, cod_madre, pcm_supplier, location,
                                  service_type_madre=None):
    """
    Busca el servicio madre en Product Setup UNA VEZ y verifica que el
    supplier coincida con el del PCM. Se reutiliza para todos los
    PCM_Detail pendientes de Fase 2 que compartan el mismo servicio madre
    (mismo cod_destino/supplier/location/service_type), evitando repetir
    la búsqueda completa del producto por cada período a aplicar. La
    navegación a RATES (barata) NO se hace acá — se repite por período en
    actualizar_rates_servicio_madre, porque la verificación post-SAVE del
    período anterior puede dejar al driver parado en el detalle de otro
    período, no en la lista.
    """
    print(f"\n  🔍 Buscando servicio madre: {cod_madre}")
    buscar_producto(driver, location, pcm_supplier, cod_madre,
                    service_type=service_type_madre)

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
    if cod_supplier_actual and pcm_supplier and cod_supplier_actual != pcm_supplier.upper():
        raise Exception(
            f"Supplier del servicio madre ({supplier_actual}) ≠ supplier del PCM ({pcm_supplier}). "
            f"No se actualiza.")


def _cerrar_ultimo_tp_dialog(driver, max_espera=4):
    """
    Clickea EXIT/CANCEL/CLOSE en el ÚLTIMO <tp-dialog> visible del documento
    y espera a que se cierre. Angular nunca saca del DOM un tp-dialog ya
    cerrado (solo deja de mostrarlo), así que el diálogo activo SIEMPRE es
    el último del documento, nunca el primero (mismo criterio confirmado en
    valorizacion_desde_madre.py). Devuelve True si se cerró, False si
    seguía abierto tras `max_espera` segundos.
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


def actualizar_rates_servicio_madre(driver, cod_madre, pcm_supplier,
                                    rate_from_str, rate_to_str,
                                    location, valores_pcm, nombre_pcm,
                                    service_type_madre=None, price_code=""):
    """
    Aplica UN período de rates sobre un servicio madre YA BUSCADO (ver
    _buscar_servicio_madre_fase2, llamada por el caller antes de este loop):
    1. Navega a RATES
    2. Busca el periodo rate_from_str – rate_to_str
    3. Si no existe, crea uno nuevo (INSERT con el Price Code del PCM)
    4. Abre el período y actualiza los valores
    5. Recarga RATES y verifica que los valores hayan quedado guardados
    """
    print(f"\n  🔄 Actualizando servicio madre: {cod_madre}  periodo {rate_from_str}–{rate_to_str}")

    # Navegar a RATES
    hamburger(driver)
    ss(driver, f"madre_hamburger_{cod_madre}")
    menu_item(driver, "RATES")
    time.sleep(4 * VELOCIDAD)
    ss(driver, f"madre_rates_lista_{cod_madre}")

    # Price Code a usar en el período del servicio madre. Viene de la columna
    # PRICE CODE del Excel (Fase 1 → PCM_Detail). Si es 'ALL' (o vacío) se opera
    # sobre el período genérico SIN price code (Unassigned/All); si es un código
    # (TR, 34, etc.) se opera sobre el período de ESE price code, creándolo si no
    # existe (en Tourplan no se puede cambiar el PC de un período existente).
    PC = (price_code or PRICE_CODE_DEFAULT or "").strip().upper()
    IS_ALL = PC in ("", "ALL", "TODOS", "*", "UNASSIGNED")

    def _sel_pc(etq):
        # Filtra la lista de rates por el price code objetivo. Para ALL no se
        # filtra: la vista por defecto ya muestra el período Unassigned.
        if not IS_ALL:
            _seleccionar_price_code(driver, PC, etiqueta=etq)

    _sel_pc(f"lista {cod_madre}")
    time.sleep(1.5)
    ss(driver, f"madre_rates_lista_pc_{cod_madre}")

    def _leer_periodos():
        # Cada fila de la lista de rates trae la FECHA (td.tpcol-rateperiod) y su
        # PRICE CODE (td.tpcol-pricecodecode). Leemos ambos para distinguir el
        # período TR del 'Unassigned' que existe con las MISMAS fechas. El orden
        # coincide con find_elements('td.tpcol-rateperiod').
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

    def _idx_periodo_tr(rows):
        # Índice de la fila cuya FECHA coincide (±2 días) Y cuyo Price Code
        # coincide con el objetivo. Para ALL el período válido es el que NO
        # tiene price code (Unassigned/blank); si no, el del código exacto.
        for i, r in enumerate(rows):
            pf, pt = _parse_rate_period(r["date"])
            if not (pf and pt and abs((pf - rf).days) <= 2 and abs((pt - rt).days) <= 2):
                continue
            pc_row = (r["pc"] or "").strip().upper()
            if IS_ALL:
                if pc_row in PC_VACIOS:
                    return i
            elif pc_row == PC:
                return i
        return None

    periodos = _leer_periodos()
    print("    Períodos: {} (ej: {})".format(
        len(periodos),
        ", ".join("{}[{}]".format(r["date"], r["pc"] or "—") for r in periodos[:4])))

    # En Tourplan NO se puede cambiar el Price Code de un período existente: el
    # período de estas fechas que ya existe tiene Price Code 'Unassigned'. Por
    # eso buscamos un período de estas fechas Y Price Code TR; si no existe, se
    # crea uno NUEVO con TR (INSERT). En re-corridas el período TR ya existe y
    # se reusa (no se duplica).
    idx_periodo = _idx_periodo_tr(periodos)

    pc_warning = ""
    ya_en_detalle = False
    if idx_periodo is None:
        etq_pc = "ALL/Unassigned" if IS_ALL else PC
        print(f"    No hay período {etq_pc} para {fmt_tp(rf)}–{fmt_tp(rt)} → COPY último período")
        pc_warning, ya_en_detalle = _copiar_ultimo_period(
            driver, rf, rt, cod_madre, "" if IS_ALL else PC)
        ss(driver, f"madre_rates_creado_{cod_madre}")
        if not ya_en_detalle:
            _sel_pc(f"post-insert {cod_madre}")
            time.sleep(1.5)
            periodos = _leer_periodos()
            idx_periodo = _idx_periodo_tr(periodos)
            print(f"    Períodos tras INSERT: {len(periodos)}  (idx={idx_periodo})")
            if idx_periodo is None:
                raise Exception(
                    f"El período {etq_pc} {fmt_tp(rf)}–{fmt_tp(rt)} no aparece en la "
                    f"lista de rates de {cod_madre} tras el INSERT.")

    # Abrir el período — si COPY DATE RANGE ya lo abrió, omitimos el click.
    # Default para cuando ya_en_detalle=True (COPY DATE RANGE nos dejó
    # directo en el detalle, sin pasar por la lista) — sin esto,
    # periodo_encontrado queda sin asignar y la verificación post-SAVE de
    # más abajo revienta con UnboundLocalError si el período no aparece al
    # recargar RATES (mismo bug confirmado en corrida real de
    # valorizacion_desde_madre.py, 2026-08).
    periodo_encontrado = f"{fmt_tp(rf)}–{fmt_tp(rt)}"
    if not ya_en_detalle:
        periodo_encontrado = periodos[idx_periodo]["date"]
        filas_td = driver.find_elements(By.CSS_SELECTOR, "td.tpcol-rateperiod")
        print(f"    Abriendo período {PC if not IS_ALL else 'ALL'}: "
              f"{periodo_encontrado} (idx {idx_periodo})")
        jc(driver, filas_td[idx_periodo])
    else:
        print("    Ya en detalle del período copiado → omitiendo apertura desde lista")

    time.sleep(5 * VELOCIDAD)
    ss(driver, f"madre_rates_detalle_{cod_madre}")

    # Reasegurar el Price Code dentro del período abierto (backup del filtro)
    _sel_pc(f"escribir {cod_madre}")
    ss(driver, f"madre_rates_pricecode_{cod_madre}")

    # Actualizar valores en la tabla de rates
    v_viejos, v_nuevos, rangos, pendientes = _escribir_rates(
        driver, valores_pcm, cod_madre)
    if pc_warning:
        pendientes.append(pc_warning)

    # ── Verificación post-SAVE ─────────────────────────────────
    # El click en SAVE puede "parecer" exitoso sin persistir nada.
    # Única prueba real: cerrar el detalle, reabrir el período y comparar.
    # IMPORTANTE: si la verificación falla NO levantamos excepción acá; en su
    # lugar devolvemos un error_post junto con v_viejos/v_nuevos para que el
    # llamador igual deje en el Excel el valor anterior y el nuevo (auditoría).
    #
    # No se reabre el menú lateral (hamburger) acá: reabrirlo con la
    # pantalla de detalle del período todavía activa dejaba colgado el
    # backdrop de navegación (.tpnavbackdrop) y ese backdrop interceptaba
    # el click/lectura siguiente — el valor quedaba bien guardado en
    # Tourplan pero la verificación leía mal y reportaba "no se actualizó"
    # (mismo bug confirmado y corregido en valorizacion_desde_madre.py). En
    # su lugar: cerrar el detalle con EXIT (_cerrar_ultimo_tp_dialog) para
    # volver a la lista de períodos subyacente, y reabrir el período desde
    # ahí — igual que la apertura inicial más arriba.
    error_post = ""
    try:
        cerrado = _cerrar_ultimo_tp_dialog(driver)
        if not cerrado:
            ss(driver, f"verif_dialog_no_cerro_{cod_madre[:10]}")
            error_post = (f"Verificación post-SAVE: el detalle del período "
                          f"{periodo_encontrado} no se cerró (EXIT) tras el SAVE")
        else:
            time.sleep(2 * VELOCIDAD)
            periodos_v = _leer_periodos()
            idx_v = _idx_periodo_tr(periodos_v)
            if idx_v is None:
                ss(driver, f"verif_sin_periodo_{cod_madre[:10]}")
                error_post = (f"Verificación post-SAVE: el período TR {periodo_encontrado} "
                              f"no aparece en la lista tras reabrir")
            else:
                filas_td = driver.find_elements(By.CSS_SELECTOR, "td.tpcol-rateperiod")
                jc(driver, filas_td[idx_v])
                time.sleep(5 * VELOCIDAD)
                # Mismo price code que al escribir, si no la verificación lee otra vista
                _sel_pc(f"verif {cod_madre}")

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
                        actual = leidos.get(rango)
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
                    ss(driver, f"verif_sin_tabla_{cod_madre[:10]}")
                    dump(driver, f"verif_sin_tabla_{cod_madre[:10]}")
                    error_post = "Verificación post-SAVE: no pude releer la grilla del período"
                elif difs:
                    ss(driver, f"verif_error_{cod_madre[:10]}")
                    dump(driver, f"verif_error_{cod_madre[:10]}")
                    error_post = (f"Los valores NO quedaron guardados en {cod_madre}: "
                                  f"{len(difs)}/{len(rangos)} rangos difieren tras recargar y "
                                  f"{MAX_INTENTOS_VERIF} intentos (ej: {'; '.join(difs[:3])})")
                else:
                    print(f"    ✔ Verificación post-SAVE OK: {len(rangos)} rangos AD "
                          f"persistidos en {cod_madre}")
                    ss(driver, f"verif_ok_{cod_madre[:10]}")
    except Exception as _ve:
        error_post = f"Verificación post-SAVE falló: {_ve}"

    return v_viejos, v_nuevos, rangos, pendientes, error_post

def _copiar_ultimo_period(driver, fecha_desde, fecha_hasta, cod, price_code=""):
    """
    Copia el período más reciente via COPY DATE RANGE y ajusta FROM y TO.
    Retorna (pc_warning, ya_en_detalle).
      ya_en_detalle=True  → Tourplan ya abrió el nuevo período; el caller omite
                            el click en la lista.
      ya_en_detalle=False → fallback a INSERT; el caller debe abrir normalmente.

    El dialog "Copy Rate" que abre acá es el MISMO componente tp-insert-rate
    / #insertrateview que usa también el fallback a INSERT. Angular nunca
    saca del DOM un <tp-dialog> ya cerrado (se van acumulando, uno o más por
    período procesado) — por eso el dialog correcto SIEMPRE es el ÚLTIMO
    tp-dialog del documento (tp-dialog:last-of-type), nunca una posición fija
    como nth-of-type(2) — mismo bug confirmado con dumps de HTML reales y
    corregido en valorizacion_desde_madre.py, portado acá. Ver detalle en
    _fill_dialog_date() y _cancelar_dialog_y_hacer_insert() más abajo.

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
        corrida real (AEEZMV, valorizacion_desde_excel.py) que el campo TO
        a veces re-formatea el valor que se le escribió ('31/03/27') a un
        formato distinto pero equivalente ('31/Mar/2027') apenas lo
        acepta — comparar como texto literal tomaba eso como una falla y
        cancelaba un COPY DATE RANGE que en realidad había funcionado
        bien."""
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
        #
        # Angular nunca saca del DOM un <tp-dialog> ya cerrado — se van
        # acumulando, uno o más por período procesado (se vieron hasta 8 en
        # una misma corrida). El diálogo recién abierto SIEMPRE es el
        # ÚLTIMO del documento, nunca una posición fija — ver
        # _cerrar_ultimo_tp_dialog() más arriba (mismo criterio ya usado en
        # la verificación post-SAVE de actualizar_rates_servicio_madre).
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

        hamburger(driver)
        menu_item(driver, "RATES")
        time.sleep(3 * VELOCIDAD)
        return _insertar_rate_period(driver, fecha_desde, fecha_hasta, cod, price_code), False

    def _diag_tp_dialogs():
        """Diagnóstico: cuántos <tp-dialog> hay en el documento en este
        momento, cuáles están visibles y qué texto tienen. Mismo diagnóstico
        que valorizacion_desde_madre.py: Angular nunca saca del DOM un
        <tp-dialog> ya cerrado, así que un selector posicional como
        'tp-dialog:nth-of-type(2)' puede agarrar un diálogo viejo y
        abandonado en vez del recién abierto (siempre el ÚLTIMO)."""
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
        print(f"    ⚠ Campo FROM no encontrado: {_e} → cancela y hace INSERT")
        print(f"    🔍 Diagnóstico tp-dialog en el documento: {_diag_tp_dialogs()}")
        ss(driver, f"rates_copy_nofrom_{cod[:10]}")
        dump(driver, f"rates_copy_nofrom_{cod[:10]}")
        return _cancelar_dialog_y_hacer_insert()

    try:
        _fill_dialog_date(
            "tp-dialog:last-of-type li:nth-of-type(3) input[type='text']",
            fecha_hasta_str, "TO")
    except Exception as _e:
        print(f"    ⚠ Campo TO no encontrado: {_e} → cancela y hace INSERT")
        print(f"    🔍 Diagnóstico tp-dialog en el documento: {_diag_tp_dialogs()}")
        ss(driver, f"rates_copy_noto_{cod[:10]}")
        dump(driver, f"rates_copy_noto_{cod[:10]}")
        return _cancelar_dialog_y_hacer_insert()

    ss(driver, f"rates_copy_lleno_{cod[:10]}")
    jc(driver, wait(driver, "tp-button.ok > button"))
    time.sleep(5 * VELOCIDAD)
    ss(driver, f"rates_copy_ok_{cod[:10]}")
    print(f"    ✅ COPY DATE RANGE OK → nuevo período {fecha_desde_str} – {fecha_hasta_str}")
    return "", True   # ya_en_detalle=True


def _insertar_rate_period(driver, fecha_desde, fecha_hasta, cod, price_code=""):
    """Clickea INSERT en la lista de rates y llena el nuevo período:
    fechas From/To y el PRICE CODE leído del PCM (sin él, el período
    queda 'Unassigned' y Tourplan no asocia los valores).

    Scopea los inputs y el botón SAVE al diálogo (tp-dialog / tp-modal),
    porque buscando en toda la página se levantan los inputs del
    encabezado del producto (Code/Supplier/Location) y se escriben las
    fechas en el lugar equivocado.

    Devuelve "" si todo OK, o un texto de advertencia si el Price Code
    no se pudo asignar (para el pipeline de pendientes).
    """
    try:
        btn_insert = waitx(driver,
            "//button[normalize-space(text())='INSERT' or normalize-space(text())='Insert']",
            t=8)
        jc(driver, btn_insert)
        time.sleep(3 * VELOCIDAD)
        ss(driver, f"rates_insert_dlg_{cod[:10]}")

        # Diagnóstico: inputs y botones DENTRO del diálogo
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
                var c = inp.closest('li,tr,tp-input,div');
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

        # El diálogo precarga la fecha FROM (último fin + 1) en un input
        # y deja el TO vacío (visto en logs: input[1]='01/Sep/2026',
        # input[2]=''). Identificar cada uno por contenido.
        pat_fecha = re.compile(r'\d{1,2}/[A-Za-z]{3}/\d{2,4}')
        idx_con_fecha = [i["idx"] for i in inputs_info
                         if pat_fecha.search(i["val"])]
        idx_vacios = [i["idx"] for i in inputs_info
                      if not (i["val"] or "").strip()]

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
            # From y To visibles con fecha: llenar ambos
            _fill(idx_con_fecha[0], fecha_desde)
            _fill(idx_con_fecha[1], fecha_hasta)
        elif len(idx_con_fecha) == 1 and idx_vacios:
            # From precargado + To vacío (caso normal)
            idx_from = idx_con_fecha[0]
            idx_to = next((i for i in idx_vacios if i > idx_from),
                          idx_vacios[0])
            val_from = next(i["val"] for i in inputs_info
                            if i["idx"] == idx_from)
            if val_from != fmt_tp(fecha_desde):
                _fill(idx_from, fecha_desde)
            _fill(idx_to, fecha_hasta)
            print(f"    From=input[{idx_from}] ('{val_from}')  "
                  f"To=input[{idx_to}] ← {fmt_tp(fecha_hasta)}")
        elif len(idx_con_fecha) == 1:
            _fill(idx_con_fecha[0], fecha_hasta)
        elif len(idx_vacios) >= 2:
            # Diálogo SIN fecha precargada (caso PCSA34): input[0] suele ser el
            # Price Code (val='Unassigned' → NO está vacío, no entra acá). Los dos
            # primeros inputs VACÍOS son From y To, en ese orden.
            idx_from, idx_to = idx_vacios[0], idx_vacios[1]
            _fill(idx_from, fecha_desde)
            _fill(idx_to, fecha_hasta)
            print(f"    (sin precarga) From=input[{idx_from}]←{fmt_tp(fecha_desde)}  "
                  f"To=input[{idx_to}]←{fmt_tp(fecha_hasta)}")
        else:
            dump(driver, f"rates_insert_sin_fechas_{cod[:10]}")
            raise Exception(
                f"No identifiqué inputs de fecha en el diálogo INSERT "
                f"(inputs: {inputs_info})")

        # ── PRICE CODE en el diálogo INSERT ──────────────────────
        # En esta versión de Tourplan NX el Price Code se elige en la
        # página de rates (grupo 'InsertServicePriceCodeGroup': radio
        # 'Selected' + dropdown #priceCode) ANTES de clickear INSERT, y el
        # nuevo período hereda ese price code. Por las dudas, si el diálogo
        # trae su propio combo de Price Code (input con 'Unassigned' y/o un
        # #priceCode propio), lo seteamos también acá. Si no lo trae, NO es
        # un error: ya quedó fijado en la página.
        pc_warning = ""
        if price_code:
            # 1) ¿Hay un #priceCode dentro del diálogo? (tabla tr/td.code)
            elegido_dlg = _seleccionar_price_code(
                driver, price_code, etiqueta=f"dialog {cod}") \
                if driver.execute_script(
                    "var d=document.querySelector("
                    "'tp-dialog #priceCode, tp-modal #priceCode, "
                    "[role=\"dialog\"] #priceCode');return !!d;") else False
            # 2) Fallback: input de texto con 'Unassigned' (combo viejo)
            tipeado = "SKIP"
            opcion = None
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
            asignado = bool(elegido_dlg) or any(
                price_code.upper() in (v or "").upper() for v in valores_dlg)
            print(f"    Price Code '{price_code}' (dialog): dropdown={elegido_dlg} "
                  f"tipeo={tipeado} opción={opcion} inputs={valores_dlg}")
            # Si el diálogo no trae campo de Price Code, NO es error: ya quedó
            # fijado en la página de rates antes del INSERT.
            sin_campo_pc = (not elegido_dlg) and (tipeado == "NO_INPUT")
            if not asignado and not sin_campo_pc:
                pc_warning = (f"Price Code '{price_code}' no se pudo asignar "
                              f"al período (el campo quedó: {valores_dlg})")
                print(f"    🚩 {pc_warning}")
        else:
            pc_warning = ("PCM sin Price Code leído: el período pudo quedar "
                          "'Unassigned' — revisar manualmente")
            print(f"    🚩 {pc_warning}")

        ss(driver, f"rates_insert_lleno_{cod[:10]}")

        # SAVE del INSERT — scopeado al diálogo
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
            raise Exception(
                f"No encontré botón SAVE en el diálogo INSERT "
                f"(botones visibles: {info.get('botones')})")
        print(f"    SAVE del INSERT clickeado ('{guardado}')")
        time.sleep(4 * VELOCIDAD)
        ss(driver, f"rates_insert_save_{cod[:10]}")
        print(f"    ✅ Período insertado {fmt_tp(fecha_desde)} – {fmt_tp(fecha_hasta)}")
        return pc_warning
    except Exception as e:
        try: ss(driver, f"rates_insert_error_{cod[:10]}")
        except: pass
        raise Exception(f"Error al insertar período: {e}")

def _leer_tabla_rates(driver):
    """Devuelve headers + filas (celdas e inputs) de la tabla de rates
    visible, o None si no hay ninguna en pantalla."""
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
    """De una tabla de rates devuelve {label_rango_AD: valor_float}.
    Mismo criterio de filas que la escritura: solo rangos pax 'N - N AD'."""
    headers = tabla["headers"]
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
        # Group Cost = primer input de la fila (idx 0); ver nota en _leer_valor.
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
    """
    Lee los rates actuales del período abierto y escribe los nuevos basados en valores_pcm.
    Solo toca filas AD (no CH/IN ni suplementos). Por cada fila AD escribe 4 columnas:
    Group Cost (0), Group Sell (3), FIT Cost (4), FIT Sell (7).
    Retorna: (valores_viejos, valores_nuevos, rangos_madre, pendientes_error)
    """
    # Ir al tab RATES si existe
    for sel in ["[class*='tab-item']", "div.tab", "a.tab", "button.tab"]:
        for el in driver.find_elements(By.CSS_SELECTOR, sel):
            if el.text.strip().upper() == "RATES" and el.is_displayed():
                jc(driver, el); time.sleep(3 * VELOCIDAD); break

    tabla = _leer_tabla_rates(driver)

    if not tabla:
        ss(driver, f"rates_sin_tabla_{cod[:10]}")
        dump(driver, f"rates_sin_tabla_{cod[:10]}")
        raise Exception(f"No encontré tabla de rates en {cod}")

    headers = tabla["headers"]
    idx_svc  = next((i for i,h in enumerate(headers) if "SERVICE" in h.upper()), 0)
    idx_cost = next((i for i,h in enumerate(headers)
                     if "GROUP COST" in h.upper() and "FIT" not in h.upper()), 1)

    rangos_madre = []   # filas AD con rango pax: las únicas que se escriben
    v_viejos = {}
    pendientes = []     # hallazgos para revisión manual

    def _leer_valor(row, celdas):
        # El costo (Group Cost) es el PRIMER input de la fila (idx 0). Los
        # headers tienen una columna 'Service Item' sin input, por eso NO se
        # puede usar el índice de header contra la lista de inputs.
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
        celdas = row["celdas"]
        if not celdas: continue
        svc_item = celdas[idx_svc] if idx_svc < len(celdas) else ""
        if not svc_item.strip(): continue
        valor_actual = _leer_valor(row, celdas)
        es_rango = bool(PAT_RANGO.search(svc_item))
        es_ad = bool(re.search(r'\bAD\b', svc_item.upper()))
        if es_rango and es_ad:
            v_viejos[svc_item] = valor_actual
            rangos_madre.append(svc_item)
            continue
        # Las demás filas NO se escriben: rangos CH/IN (child/infant) y
        # filas sin rango (HALF TWIN, SINGLE SUPPLEMENT, TRIPLE/QUAD
        # REDUCTION, ADDITIONAL ADULT). En los servicios madre no
        # deberían tener valor; si lo tienen se reporta para revisión.
        if valor_actual:
            msg = f"{svc_item} tiene valor {valor_actual} (no se tocó)"
            pendientes.append(msg)
            print(f"    🚩 REVISAR: {msg}")
        else:
            print(f"    ⏭ Fila no se toca (sin valor): {svc_item}")

    # Mapear valores PCM a rangos del servicio madre
    def pax_n(key):
        m = re.match(r'(\d+)\+', key)
        return int(m.group(1)) if m else 0

    def rango_limites(texto):
        """Devuelve (pmin, pmax) del texto de rango, ej. '5 - 9 AD' → (5, 9)."""
        m = re.search(r'(\d+)\s*[-–]\s*(\d+)', texto)
        if m:
            return int(m.group(1)), int(m.group(2))
        m = re.search(r'(\d+)', texto)
        n = int(m.group(1)) if m else 0
        return n, n

    # Reglas de mapeo PCM → servicio madre:
    # - Si pmax == 9999 y pmin >= 42 → escribir 0 (rango abierto alto)
    # - Si pmax == 9999 y pmin < 42  → escribir 999 (rango abierto bajo)
    # - Caso normal: leer todos los valores del PCM con pax en [pmin, pmax],
    #   tomar el máximo. Si no hay ningún valor en ese rango → error de
    #   inconsistencia en bases (no debería ocurrir).
    por_pax = {pax_n(k): v for k, v in valores_pcm.items()}
    v_nuevos = {}
    for rango in rangos_madre:
        pmin, pmax = rango_limites(rango)
        if pmax >= 9999:
            val = 0.0 if pmin >= 42 else 999.0
            v_nuevos[rango] = val
            print(f"    {rango}: {v_viejos.get(rango,'?')} → {val:.2f}  [regla 9999]")
            continue
        # Rango normal: máximo costo del PCM para pax en [pmin, pmax]
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
        print(f"    {rango}: {v_viejos.get(rango,'?')} → {val:.2f}  "
              f"[max de pax {pmin}-{pmax}: {vals_en_rango}]")

    # Escribir en los inputs de la tabla — mismo filtro que rangos_madre
    # (solo filas AD con rango pax) para que los índices coincidan.
    # IMPORTANTE: scopear al DIÁLOGO del período (tp-dialog). Si buscamos en
    # toda la página tomamos la grilla de la LISTA de rates que quedó detrás
    # del modal y escribimos en el lugar equivocado (la grabación del flujo
    # humano también scopea con 'tp-dialog ...').
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

    inputs_escritos = []   # refs a los Group Cost escritos (para readback diag)
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
        # La grilla tiene 8 columnas de costo por fila:
        #   0=Group Cost  1=Mup%  2=Mup$  3=Group Sell
        #   4=FIT Cost    5=Mup%  6=Mup$  7=FIT Sell
        # Se escribe el MISMO valor en las 4 columnas pedidas: Group Cost,
        # Group Sell, FIT Cost y FIT Sell (0, 3, 4, 7).
        # El input es un componente Angular (tp-number > tp-validator > input).
        # Se replica EXACTAMENTE la grabación del flujo humano que SÍ persiste:
        # foco → setear value por el setter NATIVO → disparar 'input' y 'change'
        # → blur/focusout. tp-number confirma el valor al perder foco (blur);
        # sin el blur el modelo queda en 0.0 aunque el input muestre el número.
        valor_txt = f"{v_nuevos.get(rango,0):.2f}"
        for ci in (0, 3, 4, 7):
            if ci >= len(edits):
                continue
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

    # ── DIAGNÓSTICO de escritura ───────────────────────────────────────────
    # Reportamos (a) si los valores QUEDARON en los inputs de la grilla y
    # (b) TODOS los botones Save de la página con su estado. Esto distingue
    # dos fallas muy distintas que dan el mismo síntoma ("0.0 al recargar"):
    #   · los valores no se tipean en el input correcto  → input.value vacío/0
    #   · se chequea/clickea un Save que NO es el de la grilla → Save disabled
    #     mientras el verdadero (dentro del tp-dialog) está enabled.
    try:
        diag = driver.execute_script("""
            function vis(e){ return !!(e && (e.offsetWidth || e.offsetHeight
                                       || e.getClientRects().length)); }
            // valores actuales en los inputs de costo de la grilla visible
            var inputs = Array.from(document.querySelectorAll('input'))
                .filter(function(i){ return vis(i) && !i.readOnly
                        && (i.className||'').indexOf('tpnumber') >= 0; });
            var vals = inputs.slice(0, 12).map(function(i){ return i.value; });
            // todos los Save de la página
            var saves = Array.from(document.querySelectorAll('tp-button.save > button'))
                .map(function(b, k){
                    return {idx:k, disabled:!!b.disabled, visible:vis(b),
                            inDialog: !!b.closest('tp-dialog, tp-modal, [role=\"dialog\"]')};
                });
            var dlg = document.querySelector('tp-dialog, tp-modal, [role="dialog"]');
            return {nInputsCosto: inputs.length, valsCosto: vals,
                    saves: saves, dialogVisible: vis(dlg)};
        """)
        print(f"    [diag] inputs de costo visibles: {diag['nInputsCosto']} "
              f"· (primeros 12, suelen ser CH=0): {diag['valsCosto']}")
        print(f"    [diag] dialog visible: {diag['dialogVisible']} · "
              f"botones Save: {diag['saves']}")
        # Lo que importa: los inputs AD que REALMENTE escribimos (Group Cost).
        leidos_ad = []
        for inp in inputs_escritos[:8]:
            try: leidos_ad.append(inp.get_attribute("value"))
            except Exception: leidos_ad.append("?")
        print(f"    [diag] valores en inputs AD escritos (Group Cost): {leidos_ad}")
    except Exception as _de:
        print(f"    [diag] no se pudo inspeccionar: {_de}")

    print("    📸 PANTALLA: DESPUÉS DE ESCRIBIR — ANTES DEL SAVE")
    ss(driver, f"A_DESPUES_ESCRIBIR_ANTES_SAVE_{cod[:8]}")

    # SAVE de la grilla: puede haber VARIOS tp-button.save en la página (uno a
    # nivel producto, otro dentro del diálogo del período). Hay que clickear el
    # que está ENABLED y, de preferencia, el que vive dentro del tp-dialog.
    # Si clickeamos el de producto (siempre disabled) el período no persiste.
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
            # prioridad: enabled+dialog > enabled > resto
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
            except Exception:
                continue
    except Exception:
        pass
    if not guardado:
        guardado = driver.execute_script("""
            function vis(e){ return !!(e.offsetWidth || e.offsetHeight
                                       || e.getClientRects().length); }
            var saves = Array.from(document.querySelectorAll('tp-button.save > button'))
                .filter(function(b){ return vis(b) && !b.disabled; });
            // preferir el de dentro del diálogo
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
        #     un SAVE aparte) — visto en producción: Tourplan redondea el costo
        #     al escribirlo (ej. tipeamos 241.56 y queda 241.50), así que el
        #     form vuelve a estar 'pristine' apenas se escribe, mucho antes de
        #     que busquemos el botón. Comparar contra el viejo (anterior a
        #     escribir) da un falso error en este caso — hay que comparar
        #     contra lo que HAY AHORA en la grilla, con una tolerancia que
        #     contemple ese redondeo.
        def _f(x):
            try: return float(str(x).replace(",", ".") or 0)
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
            print(f"    ✅ Sin cambios: el período ya tenía estos valores "
                  f"({len(rangos_madre)} rangos AD) — no hace falta guardar")
            return v_viejos, v_nuevos, rangos_madre, pendientes
        if _ya_aplicado(0.30):
            print(f"    ✅ Sin botón SAVE pero los valores ya están aplicados en "
                  f"la grilla (dentro de tolerancia de redondeo) — período nuevo "
                  f"confirmado sin SAVE aparte ({len(rangos_madre)} rangos AD)")
            return v_viejos, v_nuevos, rangos_madre, pendientes
        ss(driver, f"rates_save_error_{cod[:10]}")
        dump(driver, f"rates_save_error_{cod[:10]}")
        raise Exception(f"No encontré el botón SAVE de la grilla de rates en {cod}")
    print(f"    SAVE de rates clickeado ('{guardado}')")
    time.sleep(4 * VELOCIDAD)
    print("    📸 PANTALLA: DESPUÉS DEL SAVE")
    ss(driver, f"B_DESPUES_DEL_SAVE_{cod[:8]}")
    print(f"    ✅ SAVE OK — {len(rangos_madre)} rangos AD actualizados"
          + (f"  (🚩 {len(pendientes)} pendientes de revisión)" if pendientes else ""))
    return v_viejos, v_nuevos, rangos_madre, pendientes

# ── Excel ──────────────────────────────────────────────────────
def crear_excel_si_no_existe():
    """Crea el Excel de trabajo con estructura base si no existe.
    Devuelve True si lo acaba de crear, False si ya existía."""
    if os.path.exists(EXCEL_PATH):
        return False
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET

    fill_h = PatternFill("solid", start_color="1F4E79", fgColor="1F4E79")
    font_h = Font(name="Arial", bold=True, color="FFFFFF", size=9)
    aln_c  = Alignment(horizontal="center", vertical="center", wrap_text=True)

    headers_base = [
        "LOCATION","SUPPLIER","COD ORIGEN","SERVICE TYPE",
        "RATE FROM","RATE TO","PRICE CODE","MOSTRAR CAPTURAS",
        "COD DESTINO","RATE FROM D","RATE TO D",
        "PCM REF","DÍAS OP","ESTADO","ERROR","TIMESTAMP"
    ]
    for col, h in enumerate(headers_base, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = font_h; c.fill = fill_h; c.alignment = aln_c

    # Fila de ejemplo (TKCEM pertenece al supplier 6CEMRE; usar el supplier real
    # del producto, NO el del servicio madre). Es solo un ejemplo: reemplazá la
    # fila por tus datos y dejá ESTADO=PENDIENTE.
    ws.cell(row=2, column=C["location"],     value="BUE")
    ws.cell(row=2, column=C["supplier"],     value="6CEMRE")
    ws.cell(row=2, column=C["cod_origen"],   value="TKCEM")
    ws.cell(row=2, column=C["service_type"], value="TK")
    ws.cell(row=2, column=C["rate_from"],    value="01/Sep/2026")
    ws.cell(row=2, column=C["rate_to"],      value="31/Dec/2026")
    ws.cell(row=2, column=C["price_code"],   value="TR")    # TR, 34, ... o ALL
    ws.cell(row=2, column=C["mostrar_capturas"], value="NO")
    # OJO: el ejemplo NO queda en PENDIENTE a propósito, para que la plantilla
    # recién creada no se autoejecute. Completá tus datos y poné ESTADO=PENDIENTE.
    ws.cell(row=2, column=C["estado"],       value="EJEMPLO")

    for col_letter, w in (("A",10),("B",10),("C",12),("D",13),("E",13),
                          ("F",13),("G",11),("H",16),("I",12),
                          ("N",12),("O",40)):
        ws.column_dimensions[col_letter].width = w

    _init_pcm_detail_sheet(wb)
    _aplicar_comentarios_excel(wb)

    wb.save(EXCEL_PATH)
    print(f"✅ Excel creado: {EXCEL_PATH}")
    print("   Completá las filas con estado=PENDIENTE y volvé a correr el script.")
    return True

SHEET_DETAIL = "PCM_Detail"

# Columnas fijas de PCM_Detail (en orden):
# 1 TIMESTAMP  2 COD_ORIGEN  3 LOCATION  4 SUPPLIER  5 RATE_FROM  6 RATE_TO
# 7 PCM  8 SUPPLIER_PCM  9 COD_DESTINO  10 DIAS_OP
# 11 FECHA_ANT  12 FECHA_NUEVA  13 MARKUP_JSON
# 14 ESTADO_FASE1  15 ERROR_FASE1  16 ESTADO_FASE2  17 ERROR_FASE2
_DETAIL_HEADERS = [
    "TIMESTAMP","COD ORIGEN","LOCATION","SUPPLIER","RATE FROM","RATE TO",
    "PCM","SUPPLIER PCM","COD DESTINO","DIAS OP",
    "FECHA ANT","FECHA NUEVA","MARKUP JSON","PRICE CODE",
    "ESTADO F1","ERROR F1","ESTADO F2","ERROR F2",
]
_DETAIL_COL = {h: i+1 for i, h in enumerate(_DETAIL_HEADERS)}

# Bloque de markup "abierto": una columna por rango pax (1+0 … 41+0), con
# el valor leído del PCM en cada fila. Sirve para comparar rápido en Excel.
# El PCM entrega los costos por cantidad de pax ("N+0"); 1..41 cubre todos
# los casos vistos. Va justo después de los headers fijos.
_MARKUP_PAX = [f"{n}+0" for n in range(1, 42)]
_DETAIL_MARKUP_START = len(_DETAIL_HEADERS) + 1            # col 19
_DETAIL_RATES_START  = _DETAIL_MARKUP_START + len(_MARKUP_PAX)  # tras el markup

def _init_pcm_detail_sheet(wb):
    """Crea o reinicia encabezados fijos de PCM_Detail (no toca filas de datos)."""
    if SHEET_DETAIL not in wb.sheetnames:
        ws = wb.create_sheet(SHEET_DETAIL)
    else:
        ws = wb[SHEET_DETAIL]
    fill_h = PatternFill("solid", start_color="1F4E79", fgColor="1F4E79")
    font_h = Font(name="Arial", bold=True, color="FFFFFF", size=9)
    aln_c  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col, h in enumerate(_DETAIL_HEADERS, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = font_h; c.fill = fill_h; c.alignment = aln_c
    widths = [16,12,10,10,13,13, 28,12,12,14, 12,12,40,11, 12,35,14,35]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    # Headers del bloque markup abierto (un rango pax por columna)
    fill_m = PatternFill("solid", start_color="2E75B6", fgColor="2E75B6")
    font_m = Font(name="Arial", bold=True, color="FFFFFF", size=8)
    for i, rango in enumerate(_MARKUP_PAX):
        col = _DETAIL_MARKUP_START + i
        c = ws.cell(row=1, column=col, value=rango)
        c.font = font_m; c.fill = fill_m; c.alignment = aln_c
        ws.column_dimensions[get_column_letter(col)].width = 8
    return ws


# Comentarios de ayuda en los headers del Excel: qué se completa a mano
# (ENTRADA), qué escribe el script (SALIDA) y cómo leer cada estado.
_COMENTARIOS_PRODUCTOS = {
    "LOCATION":     "ENTRADA\nCiudad Tourplan del producto origen.\nEj: BUE",
    "SUPPLIER":     "ENTRADA\nCódigo de supplier del producto origen "
                    "(solo el código, sin nombre).\nEj: 6CEMRE",
    "COD ORIGEN":   "ENTRADA\nCódigo del producto origen. De su lista USED IN "
                    "salen los PCMs 'Package Service' (PKG-...).\nEj: TKCEM",
    "SERVICE TYPE": "ENTRADA\nTipo de servicio del producto origen (ej: TK). "
                    "Si el tipo no existe en el buscador, se busca sin filtro "
                    "(warning inofensivo en el log).",
    "RATE FROM":    "ENTRADA\nInicio del período a valorizar en los servicios "
                    "madre. Formato dd/Mon/yyyy.\nEj: 01/Sep/2026",
    "RATE TO":      "ENTRADA\nFin del período a valorizar.\nEj: 31/Dec/2026",
    "COD DESTINO":  "ENTRADA OPCIONAL\nFuerza el código del servicio madre. "
                    "Si queda vacío se usa el último segmento del nombre del "
                    "PCM (PKG-BUE-EX-1EURO1-XXXX → XXXX).",
    "RATE FROM D":  "ENTRADA OPCIONAL\nFecha base para el recalculate/Change "
                    "Base Date del PCM (con la regla de siguiente día hábil "
                    "si no opera esa fecha exacta). Vacía → usa RATE FROM, "
                    "igual que antes. RATE FROM sigue siendo la base para "
                    "crear/buscar el período de rates a valorizar.\n"
                    "Formato dd/Mon/yyyy.\nEj: 01/Sep/2026",
    "RATE TO D":    "Reservada (hoy no se usa).",
    "PCM REF":      "SALIDA\nPCMs procesados en Fase 1, uno por línea.",
    "DÍAS OP":      "SALIDA\nDías de operación leídos del producto. Si no se "
                    "pudieron leer, se asumen todos.",
    "ESTADO":       "ENTRADA/SALIDA\nPoner PENDIENTE para que el script "
                    "procese la fila.\nEl script escribe: PROCESANDO → "
                    "FASE1_OK (verde) o ERROR (rojo).\nPara reprocesar, "
                    "volver a escribir PENDIENTE.",
    "ERROR":        "SALIDA\nDetalle del error de Fase 1 (si lo hubo).",
    "TIMESTAMP":    "SALIDA\nFecha/hora de la última actualización de la fila.",
}

_COMENTARIOS_DETAIL = {
    "TIMESTAMP":    "SALIDA\nCuándo se leyó este PCM en Fase 1.",
    "COD ORIGEN":   "SALIDA\nProducto origen del que salió este PCM "
                    "(copiado de PRODUCTOS).",
    "LOCATION":     "SALIDA\nCopiado de PRODUCTOS.",
    "SUPPLIER":     "SALIDA\nSupplier del producto origen (PRODUCTOS).",
    "RATE FROM":    "SALIDA\nInicio del período a aplicar (PRODUCTOS).",
    "RATE TO":      "SALIDA\nFin del período a aplicar (PRODUCTOS).",
    "PCM":          "SALIDA\nNombre del Package Service procesado:\n"
                    "PKG-{ciudad}-{tipo}-{supplier}-{cod_madre}",
    "SUPPLIER PCM": "SALIDA\nSupplier leído dentro del PCM. Fase 2 valida que "
                    "coincida con el supplier del servicio madre.",
    "COD DESTINO":  "SALIDA\nCódigo del servicio madre que Fase 2 actualiza "
                    "(último segmento del nombre del PCM u override).",
    "DIAS OP":      "SALIDA\nDías de operación usados para el Change Base Date.",
    "FECHA ANT":    "SALIDA\nFecha base del PCM antes del cambio.",
    "FECHA NUEVA":  "SALIDA\nFecha base aplicada para forzar el recálculo.",
    "MARKUP JSON":  "SALIDA\nValores leídos por rango pax (JSON). Es el insumo "
                    "que Fase 2 escribe en el servicio madre. A la derecha se "
                    "abre el mismo dato en una columna por rango (1+0…41+0) "
                    "para comparar rápido en Excel.",
    "PRICE CODE":   "SALIDA\nPrice Code leído del PCM (ej: TR). Fase 2 lo "
                    "asigna al crear el período de rates del servicio madre. "
                    "Si queda vacío, Fase 2 usa el Price Code por defecto "
                    "(PRICE_CODE_DEFAULT, hoy 'TR') para que el período no "
                    "quede Unassigned.",
    "ESTADO F1":    "SALIDA\nLEIDO = lectura OK · ERROR = falló la lectura "
                    "(ver ERROR F1).",
    "ERROR F1":     "SALIDA\nDetalle del error de Fase 1.",
    "ESTADO F2":    "SALIDA (editable para reintentar)\n"
                    "PENDIENTE_APLICAR: Fase 2 lo va a procesar.\n"
                    "OK: período creado/abierto y rangos AD guardados.\n"
                    "OK_REVISAR: aplicado, pero una fila que no se toca "
                    "(CH/IN, HALF TWIN, supplements) tenía valor → revisar "
                    "manualmente (detalle en ERROR F2).\n"
                    "SVS_MADRE_NO_ENCONTRADO: el servicio madre no existe "
                    "en Tourplan.\n"
                    "ERROR: falla técnica (ver ERROR F2 y screenshots).\n"
                    "SKIP: sin markup leído.\n"
                    "Para reintentar una fila, volver a escribir "
                    "PENDIENTE_APLICAR.",
    "ERROR F2":     "SALIDA\nDetalle del error de Fase 2, o el listado de "
                    "filas con valor inesperado cuando el estado es "
                    "OK_REVISAR.",
}

def _aplicar_comentarios_excel(wb):
    """Agrega/actualiza los comentarios de ayuda en los headers de ambas hojas."""
    def _set(ws, comentarios):
        for cell in ws[1]:
            texto = comentarios.get(str(cell.value or "").strip())
            if texto:
                c = Comment(texto, "Valorización PKG", height=190, width=320)
                cell.comment = c
    if SHEET in wb.sheetnames:
        _set(wb[SHEET], _COMENTARIOS_PRODUCTOS)
    if SHEET_DETAIL in wb.sheetnames:
        _set(wb[SHEET_DETAIL], _COMENTARIOS_DETAIL)


def _detail_fill(estado):
    """Retorna (fill, font) según estado."""
    if estado in ("LEIDO", "OK"):
        return (PatternFill("solid", start_color="E2EFDA", fgColor="E2EFDA"),
                Font(name="Arial", size=9, bold=True, color="375623"))
    if estado in ("ERROR", "ERROR_FASE1", "ERROR_FASE2"):
        return (PatternFill("solid", start_color="FFE0E0", fgColor="FFE0E0"),
                Font(name="Arial", size=9, bold=True, color="C00000"))
    if estado == "PENDIENTE_APLICAR":
        return (PatternFill("solid", start_color="FFF2CC", fgColor="FFF2CC"),
                Font(name="Arial", size=9, bold=True, color="7F6000"))
    if estado == "OK_REVISAR":
        return (PatternFill("solid", start_color="FCE4D6", fgColor="FCE4D6"),
                Font(name="Arial", size=9, bold=True, color="C55A11"))
    return (PatternFill("solid", start_color="D9D9D9", fgColor="D9D9D9"),
            Font(name="Arial", size=9, bold=True, color="000000"))


def escribir_pcm_detail_fase1(cod_origen, location, supplier, rate_from, rate_to,
                               nombre_pcm, supplier_pcm, cod_destino, dias_op,
                               fecha_ant, fecha_nueva, markup_dict,
                               price_code="", estado_fase1="LEIDO",
                               error_fase1=""):
    """
    Fase 1: agrega una fila en PCM_Detail con los datos leídos del PCM.
    ESTADO_FASE2 = PENDIENTE_APLICAR (si lectura OK) o SKIP (si error).
    """
    wb = openpyxl.load_workbook(EXCEL_PATH)
    if SHEET_DETAIL not in wb.sheetnames:
        _init_pcm_detail_sheet(wb)
    ws = wb[SHEET_DETAIL]
    if ws.cell(row=1, column=1).value != "TIMESTAMP":
        _init_pcm_detail_sheet(wb)

    next_row = ws.max_row + 1 if ws.max_row > 1 else 2
    estado_f2 = "PENDIENTE_APLICAR" if estado_fase1 == "LEIDO" else "SKIP"

    markup_str = json.dumps(markup_dict, ensure_ascii=False) if markup_dict else ""

    vals = [
        datetime.now().strftime("%Y-%m-%d %H:%M"),
        cod_origen, location, supplier, rate_from, rate_to,
        nombre_pcm, supplier_pcm, cod_destino, dias_op,
        fecha_ant, fecha_nueva, markup_str, price_code,
        estado_fase1, error_fase1[:300] if error_fase1 else "",
        estado_f2, "",
    ]
    for col, v in enumerate(vals, 1):
        c = ws.cell(row=next_row, column=col, value=v)
        c.font = Font(name="Arial", size=9)
        c.alignment = Alignment(vertical="center",
                                wrap_text=(col in (_DETAIL_COL["MARKUP JSON"],
                                                   _DETAIL_COL["ERROR F1"],
                                                   _DETAIL_COL["ERROR F2"])))
        if col == _DETAIL_COL["ESTADO F1"]:
            fill, font = _detail_fill(estado_fase1)
            c.fill = fill; c.font = font
        if col == _DETAIL_COL["ESTADO F2"]:
            fill, font = _detail_fill(estado_f2)
            c.fill = fill; c.font = font

    # Markup abierto: un valor por columna de rango pax (1+0 … 41+0)
    if markup_dict:
        for i, rango in enumerate(_MARKUP_PAX):
            if rango in markup_dict:
                mc = ws.cell(row=next_row, column=_DETAIL_MARKUP_START + i,
                             value=markup_dict[rango])
                mc.font = Font(name="Arial", size=8)
                mc.alignment = Alignment(horizontal="center")

    wb.save(EXCEL_PATH); wb.close()
    print(f"    📝 PCM_Detail F1: {nombre_pcm} → {estado_fase1}/{estado_f2}")


def leer_pcm_detail_pendientes_fase2():
    """Lee filas de PCM_Detail donde ESTADO_FASE2 = PENDIENTE_APLICAR."""
    wb = openpyxl.load_workbook(EXCEL_PATH)
    if SHEET_DETAIL not in wb.sheetnames:
        wb.close(); return []
    ws = wb[SHEET_DETAIL]

    # Detectar columnas por header
    header_row = [str(c.value or "").strip() for c in ws[1]]
    col_idx = {h: i for i, h in enumerate(header_row)}

    def gc(name):
        return col_idx.get(name, _DETAIL_COL.get(name, 0) - 1)

    filas = []
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        def val(name):
            idx = gc(name)
            return str(row[idx].value or "").strip() if idx < len(row) else ""

        if FASE2_REAPLICAR_TODO:
            # Modo re-test: tomar cualquier fila que tenga COD DESTINO,
            # sin importar el ESTADO F2 (que pudo quedar en error/done).
            if not val("COD DESTINO"):
                continue
        elif val("ESTADO F2").upper() != "PENDIENTE_APLICAR":
            continue
        try:
            markup_dict = json.loads(val("MARKUP JSON")) if val("MARKUP JSON") else {}
        except Exception:
            markup_dict = {}

        filas.append({
            "_row":        i,
            "cod_origen":  val("COD ORIGEN"),
            "location":    val("LOCATION"),
            "supplier":    val("SUPPLIER"),
            "rate_from":   val("RATE FROM"),
            "rate_to":     val("RATE TO"),
            "pcm":         val("PCM"),
            "supplier_pcm":val("SUPPLIER PCM"),
            "cod_destino": val("COD DESTINO"),
            "dias_op":     val("DIAS OP"),
            "fecha_nueva": val("FECHA NUEVA"),
            "markup_dict": markup_dict,
            "price_code":  val("PRICE CODE"),
        })
    wb.close()
    return filas


def marcar_pcm_detail_row(row_num, estado_fase2, error="",
                           rangos=None, v_viejos=None, v_nuevos=None):
    """Actualiza ESTADO_FASE2, ERROR_FASE2 y columnas de rates en PCM_Detail."""
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb[SHEET_DETAIL]

    # Headers de rangos dinámicos
    if rangos:
        fill_h = PatternFill("solid", start_color="1F4E79", fgColor="1F4E79")
        font_h = Font(name="Arial", bold=True, color="FFFFFF", size=8)
        aln_c  = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for i, rango in enumerate(rangos):
            col_v = _DETAIL_RATES_START + i * 2
            col_n = _DETAIL_RATES_START + i * 2 + 1
            for col, lbl in ((col_v, f"VIEJO {rango}"), (col_n, f"NUEVO {rango}")):
                c = ws.cell(row=1, column=col, value=lbl)
                c.font = font_h; c.fill = fill_h; c.alignment = aln_c

    # Estado y error
    fill, font = _detail_fill(estado_fase2)
    c_est = ws.cell(row=row_num, column=_DETAIL_COL["ESTADO F2"])
    c_est.value = estado_fase2; c_est.fill = fill; c_est.font = font
    c_est.alignment = Alignment(horizontal="center", vertical="center")

    c_err = ws.cell(row=row_num, column=_DETAIL_COL["ERROR F2"])
    c_err.value = error[:300] if error else ""
    c_err.font  = Font(name="Arial", size=9)

    # Valores de rates
    if rangos and v_viejos and v_nuevos:
        for i, rango in enumerate(rangos):
            col_v = _DETAIL_RATES_START + i * 2
            col_n = _DETAIL_RATES_START + i * 2 + 1
            cv = ws.cell(row=row_num, column=col_v, value=v_viejos.get(rango, ""))
            cv.font = Font(name="Arial", color="C00000", size=9)
            cv.fill = PatternFill("solid", start_color="FFE0E0", fgColor="FFE0E0")
            cv.alignment = Alignment(horizontal="center")
            cn = ws.cell(row=row_num, column=col_n, value=v_nuevos.get(rango, ""))
            cn.font = Font(name="Arial", color="375623", size=9, bold=True)
            cn.fill = PatternFill("solid", start_color="E2EFDA", fgColor="E2EFDA")
            cn.alignment = Alignment(horizontal="center")

    wb.save(EXCEL_PATH); wb.close()


def leer_flag_mostrar_capturas():
    """Lee la columna MOSTRAR CAPTURAS de la hoja de entrada (primer valor no
    vacío). Devuelve True sólo si dice SI/SÍ/YES/TRUE/1. Default False."""
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
        ws = wb[SHEET] if SHEET in wb.sheetnames else wb.active
        header = [str(c.value or "").strip().upper() for c in ws[1]]
        idx = next((k for k, h in enumerate(header)
                    if h in ("MOSTRAR CAPTURAS", "CAPTURAS",
                             "MOSTRAR SCREENSHOTS", "SCREENSHOTS")), None)
        val = ""
        if idx is not None:
            for row in ws.iter_rows(min_row=2):
                if idx < len(row) and str(row[idx].value or "").strip():
                    val = str(row[idx].value).strip().upper(); break
        wb.close()
        return val in ("SI", "SÍ", "S", "YES", "Y", "TRUE", "1")
    except Exception:
        return False


def leer_pendientes():
    """
    Lee filas con estado PENDIENTE o PENDING.
    Detecta automáticamente la columna de estado buscando el header,
    con fallback a la posición fija del diccionario C.
    """
    ESTADOS_PENDIENTE = {"PENDIENTE", "PENDING", "PEND"}

    wb = openpyxl.load_workbook(EXCEL_PATH)
    # Buscar la hoja correcta (acepta PRODUCTOS, Sheet1, primera hoja)
    if SHEET in wb.sheetnames:
        ws = wb[SHEET]
    else:
        ws = wb.active

    # Detectar columnas por header (fila 1)
    col_map = {}  # nombre_campo → índice base-0
    header_row = [str(c.value or "").strip().upper() for c in ws[1]]
    HEADER_ALIASES = {
        "LOCATION": "location", "SUPPLIER": "supplier",
        "COD ORIGEN": "cod_origen", "CODIGO ORIGEN": "cod_origen", "CODE": "cod_origen",
        "SERVICE TYPE": "service_type", "SERVICETYPE": "service_type",
        "TIPO SERVICIO": "service_type", "TIPO": "service_type", "SVC TYPE": "service_type",
        "RATE FROM": "rate_from", "FROM": "rate_from", "DESDE": "rate_from",
        "RATE TO": "rate_to", "TO": "rate_to", "HASTA": "rate_to",
        "PRICE CODE": "price_code", "PRICECODE": "price_code", "PC": "price_code",
        "MOSTRAR CAPTURAS": "mostrar_capturas", "CAPTURAS": "mostrar_capturas",
        "MOSTRAR SCREENSHOTS": "mostrar_capturas", "SCREENSHOTS": "mostrar_capturas",
        "COD DESTINO": "cod_destino", "DESTINO": "cod_destino",
        "RATE FROM D": "rate_from_d", "RATE TO D": "rate_to_d",
        "PCM REF": "pcm_ref", "DÍAS OP": "dias_op", "DIAS OP": "dias_op",
        "ESTADO": "estado", "STATUS": "estado", "STATE": "estado",
        "ERROR": "error_msg", "TIMESTAMP": "timestamp",
    }
    for idx, h in enumerate(header_row):
        campo = HEADER_ALIASES.get(h)
        if campo and campo not in col_map:
            col_map[campo] = idx

    # Para campos no detectados por header, usar posición fija de C
    def get_col(campo):
        return col_map.get(campo, C[campo] - 1)

    # Índice de columna de estado
    idx_estado = get_col("estado")

    filas = []
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        if idx_estado >= len(row):
            continue
        est = str(row[idx_estado].value or "").strip().upper()
        if est in ESTADOS_PENDIENTE:
            d = {}
            for campo in C:
                ci = get_col(campo)
                d[campo] = row[ci].value if ci < len(row) else None
            d["_row"] = i
            filas.append(d)

    wb.close()
    return filas

def _estilo_estado(estado):
    if estado == "OK":
        return (PatternFill("solid", start_color="E2EFDA", fgColor="E2EFDA"),
                Font(name="Arial", bold=True, color="375623", size=9))
    elif estado == "ERROR":
        return (PatternFill("solid", start_color="FFE0E0", fgColor="FFE0E0"),
                Font(name="Arial", bold=True, color="C00000", size=9))
    return (PatternFill("solid", start_color="FFF2CC", fgColor="FFF2CC"),
            Font(name="Arial", bold=True, color="7F6000", size=9))

def marcar_procesando(row_num):
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb[SHEET]
    c = ws.cell(row=row_num, column=C["estado"])
    fill, font = _estilo_estado("PROCESANDO")
    c.value = "PROCESANDO"; c.fill = fill; c.font = font
    wb.save(EXCEL_PATH); wb.close()

def escribir_resultado(row_num, estado, pcm_ref="", dias_op="",
                       rangos=None, v_viejos=None, v_nuevos=None,
                       error=""):
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb[SHEET]

    # Asegurar headers de rates dinámicos en fila 1
    if rangos:
        for i, rango in enumerate(rangos):
            col_v = RATES_START + i * 2
            col_n = RATES_START + i * 2 + 1
            # Header viejo
            cv = ws.cell(row=1, column=col_v)
            cv.value = f"VIEJO {rango}"
            cv.font  = Font(name="Arial", bold=True, color="C00000", size=8)
            cv.fill  = PatternFill("solid", start_color="FFE0E0", fgColor="FFE0E0")
            cv.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            # Header nuevo
            cn = ws.cell(row=1, column=col_n)
            cn.value = f"NUEVO {rango}"
            cn.font  = Font(name="Arial", bold=True, color="375623", size=8)
            cn.fill  = PatternFill("solid", start_color="E2EFDA", fgColor="E2EFDA")
            cn.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Valores de rates
    if rangos and v_viejos and v_nuevos:
        for i, rango in enumerate(rangos):
            col_v = RATES_START + i * 2
            col_n = RATES_START + i * 2 + 1
            cv = ws.cell(row=row_num, column=col_v)
            cv.value = v_viejos.get(rango, "")
            cv.font  = Font(name="Arial", color="C00000", size=9)
            cv.fill  = PatternFill("solid", start_color="FFE0E0", fgColor="FFE0E0")
            cv.alignment = Alignment(horizontal="center")
            cn = ws.cell(row=row_num, column=col_n)
            cn.value = v_nuevos.get(rango, "")
            cn.font  = Font(name="Arial", color="375623", size=9, bold=True)
            cn.fill  = PatternFill("solid", start_color="E2EFDA", fgColor="E2EFDA")
            cn.alignment = Alignment(horizontal="center")

    # PCM REF: un PCM por línea dentro de la celda (legible). El valor llega
    # con saltos de línea desde _procesar_periodo_fase1.
    c_pcmref = ws.cell(row=row_num, column=C["pcm_ref"])
    c_pcmref.value = pcm_ref
    c_pcmref.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions[get_column_letter(C["pcm_ref"])].width = 30
    ws.cell(row=row_num, column=C["dias_op"]).value   = dias_op
    ws.cell(row=row_num, column=C["timestamp"]).value = datetime.now().strftime("%Y-%m-%d %H:%M")
    ws.cell(row=row_num, column=C["error_msg"]).value = error[:300] if error else ""

    c_est = ws.cell(row=row_num, column=C["estado"])
    fill, font = _estilo_estado(estado)
    c_est.value = estado; c_est.fill = fill; c_est.font = font
    c_est.alignment = Alignment(horizontal="center", vertical="center")

    wb.save(EXCEL_PATH); wb.close()

# ── Procesar una fila del Excel ───────────────────────────────
def _buscar_componente_fase1(driver, fila):
    """
    Busca el componente origen UNA VEZ (Product Search + días de operación +
    USED IN + lista de PCMs). El resultado se reutiliza para todas las filas
    del Excel que compartan el mismo componente (location+supplier+cod_origen+
    service_type) pero distinto período (RATE FROM/RATE TO), evitando repetir
    la búsqueda completa por cada fecha a valorizar del mismo producto.
    """
    loc  = str(fila["location"]     or "").strip()
    sup  = str(fila["supplier"]     or "").strip()
    cod  = str(fila["cod_origen"]   or "").strip()
    stype= str(fila.get("service_type") or "").strip().upper()

    handle_prod = driver.current_window_handle

    # ── Buscar componente origen ───────────────────────────────
    buscar_producto(driver, loc, sup, cod, service_type=stype, handle_origen=handle_prod)

    # ── Leer días de operación ─────────────────────────────────
    dias_op  = leer_dias_operacion(driver)
    dias_str = ",".join(sorted(dias_op))

    # ── USED IN ────────────────────────────────────────────────
    ir_a_used_in(driver)
    pcm_list = leer_pcm_list(driver)
    if not pcm_list:
        raise Exception("No hay PCMs de tipo Package en USED IN")

    return handle_prod, dias_op, dias_str, pcm_list


def _procesar_pcm_todos_periodos(driver, pcm_info, grupo, handle_prod,
                                  dias_op, dias_str, resultados_por_fila,
                                  start_scroll=0):
    """
    Abre UN PCM una sola vez y aplica Change Base Date + lee costos para
    TODOS los períodos (filas) del grupo sobre esa misma ventana ya abierta,
    cerrándola recién al final. Evita reabrir el mismo PCM (con el scroll
    costoso por las ~1900 filas de USED IN) una vez por cada período —
    cambiar_fecha_pcm/leer_markup_commission/leer_price_code_pcm navegan
    todas dentro de la ventana ya abierta (vía el menú hamburguesa), no
    hace falta cerrar y reabrir entre período y período.

    resultados_por_fila: dict {row_num: [nombres de PCM leídos]} que se va
    completando a medida que cada período se procesa OK; lo escribe/lee el
    caller (MAIN) para el resultado final de cada fila.

    start_scroll: posición de scroll en USED IN desde donde arrancar la
    búsqueda de este PCM (el caller la va encadenando entre PCMs del mismo
    grupo, igual que hacía el loop original por período).

    Devuelve la posición de scroll donde se encontró este PCM (o
    start_scroll si no se pudo abrir), para que el caller se la pase al
    siguiente PCM del grupo.
    """
    nombre_pcm = pcm_info["nombre"]
    print(f"\n  ▶ PCM: {nombre_pcm}  ({len(grupo)} período(s) a aplicar)")

    def _campos(fila):
        return (
            str(fila["location"]    or "").strip(),
            str(fila["supplier"]    or "").strip(),
            str(fila["cod_origen"]  or "").strip(),
            str(fila["rate_from"]   or "").strip(),
            str(fila["rate_to"]     or "").strip(),
            str(fila["cod_destino"] or "").strip(),
            str(fila.get("price_code") or "").strip(),
            str(fila.get("rate_from_d") or "").strip(),
        )

    # Re-navegar a USED IN antes de abrir: el orden del listado puede
    # cambiar tras Change Base Date en PCMs anteriores del mismo grupo.
    try:
        ir_a_used_in(driver)
        time.sleep(2)
    except Exception as _nav_e:
        print(f"    ⚠ Re-nav USED IN falló ({_nav_e}), continuando con vista actual")

    try:
        _handle_pcm, found_scroll = abrir_pcm(driver, pcm_info, start_scroll=start_scroll)
    except Exception as e:
        # No se pudo abrir el PCM: todos los períodos del grupo fallan para este PCM.
        print(f"  ❌ Error abriendo {nombre_pcm}: {e}")
        ss(driver, f"error_f1_abrir_{nombre_pcm[:12]}")
        for fila in grupo:
            loc, sup, cod, rf, rt, cd, _pc_in, _rfd = _campos(fila)
            escribir_pcm_detail_fase1(
                cod_origen=cod, location=loc, supplier=sup, rate_from=rf, rate_to=rt,
                nombre_pcm=nombre_pcm, supplier_pcm="", cod_destino=cd or "",
                dias_op=dias_str, fecha_ant="", fecha_nueva="", markup_dict=None,
                estado_fase1="ERROR", error_fase1=str(e),
            )
        try:
            cerrar_pcm(driver, handle_prod)
        except Exception:
            pass
        return start_scroll

    for fila in grupo:
        loc, sup, cod, rf, rt, cd, pc_in, rfd = _campos(fila)
        row = fila["_row"]
        # RATE FROM D (columna opcional, reutilizada): si la fila la trae
        # cargada, se usa como base del Change Base Date del PCM en vez de
        # RATE FROM. RATE FROM sigue siendo la base para crear/buscar el
        # período de rates a valorizar (rf/rt de acá arriba) — eso no cambia.
        rf_recalc = rfd or rf

        print(f"    · Período [{rf} – {rt}] (fila {row})")
        try:
            # Cambiar fecha base: usa RATE FROM D si la fila la trae cargada,
            # si no cae a RATE FROM (comportamiento de siempre)
            nueva_fecha, nueva_str, ant_str, _tiene_vencidos = cambiar_fecha_pcm(
                driver, dias_op, nombre_pcm, rate_from_str=rf_recalc)

            # Leer costos del PCM (Markup/Commission o DASHBOARD fallback)
            valores_pcm, pcm_sup, cod_madre, nombre_completo = \
                leer_markup_commission(driver, nombre_pcm)

            if not valores_pcm:
                print("    ℹ️  VOUCHER COST no existe en este PCM — se guardará con markup vacío")

            cod_final = cd if cd else cod_madre
            if not cod_final:
                raise Exception("No se pudo determinar código del servicio madre")

            # Price Code para el período del servicio madre: PRIORIDAD a la
            # columna PRICE CODE del Excel (TR, 34, ... o ALL); si está vacía,
            # se intenta leer del PCM.
            price_code = pc_in or leer_price_code_pcm(driver, nombre_pcm)

            escribir_pcm_detail_fase1(
                cod_origen   = cod,
                location     = loc,
                supplier     = sup,
                rate_from    = rf,
                rate_to      = rt,
                nombre_pcm   = nombre_pcm,
                supplier_pcm = pcm_sup,
                cod_destino  = cod_final,
                dias_op      = dias_str,
                fecha_ant    = ant_str,
                fecha_nueva  = nueva_str,
                markup_dict  = valores_pcm,
                price_code   = price_code,
                estado_fase1 = "LEIDO",
            )
            resultados_por_fila[row].append(
                nombre_pcm + (" [COMPONENTES VENCIDOS]" if _tiene_vencidos else ""))

        except Exception as e:
            print(f"  ❌ Error leyendo {nombre_pcm} (fila {row}): {e}")
            ss(driver, f"error_f1_{nombre_pcm[:12]}_row{row}")
            escribir_pcm_detail_fase1(
                cod_origen   = cod,
                location     = loc,
                supplier     = sup,
                rate_from    = rf,
                rate_to      = rt,
                nombre_pcm   = nombre_pcm,
                supplier_pcm = "",
                cod_destino  = cd or "",
                dias_op      = dias_str,
                fecha_ant    = "", fecha_nueva  = "",
                markup_dict  = None,
                estado_fase1 = "ERROR", error_fase1 = str(e),
            )

    cerrar_pcm(driver, handle_prod)
    return found_scroll


def procesar_fila_fase2(driver, pcm_row):
    """
    Fase 2 — Solo escritura en Tourplan.
    Lee una fila de PCM_Detail (PENDIENTE_APLICAR) y aplica las tasas
    en el servicio madre correspondiente.
    """
    row         = pcm_row["_row"]
    cod_origen  = pcm_row["cod_origen"]
    location    = pcm_row["location"]
    rate_from   = pcm_row["rate_from"]
    rate_to     = pcm_row["rate_to"]
    nombre_pcm  = pcm_row["pcm"]
    supplier_pcm= pcm_row["supplier_pcm"]
    cod_destino = pcm_row["cod_destino"]
    markup_dict = pcm_row["markup_dict"]
    price_code  = (pcm_row.get("price_code") or "").strip().upper() or PRICE_CODE_DEFAULT

    print(f"\n  ▶ FASE 2 — {nombre_pcm} → {cod_destino}  [{rate_from} – {rate_to}]")

    if not markup_dict:
        marcar_pcm_detail_row(row, "SKIP", error="markup_dict vacío")
        return

    # Service type del servicio madre: 3er segmento del nombre del PCM
    # (PKG-{ciudad}-{tipo}-{supplier}-{cod_madre} → tipo, ej. EX)
    partes_pcm = nombre_pcm.split("-")
    stype_madre = partes_pcm[2].strip() if len(partes_pcm) >= 5 else None

    try:
        v_viejos, v_nuevos, rangos, pendientes, error_post = \
            actualizar_rates_servicio_madre(
                driver, cod_destino, supplier_pcm,
                rate_from, rate_to, location,
                markup_dict, nombre_pcm,
                service_type_madre=stype_madre,
                price_code=price_code)

        # Si la verificación post-SAVE falló, igual dejamos en el Excel el
        # valor ANTERIOR del servicio madre (VIEJO) y el que se intentó
        # escribir (NUEVO) — para auditoría — y marcamos ERROR con el detalle.
        if error_post:
            print(f"    ❌ {nombre_pcm} → {cod_destino}: {error_post}")
            ss(driver, f"error_f2_{nombre_pcm[:12]}")
            err_full = error_post + (("; " + "; ".join(pendientes)) if pendientes else "")
            marcar_pcm_detail_row(
                row, "ERROR", error=err_full,
                rangos=rangos, v_viejos=v_viejos, v_nuevos=v_nuevos)
            return

        # Las filas que no se escriben (CH/IN, HALF TWIN, supplements…)
        # no deberían tener valor en el servicio madre. Si lo tienen, la
        # fila queda OK_REVISAR con el detalle en ERROR F2 para revisión
        # manual; los rangos AD igual se aplicaron.
        estado = "OK_REVISAR" if pendientes else "OK"
        marcar_pcm_detail_row(
            row, estado, error="; ".join(pendientes),
            rangos=rangos, v_viejos=v_viejos, v_nuevos=v_nuevos)
        if pendientes:
            print(f"    🚩 {nombre_pcm} → {cod_destino} OK_REVISAR: "
                  f"{len(pendientes)} filas con valor inesperado")
        else:
            print(f"    ✅ {nombre_pcm} → {cod_destino} OK")

    except ProductoNoEncontrado as e:
        # El servicio madre no existe en Tourplan: no es un error del script,
        # se marca aparte para revisión manual y no se reintenta.
        print(f"    ⚠ Servicio madre {cod_destino} no existe en Tourplan → SVS_MADRE_NO_ENCONTRADO")
        ss(driver, f"error_f2_{nombre_pcm[:12]}")
        marcar_pcm_detail_row(row, "SVS_MADRE_NO_ENCONTRADO", error=str(e))

    except Exception as e:
        print(f"    ❌ Error fase 2 {nombre_pcm}: {e}")
        ss(driver, f"error_f2_{nombre_pcm[:12]}")
        marcar_pcm_detail_row(row, "ERROR", error=str(e))


# ── Agrupamiento para evitar re-búsquedas ──────────────────────
def _agrupar_por_componente(filas):
    """
    Agrupa filas PENDIENTE de Fase 1 por componente
    (location+supplier+cod_origen+service_type), preservando el orden de
    primera aparición de cada grupo y el orden relativo dentro de cada uno.
    Filas con el mismo componente pero distinto período (RATE FROM/RATE TO)
    quedan juntas para reutilizar la búsqueda del producto y la lista de PCMs.
    """
    grupos, orden = {}, []
    for f in filas:
        key = (
            str(f.get("location") or "").strip().upper(),
            str(f.get("supplier") or "").strip().upper(),
            str(f.get("cod_origen") or "").strip().upper(),
            str(f.get("service_type") or "").strip().upper(),
        )
        if key not in grupos:
            grupos[key] = []
            orden.append(key)
        grupos[key].append(f)
    return [grupos[k] for k in orden]


def _agrupar_por_servicio_madre(filas):
    """
    Agrupa filas PENDIENTE_APLICAR de Fase 2 por servicio madre
    (location+cod_destino+supplier_pcm+service_type derivado del nombre del
    PCM), preservando el orden de primera aparición. Filas para el mismo
    servicio madre pero distinto período quedan juntas para reutilizar la
    búsqueda del producto (no la navegación a RATES, que se repite por
    período — ver _buscar_servicio_madre_fase2).
    """
    grupos, orden = {}, []
    for f in filas:
        nombre_pcm = f.get("pcm") or ""
        partes_pcm = nombre_pcm.split("-")
        stype_madre = partes_pcm[2].strip().upper() if len(partes_pcm) >= 5 else ""
        key = (
            str(f.get("location") or "").strip().upper(),
            str(f.get("cod_destino") or "").strip().upper(),
            str(f.get("supplier_pcm") or "").strip().upper(),
            stype_madre,
        )
        if key not in grupos:
            grupos[key] = []
            orden.append(key)
        grupos[key].append(f)
    return [grupos[k] for k in orden]


# ── MAIN ──────────────────────────────────────────────────────
print("="*60)
print(f"  TOURPLAN NX — VALORIZACIÓN PKG  v{VERSION}  [MODO={MODO}]")
print(f"  📌 VERSION {VERSION}  ·  {VERSION_FECHA}")
print("="*60)

_recien_creado = crear_excel_si_no_existe()
print(f"📄 Excel: {EXCEL_PATH}  (existía: {'NO, se creó plantilla' if _recien_creado else 'sí'})")

if not os.path.exists(EXCEL_PATH):
    raise FileNotFoundError(EXCEL_PATH)

if _recien_creado:
    print("\n" + "="*60)
    print("⛔ NO HABÍA EXCEL — se creó una PLANTILLA vacía (fila de ejemplo).")
    print(f"   Ruta: {EXCEL_PATH}")
    print("   1) Completá tus filas con ESTADO=PENDIENTE (o subí tu Excel a esa")
    print("      ruta). Recordá: Colab borra /content al reiniciar → re-subilo.")
    print("   2) Volvé a correr la celda.")
    print("="*60)
    raise SystemExit("Excel recién creado: cargá tus datos y volvé a correr.")

# Asegurar hoja PCM_Detail con columnas nuevas y comentarios de ayuda en los
# headers, aunque el Excel ya existiera
_wb_tmp = openpyxl.load_workbook(EXCEL_PATH)
_init_pcm_detail_sheet(_wb_tmp)
_aplicar_comentarios_excel(_wb_tmp)
_wb_tmp.save(EXCEL_PATH); _wb_tmp.close()

# Flag global: mostrar capturas inline (default NO → corre más rápido)
MOSTRAR_CAPTURAS = leer_flag_mostrar_capturas()
print(f"🖼  Mostrar capturas inline: {'SÍ' if MOSTRAR_CAPTURAS else 'NO'} "
      f"(columna MOSTRAR CAPTURAS)")

# ── Recolectar trabajo de cada fase ───────────────────────────
pendientes_f1  = leer_pendientes()          if MODO in ("LEER",    "COMPLETO") else []
pendientes_f2  = leer_pcm_detail_pendientes_fase2() if MODO in ("APLICAR", "COMPLETO") else []

print(f"📊 Filas PENDIENTE Fase 1: {len(pendientes_f1)}")
print(f"📊 PCMs  PENDIENTE Fase 2: {len(pendientes_f2)}")

# Diagnóstico: mostrar EXACTAMENTE lo que se leyó del Excel (para detectar si no
# está leyendo el archivo correcto o una columna corrida).
for f in pendientes_f1:
    print(f"   · F1 fila {f.get('_row')}: location={f.get('location')!r} "
          f"supplier={f.get('supplier')!r} cod_origen={f.get('cod_origen')!r} "
          f"service_type={f.get('service_type')!r} price_code={f.get('price_code')!r} "
          f"rate={f.get('rate_from')!r}–{f.get('rate_to')!r}")

if not pendientes_f1 and not pendientes_f2:
    print("\n" + "="*60)
    print("⛔ EL EXCEL NO TIENE FILAS PENDIENTE.")
    print(f"   Ruta: {EXCEL_PATH}")
    print("   Poné ESTADO=PENDIENTE en las filas a procesar (hoja PRODUCTOS para")
    print("   Fase 1) y volvé a correr. Verificá que estás editando ESTE archivo.")
    print("="*60)
    raise SystemExit("Sin filas PENDIENTE: cargá ESTADO=PENDIENTE y volvé a correr.")
else:
    driver = crear_driver()
    _abortado = False
    try:
        login(driver)

        # ── Fase 1: leer PCMs y guardar en Excel ──────────────
        if pendientes_f1:
            print(f"\n{'─'*60}")
            print(f"  ▶▶ FASE 1 — Leyendo {len(pendientes_f1)} productos")
            print(f"{'─'*60}")
            for grupo in _agrupar_por_componente(pendientes_f1):
                chequear_abort()
                fila_ref = grupo[0]
                cod_ref = str(fila_ref.get("cod_origen") or "").strip()
                try:
                    handle_prod, dias_op, dias_str, pcm_list = \
                        _buscar_componente_fase1(driver, fila_ref)
                except Exception as e:
                    print(f"\n❌ ERROR F1 (búsqueda de componente) {cod_ref}: {e}")
                    ss(driver, f"error_f1_busqueda_{cod_ref[:12]}")
                    for fila in grupo:
                        escribir_resultado(fila["_row"], "ERROR_F1", error=str(e))
                    cerrar_pcm(driver, driver.window_handles[0])
                    continue

                if len(grupo) > 1:
                    print(f"  ℹ️  {len(grupo)} período(s) para el mismo componente "
                          f"{cod_ref} — reutilizando búsqueda y lista de PCMs")

                for fila in grupo:
                    marcar_procesando(fila["_row"])

                # Por PCM (no por período): abre cada PCM una sola vez y le
                # aplica todos los períodos del grupo antes de cerrarlo — ver
                # _procesar_pcm_todos_periodos.
                resultados_por_fila = {fila["_row"]: [] for fila in grupo}
                _last_scroll = 0
                for pcm_info in pcm_list:
                    try:
                        _last_scroll = _procesar_pcm_todos_periodos(
                            driver, pcm_info, grupo, handle_prod, dias_op, dias_str,
                            resultados_por_fila, start_scroll=_last_scroll)
                    except Exception as e:
                        # _procesar_pcm_todos_periodos ya maneja sus propios errores
                        # por PCM/período; esto es un resguardo para no frenar el
                        # resto de los PCMs del grupo ante algo inesperado.
                        print(f"\n❌ ERROR F1 inesperado procesando PCM "
                              f"{pcm_info.get('nombre','?')}: {e}")
                        ss(driver, f"error_f1_pcm_inesperado_{str(pcm_info.get('nombre','x'))[:12]}")
                        try:
                            cerrar_pcm(driver, handle_prod)
                        except Exception:
                            pass

                for fila in grupo:
                    row = fila["_row"]
                    pcms_leidos = resultados_por_fila[row]
                    cod_fila = str(fila.get("cod_origen") or "").strip()
                    if not pcms_leidos:
                        print(f"\n❌ ERROR F1 {cod_fila} (fila {row}): "
                              f"Ningún PCM se pudo leer correctamente")
                        escribir_resultado(row, "ERROR_F1",
                                           error="Ningún PCM se pudo leer correctamente")
                    else:
                        escribir_resultado(
                            row_num=row, estado="FASE1_OK",
                            pcm_ref="\n".join(pcms_leidos), dias_op=dias_str,
                        )
                        print(f"\n✅ {cod_fila} Fase 1 OK — {len(pcms_leidos)} "
                              f"PCMs leídos (fila {row})")

        # ── Fase 2: aplicar rates en servicios madre ──────────
        # Si corrimos Fase 1 en este ciclo, recargar los pendientes de Fase 2
        if MODO == "COMPLETO" and pendientes_f1:
            pendientes_f2 = leer_pcm_detail_pendientes_fase2()
            print(f"\n  📋 PCMs listos para Fase 2: {len(pendientes_f2)}")

        if FASE2_LIMIT and pendientes_f2:
            print(f"  ⚡ FASE2_LIMIT={FASE2_LIMIT} → proceso solo los primeros "
                  f"{FASE2_LIMIT} de {len(pendientes_f2)} PCMs (modo test rápido)")
            pendientes_f2 = pendientes_f2[:FASE2_LIMIT]

        if pendientes_f2:
            print(f"\n{'─'*60}")
            print(f"  ▶▶ FASE 2 — Aplicando rates en {len(pendientes_f2)} PCMs")
            print(f"{'─'*60}")
            for grupo in _agrupar_por_servicio_madre(pendientes_f2):
                chequear_abort()
                ref = grupo[0]
                cod_destino_ref = str(ref.get("cod_destino") or "").strip()
                location_ref = str(ref.get("location") or "").strip()
                supplier_pcm_ref = str(ref.get("supplier_pcm") or "").strip()
                nombre_pcm_ref = ref.get("pcm") or ""
                partes_pcm_ref = nombre_pcm_ref.split("-")
                stype_madre_ref = partes_pcm_ref[2].strip() if len(partes_pcm_ref) >= 5 else None

                try:
                    _buscar_servicio_madre_fase2(
                        driver, cod_destino_ref, supplier_pcm_ref, location_ref,
                        service_type_madre=stype_madre_ref)
                except ProductoNoEncontrado as e:
                    print(f"    ⚠ Servicio madre {cod_destino_ref} no existe en Tourplan "
                          f"→ SVS_MADRE_NO_ENCONTRADO")
                    ss(driver, f"error_f2_busqueda_{cod_destino_ref[:12]}")
                    for pcm_row in grupo:
                        marcar_pcm_detail_row(pcm_row["_row"], "SVS_MADRE_NO_ENCONTRADO", error=str(e))
                    continue
                except Exception as e:
                    print(f"\n❌ ERROR F2 (búsqueda de servicio madre) {cod_destino_ref}: {e}")
                    ss(driver, f"error_f2_busqueda_{cod_destino_ref[:12]}")
                    for pcm_row in grupo:
                        marcar_pcm_detail_row(pcm_row["_row"], "ERROR", error=str(e))
                    continue

                if len(grupo) > 1:
                    print(f"  ℹ️  {len(grupo)} período(s) para el mismo servicio madre "
                          f"{cod_destino_ref} — reutilizando búsqueda")

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

    try:
        from google.colab import files
        files.download(EXCEL_PATH)
        print("📥 Excel descargado.")
    except Exception:
        pass
