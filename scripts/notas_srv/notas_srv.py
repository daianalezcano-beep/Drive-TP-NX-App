# ============================================================
# TOURPLAN NX — NOTAS SRV (Product Notes)
# Inserta/edita notas de traducción/descriptivo en productos, SIN pisar
# notas existentes salvo que EDICION="SI".
# Copia adaptada para tp-nx-app (app local). El original para Google Colab
# vive sin cambios en el repo copy-products, rama notas-SRV. Mismos
# cambios que el resto de los scripts vendorizados: config por variables
# de entorno, Chrome delegado a common/chrome_bootstrap.py, sin
# --headless (corre con pantalla, no en el contenedor de Colab), perfil
# de Chrome aislado, log en carpeta temporal del OS. El resto (selectores,
# login, navegación a Product Notes, lectura/escritura de notas, CKEditor)
# es igual — ver ESTADO.md en el repo origen para el historial completo de
# bugs y fixes ya resueltos antes de tocar esa parte.
# ------------------------------------------------------------
#   VERSION : 1.21
#   FECHA   : 2026-09-08
# ============================================================

VERSION       = "1.21"
VERSION_FECHA = "2026-09-08"

# ── CAMBIOS ──
# v1.0: primera versión.
# v1.1: corregida abrir_product_notes() (estaba anidada mal en el menú) y
#       seleccionar_codigo_en_dialog() (usaba tr.selectedRow, que no aplica
#       a este dialog) — ambas con datos de una grabación real (Chrome
#       DevTools Recorder) tras el error "no se pudo abrir Product Notes"
#       en la primera corrida.
# v1.2: abrir_product_notes() reescrita para buscar "Content"/"Product
#       Notes" por TEXTO en vez de por posición (nth-of-type) — el menú de
#       primer nivel varía según los módulos habilitados de cada producto,
#       confirmado en corrida real donde "Content" apareció en 7mo lugar en
#       vez de 3ro. También corrige un StaleElementReferenceException por
#       leer .text de un elemento después de clickearlo.
# v1.3: "Content" se expandía pero "Product Notes" no aparecía adentro en
#       corrida real contra producción. Se agrega poll (en vez de sleep
#       fijo de 1seg) para el render del submenú, más lento en producción
#       que en Test; y se evita re-clickear "Content" si ya está expandido
#       (el click es un toggle — si el estado del menú persistió de una
#       fila anterior, volver a clickearlo lo cerraba).
# v1.4: causa real encontrada con un dump completo del DOM: "Content"
#       nunca se expandía porque el click caía en `.menu-heading` (el
#       contenedor con el texto) en vez de en `.click-area` (un <div>
#       hijo puesto ahí específicamente para el manejador de click de
#       Angular) — `li.querySelector('div div')` matcheaba el contenedor
#       equivocado por un detalle no obvio de cómo evalúa ancestros un
#       selector descendiente. Corregido usando `:scope > div` para
#       bajar un nivel por vez (heading, después click-area dentro de
#       heading), igual que el selector confirmado por la grabación
#       (`li > div > div`).
# v1.5: corrida real en MODO="aplicar" llegó hasta escribir la nota y
#       falló con "Cannot read properties of undefined (reading
#       'getSelection')" — el editor CKEDITOR no estaba realmente listo
#       (status !== 'ready', su document interno recién armándose) cuando
#       se llamaba a setData()/fire('change'). esperar_editor() ahora
#       también espera el status 'ready' de la instancia, no solo que el
#       elemento #noteeditorview exista en el DOM. escribir_ckeditor() y
#       leer_ckeditor() quedan además con try/catch por paso, para
#       diagnosticar más preciso si vuelve a fallar.
# v1.6: se agrega activar_editor_click() — clickea el cuadro de texto
#       (el <iframe> del editor) y llama a editor.focus() vía la API
#       antes de escribir, replicando lo que haría una persona para
#       habilitar la edición. Algunos editores WYSIWYG no quedan
#       realmente interactivos hasta el primer click/foco aunque CKEDITOR
#       ya reporte status 'ready' — esto puede ser la pieza que faltaba
#       si el fix de v1.5 no alcanza solo.
# v1.7: corrida real con diagnóstico por paso confirmó que el error de
#       getSelection ocurre específicamente dentro de
#       CKEDITOR.instances[...].setData() — ni el fix de v1.5 (esperar
#       status ready) ni el de v1.6 (click/focus) lo resolvieron, porque
#       el problema está adentro de la propia API de CKEDITOR (bug
#       conocido de CKEditor 4 clásico con navegadores modernos: setData
#       intenta guardar un snapshot de la selección para el historial de
#       deshacer, y esa lectura puede fallar). USAR_API_CKEDITOR pasa a
#       False por default: ahora se escribe/lee directo en el <body> del
#       iframe del editor, bypaseando la API de CKEDITOR por completo
#       (el plan B que ya anticipaba la especificación original). Riesgo
#       no descartado: al bypasear la API también se bypasea su lógica de
#       Plain Text vs HTML — revisar manualmente el formato de la primera
#       nota que se guarde así.
# v1.8: la usuaria confirmó por captura de pantalla real que el editor de
#       una nota NUEVA (INSERT) para estos códigos ya es texto plano
#       directamente, sin ningún selector de formato en la vista. Se
#       elimina forzar_formato_plain_text() (buscaba un control que nunca
#       existió, al menos para INSERT) — código muerto que solo generaba
#       una advertencia confusa en cada corrida.
# v1.9: primera corrida real sobre el batch completo (305 filas) —
#       ¡INSERT funciona de punta a punta! Pero falla intermitentemente
#       (no siempre, no en el mismo producto/código) con "No apareció
#       fila filtrada" en seleccionar_codigo_en_dialog(). Causa: carrera
#       entre escribir el filtro y que la lista completa del diálogo
#       termine de cargar — si se escribe antes de tiempo, el filtro
#       queda aplicado sobre una lista vacía y no se reintenta solo.
#       Se agrega esperar a que la lista completa tenga filas antes de
#       filtrar, más un reintento (borrar y reescribir) si igual sale
#       vacío.
# v1.10: fix preventivo (detectado al armar exportar_notas.py, mismo
#        patrón compartido) — la tabla de Product Notes usa Angular CDK
#        virtual scroll cuando el producto tiene varias notas cargadas, y
#        parsear_notas() solo leía las filas ya renderizadas en el DOM
#        sin scrollear. Con pocos códigos (los 5 de Nota SRV) el síntoma
#        práctico era leve — como el diálogo de INSERT ya excluye
#        categorías usadas, en la práctica esto se manifestaba como un
#        "No apareció fila filtrada" en vez de una inserción duplicada —
#        pero igual podía dar una detección incorrecta de "YA EXISTE" en
#        productos con más notas. Reemplazado por buscar_nota_por_codigo():
#        scrollea progresivamente el contenedor scrolleable (cascada:
#        viewport CDK → clase confirmada .productnoteslistview → ancestro
#        con overflow → fallback documento), releyendo tbody tr en cada
#        paso (nunca cachea WebElements) y comparando la columna CAT
#        contra el código buscado, con salida temprana tras 8 pasos sin
#        códigos nuevos vistos. Basado en el skill
#        recorriendo-grillas-virtuales-de-tourplan.
# v1.11: el mismo fix (v1.10) se aplicó primero en exportar_notas.py y
#        seguía dando falso "no existe" en corrida real (código UES sobre
#        IGR/1RIOTU/PARBR1). Causa: _detectar_contenedor_scroll() buscaba
#        el primer `cdk-virtual-scroll-viewport` de TODA la página con
#        `document.querySelector` — si hay cualquier otro viewport CDK en
#        pantalla sin relación con esta tabla, se scrollea ese elemento
#        equivocado y nunca aparece ninguna fila nueva. Corregido acá
#        también (mismo código compartido): la búsqueda del viewport CDK
#        y de `.productnoteslistview` ahora está limitada a los
#        ANCESTROS de la tabla, nunca al documento completo. Se agrega
#        además diagnóstico por consola: contenedor detectado (tag,
#        clase, scrollHeight/clientHeight) y, si el código no aparece, la
#        lista completa de códigos CAT vistos durante el recorrido.
# v1.12: el diagnóstico de v1.11 (aplicado primero en exportar_notas.py)
#        reveló la causa REAL, distinta de las hipótesis de v1.10/v1.11:
#        un dump completo del DOM real (producto IGR/1RIOTU/PARBR1,
#        código UES) confirmó que Product Notes NO usa Angular CDK
#        virtual scroll — no hay ningún `cdk-virtual-scroll-viewport` en
#        la página, y las ~20 notas del producto (incluyendo UES) están
#        TODAS presentes en el DOM a la vez. El diagnóstico de "virtual
#        scroll" de v1.10/v1.11 era incorrecto. Los dos bugs reales: (1)
#        encontrar_tabla_notas() devolvía el <table> del HEADER
#        (CAT/DESCRIPTION), que no tiene ningún <tbody> con filas — los
#        datos viven en tablas SEPARADAS, una por nota, envueltas cada
#        una en <div class="tprow-even"/"tprow-odd">, hermanas dentro de
#        un contenedor común (.productnoteslistview) — por eso "tbody tr"
#        nunca encontraba nada. (2) _fila_a_nota() leía columnas por
#        índice posicional de <td>, pero cada fila real tiene más <td>
#        (celdas de expansor + ícono de lápiz intercaladas) — el mapeo de
#        columnas era incorrecto incluso si se hubiese encontrado la
#        fila. Corregido: encontrar_tabla_notas() devuelve el contenedor
#        común, no la tabla de header; _fila_a_nota() lee cada campo por
#        su clase CSS estable (td.tpcol-category, td.tpcol-description,
#        etc.). Se elimina la lógica de detección de contenedor
#        scrolleable y scroll progresivo — no resolvía nada real.
# v1.13: bug real distinto reportado por la usuaria (mismo patrón ya visto
#        antes en Copy Products al completar Location): en
#        seleccionar_codigo_en_dialog(), tras filtrar el diálogo de INSERT
#        por el código de nota, el código tomaba SIEMPRE la primera fila
#        filtrada (`candidatos[0]`) asumiendo que era el código exacto
#        buscado — pero el filtro del diálogo no garantiza esa posición
#        (puede matchear por substring contra la descripción u otros
#        códigos parecidos), así que podía insertar la nota bajo un
#        código equivocado sin ningún error visible. Corregido: ahora se
#        busca, entre las filas que deja visibles el filtro, la que tiene
#        una celda con texto EXACTO (case-insensitive) igual a
#        codigo_nota — mismo patrón (`matchRow`, comparar contra
#        `td.children.length===0`) que ya usa buscar_producto() en este
#        archivo para el mismo tipo de problema al matchear código de
#        producto.
# v1.14: mismo tipo de bug (asumir posición en vez de código exacto),
#        encontrado al revisar el resto de los scripts vendorizados tras
#        el fix de v1.13: en buscar_producto(), cuando se completa
#        Location, si la fila que Tourplan marca como `tr.selectedRow`
#        (el match que el propio Tourplan resalta) no aparece a tiempo,
#        el fallback clickeaba SIEMPRE la primera fila de la tabla de
#        sugerencias (`rows[0]`) sin verificar que fuera el Location
#        exacto buscado. Corregido con el mismo patrón que v1.13: buscar,
#        entre las filas sugeridas, la que tiene una celda con texto
#        EXACTO (case-insensitive) igual al Location pedido; si ninguna
#        coincide, no se clickea nada a ciegas (se avisa por consola en
#        vez de arriesgar seguir con un Location equivocado). Mismo
#        cambio portado a exportar_notas.py, tourplan_copiar_pcm.py,
#        valorizacion_desde_excel.py, valorizacion_desde_madre.py (dos
#        veces: buscar_producto() y buscar_servicios_madre()) y
#        tourplan_valorizacion_pkg_v3.py.
# v1.15: reportado por la usuaria: la nota se editaba en pantalla pero no
#        quedaba guardada. Causa: click_save() para nota NUEVA (INSERT)
#        usaba wait_click() con solo 8seg y sin reintento, bajo el
#        supuesto (desde v1.9) de que ahí el botón SAVE viene habilitado
#        desde el vamos — supuesto que dejó de cumplirse (el botón puede
#        tardar unos segundos en habilitarse también en INSERT). Ahora
#        click_save() usa siempre el mismo polling activo
#        (wait_save_habilitado(), ya usado para EDIT) para ambos casos,
#        con SAVE_POLL_TIMEOUT_INSERT=15 como timeout propio de INSERT.
# v1.16: la usuaria reportó que el error de guardado lo vio en EDIT, no en
#        INSERT (donde apuntaba el fix de v1.15) — sospecha de timing en
#        SAVE. Al revisar wait_save_habilitado() se encontró que solo
#        chequea el PRIMER botón SAVE encontrado (elems[0]) en toda la
#        página, mientras que la misma función en otros scripts
#        vendorizados (modificar_description_comment.py,
#        flag_products_as_deleted.py) recorre TODOS los botones SAVE
#        encontrados — puede haber más de uno en pantalla (ver skill
#        escribiendo-en-formularios-angular-de-tourplan, ambigüedad de
#        múltiples SAVE). Todavía no confirmado con datos reales si ese es
#        el problema en EDIT, así que NO se cambió esa lógica todavía. Se
#        agrega dump_save_candidates() (solo diagnóstico, no cambia
#        comportamiento): si click_save() agota el timeout sin encontrar
#        el SAVE habilitado, ahora el mensaje de ERROR que queda en el
#        Excel (columna ESTADO/DETALLE_PROCESO) incluye la lista completa
#        de botones SAVE vistos en pantalla (visible/disabled/si están
#        dentro de un dialog) — con eso se podrá confirmar en la próxima
#        corrida real si el bug es la ambigüedad de botones o algo distinto.
# v1.17: la usuaria confirmó el patrón sospechado desde el principio: en una
#        corrida fallida, el script reportó "EDITADA" (éxito) pero el
#        contenido NO había quedado guardado. Causa raíz: process_nota()
#        daba por bueno el guardado con solo que esperar_cierre_editor()
#        confirmara que el editor se cerró — eso prueba que Tourplan
#        aceptó CERRAR el diálogo, no que haya persistido el cambio. Se
#        agrega verificar_guardado(): tras cerrar el editor, relee la nota
#        desde la tabla (misma buscar_nota_por_codigo() que ya se usa para
#        detectar si existe) — en INSERT alcanza con confirmar que la nota
#        ahora existe; en EDIT además exige que UPDATED/UPDATED BY hayan
#        CAMBIADO respecto a antes de guardar. Si la verificación falla,
#        ahora se reporta ERROR (con el motivo) en vez de INSERTADA/EDITADA.
# v1.20: entre v1.17 y v1.19 (ambas revertidas a pedido de la usuaria) se
#        probaron variantes de verificar_guardado() que decidían
#        automáticamente ERROR vs EDITADA/INSERTADA comparando
#        UPDATED/UPDATED BY — sin terminar de coincidir con lo que la
#        usuaria veía revisando Tourplan directo. Se pausa esa decisión
#        automática. En su lugar, verificar_guardado() se reemplaza por
#        leer_estado_nota_fresco(): SOLO exporta el dato (no decide nada,
#        nunca marca ERROR por esto). Fuerza la misma recarga real que usa
#        exportar_notas.py (buscar_producto()+abrir_product_notes(),
#        navegación real vía driver.get(), no la tabla ya cargada en el
#        navegador) y agrega al DETALLE_PROCESO del Excel el CREATED/
#        UPDATED/UPDATED BY leído así — en EDIT, además del valor de ANTES
#        de guardar — para que la usuaria compare fila por fila contra
#        Tourplan real y confirme si esta fecha es o no un indicador
#        confiable, antes de volver a automatizar la decisión.
# v1.21: la usuaria reportó que, con el cambio de v1.20, Crear/Modificar
#        Notas dejó de encadenar las notas del mismo producto (volvía a
#        buscar_producto() por cada nota individual), a diferencia de
#        Exportar Notas que sí encadena — corridas grandes tardaban mucho
#        más. Causa: la relectura de validación (antes leer_estado_nota_
#        fresco(), llamada desde DENTRO de process_nota() apenas se
#        guardaba cada nota) forzaba su propia recarga real de la página
#        por cada nota, sin enterarse de que ya había otras notas
#        pendientes del mismo producto — rompía el agrupamiento de
#        _agrupar_por_producto() que ya hace MAIN() para el guardado.
#        Separado en dos pasadas independientes, ambas encadenadas por
#        producto: (1) MAIN() guarda TODAS las notas de todos los
#        productos primero, como en v1.16 (sin ninguna relectura de
#        validación en el medio) — igual que antes de v1.20; (2) recién
#        terminado TODO el guardado, verificar_notas_guardadas() reagrupa
#        por producto (misma _agrupar_por_producto()) las filas que
#        quedaron INSERTADA/EDITADA y, por cada producto, hace UNA sola
#        recarga real (no una por nota) para releer CREATED/UPDATED/
#        UPDATED BY de todas sus notas y agregarlo al DETALLE_PROCESO ya
#        escrito — mismo patrón de encadenamiento que ya usa
#        exportar_notas.py. process_nota() ahora solo guarda el
#        UPDATED/UPDATED BY de ANTES de editar en el propio dict de la
#        fila (_updated_antes/_updated_by_antes) para que la segunda
#        pasada arme el detalle ANTES/DESPUÉS sin tener que volver a leerlo.

# ── CONFIGURACIÓN — via variables de entorno (con default = valor original) ──
import os

from common.user_config import (
    CREDENTIALS_PATH as _CREDENTIALS_PATH_DEFAULT,
    TOKEN_PATH as _TOKEN_PATH_DEFAULT,
)

MODO       = os.environ.get("TOURPLAN_MODO", "lectura")  # "lectura" (dry-run, no escribe nada en Tourplan)
                                                           # "aplicar"/cualquier otro valor = ejecuta los cambios
EDICION    = os.environ.get("TOURPLAN_EDICION", "NO")    # Global para toda la corrida, NO varía por fila.
                           # "NO": si la nota ya existe, no se toca — solo se
                           #       reporta para revisión manual.
                           # "SI": si la nota ya existe, se sobreescribe —
                           #       pero antes se lee y guarda el contenido
                           #       anterior en CONTENIDO_ANTERIOR.

SHEET_URL        = os.environ.get("TOURPLAN_SHEET_URL", "")
HOJA             = os.environ.get("TOURPLAN_HOJA", "NOTAS_SRV")
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

# Camino de escritura/lectura de la nota. False (default, desde v1.7):
# escribe directo en el <body> del iframe del editor — la API de CKEDITOR
# (setData()) falla de forma consistente en corrida real contra Tourplan
# real ("Cannot read properties of undefined (reading 'getSelection')").
# True: usa la API de CKEDITOR, con fallback automático al camino directo
# si falla — dejarlo como opción por si una futura versión de Tourplan
# corrige el bug.
USAR_API_CKEDITOR = False

# ── PASO 0: Entorno ───────────────────────────────────────────
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

# ── Códigos de nota soportados por este flujo ─────────────────
# Solo se insertan/editan acá códigos con Category Type = "Product" y
# Format Type = "Plain Text" (ver sección "Formato Plain Text vs HTML" en la
# descripción de la tarea — el Format Type del catálogo es solo el DEFAULT
# de Tourplan, no una garantía; igual se fuerza explícitamente en el editor).
# Si se agregan códigos nuevos, confirmar sus atributos reales en Tourplan
# antes de sumarlos — no asumir por el nombre.
CODIGOS_NOTA_VALIDOS = {
    # codigo: (descripcion, category_type, format_type_default_tourplan)
    "NAL": ("Nota SRV Aleman",   "Product", "Plain Text"),
    "NES": ("Nota SRV Español",  "Product", "Plain Text"),
    "NFR": ("Nota SRV Frances",  "Product", "Plain Text"),
    "NIN": ("Nota SRV Ingles",   "Product", "Plain Text"),
    "NIT": ("Nota SRV Italiano", "Product", "Plain Text"),
    # Códigos "Descriptivo" — confirmados por la usuaria con captura real
    # del catálogo de Note Category de Tourplan (Category Type = Product,
    # Format Type = Plain Text, igual que los 5 de arriba).
    "UAL": ("Descriptivo Aleman",  "Product", "Plain Text"),
    "UES": ("Descriptivo Español", "Product", "Plain Text"),
    "UFR": ("Descriptivo Frances", "Product", "Plain Text"),
    "UIN": ("Descriptivo Ingles",  "Product", "Plain Text"),
    "UIT": ("Descriptivo Italiano","Product", "Plain Text"),
    # Resto de los códigos Plain Text del mismo catálogo (misma captura,
    # 2026-08-28) — todos Category Type = Product, Format Type = Plain
    # Text. Los códigos HTML del mismo catálogo (OBS, NAP, SDV, Hyperlinks,
    # Tax Brasil, etc.) quedan deliberadamente AFUERA: el mecanismo de
    # escritura de este script (bypass directo del <body contenteditable>
    # del iframe, sin pasar por la API de CKEDITOR) está confirmado en
    # corrida real solo para editores Plain Text — nunca se probó contra
    # un editor HTML (que tiene su propia barra de formato), así que no se
    # asume que escribiría bien ahí.
    "DRT": ("Direccion Rent a Car",         "Product", "Plain Text"),
    "LWE": ("Luggage Waiver Ingles",        "Product", "Plain Text"),
    "LWF": ("Luggage Waiver Frances",       "Product", "Plain Text"),
    "LWG": ("Luggage Waiver Aleman",        "Product", "Plain Text"),
    "LWI": ("Luggage Waiver Italian",       "Product", "Plain Text"),
    "LWS": ("Luggage Waiver Español",       "Product", "Plain Text"),
    "PCR": ("Producto Coordinates",         "Product", "Plain Text"),
    "REM": ("Remodelacion Hotel",           "Product", "Plain Text"),
    "REO": ("Nota Cliente Solo Voucher",    "Product", "Plain Text"),
    "SC2": ("External Option -Mapeo Especif","Product", "Plain Text"),
    "TAL": ("Titulo Aleman",                "Product", "Plain Text"),
    "TES": ("Titulo Español",               "Product", "Plain Text"),
    "TFR": ("Titulo Frances",               "Product", "Plain Text"),
    "TIN": ("Titulo Ingles",                "Product", "Plain Text"),
    "TIT": ("Titulo Italiano",              "Product", "Plain Text"),
}

SEL = {
    "NAV_ICON":    "nav img",
    "INSERT_BTN":  "div.actionbuttons button",
    "DIALOG":      "tp-dialog",
    "NOTE_EDITOR": "#noteeditorview",
    "SAVE_BTN":    "tp-button.save > button",
    "EDIT_ICON":   "i.fa-pencil-square-o.note-icon",
}

SAVE_POLL_INTERVAL = 0.4
SAVE_POLL_TIMEOUT_EDIT = 25  # seg — nota existente: el botón tarda unos
                              # segundos en habilitarse tras el cambio.
SAVE_POLL_TIMEOUT_INSERT = 15  # seg — nota nueva: aunque se asumía que el
                                 # botón viene habilitado desde el vamos
                                 # (v1.14 y antes), reportado en corrida real
                                 # que también puede demorar unos segundos —
                                 # ahora usa el mismo polling activo que EDIT.


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


def wait_save_habilitado(driver, save_selector, timeout, interval=SAVE_POLL_INTERVAL):
    """Polling activo del botón SAVE — no un sleep fijo. Distingue 'no
    existe en el DOM' de 'existe pero nunca se habilitó'."""
    fin = time.time() + timeout
    encontrado_alguna_vez = False
    while time.time() < fin:
        elems = driver.find_elements(By.CSS_SELECTOR, save_selector)
        if elems:
            encontrado_alguna_vez = True
            try:
                if elems[0].is_displayed() and elems[0].is_enabled():
                    return elems[0], None
            except Exception:
                pass
        time.sleep(interval)
    if not encontrado_alguna_vez:
        return None, "el botón SAVE no se encontró en la página"
    return None, "el botón SAVE nunca se habilitó (timeout)"


def dump_save_candidates(driver):
    """Diagnóstico (no cambia ningún comportamiento de guardado): lista
    TODOS los botones cuya clase o texto contiene 'save', con su estado
    real (visible/disabled) y si viven dentro de un tp-dialog. Se agrega
    para confirmar con datos reales si un timeout de SAVE en EDIT se debe
    a que hay más de un botón SAVE en pantalla y wait_save_habilitado()
    (que hoy solo mira el primero, elems[0]) está esperando el equivocado
    — mismo tipo de ambigüedad ya documentada en otros scripts vendorizados
    (ver skill escribiendo-en-formularios-angular-de-tourplan). Point de
    partida para decidir si hace falta corregir wait_save_habilitado()."""
    try:
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
                var inDialog = !!el.closest('tp-dialog, tp-modal, [role="dialog"]');
                out.push(el.tagName.toLowerCase() + '.' + cls + '=' +
                          JSON.stringify(txt.slice(0, 40)) +
                          ' visible=' + vis + ' disabled=' + !!el.disabled +
                          ' inDialog=' + inDialog);
            }
            return out;
        """)
    except Exception as e:
        return [f"(no se pudo inspeccionar: {e})"]


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
    Flujo (ver skill buscando-productos-en-tourplan):
    1. #/home → #/product (siempre pasar por #/home antes).
    2. Click lupa → modal Product Search (tabs SELECTION/RESULTS).
    3. Service Type en el sidebar izquierdo (prefijo numérico, no sigla).
    4. Location → Supplier → Code.
    5. SEARCH → click en la fila de resultado que matchea código+service type.
    6. Verificar contexto abriendo el menú hamburguesa (UTILITIES/RATES/...).

    Se usa esta función (y no buscar_y_abrir_option) porque el script
    necesita ENTRAR de lleno al registro para llegar al submenú Product
    Notes — no alcanza con clickear la fila de resultados.
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
    buscar_producto(). "Product Notes" vive anidado dentro de "Content"
    (confirmado por grabación + por la usuaria), pero NO se puede ubicar
    "Content" por posición fija (`nth-of-type`): el menú de primer nivel
    varía según qué módulos tiene habilitados cada producto — en una
    corrida real, un producto mostró GENERAL/POLICIES/PICKUP POINTS ANTES
    de CONTENT, corriendo su posición del 3er al 7mo lugar. Por eso se
    busca siempre por TEXTO ("Content", luego "Product Notes" adentro),
    nunca por índice.

    Además, cada búsqueda hace todo (encontrar + click) en un único
    `execute_script`, sin guardar la referencia del WebElement para
    usarla después en Python — leer `.text` de un elemento DESPUÉS de
    clickearlo (cuando el click ya disparó una navegación/re-render de
    Angular) tira `StaleElementReferenceException` (confirmado en corrida
    real, 2026-08).
    """
    nav_icon = wait(driver, SEL["NAV_ICON"], t=timeout)
    jc(driver, nav_icon)
    time.sleep(1.5 * VELOCIDAD)

    def _click_item_de_nivel(rx_texto):
        """
        Busca por texto entre los ítems de PRIMER NIVEL del menú.

        CONFIRMADO con un dump real del DOM (2026-08): cada <li> de primer
        nivel tiene esta estructura:
            <li>
              <div class="menu-heading">     <- contiene el texto (label)
                <div class="click-area"></div>  <- ACÁ hay que clickear
                <label>Content</label>
              </div>
            </li>
        `li.querySelector('div div')` (usado en versiones previas de este
        helper, y en _en_contexto_producto() de buscar_producto() solo
        para LEER texto) en realidad devuelve `.menu-heading` — no
        `.click-area` — porque `.menu-heading` YA cuenta como "div con
        ancestro div" gracias a un div contenedor fuera del propio <li>
        (un detalle no obvio de cómo evalúa ancestros un selector
        descendiente con querySelector). Leer el texto de `.menu-heading`
        funciona bien (por eso _en_contexto_producto() nunca mostró el
        bug), pero CLICKEARLO no hace nada — el manejador de click real
        de Angular está en `.click-area`, un <div> hijo directo puesto
        ahí a propósito como área de click. Por eso "Content" nunca se
        expandía pese a que el texto se encontraba bien.

        Por eso acá se usa `:scope > div` (hijo directo) dos veces: una
        para llegar a `.menu-heading` (leer texto) y otra para llegar a
        `.click-area` dentro de ese (clickear) — en vez del descendiente
        genérico `div div`.
        """
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
        """Si 'Content' ya está expandido (ej. quedó así de una fila
        anterior — la navegación entre filas es por hash dentro de la
        misma SPA, no una recarga completa, así que el estado del menú
        puede persistir), NO hay que re-clickearlo: el click es un toggle
        y lo volvería a cerrar."""
        return bool(driver.execute_script("""
            var abierto = document.querySelector('li.tpnavmenuopen');
            if (!abierto) return false;
            var heading = abierto.querySelector(':scope > div') || abierto;
            var t = (heading.innerText || '').trim().split('\\n')[0];
            return /^\\s*content\\s*$/i.test(t);
        """))

    def _esperar_y_click_submenu(rx_texto, timeout_sub=8):
        """Poll en vez de un sleep fijo — Tourplan puede tardar más en
        renderizar el submenú en producción que en Test."""
        fin = time.time() + timeout_sub * VELOCIDAD
        while time.time() < fin:
            encontrado = _click_en_submenu_abierto(rx_texto)
            if encontrado:
                return encontrado
            time.sleep(0.3)
        return None

    # Atajo: por si "Product Notes" ya está visible sin expandir nada (el
    # menú puede recordar el último submenú abierto de una fila anterior).
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
    presentes en el DOM al mismo tiempo, sin necesidad de scrollear nada.
    El intento de scroll era una solución para un problema que no
    existía. Alcanza con leer `tbody tr` directamente del contenedor
    devuelto por `encontrar_tabla_notas()` (que ya lo ubica
    correctamente, ver ese fix).

    Devuelve (nota_o_None, vistos_ordenados) — vistos_ordenados es la
    lista de todos los códigos CAT realmente encontrados, para poder
    diagnosticar sin adivinar si vuelve a fallar."""
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


# ── INSERT de nota nueva ───────────────────────────────────────
def click_insert(driver):
    btn = driver.execute_script("""
        var btns = Array.from(document.querySelectorAll('div.actionbuttons button'));
        for (var b of btns) {
            if (!b.offsetParent) continue;
            var t = (b.innerText || '').trim().toUpperCase();
            if (t === 'INSERT') return b;
        }
        return null;
    """)
    if btn is None:
        raise Exception("Botón INSERT no encontrado en div.actionbuttons")
    jc(driver, btn)
    time.sleep(1.5 * VELOCIDAD)


def seleccionar_codigo_en_dialog(driver, codigo_nota, timeout=10, reintentos=2):
    """
    Espera a que la lista completa (sin filtrar) del diálogo tenga al
    menos una fila ANTES de escribir el filtro — evita la carrera
    confirmada en corrida real (2026-08, batch de 305 filas): si se
    escribe el filtro antes de que la lista termine de cargar, el filtro
    se aplica sobre una lista todavía vacía y Tourplan no lo reaplica solo
    cuando los datos llegan — la lista visible queda vacía para siempre en
    esa apertura del diálogo. Reintento (borrar y reescribir) por si la
    carrera ocurre de todos modos.

    BUG CONFIRMADO (2026-09): entre las filas que el filtro deja
    visibles, NO alcanza con clickear la PRIMERA (`candidatos[0]`) — el
    filtro del diálogo no garantiza que el código exacto buscado quede
    primero (puede matchear por substring contra descripción u otros
    códigos parecidos), igual que el bug ya visto antes en Copy Products
    al completar Location. Se busca la fila cuyo texto de celda coincide
    EXACTO (case-insensitive) con `codigo_nota` — mismo patrón ya usado en
    buscar_producto() de este archivo (`matchRow`) para evitar el mismo
    problema al matchear código de producto — en vez de asumir la
    posición.
    """
    dlg = wait(driver, SEL["DIALOG"])
    inp = dlg.find_element(By.CSS_SELECTOR, "input")

    fin_carga = time.time() + timeout * VELOCIDAD
    while time.time() < fin_carga:
        if dlg.find_elements(By.CSS_SELECTOR, "table tbody tr"):
            break
        time.sleep(0.2)

    row = None
    for intento in range(reintentos):
        inp.click()
        if intento > 0:
            set_val(driver, inp, "")
            time.sleep(0.3 * VELOCIDAD)
        set_val(driver, inp, codigo_nota.lower())
        time.sleep(1.0 * VELOCIDAD)
        fin = time.time() + timeout * VELOCIDAD
        while time.time() < fin:
            row = driver.execute_script("""
                var dlg = arguments[0], cod = arguments[1].toUpperCase();
                var trs = Array.from(dlg.querySelectorAll('table tbody tr'));
                for (var tr of trs) {
                    var tds = Array.from(tr.querySelectorAll('td'));
                    var exacto = tds.some(function(td){
                        return td.children.length === 0 &&
                               td.innerText.trim().toUpperCase() === cod;
                    });
                    if (exacto) {
                        return tr.querySelector('td.description') || tr;
                    }
                }
                return null;
            """, dlg, codigo_nota)
            if row is not None:
                break
            time.sleep(0.3)
        if row is not None:
            break
        print(f"    ⚠ Intento {intento + 1}/{reintentos}: no apareció fila con código "
              f"exacto '{codigo_nota}' — reintentando")

    if row is None:
        raise Exception(f"No apareció fila con código exacto '{codigo_nota}' "
                         f"tras {reintentos} intentos")
    jc(driver, row)
    time.sleep(1.5 * VELOCIDAD)


def esperar_editor(driver, timeout=15):
    """
    Espera a que aparezca #noteeditorview Y a que la instancia de CKEDITOR
    esté realmente lista (status === 'ready', con su document interno ya
    armado) antes de devolver el control.

    BUG CONFIRMADO EN CORRIDA REAL (2026-08): sin este segundo chequeo,
    `escribir_ckeditor()` fallaba con
    `Cannot read properties of undefined (reading 'getSelection')` — el
    elemento #noteeditorview ya estaba en el DOM, pero CKEDITOR todavía no
    había terminado de inicializar su `document` interno (el iframe recién
    cargando), y setData()/fire('change') dependen de poder leer la
    selección actual (usada internamente por el plugin de undo, entre
    otros) para funcionar.
    """
    wait(driver, SEL["NOTE_EDITOR"], t=timeout)
    time.sleep(1.0 * VELOCIDAD)

    fin = time.time() + timeout * VELOCIDAD
    listo = False
    while time.time() < fin:
        listo = driver.execute_script("""
            var nombres = Object.keys(CKEDITOR.instances || {});
            if (!nombres.length) return false;
            var ed = CKEDITOR.instances[nombres[0]];
            return !!(ed && ed.status === 'ready' && ed.document);
        """)
        if listo:
            break
        time.sleep(0.3)
    if not listo:
        print("    ⚠ CKEDITOR no llegó a status 'ready' dentro del timeout — "
              "puede fallar la escritura")


def activar_editor_click(driver):
    """
    Clickea el cuadro de texto de la nota para habilitar su edición, tal
    como haría una persona — algunos editores WYSIWYG no quedan realmente
    interactivos hasta el primer click/foco, aunque CKEDITOR ya reporte
    `status: 'ready'` internamente (ver esperar_editor()).

    Dos caminos, por las dudas uno falle:
    1. Click físico sobre el <iframe> del editor — visible en el
       documento principal, no hace falta switch_to.frame para clickearlo
       desde afuera (solo haría falta para escribir texto adentro con
       send_keys, que no es el camino que usa este script).
    2. `editor.focus()` vía la API de CKEDITOR — el equivalente funcional
       a un click, por si el click físico no alcanza (ej. el iframe no es
       clickeable todavía por algún overlay).
    """
    try:
        iframe = driver.find_element(By.CSS_SELECTOR, f"{SEL['NOTE_EDITOR']} iframe")
        jc(driver, iframe)
        time.sleep(0.3 * VELOCIDAD)
    except Exception:
        pass

    driver.execute_script("""
        var nombres = Object.keys(CKEDITOR.instances || {});
        if (nombres.length) {
            try { CKEDITOR.instances[nombres[0]].focus(); } catch (e) {}
        }
    """)
    time.sleep(0.3 * VELOCIDAD)


def texto_a_html_pre(texto):
    """Un texto de una sola línea no necesita conversión. Si el Excel trae
    saltos de línea reales, convertirlos a <br> (formato Plain Text de
    Tourplan: un <pre> con <br> entre líneas)."""
    partes = texto.replace("\r\n", "\n").split("\n")
    return "<br>".join(partes) if len(partes) > 1 else texto


def leer_iframe_directo(driver):
    """Lee el HTML actual del <body contenteditable> del iframe del
    editor, sin pasar por la API de CKEDITOR."""
    iframe = wait(driver, f"{SEL['NOTE_EDITOR']} iframe")
    driver.switch_to.frame(iframe)
    try:
        return driver.execute_script("return document.body.innerHTML;")
    finally:
        driver.switch_to.default_content()


def escribir_iframe_directo(driver, texto_html):
    """
    Escribe directo en el <body contenteditable="true"> del iframe del
    editor, sin pasar por la API de CKEDITOR.

    BUG CONFIRMADO EN CORRIDA REAL (2026-08): `CKEDITOR.instances[...].
    setData()` falla de forma consistente y reproducible con
    `Cannot read properties of undefined (reading 'getSelection')` —
    incluso con el editor en status 'ready' (esperar_editor()) y tras
    clickear/enfocar el cuadro (activar_editor_click()). Es un problema
    conocido de CKEditor 4 clásico en navegadores modernos: setData()
    intenta internamente guardar un snapshot de la selección para el
    historial de deshacer, y esa lectura de selección puede fallar según
    el estado del iframe. Este camino evita el problema de raíz
    escribiendo el HTML directo en el DOM del iframe (tal como anticipaba
    la especificación original de esta tarea como plan B).

    Nota sobre Plain Text vs HTML: para una nota NUEVA (INSERT), se
    confirmó por captura de pantalla real que el editor de estos códigos
    (NAL/NES/NFR/NIN/NIT) es directamente texto plano, sin ningún selector
    de formato en la vista — no hace falta forzar nada. Por eso se sacó
    el intento de togglear el modo que tenía este script en versiones
    anteriores (buscaba un control que nunca existió, al menos para el
    caso de INSERT). Sin confirmar todavía si una nota EXISTENTE (flujo
    EDIT) puede abrir en un modo distinto — si algún día aparece un
    problema de formato ahí, revisar ese caso puntual.
    """
    iframe = wait(driver, f"{SEL['NOTE_EDITOR']} iframe")
    driver.switch_to.frame(iframe)
    try:
        driver.execute_script("""
            document.body.innerHTML = arguments[0];
            document.body.dispatchEvent(new Event('input', {bubbles: true}));
            document.body.dispatchEvent(new Event('keyup', {bubbles: true}));
        """, texto_html)
    finally:
        driver.switch_to.default_content()

    try:
        driver.execute_script("""
            var nombres = Object.keys(CKEDITOR.instances || {});
            if (nombres.length) { CKEDITOR.instances[nombres[0]].updateElement(); }
        """)
    except Exception as e:
        print(f"    ⚠ No se pudo sincronizar updateElement() tras la escritura directa "
              f"(no necesariamente fatal — Tourplan puede leer directo del iframe al "
              f"guardar): {e}")


def leer_ckeditor(driver):
    if USAR_API_CKEDITOR:
        resultado = driver.execute_script("""
            var nombres = Object.keys(CKEDITOR.instances || {});
            if (!nombres.length) return {ok:false, motivo:'CKEDITOR.instances vacío'};
            // TODO: si hay más de una instancia activa a la vez, confirmar cuál
            // corresponde a la nota actual — hoy se asume nombres[0].
            try {
                return {ok:true, data: CKEDITOR.instances[nombres[0]].getData()};
            } catch (e) {
                return {ok:false, motivo:'getData: ' + (e && e.message || e)};
            }
        """)
        if resultado and resultado.get("ok"):
            return resultado.get("data")
        print(f"    ⚠ No se pudo leer vía API de CKEDITOR "
              f"({(resultado or {}).get('motivo')}) — probando lectura directa del iframe")
    try:
        return leer_iframe_directo(driver)
    except Exception as e:
        print(f"    ⚠ Tampoco se pudo leer directo del iframe: {e}")
        return None


def escribir_ckeditor(driver, texto_html):
    """
    Por defecto escribe directo en el iframe (USAR_API_CKEDITOR=False),
    porque la API de CKEDITOR (setData()) falla de forma consistente en
    corrida real — ver escribir_iframe_directo() para el detalle. Se deja
    la API como camino opcional (con fallback automático al directo si
    falla) por si una futura actualización de Tourplan corrige el bug.
    """
    if USAR_API_CKEDITOR:
        resultado = driver.execute_script("""
            var nombres = Object.keys(CKEDITOR.instances || {});
            if (!nombres.length) return {ok:false, motivo:'CKEDITOR.instances vacío'};
            var editor = CKEDITOR.instances[nombres[0]];
            try {
                editor.setData(arguments[0]);
            } catch (e) {
                return {ok:false, motivo:'setData: ' + (e && e.message || e)};
            }
            try {
                editor.updateElement();
            } catch (e) {
                return {ok:false, motivo:'updateElement: ' + (e && e.message || e)};
            }
            try {
                editor.fire('change');
            } catch (e) {
                return {ok:false, motivo:'fire(change): ' + (e && e.message || e)};
            }
            return {ok:true};
        """, texto_html)
        if resultado and resultado.get("ok"):
            return
        print(f"    ⚠ Falló la API de CKEDITOR "
              f"({(resultado or {}).get('motivo')}) — probando escritura directa en el iframe")
    escribir_iframe_directo(driver, texto_html)


def click_save(driver, con_polling):
    """Ambos casos (nota nueva y nota existente editada) usan polling activo
    del botón SAVE — reportado en corrida real que la nota nueva también
    puede demorar unos segundos en habilitarse, no solo la edición como se
    asumía antes (hasta v1.14: 'SAVE viene habilitado desde el vamos' para
    INSERT, con un wait_click de solo 8seg sin reintento). Se mantiene el
    parámetro con_polling para elegir el timeout correcto según el caso."""
    timeout = SAVE_POLL_TIMEOUT_EDIT if con_polling else SAVE_POLL_TIMEOUT_INSERT
    btn, err = wait_save_habilitado(driver, SEL["SAVE_BTN"], timeout=timeout)
    if btn is None:
        candidatos = dump_save_candidates(driver)
        raise Exception(f"No se pudo guardar: {err} — botones SAVE vistos en pantalla: {candidatos}")
    jc(driver, btn)
    time.sleep(2 * VELOCIDAD)


def esperar_cierre_editor(driver, timeout=10):
    fin = time.time() + timeout * VELOCIDAD
    while time.time() < fin:
        if not driver.find_elements(By.CSS_SELECTOR, SEL["NOTE_EDITOR"]):
            return True
        time.sleep(0.3)
    return False


# ── Sheet — I/O ───────────────────────────────────────────────
def load_sheet(hoja):
    ws = conectar_sheets(SHEET_URL, hoja, CREDENTIALS_PATH, TOKEN_PATH)
    rows, columnas = cargar_sheet(ws)
    print(f"Sheet cargado: {len(rows)} fila(s) en '{hoja}'")
    return ws, rows, columnas


def update_row(ws, row_idx, columnas, estado=None, detalle=None, contenido_anterior=None):
    """Escribe INMEDIATAMENTE tras cada fila — no acumular para el final."""
    valores = {}
    if estado is not None:
        valores["ESTADO"] = estado
    if detalle is not None and "DETALLE_PROCESO" in columnas:
        valores["DETALLE_PROCESO"] = detalle
    if contenido_anterior is not None and "CONTENIDO_ANTERIOR" in columnas:
        valores["CONTENIDO_ANTERIOR"] = contenido_anterior
    actualizar_fila_sheet(ws, row_idx, columnas, valores)


# ── Agrupar filas por producto (encadenamiento) ──────────────────
def _agrupar_por_producto(filas):
    """
    Agrupa filas PENDIENTE por producto (Location+Supplier+Code+
    Service_Type), preservando el orden de primera aparición. Filas del
    mismo producto pero distinto Codigo_Nota quedan juntas para reutilizar
    la búsqueda del producto y la apertura de Product Notes — no hace
    falta repetirlas por cada nota del mismo producto. Mismo patrón que
    ya usan los scripts de Valorización (agrupar por servicio madre/
    componente antes de procesar).
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
    compartan el mismo producto, sin repetir buscar_producto() ni
    abrir_product_notes() por cada Codigo_Nota.
    """
    buscar_producto(driver, location, supplier, code, service_type=service_type)
    abrir_product_notes(driver)


def verificar_notas_guardadas(driver, ws, columnas, filas):
    """
    SEGUNDA PASADA, después de terminar TODO el guardado — SOLO
    diagnóstico, no cambia el ESTADO de ninguna fila. A pedido explícito de
    la usuaria, para poder validar a mano (comparando muchas filas del
    Excel contra Tourplan real) si UPDATED/UPDATED BY son un indicador
    confiable de que el guardado se aplicó de verdad.

    Versiones anteriores probadas (v1.17-v1.19, revertidas) decidían
    automáticamente ERROR vs EDITADA/INSERTADA comparando UPDATED/UPDATED
    BY, con resultados que no coincidían con lo que la usuaria veía en
    Tourplan. Se pausa esa decisión automática — acá solo se EXPORTA el
    dato al DETALLE_PROCESO.

    Un primer intento (v1.20) hacía esta relectura DENTRO de process_nota,
    apenas guardada cada nota — pero eso forzaba una recarga real de la
    página (buscar_producto()+abrir_product_notes()) POR CADA NOTA,
    rompiendo el encadenamiento por producto que reutiliza la búsqueda
    entre notas del mismo grupo (mismo problema que ya resolvía
    _agrupar_por_producto() para el guardado) y volviendo la corrida mucho
    más lenta con lotes grandes. Por eso esta validación se separó en una
    pasada aparte, ejecutada DESPUÉS de que MAIN termine de guardar todas
    las notas — reagrupando otra vez por producto (misma
    _agrupar_por_producto()) para encadenar la relectura igual que hace
    exportar_notas.py, en vez de recargar por cada nota individual.

    `filas`: lista de dicts de fila (mismos que ya tiene MAIN, con
    __row_idx__) que quedaron INSERTADA o EDITADA en la primera pasada.
    Si la nota es una EDIT, el dict de fila ya trae "_updated_antes"/
    "_updated_by_antes" (los guardó process_nota() antes de editar).
    """
    if not filas:
        return
    print(f"\n{'=' * 60}")
    print(f"🔎 Segunda pasada — validación de UPDATED ({len(filas)} nota(s), solo diagnóstico)")
    print(f"{'=' * 60}")

    grupos = _agrupar_por_producto(filas)
    for grupo in grupos:
        chequear_abort()
        ref          = grupo[0]
        location     = str(ref.get("Location")     or "").strip()
        supplier     = str(ref.get("Supplier")     or "").strip()
        service_type = str(ref.get("Service_Type") or "").strip()
        code         = str(ref.get("Code")         or "").strip()

        try:
            buscar_producto(driver, location, supplier, code, service_type=service_type)
            abrir_product_notes(driver)
            tabla = encontrar_tabla_notas(driver)
            if tabla is None:
                raise Exception("no se pudo releer la tabla de Product Notes")
            esperar_filas_estables(driver, tabla)
        except Exception as e:
            print(f"    ⚠ No se pudo recargar {location}/{supplier}/{code} para validar — {e}")
            for row in grupo:
                _agregar_detalle_validacion(ws, columnas, row["__row_idx__"],
                                             f"[VALIDACIÓN UPDATED] no se pudo recargar el producto — {e}")
            continue

        for row in grupo:
            chequear_abort()
            row_idx = row["__row_idx__"]
            codigo_nota = str(row.get("Codigo_Nota") or "").strip().upper()
            nota, _vistos = buscar_nota_por_codigo(driver, tabla, codigo_nota)
            if nota is None:
                detalle_validacion = f"[VALIDACIÓN UPDATED] la nota {codigo_nota} no aparece tras recargar"
            else:
                updated_antes = row.get("_updated_antes")
                if updated_antes is not None:
                    detalle_validacion = (
                        f"[VALIDACIÓN UPDATED] ANTES: UPDATED={updated_antes} "
                        f"({row.get('_updated_by_antes')}) — DESPUÉS (recarga real): "
                        f"UPDATED={nota['updated']} ({nota['updated_by']})")
                else:
                    detalle_validacion = (
                        f"[VALIDACIÓN UPDATED] tras recarga real: "
                        f"CREATED={nota['created']} ({nota['created_by']}) "
                        f"UPDATED={nota['updated']} ({nota['updated_by']})")
            print(f"  Fila {row_idx}: {detalle_validacion}")
            _agregar_detalle_validacion(ws, columnas, row_idx, detalle_validacion)


def _agregar_detalle_validacion(ws, columnas, row_idx, detalle_validacion):
    """Agrega `detalle_validacion` al DETALLE_PROCESO ya escrito en la
    primera pasada, en vez de pisarlo — lee el valor actual de la celda
    directo del Sheet antes de escribir la versión combinada."""
    if "DETALLE_PROCESO" not in columnas:
        return
    col_num = columnas.index("DETALLE_PROCESO") + 1
    actual = ws.cell(row_idx, col_num).value
    nuevo = f"{actual} | {detalle_validacion}" if actual else detalle_validacion
    update_row(ws, row_idx, columnas, detalle=nuevo)


# ── Procesamiento de una nota, sobre un producto ya abierto ──────
def process_nota(driver, row):
    """Procesa UN Codigo_Nota sobre un producto YA ABIERTO en Product
    Notes (ver _buscar_y_abrir_producto(), llamada por el caller una sola
    vez por grupo de filas del mismo producto — no repite la búsqueda del
    producto ni la apertura de Product Notes).

    Devuelve (estado, detalle, contenido_anterior).
    estado: "PENDIENTE" (dry-run, no se toca) | "INSERTADA" |
            "YA EXISTE - no modificada" | "EDITADA" | "ERROR: <detalle>"
    """
    location     = str(row.get("Location") or "").strip()
    supplier     = str(row.get("Supplier") or "").strip()
    service_type = str(row.get("Service_Type") or "").strip()
    code         = str(row.get("Code") or "").strip()
    codigo_nota  = str(row.get("Codigo_Nota") or "").strip().upper()
    texto        = str(row.get("Texto") or "")
    row_idx      = row["__row_idx__"]

    if not (location and supplier and service_type and code and codigo_nota and texto):
        return ("ERROR: faltan campos obligatorios "
                "(Location/Supplier/Service_Type/Code/Codigo_Nota/Texto)", None, None)

    # ── Validación de código de nota ANTES de tocar la UI ──────
    info_codigo = CODIGOS_NOTA_VALIDOS.get(codigo_nota)
    if info_codigo is None:
        return (f"ERROR: código de nota no soportado o no es tipo Product/Plain Text "
                f"('{codigo_nota}' no está en CODIGOS_NOTA_VALIDOS)", None, None)
    _descripcion, category_type, format_type = info_codigo
    if category_type != "Product" or format_type != "Plain Text":
        return (f"ERROR: código de nota no soportado o no es tipo Product/Plain Text "
                f"(category_type={category_type}, format_type={format_type})", None, None)

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

    # Se guarda el UPDATED/UPDATED BY de ANTES de tocar la nota directo en
    # el dict de la fila (mismo objeto que ya tiene MAIN, por referencia) —
    # así la segunda pasada de validación (verificar_notas_guardadas(), que
    # corre después de terminar TODO el guardado, para no romper el
    # encadenamiento por producto) puede armar el detalle ANTES/DESPUÉS sin
    # tener que volver a buscar el valor de antes de editar.
    if existente:
        row["_updated_antes"] = existente["updated"]
        row["_updated_by_antes"] = existente["updated_by"]

    if MODO == "lectura":
        if existente:
            detalle = (f"[LECTURA] Ya existe {codigo_nota} — CREATED={existente['created']} "
                       f"({existente['created_by']}) UPDATED={existente['updated']} "
                       f"({existente['updated_by']})")
        else:
            detalle = (f"[LECTURA] {codigo_nota} no existe — se insertaría "
                       f"(códigos CAT vistos: {vistos})")
        print(f"  {detalle}")
        # No se toca ESTADO en modo lectura — la fila sigue PENDIENTE para
        # cuando se corra en modo aplicar.
        return "PENDIENTE", detalle, None

    texto_html = texto_a_html_pre(texto)

    if not existente:
        try:
            click_insert(driver)
            seleccionar_codigo_en_dialog(driver, codigo_nota)
            esperar_editor(driver)
            activar_editor_click(driver)
            escribir_ckeditor(driver, texto_html)
            time.sleep(0.8 * VELOCIDAD)
            releido = leer_ckeditor(driver)
            if releido is None or texto_html.strip() not in (releido or "").strip():
                print("    ⚠ Contenido releído no calza exactamente con lo escrito — revisar manualmente")
            click_save(driver, con_polling=False)
            if not esperar_cierre_editor(driver):
                ss(driver, f"editor_no_cerro_f{row_idx}")
                return "ERROR: el editor no se cerró tras SAVE (posible error de validación)", None, None
            return "INSERTADA", f"Nota {codigo_nota} insertada", None
        except Exception as e:
            ss(driver, f"insert_error_f{row_idx}")
            return f"ERROR: falla insertando nota — {e}", None, None

    if EDICION == "NO":
        detalle = (f"CREATED={existente['created']} ({existente['created_by']}) "
                   f"UPDATED={existente['updated']} ({existente['updated_by']})")
        return "YA EXISTE - no modificada", detalle, None

    # EDICION == "SI" — sobreescribir, guardando antes el contenido anterior
    try:
        icono = existente["row_el"].find_element(By.CSS_SELECTOR, SEL["EDIT_ICON"])
        jc(driver, icono)
        esperar_editor(driver)
        activar_editor_click(driver)
        contenido_anterior = leer_ckeditor(driver)
        escribir_ckeditor(driver, texto_html)
        time.sleep(0.8 * VELOCIDAD)
        releido = leer_ckeditor(driver)
        if releido is None or texto_html.strip() not in (releido or "").strip():
            print("    ⚠ Contenido releído no calza exactamente con lo escrito — revisar manualmente")
        click_save(driver, con_polling=True)
        if not esperar_cierre_editor(driver):
            ss(driver, f"editor_no_cerro_edit_f{row_idx}")
            return ("ERROR: el editor no se cerró tras SAVE (posible error de validación)",
                     None, contenido_anterior)
        return "EDITADA", f"Nota {codigo_nota} sobreescrita", contenido_anterior
    except Exception as e:
        ss(driver, f"edit_error_f{row_idx}")
        return f"ERROR: falla editando nota — {e}", None, None


# ── MAIN ──────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print(f"  NOTAS SRV (Product Notes)  v{VERSION}  [{VERSION_FECHA}]  "
          f"MODO={MODO}  EDICION={EDICION}")
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
    filas_para_validar = []
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

                estado, detalle, contenido_anterior = "ERROR", "Error desconocido", None
                try:
                    estado, detalle, contenido_anterior = process_nota(driver, row)
                except Exception:
                    estado = "ERROR"
                    detalle = traceback.format_exc(limit=3)

                print(f"  Estado: {estado}" + (f" — {detalle}" if detalle else ""))
                update_row(ws, row_idx, columnas,
                           estado=estado, detalle=detalle, contenido_anterior=contenido_anterior)

                if estado in ("INSERTADA", "EDITADA"):
                    filas_para_validar.append(row)

                # Re-sincronización: una nota en ERROR puede dejar un
                # diálogo/editor a medio abrir en pantalla — antes, cada
                # fila hacía buscar_producto() desde cero, lo que de paso
                # "limpiaba" cualquier resto trabado; con el encadenamiento
                # ya no. Si quedan más notas de este mismo producto, volver
                # a buscar y abrir el producto para arrancar la próxima
                # nota desde un estado limpio, en vez de asumir que la
                # pantalla quedó bien tras el fallo.
                hay_mas_notas = i < len(grupo) - 1
                if estado.startswith("ERROR") and hay_mas_notas:
                    print("  ↩ Nota en ERROR — re-sincronizando (nueva búsqueda del "
                          "producto) antes de la próxima nota del mismo producto")
                    try:
                        _buscar_y_abrir_producto(driver, location, supplier, code, service_type)
                    except Exception as e:
                        # Si ni la re-sincronización funciona, no tiene sentido
                        # seguir intentando notas sobre este producto.
                        ss(driver, f"resync_error_grupo_{code[:12]}")
                        estado_resync = (f"ERROR: no se pudo re-sincronizar el producto "
                                         f"tras falla de nota — {e}")
                        for row_restante in grupo[i + 1:]:
                            print(f"  Estado: {estado_resync}")
                            update_row(ws, row_restante["__row_idx__"], columnas,
                                       estado=estado_resync, detalle=None)
                        break

        # Segunda pasada, DESPUÉS de terminar todo el guardado — ver
        # verificar_notas_guardadas(). Separada a propósito del loop de
        # arriba para no romper su encadenamiento por producto con una
        # recarga por cada nota individual.
        verificar_notas_guardadas(driver, ws, columnas, filas_para_validar)

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
