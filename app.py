"""Tourplan NX - Herramientas: app local (Streamlit) que unifica los scripts
de automatizacion de Tourplan NX en una sola herramienta con UI.

Scripts activos, agrupados por categoría:
- Productos: Copy Products, Flag as Deleted, Modificar Description y Comment.
- PCM: Copy y Linkeo de PCM.
- Valorización: Desde Componentes, Desde Servicio Madre, Desde Excel, Tarifario estático.
- Notas: Notas SRV (insertar/editar Product Notes), Exportar Notas (solo lectura).
- Relevamiento: Vigencias (lista de períodos de RATES sin costos, solo lectura).

Cada script corre como subproceso propio (ver README.md - "Por que
subprocess"), parametrizado por variables de entorno. El usuario/password
de Tourplan se guardan en ~/.tourplan-nx-app/config.json (ver
common/user_config.py) para no tener que tipearlos en cada corrida.

Layout (rediseño UI, rama rediseno-ui-streamlit): header fijo arriba +
sidebar de navegación + centro (parámetros/consola del script elegido) +
panel derecho con pestañas Cola/Config. Ver _inyectar_css() para todo el
CSS custom, concentrado en un solo lugar porque los selectores de
Streamlit (data-testid, etc.) son frágiles entre versiones.
"""
import html
import os
import subprocess
import sys
import tempfile
import threading
import time
from pathlib import Path

import streamlit as st

from common.abort import ABORT_EXIT_CODE
from common import user_config
from common.sheets_client import conectar_sheets, cargar_sheet

REPO_ROOT = Path(__file__).resolve().parent

# URL default de producción para todos los scripts (editable en la UI). Las
# pruebas contra el ambiente de Test se hacen normalmente en Colab, antes de
# que un script se sume a esta app — el campo viene precargado con
# producción para que correr contra Test sea una decisión deliberada
# (cambiar el campo a mano), no un olvido.
PRODUCCION_URL = "https://tourplannx.eurotur.com.ar/tourplannx"

# modo_options: lista de (etiqueta visible, valor real que se manda como
# TOURPLAN_MODO). La primera opción es siempre la más segura (default del
# radio). Si un script no tiene modo de solo lectura, se omite esta clave.
#
# MODO_LECTURA_ESCRITURA es genérico a propósito — lo comparten varios
# scripts (Copy Products, Modificar Description y Comment, Valorización
# desde Excel, Notas SRV) que no tienen nada que ver entre sí más que el
# patrón lectura/aplicar, así que el texto no puede nombrar una acción
# puntual (antes decía "copia los productos", que solo es cierto para
# Copy Products y aparecía igual en los otros 3). Copy Products usa su
# propia variante de abajo con el detalle específico.
MODO_LECTURA_ESCRITURA = [
    ("Solo revisar (lectura, no modifica nada)", "lectura"),
    ("Aplicar cambios", "completo"),
]

MODO_COPY_PRODUCTS = [
    ("Solo revisar (lectura, no modifica nada)", "lectura"),
    ("Aplicar cambios (copia los productos)", "completo"),
]

MODO_FASES = [
    ("Solo leer (Fase 1: buscar y leer costos, no escribe nada)", "LEER"),
    ("Solo aplicar (Fase 2: escribir las tarifas ya leídas)", "APLICAR"),
    ("Completo (Fase 1 + Fase 2 en la misma corrida)", "COMPLETO"),
]

# Parámetro global extra de Notas SRV — además de MODO (lectura/aplicar),
# controla qué hacer si el código de nota YA existe en el producto. No es
# por fila: aplica a toda la corrida (ver TOURPLAN_EDICION más abajo).
EDICION_NO_SI = [
    ("No tocar notas existentes (reporta para revisión manual)", "NO"),
    ("Sobreescribir notas existentes (guarda antes el contenido anterior)", "SI"),
]

# Cada entrada tiene una "category" (agrupa la navegación del sidebar) y un
# "script_path" — render_script_tab avisa en la UI en vez de romper si
# algún script_path todavía no existiera en el repo.
# Flag as Deleted: qué hacer con cada producto (TOURPLAN_ELIMINAR). Reemplaza al
# selector de Modo para ese script (no tiene modo lectura en la app).
ELIMINAR_NO_SI = [
    ("Flag as deleted", "NO"),
    ("Delete (si falla, flag as deleted)", "SI"),
]

# Catálogo de notas para Exportar Notas (TOURPLAN_CODIGOS_NOTA) — mismo
# catálogo que FAMILIAS_NOTA/CODIGOS_SUELTOS en exportar_notas.py (no hay
# import cruzado entre la app y los scripts, cada uno corre como
# subproceso propio — ver duplicado ahí con el mismo detalle de por qué
# Luggage Waiver usa sufijos en inglés). La UI arma un desplegable de
# familias + código suelto (Paso 1) y, para las que tienen variante de
# idioma, checkboxes de idioma (Paso 2) — ver render_script_tab().
FAMILIAS_NOTA = {
    "Nota SRV":    {"Aleman": "NAL", "Espanol": "NES", "Frances": "NFR", "Ingles": "NIN", "Italiano": "NIT"},
    "Descriptivo": {"Aleman": "UAL", "Espanol": "UES", "Frances": "UFR", "Ingles": "UIN", "Italiano": "UIT"},
    "Titulo":      {"Aleman": "TAL", "Espanol": "TES", "Frances": "TFR", "Ingles": "TIN", "Italiano": "TIT"},
    "Luggage Waiver": {"Aleman": "LWG", "Espanol": "LWS", "Frances": "LWF", "Ingles": "LWE", "Italiano": "LWI"},
}

CODIGOS_SUELTOS = {
    "Direccion Rent a Car":               "DRT",
    "Producto Coordinates":               "PCR",
    "Remodelacion Hotel":                 "REM",
    "Nota Cliente Solo Voucher":          "REO",
    "External Option - Mapeo Especifico": "SC2",
}

SCRIPTS = {
    "copy_products": {
        "label": "Copy Products (01)",
        "category": "Productos",
        # Reemplaza a los copy_products/rename_products viejos (eliminados).
        # Fuente: https://github.com/daianalezcano-beep/Copy-products/tree/claude/new-option-copy-script-25yt7w
        "help": "Copy Product con la posibilidad de modificar todos los campos",
        "script_path": REPO_ROOT / "scripts" / "copy_products" / "copy_products.py",
        "base_url": PRODUCCION_URL,
        "sheet": "PRODUCTOS",
        "modo_options": MODO_COPY_PRODUCTS,
    },
    "flag_as_deleted": {
        "label": "Flag as Deleted (02)",
        "category": "Productos",
        "help": "Marca productos como \"Flag Product as Deleted\" en Tourplan.",
        "script_path": REPO_ROOT / "scripts" / "flag_as_deleted" / "flag_products_as_deleted.py",
        "base_url": PRODUCCION_URL,
        "sheet": "PRODUCTOS",
        "no_modo_info": "ℹ️ Este script aplica los cambios directo en Tourplan (sin modo de solo lectura).",
        "eliminar_options": ELIMINAR_NO_SI,
    },
    "modificar_description_comment": {
        "label": "Modificar Description y Comment (03)",
        "category": "Productos",
        # Fuente: https://github.com/daianalezcano-beep/Copy-products/tree/claude/nuevo-branch-repositorio-p32d22
        "help": "Modifica los campos description y/o comment de un producto en Tourplan.",
        "script_path": REPO_ROOT / "scripts" / "modificar_description_comment" / "modificar_description_comment.py",
        "base_url": PRODUCCION_URL,
        "sheet": "PRODUCTOS",
        "modo_options": MODO_LECTURA_ESCRITURA,
    },
    "copy_pcm_linkeo": {
        "label": "Copy y Linkeo PCM (10)",
        "category": "PCM",
        "help": "Copia un PCM (Package Header) desde un servicio madre y lo linkea a un product code nuevo.",
        "script_path": REPO_ROOT / "scripts" / "copy_pcm_linkeo" / "tourplan_copiar_pcm.py",
        "base_url": PRODUCCION_URL,
        "sheet": "PRODUCTOS",
    },
    "valorizacion_pkg": {
        "label": "Valorización desde Componente (20)",
        "category": "Valorización",
        "help": "Valoriza servicios madre PKG a partir de un componente (COD ORIGEN): lee costos de los PCM hijos y escribe las tarifas en el servicio madre.",
        "script_path": REPO_ROOT / "scripts" / "valorizacion_pkg" / "tourplan_valorizacion_pkg_v3.py",
        "base_url": PRODUCCION_URL,
        "sheet": "PRODUCTOS",
        "modo_options": MODO_FASES,
    },
    "valorizacion_madre": {
        "label": "Valorización desde Servicio Madre (21)",
        "category": "Valorización",
        "help": "Valoriza servicios madre PKG partiendo directo del servicio madre (sin componente): busca sus PCM tipo \"Package Header\" y aplica las tarifas.",
        "script_path": REPO_ROOT / "scripts" / "valorizacion_madre" / "valorizacion_desde_madre.py",
        "base_url": PRODUCCION_URL,
        "sheet": "PRODUCTOS",
        "modo_options": MODO_FASES,
    },
    "valorizacion_excel": {
        "label": "Valorización desde Excel (22)",
        "category": "Valorización",
        # Fuente: https://github.com/daianalezcano-beep/Valorizacion-EX-TF-dsd-servicio-madre/tree/claude/nuevo-script-471bvm
        "help": "Valoriza servicios con un solo pax break, desde un valor cargado en el excel.",
        "script_path": REPO_ROOT / "scripts" / "valorizacion_excel" / "valorizacion_desde_excel.py",
        "base_url": PRODUCCION_URL,
        "sheet": "TARIFAS",
        "modo_options": MODO_LECTURA_ESCRITURA,
    },
    "valorizacion_madre_numericos": {
        "label": "Valorización Tarifario estático (23)",
        "category": "Valorización",
        "help": "Variante de \"Valorización desde Servicio Madre\" con el filtro de códigos invertido: procesa únicamente los servicios madre PKG cuyo código empieza con un número (los que ese script descarta). Busca sus PCM tipo \"Package Header\" y aplica las tarifas.",
        "script_path": REPO_ROOT / "scripts" / "valorizacion_madre_numericos" / "valorizacion_desde_madre_numericos.py",
        "base_url": PRODUCCION_URL,
        "sheet": "PRODUCTOS",
        "modo_options": MODO_FASES,
    },
    "modificar_rate_name_text": {
        "label": "Modificar Rate Text / Rate Name (24)",
        "category": "Valorización",
        # Fuente: https://github.com/daianalezcano-beep/Tourplan-Valorizacion-EX-TF
        "help": "Modifica el Rate Name y/o el Rate Text de tarifas existentes, identificando el período (vigencia) y el price code dentro de la grilla de RATES del producto.",
        "script_path": REPO_ROOT / "scripts" / "modificar_rate_name_text" / "modificar_rate_name_text.py",
        "base_url": PRODUCCION_URL,
        "sheet": "Datos",
        "modo_options": MODO_LECTURA_ESCRITURA,
    },
    "notas_srv": {
        "label": "Crear/Modificar (30)",
        "category": "Notas",
        # Fuente: https://github.com/daianalezcano-beep/copy-products/tree/notas-SRV
        "help": "Inserta/edita notas Plain Text (Product Notes) en productos de Tourplan NX: Nota SRV, Descriptivo, Luggage Waiver, Título y otros 25 códigos confirmados. No pisa notas existentes salvo que se pida explícitamente.",
        "script_path": REPO_ROOT / "scripts" / "notas_srv" / "notas_srv.py",
        "base_url": PRODUCCION_URL,
        "sheet": "NOTAS_SRV",
        "modo_options": MODO_LECTURA_ESCRITURA,
        "edicion_options": EDICION_NO_SI,
    },
    "exportar_notas": {
        "label": "Exportar (31)",
        "category": "Notas",
        # Fuente: https://github.com/daianalezcano-beep/copy-products/tree/notas-SRV
        "help": "Exporta notas (Product Notes) de una lista de product codes. Elegí abajo qué notas exportar — los resultados se agregan a la pestaña EXPORTAR_NOTAS_RESULTADOS del mismo Sheet, una fila por producto × nota.",
        "script_path": REPO_ROOT / "scripts" / "exportar_notas" / "exportar_notas.py",
        "base_url": PRODUCCION_URL,
        "sheet": "EXPORTAR_NOTAS",
        "sheet_resultados": "EXPORTAR_NOTAS_RESULTADOS",
        "no_modo_info": "ℹ️ Este script es de solo lectura: nunca inserta, edita ni guarda nada en Tourplan.",
        "notas_a_exportar": True,
    },
    "extraccion_vigencias": {
        "label": "Vigencias (40)",
        "category": "Relevamiento",
        # Fuente: https://github.com/daianalezcano-beep/generico-vs-especificos/tree/claude/rates-extraction-script-lm4tyn
        "help": "Relevamiento de vigencias: dado un supplier/location/service type (o código puntual), exporta el estado de los períodos de RATES de cada producto (fechas, Price Code, moneda, estado) sin leer costos. Sirve para detectar tarifas vencidas o sin cargar.",
        "script_path": REPO_ROOT / "scripts" / "extraccion_vigencias" / "extraccion_vigencias.py",
        "base_url": PRODUCCION_URL,
        "sheet": "PRODUCTOS",
        "no_modo_info": "ℹ️ Este script es de solo lectura: exporta la vigencia de cada período de RATES a la hoja VIGENCIAS, sin insertar, editar ni guardar nada en Tourplan.",
    },
}

# Colores de las 4 tarjetas de conteo del panel derecho (bg, texto). "Otro"
# no tiene tarjeta — se muestra como texto chico al lado de la barra de
# progreso, igual que antes del rediseño.
_COLORES_ESTADO = {
    "Pendiente":  ("#FFF4D6", "#7A5200"),
    "En proceso": ("#DCEBFF", "#0B4A9E"),
    "OK":         ("#DDF5E6", "#11643A"),
    "Error":      ("#FFE0E0", "#A21B1B"),
}


def _clasificar_estado(valor):
    """Agrupa el valor crudo de la columna ESTADO (que en varios scripts
    incluye el detalle del error, ej. "ERROR: Faltan campos...") en un
    puñado de categorías fijas para poder contarlas.

    Una celda VACÍA no cuenta como "Pendiente": ningún script la toma
    como tal (su propio leer_pendientes()/filtro exige el valor literal
    PENDIENTE/PENDING/PEND, nunca una celda en blanco) — contarla acá
    como pendiente inflaba el contador con filas que el script de verdad
    nunca iba a procesar. Cae en "Otro".

    EXPORTADO/PARCIAL son el vocabulario propio de Exportar Notas (ESTADO
    agregado por producto, no por fila de Tourplan como en los demás
    scripts): EXPORTADO se cuenta como OK (se exportaron bien todas las
    notas configuradas) y PARCIAL como Error (al menos una nota falló)."""
    v = (valor or "").strip().upper()
    if v in ("PENDIENTE", "PENDING", "PEND"):
        return "Pendiente"
    if v == "PROCESANDO":
        return "En proceso"
    if v.startswith("ERROR"):
        return "Error"
    if v == "OK" or v.startswith("OK "):
        return "OK"
    if v == "EXPORTADO":
        return "OK"
    if v == "PARCIAL":
        return "Error"
    return "Otro"


def _contar_estados(filas, columna="ESTADO"):
    conteo = {"Pendiente": 0, "En proceso": 0, "OK": 0, "Error": 0, "Otro": 0}
    for fila in filas:
        conteo[_clasificar_estado(fila.get(columna))] += 1
    return conteo


def _refrescar_conteo(state, sheet_url, hoja):
    """Lee el Sheet ahora mismo y actualiza el conteo guardado en el
    estado del script. No propaga la excepción — un fallo de refresco no
    tiene que interrumpir una corrida en curso."""
    try:
        ws = conectar_sheets(
            sheet_url, hoja, user_config.CREDENTIALS_PATH, user_config.TOKEN_PATH)
        filas, _ = cargar_sheet(ws)
        state["conteo"] = _contar_estados(filas)
        state["conteo_ts"] = time.time()
        return True
    except Exception as e:
        state["conteo_error"] = str(e)
        return False


def _clasificar_estado_nota(valor):
    """Clasifica el ESTADO de una fila de EXPORTAR_NOTAS_RESULTADOS (una
    fila por producto × nota exportada) — vocabulario propio, distinto al
    de _clasificar_estado: EXPORTADA/NO EXISTE/ERROR: <detalle>."""
    v = (valor or "").strip().upper()
    if v == "EXPORTADA":
        return "OK"
    if v == "NO EXISTE" or v.startswith("ERROR"):
        return "Error"
    return "Otro"


def _contar_estados_notas(filas):
    conteo = {"OK": 0, "Error": 0, "Otro": 0}
    for fila in filas:
        conteo[_clasificar_estado_nota(fila.get("ESTADO"))] += 1
    return conteo


def _refrescar_conteo_notas(state, sheet_url, hoja):
    """Igual que _refrescar_conteo pero para la hoja de resultados por
    nota (EXPORTAR_NOTAS_RESULTADOS). Guarda las filas crudas (no un
    conteo ya agregado): el conteo que de verdad importa es el de la
    selección de notas vigente en la UI (ver _render_centro), que acá
    todavía no se conoce — eso filtra por Codigo_Nota al mostrarlo."""
    try:
        ws = conectar_sheets(
            sheet_url, hoja, user_config.CREDENTIALS_PATH, user_config.TOKEN_PATH)
        filas, _ = cargar_sheet(ws)
        state["notas_filas"] = filas
        return True
    except Exception as e:
        state["conteo_notas_error"] = str(e)
        return False


def _chequear_pendientes_globales():
    """Chequeo único al abrir la app (no se repite en cada rerun de
    Streamlit — ver _asegurar_pendientes_globales): para cada script con
    una URL de Sheet guardada en Configuración, cuenta filas
    Pendiente/En proceso. Un error puntual leyendo un Sheet no bloquea el
    chequeo de los demás."""
    pendientes = []
    for key, cfg in SCRIPTS.items():
        if not cfg["script_path"].exists():
            continue
        sheet_url = user_config.sheet_url_default(key)
        if not sheet_url:
            continue
        try:
            ws = conectar_sheets(
                sheet_url, cfg["sheet"], user_config.CREDENTIALS_PATH, user_config.TOKEN_PATH)
            filas, _ = cargar_sheet(ws)
        except Exception:
            continue
        conteo = _contar_estados(filas)
        if conteo["Pendiente"] or conteo["En proceso"]:
            pendientes.append({
                "key": key,
                "label": cfg["label"],
                "pendiente": conteo["Pendiente"],
                "procesando": conteo["En proceso"],
            })
    return pendientes


def _asegurar_pendientes_globales():
    """Calcula los pendientes una sola vez por sesión de navegador (igual
    que antes del rediseño) y los deja en session_state. Ya no se muestra
    como aviso arriba de toda la página: alimenta el chip ámbar del
    header y la caja "Corridas sin terminar" del panel derecho."""
    if "_pendientes_globales" not in st.session_state:
        st.session_state["_pendientes_globales"] = _chequear_pendientes_globales()


def _hay_algo_corriendo():
    """True si hay al menos un script con una corrida en curso en esta
    sesión — se deriva solo leyendo el estado de los subprocesos
    (state["running"] de cada _state_{key}), la app no mantiene una
    sesión propia contra Tourplan. Se usa para el punto del header y
    para decidir el run_every de los fragments."""
    for k, v in st.session_state.items():
        if k.startswith("_state_") and isinstance(v, dict) and v.get("running"):
            return True
    return False


def _scripts_en_curso():
    """Lista (key, label, started_at) de los scripts corriendo ahora
    mismo, en cualquier pestaña — no solo el script visible."""
    en_curso = []
    for state_key, s in st.session_state.items():
        if not state_key.startswith("_state_") or not isinstance(s, dict):
            continue
        if s.get("running"):
            k = state_key[len("_state_"):]
            cfg_k = SCRIPTS.get(k)
            if cfg_k is None:
                continue
            en_curso.append((k, cfg_k["label"], s.get("started_at")))
    return en_curso


def _registrar_corridas_terminadas():
    """Vigilante de solo lectura (regla de rediseño): recorre todos los
    _state_{key} y, para cada corrida que ya terminó (finished=True) y
    todavía no está anotada en el historial de la sesión, la agrega una
    sola vez a st.session_state["_corridas_terminadas"]. No consulta
    Sheets, no modifica running/finished/returncode ni ningún otro campo
    de state (aparte de la propia bandera de "ya anotada"), y no muestra
    nada — solo lee y deja un registro para que el panel derecho lo
    pinte."""
    historial = st.session_state.setdefault("_corridas_terminadas", [])
    for state_key, state in st.session_state.items():
        if not state_key.startswith("_state_") or not isinstance(state, dict):
            continue
        if not state.get("finished") or state.get("_anotada_en_historial"):
            continue
        key = state_key[len("_state_"):]
        cfg = SCRIPTS.get(key)
        if cfg is None:
            continue
        historial.append({
            "key": key,
            "label": cfg["label"],
            "returncode": state.get("returncode"),
            "hora": time.strftime("%H:%M"),
        })
        state["_anotada_en_historial"] = True


def _formatear_hace(segundos):
    minutos = max(0, int(segundos // 60))
    if minutos < 1:
        return "hace instantes"
    return f"hace {minutos} min"


def _leer_proceso(proc, state):
    """Corre en un hilo aparte: lee el stdout del subproceso sin bloquear el
    script de Streamlit, para que el botón Abortar pueda reaccionar mientras
    el script sigue corriendo."""
    for line in proc.stdout:
        state["log_lines"].append(line)
    proc.wait()
    state["returncode"] = proc.returncode
    state["finished"] = True


def _inyectar_css():
    """Único lugar con CSS custom de toda la app (pedido explícito del
    rediseño: los selectores de Streamlit — data-testid, atributos
    "kind", etc. — son frágiles y cambian entre versiones, mejor tenerlos
    todos juntos y documentados acá que desperdigados).

    A propósito NO hay CSS que dependa de la POSICIÓN de una columna (ej.
    "la segunda columna de tal fila"): Streamlit no da un selector
    estable para eso, y un cambio así podría romper layouts sin relación
    (como las columnas de idiomas de Exportar Notas, que también son
    st.columns). Por eso el ancho del panel derecho se logra con la
    proporción de st.columns (ver _render_script_tab), no con CSS."""
    st.markdown(
        """
        <style>
        /* Streamlit reserva arriba de todo su propia barra nativa
           (Running/Deploy/hamburguesa) más un padding-top grande en el
           contenedor principal para no taparla — con nuestro header
           propio (más abajo, position: sticky) eso deja un hueco vacío
           antes de llegar a él. Se esconde esa barra nativa (no hace
           falta en esta app) y se pone el padding-top en 0. A
           propósito NO se toca stSidebarCollapsedControl (el botón
           para contraer/expandir el sidebar): ese sigue visible, a
           diferencia de apps que ocultan el sidebar entero.
           Selectores verificados contra Streamlit 1.37/1.65 — si una
           versión futura cambia estos data-testid, lo único que pasa
           es que vuelve el hueco de arriba, no se rompe nada más. */
        header[data-testid="stHeader"],
        [data-testid="stToolbar"],
        [data-testid="stDecoration"] {
            display: none !important;
        }
        .block-container,
        [data-testid="stMainBlockContainer"],
        [data-testid="stMain"],
        [data-testid="stAppViewContainer"] > section {
            padding-top: 0 !important;
        }
        /* Los bloques <style> que inyectamos con st.markdown (este
           mismo, y los de las tarjetas/consola/cajas) no deberían dejar
           un renglón vacío en el flujo normal de la página. */
        [data-testid="stElementContainer"]:has(style) {
            display: none !important;
        }

        /* Sidebar: ancho aproximado ~236px (selector estable y de uso
           común para esto; no afecta nada fuera del propio sidebar). */
        [data-testid="stSidebar"] {
            width: 236px !important;
            min-width: 236px !important;
        }

        /* El sidebar reserva arriba una franja fija de 60px
           (stSidebarHeader) solo para el botón de contraer, que mide
           28px — eso es lo que empuja "Scripts" hacia abajo. Se achica
           la franja sin esconder el botón (sigue andando igual para
           contraer/expandir). */
        [data-testid="stSidebarHeader"] {
            height: 40px !important;
        }

        /* Botones primarios: Ejecutar y el script seleccionado del
           sidebar. Streamlit marca el <button> con kind="primary". */
        button[kind="primary"] {
            background-color: #D63030 !important;
            border-color: #D63030 !important;
        }
        button[kind="primary"]:hover {
            background-color: #B82828 !important;
            border-color: #B82828 !important;
        }

        /* Header fijo */
        .tp-header {
            position: sticky;
            top: 0;
            z-index: 999;
            display: flex;
            align-items: center;
            justify-content: space-between;
            height: 48px;
            padding: 0 6px;
            background: #FFFFFF;
            border-bottom: 1px solid #D5D9E2;
            margin-bottom: 0.6rem;
        }
        .tp-header-izq {
            font-size: 1.15rem;
            font-weight: 700;
        }
        .tp-header-der {
            display: flex;
            align-items: center;
            gap: 16px;
        }
        .tp-chip {
            padding: 4px 12px;
            border-radius: 999px;
            font-size: 0.8rem;
            font-weight: 600;
            white-space: nowrap;
        }
        .tp-chip-ambar {
            background: #FFF4D6;
            color: #7A5200;
        }
        .tp-sesion {
            font-size: 0.85rem;
            color: #31333F;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            white-space: nowrap;
        }
        .tp-punto {
            display: inline-block;
            width: 8px;
            height: 8px;
            border-radius: 50%;
            flex-shrink: 0;
        }
        .tp-punto-verde { background: #1DB954; }
        .tp-punto-gris  { background: #9AA0AC; }
        .tp-punto-azul  { background: #0B4A9E; }
        .tp-badge {
            padding: 4px 12px;
            border-radius: 4px;
            font-size: 0.75rem;
            font-weight: 700;
            text-transform: uppercase;
            white-space: nowrap;
        }
        .tp-badge-prod { background: #FFE0E0; color: #A21B1B; }
        .tp-badge-otro { background: #E5E7EB; color: #374151; }

        /* Tarjetas de conteo (panel derecho, pestaña Cola) */
        .tp-tarjetas {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            margin-bottom: 0.4rem;
        }
        .tp-tarjeta {
            flex: 1 1 calc(50% - 8px);
            min-width: 100px;
            border-radius: 8px;
            padding: 8px 10px;
        }
        .tp-tarjeta-etiqueta {
            font-size: 0.72rem;
            font-weight: 600;
        }
        .tp-tarjeta-valor {
            font-size: 1.35rem;
            font-weight: 700;
            line-height: 1.5rem;
        }

        /* Consola */
        .tp-consola {
            display: flex;
            flex-direction: column-reverse;
            overflow-y: auto;
            height: 360px;
            background: #F0F2F6;
            color: #31333F;
            font-family: "Source Code Pro", "SFMono-Regular", Consolas, "Courier New", monospace;
            font-size: 12.5px;
            line-height: 21px;
            padding: 8px 12px;
            border-radius: 6px;
            border: 1px solid #D5D9E2;
        }
        .tp-consola-linea {
            white-space: pre-wrap;
            word-break: break-word;
        }
        .tp-consola-vacia {
            color: #6b7280;
        }

        /* Cajas del panel derecho ("Ejecutando ahora") */
        .tp-caja {
            border-radius: 8px;
            padding: 10px 12px;
            margin-bottom: 0.6rem;
        }
        .tp-caja-ejecutando {
            background: #EEF5FF;
            border: 1px solid #B9D4F7;
        }
        .tp-caja-titulo {
            font-weight: 700;
            margin-bottom: 4px;
        }
        .tp-caja-subtitulo {
            font-weight: 600;
            margin-top: 8px;
            margin-bottom: 4px;
        }
        .tp-caja-fila {
            font-size: 0.85rem;
            padding: 2px 0;
        }
        .tp-caja-vacia {
            color: #6b7280;
        }
        .tp-caja-hace {
            color: #6b7280;
            font-size: 0.8rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _render_tarjetas_estado(conteo):
    partes = []
    for etiqueta in ("Pendiente", "En proceso", "OK", "Error"):
        bg, fg = _COLORES_ESTADO[etiqueta]
        partes.append(
            f'<div class="tp-tarjeta" style="background:{bg};color:{fg};">'
            f'<div class="tp-tarjeta-etiqueta">{html.escape(etiqueta)}</div>'
            f'<div class="tp-tarjeta-valor">{conteo[etiqueta]}</div>'
            f'</div>'
        )
    st.markdown(f'<div class="tp-tarjetas">{"".join(partes)}</div>', unsafe_allow_html=True)


def _render_consola_html(log_lines):
    lineas = log_lines[-500:]
    if lineas:
        contenido = "".join(
            f'<div class="tp-consola-linea">{html.escape(linea.rstrip(chr(10)))}</div>'
            for linea in reversed(lineas)
        )
    else:
        contenido = (
            '<div class="tp-consola-linea tp-consola-vacia">'
            'Sin ejecución todavía. El log aparece acá y sigue la última línea.'
            '</div>'
        )
    st.markdown(f'<div class="tp-consola">{contenido}</div>', unsafe_allow_html=True)


def render_header(selected_key):
    """Header fino y fijo arriba de todo — fragment con run_every
    dinámico (1 si hay algo corriendo en cualquier script, None si no),
    recalculado en cada rerun completo."""
    run_every = 1 if _hay_algo_corriendo() else None

    @st.fragment(run_every=run_every)
    def _fragment():
        pendientes = st.session_state.get("_pendientes_globales") or []
        chip_pendientes = (
            '<span class="tp-chip tp-chip-ambar">Pendientes de ejecutar</span>'
            if pendientes else ""
        )

        if _hay_algo_corriendo():
            sesion_html = (
                '<span class="tp-sesion"><span class="tp-punto tp-punto-verde"></span>'
                'Sesión Tourplan activa</span>'
            )
        else:
            sesion_html = (
                '<span class="tp-sesion"><span class="tp-punto tp-punto-gris"></span>'
                'Sin sesión activa</span>'
            )

        cfg = SCRIPTS[selected_key]
        base_url_actual = st.session_state.get(f"url_{selected_key}", cfg["base_url"])
        if base_url_actual == PRODUCCION_URL:
            badge_entorno = '<span class="tp-badge tp-badge-prod">PRODUCCIÓN</span>'
        else:
            badge_entorno = '<span class="tp-badge tp-badge-otro">URL DISTINTA</span>'

        st.markdown(
            '<div class="tp-header">'
            '<div class="tp-header-izq">Tourplan NX · Herramientas</div>'
            f'<div class="tp-header-der">{chip_pendientes}{sesion_html}{badge_entorno}</div>'
            '</div>',
            unsafe_allow_html=True,
        )

    _fragment()


def _render_caja_ejecutando(en_curso):
    """Caja "Ejecutando ahora" + "Terminaron en esta sesión" — ambas son
    puro texto (sin botones), así que se arman como un solo bloque HTML
    con color propio, sin depender de ningún selector de Streamlit."""
    filas = []
    if en_curso:
        ahora = time.time()
        for _k, label, started_at in en_curso:
            hace = _formatear_hace(ahora - started_at) if started_at else "hace instantes"
            filas.append(
                '<div class="tp-caja-fila"><span class="tp-punto tp-punto-azul"></span> '
                f'{html.escape(label)} <span class="tp-caja-hace">{html.escape(hace)}</span></div>'
            )
    else:
        filas.append('<div class="tp-caja-fila tp-caja-vacia">Nada en ejecución.</div>')

    _registrar_corridas_terminadas()
    terminadas = st.session_state.get("_corridas_terminadas", [])
    filas_terminadas = []
    for item in reversed(terminadas[-10:]):
        rc = item["returncode"]
        if rc == 0:
            texto, color = "Terminó", "#11643A"
        elif rc == ABORT_EXIT_CODE:
            texto, color = "Abortado", "#7A5200"
        else:
            texto, color = "Con error", "#A21B1B"
        filas_terminadas.append(
            f'<div class="tp-caja-fila"><span class="tp-punto" style="background:{color};"></span> '
            f'{html.escape(item["label"])} — '
            f'<span style="color:{color};font-weight:600;">{texto}</span> '
            f'<span class="tp-caja-hace">{html.escape(item["hora"])}</span></div>'
        )

    bloque_terminadas = ""
    if filas_terminadas:
        bloque_terminadas = (
            '<div class="tp-caja-subtitulo">Terminaron en esta sesión</div>'
            + "".join(filas_terminadas)
        )

    st.markdown(
        '<div class="tp-caja tp-caja-ejecutando">'
        '<div class="tp-caja-titulo">Ejecutando ahora</div>'
        + "".join(filas)
        + bloque_terminadas
        + '</div>',
        unsafe_allow_html=True,
    )


def _render_caja_corridas_sin_terminar(en_curso):
    """Contenido que antes era render_aviso_pendientes() (el banner
    arriba de toda la página) — mismo cálculo (_pendientes_globales, una
    sola vez por sesión), mismo botón "Ir", ahora en un st.container con
    borde dentro del panel derecho. Oculta los scripts que están
    corriendo en este momento (esos ya se ven en "Ejecutando ahora")."""
    st.markdown("**Corridas sin terminar**")
    pendientes = st.session_state.get("_pendientes_globales") or []
    keys_en_curso = {k for k, _label, _started in en_curso}
    pendientes_visibles = [p for p in pendientes if p["key"] not in keys_en_curso]

    with st.container(border=True):
        if not pendientes_visibles:
            st.caption("Nada pendiente de corridas anteriores.")
            return
        for item in pendientes_visibles:
            col_txt, col_btn = st.columns([4, 1])
            with col_txt:
                partes = []
                if item["pendiente"]:
                    partes.append(f"{item['pendiente']} PENDIENTE")
                if item["procesando"]:
                    partes.append(f"{item['procesando']} en PROCESANDO (quedó a medias)")
                st.markdown(f"**{item['label']}** — {', '.join(partes)}")
            with col_btn:
                if st.button("Ir", key=f"ir_a_{item['key']}", use_container_width=True):
                    st.session_state["selected_script"] = item["key"]
                    st.rerun()


def _render_tab_cola(key, cfg, state, sheet_url):
    # Refresco periódico mientras corre (mismo intervalo de siempre — no
    # se toca, ver _refrescar_conteo/_refrescar_conteo_notas).
    if state["running"] and time.time() - state.get("conteo_ts", 0) > 5:
        _refrescar_conteo(state, sheet_url, cfg["sheet"])
        if cfg.get("sheet_resultados"):
            _refrescar_conteo_notas(state, sheet_url, cfg["sheet_resultados"])

    if state.get("conteo"):
        conteo = state["conteo"]
        _render_tarjetas_estado(conteo)
        total = sum(conteo.values())
        if total > 0:
            # OK + Error + En proceso, no "total - Pendiente": con celdas
            # ESTADO vacías contando ahora como "Otro" (ver
            # _clasificar_estado), "total - Pendiente" las contaría como
            # procesadas sin haberlo sido.
            avanzadas = conteo["OK"] + conteo["Error"] + conteo["En proceso"]
            st.progress(avanzadas / total, text=f"{avanzadas}/{total} procesadas")
            st.caption(f"Otro: {conteo['Otro']}")
    elif state.get("conteo_error"):
        st.error(f"No pude leer el Sheet: {state['conteo_error']}")

    if cfg.get("sheet_resultados"):
        # Progreso de notas para la selección vigente (ver _render_centro):
        # cuántas de ESOS códigos ya están exportadas en
        # EXPORTAR_NOTAS_RESULTADOS, sobre el total esperado (productos
        # del Sheet × notas elegidas).
        notas_filas = state.get("notas_filas")
        codigos_a_exportar = state.get("codigos_a_exportar") or []
        if notas_filas is not None and codigos_a_exportar:
            codigos_set = set(codigos_a_exportar)
            filas_sel = [
                f for f in notas_filas
                if (f.get("Codigo_Nota") or "").strip().upper() in codigos_set
            ]
            conteo_notas = _contar_estados_notas(filas_sel)
            total_productos = sum(state["conteo"].values()) if state.get("conteo") else 0
            esperado = total_productos * len(codigos_set)

            st.caption(
                f"Notas exportadas para la selección — "
                f"{len(codigos_set)} nota(s) × {total_productos} producto(s) "
                f"= {esperado} esperadas:"
            )
            cols_notas = st.columns(3)
            for col, (etiqueta, n) in zip(cols_notas, conteo_notas.items()):
                col.metric(etiqueta, n)
            if esperado > 0:
                st.progress(
                    min(conteo_notas["OK"] / esperado, 1.0),
                    text=f"{conteo_notas['OK']}/{esperado} notas exportadas OK",
                )
        elif notas_filas is not None:
            st.caption("Elegí al menos una nota arriba para ver el progreso de notas.")
        elif state.get("conteo_notas_error"):
            st.error(f"No pude leer el histórico de notas: {state['conteo_notas_error']}")

    st.divider()

    en_curso = _scripts_en_curso()
    _render_caja_ejecutando(en_curso)
    _render_caja_corridas_sin_terminar(en_curso)


def _render_panel_derecho(key, cfg, state, sheet_url):
    with st.container(height=640, border=False):
        tab_cola, tab_config = st.tabs(["Cola", "Config"])

        with tab_config:
            # Fuera del fragment a propósito: Guardar dispara un rerun
            # completo normal, así el centro (credenciales, URL default
            # del Sheet) se actualiza al toque en vez de quedar un paso
            # atrás hasta el próximo tick del fragment de al lado.
            render_configuracion()

        with tab_cola:
            run_every = 1 if _hay_algo_corriendo() else None

            @st.fragment(run_every=run_every)
            def _fragment_cola():
                _render_tab_cola(key, cfg, state, sheet_url)

            _fragment_cola()


def _render_centro(key, cfg, state):
    st.subheader(cfg["label"])
    st.caption(cfg["help"])

    col_exp, col_run, col_abort = st.columns([5, 1, 1])

    with col_exp:
        with st.expander("Parámetros", expanded=not state["running"]):
            col_sheet, col_refresh = st.columns([4, 1])
            with col_sheet:
                sheet_url = st.text_input(
                    "URL del Google Sheet",
                    value=user_config.sheet_url_default(key),
                    key=f"sheet_url_{key}",
                    disabled=state["running"],
                    help="Se precarga con la URL guardada en Configuración para este script, si hay una. Siempre editable.",
                )
            with col_refresh:
                st.markdown("<div style='height: 1.9em'></div>", unsafe_allow_html=True)
                refrescar_clicked = st.button(
                    "🔄 Refrescar",
                    key=f"refrescar_{key}",
                    disabled=state["running"] or not sheet_url,
                    use_container_width=True,
                    help="Lee el Sheet ahora y cuenta cuántas filas están Pendiente/OK/Error.",
                )
            if refrescar_clicked:
                _refrescar_conteo(state, sheet_url, cfg["sheet"])
                if cfg.get("sheet_resultados"):
                    _refrescar_conteo_notas(state, sheet_url, cfg["sheet_resultados"])

            base_url = st.text_input(
                "URL de Tourplan",
                value=cfg["base_url"],
                key=f"url_{key}",
                disabled=state["running"],
                help="Viene precargada con producción. Cambiala solo si querés probar deliberadamente contra el ambiente de Test.",
            )

            modo_options = cfg.get("modo_options")
            if modo_options:
                modo_label = st.radio(
                    "Modo",
                    [label for label, _ in modo_options],
                    key=f"modo_{key}",
                    disabled=state["running"],
                )
                modo_env = dict(modo_options)[modo_label]
            else:
                st.info(cfg.get(
                    "no_modo_info",
                    "ℹ️ Este script no tiene modo de solo lectura: cada fila PENDIENTE se "
                    "copia y linkea directo, sin vista previa antes de guardar.",
                ))
                modo_env = None

            edicion_options = cfg.get("edicion_options")
            if edicion_options:
                edicion_label = st.radio(
                    "Si el código de nota ya existe en el producto",
                    [label for label, _ in edicion_options],
                    key=f"edicion_{key}",
                    disabled=state["running"],
                )
                edicion_env = dict(edicion_options)[edicion_label]
            else:
                edicion_env = None

            eliminar_options = cfg.get("eliminar_options")
            if eliminar_options:
                eliminar_label = st.radio(
                    "Acción sobre cada producto",
                    [label for label, _ in eliminar_options],
                    key=f"eliminar_{key}",
                    disabled=state["running"],
                )
                eliminar_env = dict(eliminar_options)[eliminar_label]
            else:
                eliminar_env = None

            codigos_nota_env = None
            if cfg.get("notas_a_exportar"):
                # Paso 1 — buscador: familias (con variante de idioma) + códigos
                # sueltos (sin idioma), todos al mismo nivel. Paso 2 — solo para
                # las familias elegidas: checkboxes de idioma (pedido explícito
                # de la usuaria: buscador en el paso 1, checkboxes en el paso 2).
                items_paso1 = st.multiselect(
                    "Qué notas exportar",
                    options=list(FAMILIAS_NOTA.keys()) + list(CODIGOS_SUELTOS.keys()),
                    key=f"notas_familias_{key}",
                    disabled=state["running"],
                    help="Elegí una o varias. Para las que tienen variante de idioma "
                         "(Nota SRV, Descriptivo, Titulo, Luggage Waiver), después "
                         "tildás abajo qué idiomas.",
                )
                codigos_a_exportar = []
                for item in items_paso1:
                    if item in FAMILIAS_NOTA:
                        st.caption(f"**{item}** — idiomas:")
                        idiomas = FAMILIAS_NOTA[item]
                        cols = st.columns(len(idiomas))
                        for col, idioma in zip(cols, idiomas):
                            tildado = col.checkbox(
                                idioma, value=True,
                                key=f"notas_idioma_{key}_{item}_{idioma}",
                                disabled=state["running"],
                            )
                            if tildado:
                                codigos_a_exportar.append(idiomas[idioma])
                    else:
                        codigos_a_exportar.append(CODIGOS_SUELTOS[item])
                codigos_nota_env = ",".join(codigos_a_exportar) if codigos_a_exportar else None
                if items_paso1 and not codigos_nota_env:
                    st.caption("⚠️ No queda ningún idioma tildado — destildá menos o elegí otra nota.")
                # Se guarda en state para que el panel derecho (_render_tab_cola,
                # que corre en otra columna/fragment) pueda calcular el
                # progreso de notas sobre la selección vigente.
                state["codigos_a_exportar"] = codigos_a_exportar

            username, password = user_config.tp_credenciales_default()
            if not (username and password):
                st.warning(
                    "Falta cargar tu usuario/contraseña de Tourplan en ⚙️ Configuración "
                    "(se completan solos acá una vez guardados)."
                )

    campos_completos = bool(sheet_url and username and password and base_url)
    if cfg.get("notas_a_exportar"):
        campos_completos = campos_completos and bool(codigos_nota_env)

    with col_run:
        st.markdown("<div style='height: 1.9em'></div>", unsafe_allow_html=True)
        run_clicked = st.button(
            "Ejecutar",
            key=f"run_{key}",
            disabled=state["running"] or not campos_completos,
            type="primary",
            use_container_width=True,
        )
    with col_abort:
        st.markdown("<div style='height: 1.9em'></div>", unsafe_allow_html=True)
        abort_clicked = st.button(
            "⏹ Abortar",
            key=f"abort_{key}",
            disabled=not state["running"] or state["abort_requested"],
            use_container_width=True,
            help="Corta al terminar la fila en curso, hace logout de Tourplan y deja "
                 "las filas restantes en PENDIENTE para retomar en otra corrida.",
        )
    if not campos_completos and not state["running"]:
        falta_notas = cfg.get("notas_a_exportar") and not codigos_nota_env
        st.caption(
            ("Elegí al menos una nota para exportar y c" if falta_notas else "C")
            + "ompletá la URL del Sheet y la URL de Tourplan (y tu usuario/password "
              "en ⚙️ Configuración) para poder ejecutar."
        )

    if run_clicked:
        state["mensaje_resultado"] = None
        run_dir = Path(tempfile.mkdtemp(prefix=f"tourplan_{key}_"))
        ss_dir = run_dir / "screenshots"
        ss_dir.mkdir(exist_ok=True)
        stop_file = run_dir / "ABORTAR.flag"

        env = os.environ.copy()
        env.update({
            "TOURPLAN_USERNAME": username,
            "TOURPLAN_PASSWORD": password,
            "TOURPLAN_BASE_URL": base_url,
            "TOURPLAN_SHEET_URL": sheet_url,
            "TOURPLAN_HOJA": cfg["sheet"],
            "TOURPLAN_CREDENTIALS_PATH": user_config.CREDENTIALS_PATH,
            "TOURPLAN_TOKEN_PATH": user_config.TOKEN_PATH,
            "TOURPLAN_HEADLESS": "1" if user_config.headless_default() else "0",
            "TOURPLAN_SS_DIR": str(ss_dir),
            "TOURPLAN_STOP_FILE": str(stop_file),
            "PYTHONPATH": str(REPO_ROOT) + os.pathsep + env.get("PYTHONPATH", ""),
            "PYTHONUNBUFFERED": "1",
            **({"TOURPLAN_MODO": modo_env} if modo_env else {}),
            **({"TOURPLAN_EDICION": edicion_env} if edicion_env else {}),
            **({"TOURPLAN_ELIMINAR": eliminar_env} if eliminar_env else {}),
            **({"TOURPLAN_CODIGOS_NOTA": codigos_nota_env} if codigos_nota_env else {}),
            "PYTHONIOENCODING": "utf-8",
        })

        proc = subprocess.Popen(
            [sys.executable, str(cfg["script_path"])],
            cwd=str(REPO_ROOT),
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
        )

        state.update({
            "running": True,
            "finished": False,
            "log_lines": [],
            "returncode": None,
            "proc": proc,
            "stop_file": stop_file,
            "abort_requested": False,
            "started_at": time.time(),
        })
        threading.Thread(target=_leer_proceso, args=(proc, state), daemon=True).start()
        st.rerun()

    if abort_clicked:
        state["abort_requested"] = True
        try:
            state["stop_file"].touch()
        except Exception:
            pass

    if state["abort_requested"] and state["running"]:
        st.warning(
            "⏸️ Abortando: termina la fila en curso, hace logout de Tourplan y corta. "
            "Puede tardar hasta un minuto — las filas no procesadas quedan en PENDIENTE."
        )

    run_every = 1 if _hay_algo_corriendo() else None

    @st.fragment(run_every=run_every)
    def _fragment_consola():
        # Detección de corrida terminada (antes vivía en el bucle
        # time.sleep(1)+st.rerun() de abajo de todo; ahora el mismo
        # chequeo corre en cada tick del fragment). El st.rerun() de acá
        # es con scope default ("app"): un rerun COMPLETO para que los
        # botones Ejecutar/Abortar y el header se rehabiliten — ver nota
        # de mensaje_resultado más abajo.
        if state["running"] and state["finished"]:
            state["running"] = False
            rc = state["returncode"]
            _refrescar_conteo(state, sheet_url, cfg["sheet"])
            if cfg.get("sheet_resultados"):
                _refrescar_conteo_notas(state, sheet_url, cfg["sheet_resultados"])

            if rc == 0:
                state["mensaje_resultado"] = (
                    "success",
                    f"Terminó OK (código de salida {rc}). Revisá el resultado en el Sheet.",
                )
            elif rc == ABORT_EXIT_CODE:
                state["mensaje_resultado"] = (
                    "info",
                    "⏸️ Abortado. Las filas que no llegó a procesar quedaron en PENDIENTE — "
                    "volvé a correr sobre el mismo Sheet más adelante para retomar.",
                )
            else:
                state["mensaje_resultado"] = (
                    "error",
                    f"El proceso terminó con error (código de salida {rc}). "
                    "Revisá el log arriba.",
                )
            st.rerun()

        col_titulo, col_descarga = st.columns([4, 1])
        with col_titulo:
            estado_txt = "" if state["running"] else " &nbsp; En espera"
            st.markdown(f"**Consola**{estado_txt}", unsafe_allow_html=True)
        with col_descarga:
            st.download_button(
                "Descargar",
                data="".join(state["log_lines"]),
                file_name=f"{key}_log.txt",
                mime="text/plain",
                key=f"descargar_{key}",
                use_container_width=True,
            )

        _render_consola_html(state["log_lines"])
        st.caption("Siguiendo la última línea")

        # El mensaje queda guardado en state (no es un chequeo puntual
        # de running/finished) para que siga visible después del
        # st.rerun() completo de arriba, que ya corre con running=False.
        mensaje = state.get("mensaje_resultado")
        if mensaje:
            tipo, texto = mensaje
            getattr(st, tipo)(texto)

    _fragment_consola()

    return sheet_url


def render_script_tab(key, cfg):
    state_key = f"_state_{key}"
    if state_key not in st.session_state:
        st.session_state[state_key] = {
            "running": False,
            "finished": False,
            "log_lines": [],
            "returncode": None,
            "proc": None,
            "stop_file": None,
            "abort_requested": False,
            "started_at": None,
            "mensaje_resultado": None,
            "_anotada_en_historial": False,
        }
    state = st.session_state[state_key]

    render_header(key)

    if cfg.get("warning"):
        st.warning(cfg["warning"])

    if not cfg["script_path"].exists():
        st.info("Este script todavía no fue vendorizado en el repo — va a estar disponible en la app cuando se pegue el código.")
        return

    col_centro, col_derecha = st.columns([3, 1], gap="medium")

    with col_centro:
        sheet_url = _render_centro(key, cfg, state)

    with col_derecha:
        _render_panel_derecho(key, cfg, state, sheet_url)


def render_configuracion():
    """Formulario de Configuración: se completa una vez por PC. Guarda en
    ~/.tourplan-nx-app/config.json (ver common/user_config.py) — nunca en
    el repo compartido, para no pisarse entre personas. Ahora vive en la
    pestaña "Config" del panel derecho (antes era una "vista" separada a
    pantalla completa, accesible desde un botón del sidebar)."""
    st.caption(
        "Se guarda en esta computadora (no se sube al repositorio ni se comparte "
        "con el resto del equipo)."
    )

    cfg = user_config.cargar()

    with st.form("form_configuracion"):
        st.subheader("Credenciales de Tourplan")
        st.caption(
            "Se completan automáticamente en cada script — ya no hace falta "
            "tipearlas antes de ejecutar."
        )
        col_user, col_pass = st.columns(2)
        with col_user:
            tp_usuario = st.text_input("Usuario Tourplan", value=cfg.get("tp_usuario", ""))
        with col_pass:
            tp_password = st.text_input(
                "Password Tourplan", value=cfg.get("tp_password", ""), type="password")

        headless = st.checkbox(
            "Correr sin ventana de Chrome visible (headless)",
            value=bool(cfg.get("headless", False)),
            help="Destildado (default): se abre una ventana de Chrome real mientras "
                 "corre cada script. Tildado: corre sin abrir ventana.",
        )

        st.subheader("URL de Sheet por script")
        st.caption(
            "Opcional — dejá vacío el que no uses. Se precarga al elegir ese script "
            "para ejecutar, pero siempre se puede pegar otra URL puntual ahí."
        )

        sheet_urls_guardadas = cfg.get("sheet_urls", {})
        nuevas_urls = {}
        current_category = None
        for key, script_cfg in SCRIPTS.items():
            if script_cfg["category"] != current_category:
                current_category = script_cfg["category"]
                st.markdown(f"**{current_category}**")
            nuevas_urls[key] = st.text_input(
                script_cfg["label"],
                value=sheet_urls_guardadas.get(key, ""),
                key=f"cfg_sheet_url_{key}",
            )

        guardado = st.form_submit_button("Guardar", type="primary", use_container_width=True)

    if guardado:
        user_config.guardar({
            "tp_usuario": tp_usuario,
            "tp_password": tp_password,
            "headless": headless,
            "sheet_urls": nuevas_urls,
        })
        st.success("Configuración guardada.")


def render_sidebar_nav():
    """Sidebar agrupado por categoría. Guarda la selección en
    st.session_state para que no se resetee al interactuar con los inputs
    del script elegido (esos widgets viven en el área principal, no acá).
    Ya no tiene botón de Configuración: Config pasa a ser una pestaña del
    panel derecho, no una vista aparte."""
    if "selected_script" not in st.session_state:
        st.session_state["selected_script"] = next(iter(SCRIPTS))

    st.sidebar.title("Scripts")
    current_category = None
    for key, cfg in SCRIPTS.items():
        if cfg["category"] != current_category:
            current_category = cfg["category"]
            st.sidebar.markdown(f"**{current_category}**")
        selected = st.session_state["selected_script"] == key
        if st.sidebar.button(
            cfg["label"],
            key=f"nav_{key}",
            type="primary" if selected else "secondary",
            use_container_width=True,
        ):
            st.session_state["selected_script"] = key
            st.rerun()


def main():
    st.set_page_config(page_title="Tourplan NX - Herramientas", layout="wide")
    _inyectar_css()
    _asegurar_pendientes_globales()

    render_sidebar_nav()
    selected_key = st.session_state["selected_script"]
    render_script_tab(selected_key, SCRIPTS[selected_key])


if __name__ == "__main__":
    main()
