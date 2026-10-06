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
"""
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
    nota (EXPORTAR_NOTAS_RESULTADOS) — guarda en state["conteo_notas"]."""
    try:
        ws = conectar_sheets(
            sheet_url, hoja, user_config.CREDENTIALS_PATH, user_config.TOKEN_PATH)
        filas, _ = cargar_sheet(ws)
        state["conteo_notas"] = _contar_estados_notas(filas)
        return True
    except Exception as e:
        state["conteo_notas_error"] = str(e)
        return False


def _chequear_pendientes_globales():
    """Chequeo único al abrir la app (no se repite en cada rerun de
    Streamlit — ver render_aviso_pendientes): para cada script con una URL
    de Sheet guardada en Configuración, cuenta filas Pendiente/En proceso.
    Un error puntual leyendo un Sheet no bloquea el chequeo de los demás."""
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


def render_aviso_pendientes():
    """Aviso arriba de todo si quedaron filas sin procesar de una corrida
    anterior (abortada o cortada). El chequeo es una sola vez por sesión de
    navegador (al abrir la app) — no depende del botón Refrescar de cada
    script, que es para ver el progreso de una corrida activa."""
    if "_pendientes_globales" not in st.session_state:
        st.session_state["_pendientes_globales"] = _chequear_pendientes_globales()

    pendientes = st.session_state["_pendientes_globales"]
    if not pendientes:
        return

    st.warning("⚠️ Tenés corridas sin terminar:")
    for item in pendientes:
        col_txt, col_btn = st.columns([4, 1])
        with col_txt:
            partes = []
            if item["pendiente"]:
                partes.append(f"{item['pendiente']} PENDIENTE")
            if item["procesando"]:
                partes.append(f"{item['procesando']} en PROCESANDO (quedó a medias)")
            st.markdown(f"**{item['label']}** — {', '.join(partes)}")
        with col_btn:
            if st.button("▶ Ir a este script", key=f"ir_a_{item['key']}", use_container_width=True):
                st.session_state["selected_script"] = item["key"]
                st.session_state["vista"] = "script"
                st.rerun()
    st.divider()


def _leer_proceso(proc, state):
    """Corre en un hilo aparte: lee el stdout del subproceso sin bloquear el
    script de Streamlit, para que el botón Abortar pueda reaccionar mientras
    el script sigue corriendo."""
    for line in proc.stdout:
        state["log_lines"].append(line)
    proc.wait()
    state["returncode"] = proc.returncode
    state["finished"] = True


def render_script_tab(key, cfg):
    st.subheader(cfg["label"])
    st.caption(cfg["help"])
    if cfg.get("warning"):
        st.warning(cfg["warning"])

    if not cfg["script_path"].exists():
        st.info("Este script todavía no fue vendorizado en el repo — va a estar disponible en la app cuando se pegue el código.")
        return

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
        }
    state = st.session_state[state_key]

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

    if state.get("conteo"):
        conteo = state["conteo"]
        cols = st.columns(5)
        for col, (etiqueta, n) in zip(cols, conteo.items()):
            col.metric(etiqueta, n)

        total = sum(conteo.values())
        if total > 0:
            # OK + Error + En proceso, no "total - Pendiente": con celdas
            # ESTADO vacías contando ahora como "Otro" (ver
            # _clasificar_estado), "total - Pendiente" las contaría como
            # procesadas sin haberlo sido.
            avanzadas = conteo["OK"] + conteo["Error"] + conteo["En proceso"]
            st.progress(
                avanzadas / total,
                text=f"{avanzadas}/{total} procesadas",
            )
    elif state.get("conteo_error"):
        st.error(f"No pude leer el Sheet: {state['conteo_error']}")

    if cfg.get("sheet_resultados"):
        if state.get("conteo_notas"):
            conteo_notas = state["conteo_notas"]
            st.caption("Notas exportadas (histórico acumulado en EXPORTAR_NOTAS_RESULTADOS):")
            cols_notas = st.columns(3)
            for col, (etiqueta, n) in zip(cols_notas, conteo_notas.items()):
                col.metric(etiqueta, n)

            total_notas = sum(conteo_notas.values())
            if total_notas > 0:
                st.progress(
                    conteo_notas["OK"] / total_notas,
                    text=f"{conteo_notas['OK']}/{total_notas} notas exportadas OK",
                )
        elif state.get("conteo_notas_error"):
            st.error(f"No pude leer el histórico de notas: {state['conteo_notas_error']}")

    username, password = user_config.tp_credenciales_default()
    if not (username and password):
        st.warning(
            "Falta cargar tu usuario/contraseña de Tourplan en ⚙️ Configuración "
            "(se completan solos acá una vez guardados)."
        )

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

    campos_completos = bool(sheet_url and username and password and base_url)
    if cfg.get("notas_a_exportar"):
        campos_completos = campos_completos and bool(codigos_nota_env)

    col_run, col_abort = st.columns(2)
    with col_run:
        run_clicked = st.button(
            "Ejecutar",
            key=f"run_{key}",
            disabled=state["running"] or not campos_completos,
            type="primary",
            use_container_width=True,
        )
    with col_abort:
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

    log_placeholder = st.empty()
    if state["log_lines"]:
        log_placeholder.code("".join(state["log_lines"][-500:]), language=None)

    if state["running"] and state["finished"]:
        state["running"] = False
        rc = state["returncode"]
        _refrescar_conteo(state, sheet_url, cfg["sheet"])
        if cfg.get("sheet_resultados"):
            _refrescar_conteo_notas(state, sheet_url, cfg["sheet_resultados"])

        if rc == 0:
            st.success(f"Terminó OK (código de salida {rc}). Revisá el resultado en el Sheet.")
        elif rc == ABORT_EXIT_CODE:
            st.info(
                "⏸️ Abortado. Las filas que no llegó a procesar quedaron en PENDIENTE — "
                "volvé a correr sobre el mismo Sheet más adelante para retomar."
            )
        else:
            st.error(
                f"El proceso terminó con error (código de salida {rc}). "
                "Revisá el log arriba."
            )

    if state["running"]:
        if time.time() - state.get("conteo_ts", 0) > 5:
            _refrescar_conteo(state, sheet_url, cfg["sheet"])
            if cfg.get("sheet_resultados"):
                _refrescar_conteo_notas(state, sheet_url, cfg["sheet_resultados"])
        time.sleep(1)
        st.rerun()


def render_configuracion():
    """Pantalla de Configuración: se completa una vez por PC. Guarda en
    ~/.tourplan-nx-app/config.json (ver common/user_config.py) — nunca en
    el repo compartido, para no pisarse entre personas."""
    st.header("Configuración")
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
    del script elegido (esos widgets viven en el área principal, no acá)."""
    if "selected_script" not in st.session_state:
        st.session_state["selected_script"] = next(iter(SCRIPTS))
    if "vista" not in st.session_state:
        st.session_state["vista"] = "script"

    if st.sidebar.button(
        "⚙️ Configuración",
        key="nav_configuracion",
        type="primary" if st.session_state["vista"] == "config" else "secondary",
        use_container_width=True,
    ):
        st.session_state["vista"] = "config"
        st.rerun()

    st.sidebar.divider()
    st.sidebar.title("Scripts")
    current_category = None
    for key, cfg in SCRIPTS.items():
        if cfg["category"] != current_category:
            current_category = cfg["category"]
            st.sidebar.markdown(f"**{current_category}**")
        selected = st.session_state["vista"] == "script" and st.session_state["selected_script"] == key
        if st.sidebar.button(
            cfg["label"],
            key=f"nav_{key}",
            type="primary" if selected else "secondary",
            use_container_width=True,
        ):
            st.session_state["selected_script"] = key
            st.session_state["vista"] = "script"
            st.rerun()


def main():
    st.set_page_config(page_title="Tourplan NX - Herramientas", layout="centered")
    st.title("Tourplan NX - Herramientas")
    st.caption(
        "Corre siempre contra producción. "
        "Usuario/password de Tourplan se guardan localmente en esta PC (⚙️ Configuración) — "
        "nunca se suben al repositorio ni se comparten con el resto del equipo."
    )

    render_aviso_pendientes()

    render_sidebar_nav()
    if st.session_state["vista"] == "config":
        render_configuracion()
    else:
        selected_key = st.session_state["selected_script"]
        render_script_tab(selected_key, SCRIPTS[selected_key])


if __name__ == "__main__":
    main()
