"""Tourplan NX - Herramientas: app local (Streamlit) que unifica los scripts
de automatizacion de Tourplan NX en una sola herramienta con UI.

Scripts activos, agrupados por categoría:
- Productos: Copy Products, Flag as Deleted, Modificar Description y Comment.
- PCM: Copy y Linkeo de PCM.
- Valorización: Desde Componentes, Desde Servicio Madre, Desde Excel, Tarifario estático.
- Notas: Notas SRV (insertar/editar Product Notes), Exportar Notas (solo lectura).
- Relevamiento: Vigencias (lista de períodos de RATES sin costos, solo lectura).

Cada script corre como subproceso propio (ver README.md - "Por que
subprocess"), parametrizado por variables de entorno. Las credenciales
nunca se escriben a disco: viajan solo en el entorno del subproceso.
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
MODO_LECTURA_ESCRITURA = [
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
SCRIPTS = {
    "copy_products": {
        "label": "Copy Products (05)",
        "category": "Productos",
        # Reemplaza a los copy_products/rename_products viejos (eliminados).
        # Fuente: https://github.com/daianalezcano-beep/Copy-products/tree/claude/new-option-copy-script-25yt7w
        "help": "Copy Product con la posibilidad de modificar todos los campos",
        "script_path": REPO_ROOT / "scripts" / "copy_products" / "copy_products.py",
        "base_url": PRODUCCION_URL,
        "sheet": "PRODUCTOS",
        "modo_options": MODO_LECTURA_ESCRITURA,
    },
    "flag_as_deleted": {
        "label": "Flag as Deleted (10)",
        "category": "Productos",
        "help": "Marca productos como \"Flag Product as Deleted\" en Tourplan.",
        "script_path": REPO_ROOT / "scripts" / "flag_as_deleted" / "flag_products_as_deleted.py",
        "base_url": PRODUCCION_URL,
        "sheet": "PRODUCTOS",
        "modo_options": MODO_LECTURA_ESCRITURA,
    },
    "modificar_description_comment": {
        "label": "Modificar Description y Comment (15)",
        "category": "Productos",
        # Fuente: https://github.com/daianalezcano-beep/Copy-products/tree/claude/nuevo-branch-repositorio-p32d22
        "help": "Modifica los campos description y/o comment de un producto en Tourplan.",
        "script_path": REPO_ROOT / "scripts" / "modificar_description_comment" / "modificar_description_comment.py",
        "base_url": PRODUCCION_URL,
        "sheet": "PRODUCTOS",
        "modo_options": MODO_LECTURA_ESCRITURA,
    },
    "copy_pcm_linkeo": {
        "label": "Copy y Linkeo PCM (20)",
        "category": "PCM",
        "help": "Copia un PCM (Package Header) desde un servicio madre y lo linkea a un product code nuevo.",
        "script_path": REPO_ROOT / "scripts" / "copy_pcm_linkeo" / "tourplan_copiar_pcm.py",
        "base_url": PRODUCCION_URL,
        "sheet": "PRODUCTOS",
    },
    "valorizacion_pkg": {
        "label": "Valorización desde Componente (25)",
        "category": "Valorización",
        "help": "Valoriza servicios madre PKG a partir de un componente (COD ORIGEN): lee costos de los PCM hijos y escribe las tarifas en el servicio madre.",
        "script_path": REPO_ROOT / "scripts" / "valorizacion_pkg" / "tourplan_valorizacion_pkg_v3.py",
        "base_url": PRODUCCION_URL,
        "sheet": "PRODUCTOS",
        "modo_options": MODO_FASES,
    },
    "valorizacion_madre": {
        "label": "Valorización desde Servicio Madre (30)",
        "category": "Valorización",
        "help": "Valoriza servicios madre PKG partiendo directo del servicio madre (sin componente): busca sus PCM tipo \"Package Header\" y aplica las tarifas.",
        "script_path": REPO_ROOT / "scripts" / "valorizacion_madre" / "valorizacion_desde_madre.py",
        "base_url": PRODUCCION_URL,
        "sheet": "PRODUCTOS",
        "modo_options": MODO_FASES,
    },
    "valorizacion_excel": {
        "label": "Valorización desde Excel (35)",
        "category": "Valorización",
        # Fuente: https://github.com/daianalezcano-beep/Valorizacion-EX-TF-dsd-servicio-madre/tree/claude/nuevo-script-471bvm
        "help": "Valoriza servicios con un solo pax break, desde un valor cargado en el excel.",
        "script_path": REPO_ROOT / "scripts" / "valorizacion_excel" / "valorizacion_desde_excel.py",
        "base_url": PRODUCCION_URL,
        "sheet": "TARIFAS",
        "modo_options": MODO_LECTURA_ESCRITURA,
    },
    "valorizacion_madre_numericos": {
        "label": "Valorización Tarifario estático (40)",
        "category": "Valorización",
        "help": "Variante de \"Valorización desde Servicio Madre\" con el filtro de códigos invertido: procesa únicamente los servicios madre PKG cuyo código empieza con un número (los que ese script descarta). Busca sus PCM tipo \"Package Header\" y aplica las tarifas.",
        "script_path": REPO_ROOT / "scripts" / "valorizacion_madre_numericos" / "valorizacion_desde_madre_numericos.py",
        "base_url": PRODUCCION_URL,
        "sheet": "PRODUCTOS",
        "modo_options": MODO_FASES,
    },
    "notas_srv": {
        "label": "Crear/Modificar (45)",
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
        "label": "Exportar (50)",
        "category": "Notas",
        # Fuente: https://github.com/daianalezcano-beep/copy-products/tree/notas-SRV
        "help": "Exporta a una columna del Excel el contenido de una nota (Product Notes) ya existente en un producto. El código de nota no está limitado a una lista fija.",
        "script_path": REPO_ROOT / "scripts" / "exportar_notas" / "exportar_notas.py",
        "base_url": PRODUCCION_URL,
        "sheet": "EXPORTAR_NOTAS",
        "no_modo_info": "ℹ️ Este script es de solo lectura: exporta el contenido de la nota indicada a la columna Texto_Exportado del Excel, sin insertar, editar ni guardar nada en Tourplan.",
    },
    "extraccion_vigencias": {
        "label": "Vigencias (55)",
        "category": "Relevamiento",
        # Fuente: https://github.com/daianalezcano-beep/generico-vs-especificos/tree/claude/rates-extraction-script-lm4tyn
        "help": "Relevamiento de vigencias: dado un supplier/location/service type (o código puntual), exporta el estado de los períodos de RATES de cada producto (fechas, Price Code, moneda, estado) sin leer costos. Sirve para detectar tarifas vencidas o sin cargar.",
        "script_path": REPO_ROOT / "scripts" / "extraccion_vigencias" / "extraccion_vigencias.py",
        "base_url": PRODUCCION_URL,
        "sheet": "PRODUCTOS",
        "no_modo_info": "ℹ️ Este script es de solo lectura: exporta la vigencia de cada período de RATES a la hoja VIGENCIAS, sin insertar, editar ni guardar nada en Tourplan.",
    },
}


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
            "result_path": None,
            "returncode": None,
            "proc": None,
            "stop_file": None,
            "abort_requested": False,
        }
    state = st.session_state[state_key]

    uploaded = st.file_uploader(
        "Excel de entrada (.xlsx)", type=["xlsx"], key=f"upload_{key}", disabled=state["running"])

    col1, col2 = st.columns(2)
    with col1:
        username = st.text_input("Usuario Tourplan", key=f"user_{key}", disabled=state["running"])
    with col2:
        password = st.text_input(
            "Password Tourplan", type="password", key=f"pass_{key}", disabled=state["running"])

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

    campos_completos = bool(uploaded and username and password and base_url)

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
        st.caption("Completá Excel, usuario, password y URL para poder ejecutar.")

    if run_clicked:
        run_dir = Path(tempfile.mkdtemp(prefix=f"tourplan_{key}_"))
        safe_name = Path(uploaded.name).name  # evita path traversal desde el nombre subido
        excel_path = run_dir / safe_name
        excel_path.write_bytes(uploaded.getbuffer())
        ss_dir = run_dir / "screenshots"
        ss_dir.mkdir(exist_ok=True)
        stop_file = run_dir / "ABORTAR.flag"

        env = os.environ.copy()
        env.update({
            "TOURPLAN_USERNAME": username,
            "TOURPLAN_PASSWORD": password,
            "TOURPLAN_BASE_URL": base_url,
            "TOURPLAN_EXCEL_PATH": str(excel_path),
            "TOURPLAN_HOJA": cfg["sheet"],
            "TOURPLAN_SS_DIR": str(ss_dir),
            "TOURPLAN_STOP_FILE": str(stop_file),
            "PYTHONPATH": str(REPO_ROOT) + os.pathsep + env.get("PYTHONPATH", ""),
            "PYTHONUNBUFFERED": "1",
            **({"TOURPLAN_MODO": modo_env} if modo_env else {}),
            **({"TOURPLAN_EDICION": edicion_env} if edicion_env else {}),
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
            "result_path": None,
            "returncode": None,
            "proc": proc,
            "excel_path": excel_path,
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
        excel_path = state.get("excel_path")
        if excel_path and Path(excel_path).exists():
            state["result_path"] = str(excel_path)

        if rc == 0 and state["result_path"]:
            st.success(f"Terminó OK (código de salida {rc}).")
        elif rc == 0:
            st.warning("El proceso terminó OK pero no encontré el Excel resultado.")
        elif rc == ABORT_EXIT_CODE:
            st.info(
                "⏸️ Abortado. Las filas que no llegó a procesar quedaron en PENDIENTE — "
                "descargá el Excel para ver hasta dónde llegó y volvé a subirlo más adelante "
                "para retomar."
            )
        else:
            st.error(
                f"El proceso terminó con error (código de salida {rc}). "
                "Revisá el log arriba."
            )

    if state["running"]:
        time.sleep(1)
        st.rerun()

    if state.get("result_path") and os.path.exists(state["result_path"]):
        with open(state["result_path"], "rb") as f:
            st.download_button(
                "Descargar Excel con el resultado",
                data=f.read(),
                file_name=Path(state["result_path"]).name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{key}",
            )


def render_sidebar_nav():
    """Sidebar agrupado por categoría. Guarda la selección en
    st.session_state para que no se resetee al interactuar con los inputs
    del script elegido (esos widgets viven en el área principal, no acá)."""
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
    st.set_page_config(page_title="Tourplan NX - Herramientas", layout="centered")
    st.title("Tourplan NX - Herramientas")
    st.caption(
        "Corre siempre contra producción. "
        "Las credenciales no se guardan en disco ni se comparten entre corridas."
    )

    render_sidebar_nav()
    selected_key = st.session_state["selected_script"]
    render_script_tab(selected_key, SCRIPTS[selected_key])


if __name__ == "__main__":
    main()
