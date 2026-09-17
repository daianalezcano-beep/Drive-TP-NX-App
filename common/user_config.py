"""
user_config.py — Configuración local por persona/PC de TP-NX-App.

Guarda la URL default de Sheet por script (una entrada por key del
dict SCRIPTS de app.py) y los paths de credenciales OAuth de Google,
en ~/.tourplan-nx-app/ — FUERA del repo/carpeta compartida por todo el
equipo. app.py y common/ viven en una carpeta de red única para todo
el equipo; un config al lado del código (como config_store.py de TP
Documentación, donde cada quien tiene su propia instalación completa)
se pisaría entre personas acá. Mismo tipo de solución que ya usa
run_app.bat para el venv (instalar fuera del repo por el límite de
rutas de Windows), aplicado a los datos personales en vez de al
entorno de Python.

No hay selector de usuario en la UI: siempre es 1 persona = 1 PC = 1
instancia local, así que la identidad la da la carpeta del perfil de
Windows/Mac de quien corre la app, no un login dentro de la app.
"""

import json
import os

CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".tourplan-nx-app")
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")
CREDENTIALS_PATH = os.path.join(CONFIG_DIR, "credentials.json")
TOKEN_PATH = os.path.join(CONFIG_DIR, "token.json")


def cargar():
    """Devuelve {script_key: sheet_url} guardado en esta PC. Vacío si
    todavía no se guardó nada desde la pantalla de Configuración."""
    if not os.path.exists(CONFIG_PATH):
        return {}
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def guardar(sheet_urls):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(sheet_urls, f, ensure_ascii=False, indent=2)


def sheet_url_default(script_key):
    return cargar().get(script_key, "")
