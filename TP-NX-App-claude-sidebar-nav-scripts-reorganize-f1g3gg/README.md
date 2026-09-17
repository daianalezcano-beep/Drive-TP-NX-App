# Tourplan NX — Herramientas (app local)

App local (Streamlit) que unifica los scripts de automatización de Tourplan
NX en una sola herramienta con interfaz web, para que el equipo no técnico
pueda usarlos sin tocar código ni Google Colab. Corre en la propia PC de
cada persona.

**Estado:** Flag as Deleted, Copy y Linkeo PCM, Valorización desde
Componente y Valorización desde Servicio Madre están validados de punta a
punta en Windows real. Copy Products (versión nueva, con más campos
editables), Modificar Description y Comment, y Valorización desde Excel son
entradas nuevas del sidebar, vendorizadas en `scripts/` con los mismos dos
ajustes quirúrgicos que el resto (variables de entorno, Chrome
multiplataforma) pero **todavía sin correr contra Tourplan real** — vinieron
de Colab ya probadas ahí, pero no desde esta app. Rename Products se dio de
baja: quedó absorbido por la versión nueva de Copy Products. Notas que
siguen valiendo aunque ya no se muestren como advertencia en la UI:

- **Flag as Deleted** corre contra producción, sin ambiente de test.
  Actualizado a **v1.1** (2026-07-23) desde el repo original
  (Flag-as-deleted, antes la app tenía v1.0): SUPPLIER_CODE ahora también
  filtra la búsqueda (antes solo verificaba la fila en resultados),
  glosario fijo para SERVICE_TYPE, manejo del campo CLASS antes de Save, y
  espera del botón SAVE por polling activo en vez de un solo intento.
  **v1.1 ya validado contra producción real (2026-07-24)** — confirmado
  funcionando por la usuaria.
- **Copy y Linkeo PCM** no tiene modo lectura: cada fila PENDIENTE se copia
  y linkea directo, sin vista previa.
- **Valorización desde Componente** y **Valorización desde Servicio
  Madre** son scripts grandes en dos fases (leer costos → aplicar tarifas).
- **Valorización desde Componente** tiene una optimización de navegación de
  dos niveles (ver sección abajo): agrupa filas del mismo componente/
  servicio madre para no repetir la búsqueda completa por cada período, y
  además abre cada PCM una sola vez por grupo en vez de una vez por
  período. El primer nivel ya se corrió contra producción real; el
  segundo (abrir cada PCM una sola vez) todavía no se volvió a probar. Si
  tu próxima corrida tiene 2+ filas para el mismo componente o servicio
  madre, prestale atención especial al resultado.
- **Valorización desde Servicio Madre** recibió los 2 fixes de escritura de
  rates descriptos abajo (re-búsqueda de fila por texto + falso error de
  SAVE en períodos nuevos), portados desde Valorización desde Componente
  donde se detectaron y confirmaron contra producción real. Portados el
  mismo día que se confirmaron — **todavía no se corrieron en Tourplan
  real desde este script** (solo se validó que compila y llega limpio al
  mismo punto de siempre en este entorno). Probarlos en la próxima corrida
  real, idealmente con un servicio madre que necesite crear un período
  nuevo (para ejercitar el camino de COPY DATE RANGE).
- **Valorización desde Servicio Madre** también recibió (recién ahora) el
  primer nivel de optimización de navegación en Fase 2: agrupa filas del
  mismo servicio madre para no repetir la búsqueda completa por cada
  período — hasta ahora buscaba el servicio madre entero de nuevo por
  cada período, siempre, sin excepción (a diferencia de Valorización desde
  Componente, que ya tenía esto). Detectado en una corrida real: la
  usuaria notó que para servicios con 2+ períodos se creaba/aplicaba el
  primer período y se iba a buscar el siguiente SERVICIO en vez de seguir
  con los demás períodos del mismo. **Sin correr todavía contra Tourplan
  real** con este cambio.
- **Notas SRV** y **Exportar Notas** son la categoría nueva "Notas",
  vendorizadas desde el repo `copy-products` (rama `notas-SRV`) con los
  mismos dos ajustes de entorno que el resto — **todavía sin correr contra
  Tourplan desde esta app**. En el repo origen, `notas_srv.py` (INSERT) ya
  se confirmó funcionando de punta a punta en una corrida real sobre un
  batch de 305 filas; la lectura de la tabla de notas existentes
  (`encontrar_tabla_notas()`/`_fila_a_nota()`, usada para detectar "ya
  existe" y por `exportar_notas.py` para exportar) se reescribió recién en
  la última sesión (v1.12/v1.3) tras un diagnóstico por dump real de DOM, y
  todavía no se re-probó contra Tourplan; el flujo de EDITAR una nota
  existente (`EDICION="SI"`) tampoco se probó nunca. Ver `ESTADO.md` en el
  repo origen para el historial completo — no tocar la lógica de
  automatización sin leerlo antes.
- **Vigencias** es la categoría nueva "Relevamiento", vendorizada desde el
  repo `generico-vs-especificos` (rama `claude/rates-extraction-script-lm4tyn`)
  con los mismos ajustes de entorno que el resto. Los dos problemas de la
  primera corrida (`AttributeError` de `importlib.util`, y Chrome
  corriendo en modo `--headless` heredado de Colab sin abrir ventana) se
  corrigieron y **ya se confirmó funcionando contra Tourplan real
  (2026-08-29)**. Es de solo lectura: nunca inserta, edita ni guarda nada
  en Tourplan.

600BOX Creator se descartó del alcance (era una prueba de lo que fue Rename
Products, luego absorbido por la versión nueva de Copy Products). Los repos
originales (Copy-products, Flag-as-deleted, Copy-PCM-Linkeo,
Tourplan-Valorizacion-EX-TF, Valorizacion-EX-TF-dsd-servicio-madre) **no se
modifican** — los scripts que corren acá son copias parcheadas, vendorizadas
en `scripts/`.

## Cómo correrla

```bash
python -m venv .venv
source .venv/bin/activate  # en Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

En Windows, doble clic en `run_app.bat` hace lo mismo automáticamente
(crea el entorno la primera vez, instala dependencias, abre la app en el
navegador). Requiere tener Python instalado — si no lo tenés, `run_app.bat`
te va a avisar con el link de descarga. La empaquetada como `.exe` (sin
necesitar Python preinstalado) queda para una etapa siguiente.

**Google Chrome tiene que estar instalado en la PC.** La app no lo instala
sola en Windows/Mac (no hay garantía de permisos de administrador) — si no
lo encuentra, muestra un mensaje pidiendo instalarlo manualmente desde
https://www.google.com/chrome/.

**Mientras corre, se abre una ventana de Chrome de verdad** (a diferencia de
Colab, que corría "sin pantalla"). Es esperado — dejala en paz mientras la
app trabaja, se cierra sola al terminar.

## Qué cambió respecto de los scripts originales de Colab

Cada script vendorizado en `scripts/` tiene **solo dos ediciones
quirúrgicas** respecto del original (selectores, login y lógica de negocio
quedan intactos):

1. **Configuración por variables de entorno** en vez de hardcodeada
   (`TOURPLAN_USERNAME`, `TOURPLAN_PASSWORD`, `TOURPLAN_BASE_URL`,
   `TOURPLAN_MODO`, `TOURPLAN_EXCEL_PATH`, `TOURPLAN_HOJA`,
   `TOURPLAN_SS_DIR` y, solo en Notas SRV, `TOURPLAN_EDICION`), con el
   mismo valor original como default — mismo patrón que ya usaba
   `tourplan_copiar_pcm.py`.
2. **Detección de Chrome multiplataforma**, delegada a
   `common/chrome_bootstrap.py` en vez del bloque `apt-get`/`.deb`
   Linux-only original. En Windows/Mac nunca instala Chrome solo; en Linux
   mantiene el auto-instalador original (útil para poder probar la app en
   este entorno de desarrollo).

## Por qué subprocess (y no refactor a funciones)

Cada script corre como **proceso hijo** en vez de importarse como función.
Motivo: varios de los scripts originales (especialmente los más grandes,
como el de Valorización PKG) ejecutan todo su código a nivel de módulo, con
`SystemExit` incluido — refactorizarlos a funciones puras sería el camino de
mayor riesgo para lógica ya validada en producción. Con subprocess:

- Si Selenium/Chrome crashea fuerte, no se cae la app.
- El log en vivo es simple: se lee stdout línea por línea y se muestra en
  Streamlit a medida que llega.
- Las credenciales viajan solo en el entorno del subproceso — nunca se
  escriben a disco ni se loguean.

## Botón Abortar — cancelar una corrida en curso

Mientras un script está corriendo, la app muestra un botón **⏹ Abortar** al
lado de "Ejecutar". Está pensado para cuando hay que pausar el proceso para
retomarlo más tarde pero hace falta cerrar la app ahora: en vez de matar el
proceso a la fuerza (lo que puede dejar la sesión de Tourplan colgada,
consumiendo una de las licencias concurrentes limitadas), el aborto es
cooperativo — corta al terminar la fila en curso, hace el `logout()` normal
de Tourplan y recién ahí termina.

Detalle técnico (ver `common/abort.py`): al apretar Abortar, la app deja un
archivo bandera; cada script chequea su existencia entre fila y fila
(`chequear_abort()`) y, si aparece, corta el loop con una excepción que
hereda de `BaseException` (no de `Exception`, para que no la trague ningún
`except Exception` por fila que el script ya tenga) y salta directo al
`finally` que ya hace logout + cierre del navegador. Las filas que no
llegaron a procesarse quedan en `ESTADO=PENDIENTE` en el Excel — que ya se
guarda incrementalmente fila por fila — así que después de abortar el botón
de descarga se habilita igual (antes solo se habilitaba si la corrida
terminaba sin errores) para ver hasta dónde llegó y retomar más tarde
volviendo a subir ese mismo Excel.

Como el chequeo ocurre entre filas (no en medio de una espera de Selenium),
puede tardar hasta que termine la fila en curso — no es instantáneo.

## Modo (Lectura / Aplicar cambios)

Al auditar los scripts originales encontramos que en varios de ellos
`"escritura"` y `"completo"` ejecutan **exactamente el mismo código** — la
única distinción real es `"lectura"` (dry-run) vs. cualquier otra cosa. Por
eso la UI muestra solo 2 opciones reales ("Solo revisar" / "Aplicar
cambios") en vez de repetir una 3ra opción falsa. Esto no cambia el
comportamiento de ningún script, solo simplifica lo que ve el usuario.

## Notas SRV / Exportar Notas — categoría "Notas"

Dos scripts hermanos para trabajar con **Product Notes** de Tourplan NX
(vendorizados desde `copy-products`, rama `notas-SRV`; ver `ESTADO.md` de
ese repo para el detalle completo de selectores confirmados y bugs ya
resueltos — en particular la reescritura reciente de cómo se lee la tabla
de notas, que no usa Angular CDK virtual scroll como se pensó al
principio, sino tablas hermanas dentro de un contenedor común):

- **Notas SRV** inserta/edita notas — 25 códigos habilitados en
  `CODIGOS_NOTA_VALIDOS`: `NAL`/`NES`/`NFR`/`NIN`/`NIT` (Nota SRV),
  `UAL`/`UES`/`UFR`/`UIN`/`UIT` (Descriptivo), `LWE`/`LWF`/`LWG`/`LWI`/`LWS`
  (Luggage Waiver), `TAL`/`TES`/`TFR`/`TIN`/`TIT` (Título), y `DRT`, `PCR`,
  `REM`, `REO`, `SC2`. Todos confirmados en Tourplan con Category Type
  "Product" y Format Type "Plain Text" (captura real del catálogo de Note
  Category, 2026-08-28). El mismo catálogo tiene además varios códigos
  **HTML** (`OBS`, `NAP`, `SDV`, los `HY*` de hyperlinks, los `TB*` de Tax
  Brasil, etc.) que se dejaron **deliberadamente afuera**: el mecanismo de
  escritura de este script (bypass directo del `<body contenteditable>`
  del iframe, sin pasar por la API de CKEDITOR — ver "Problemas conocidos"
  en `ESTADO.md` del repo origen) está confirmado en corrida real solo
  para editores Plain Text, nunca se probó contra un editor HTML. Sumar
  esos códigos requeriría antes confirmar que el mismo mecanismo escribe
  bien ahí (o programar un camino distinto). Cualquier código que no esté
  en `CODIGOS_NOTA_VALIDOS` queda en `ERROR` sin tocar Tourplan, a
  propósito: no se asume el tipo de un código nuevo por su nombre. Por
  default nunca pisa una nota que ya existe: solo la reporta para revisión
  manual. Además del `Modo` (lectura/aplicar) común a otros scripts, tiene
  un segundo parámetro global de la corrida — no por fila —, "Si el
  código de nota ya existe en el producto": **No tocar** (default,
  reporta) o **Sobreescribir** (guarda el contenido anterior en la columna
  `CONTENIDO_ANTERIOR` del Excel antes de pisarlo). El Excel de entrada
  usa formato largo: una fila = un producto + un código de nota — ver
  `COMO_COMPLETAR_EL_EXCEL.md` del repo origen para el detalle y ejemplos
  (ese archivo, en el repo origen, todavía lista solo los 5 códigos "Nota
  SRV" — desactualizado respecto de `CODIGOS_NOTA_VALIDOS`).
- **Exportar Notas** lee el contenido de una nota que ya existe y lo
  vuelca a la columna `Texto_Exportado` del Excel — es de **solo lectura**,
  nunca inserta, edita ni guarda nada en Tourplan, así que no tiene modo
  lectura/aplicar (siempre "ejecuta", que acá solo significa "exportar").
  A diferencia de Notas SRV, el código de nota a buscar no está limitado a
  una lista fija — exporta lo que venga en la columna `Codigo_Nota` de
  cada fila.

**Encadenamiento por producto (ambos scripts):** si un mismo producto
necesita varias notas (varias filas con el mismo Location+Supplier+Code+
Service_Type pero distinto `Codigo_Nota`), tanto Notas SRV como Exportar
Notas agrupan esas filas (`_agrupar_por_producto()`) y buscan el producto
+ abren Product Notes **una sola vez por grupo** (`_buscar_y_abrir_
producto()`); cada `Codigo_Nota` del grupo se procesa después sobre esa
misma apertura (`process_nota()`), sin repetir la búsqueda. Mismo patrón
que ya usan los scripts de Valorización. Si un grupo entero falla al
buscar/abrir el producto, todas sus filas quedan en `ERROR` sin intentar
nada más.

**Re-sincronización — distinta en cada script, porque el riesgo que
mitiga es distinto:**
- **Notas SRV:** un SAVE exitoso (INSERT o EDIT) cierra el editor solo,
  así que entre notas de un mismo grupo la pantalla ya queda limpia sin
  hacer nada extra. El riesgo real es que una nota termine en `ERROR` y
  deje un diálogo/editor a medio abrir — antes, cada fila hacía
  `buscar_producto()` desde cero, lo que de paso "limpiaba" eso; con el
  encadenamiento ya no. Fix: si una nota del grupo termina en `ERROR` y
  quedan más notas de ese mismo producto, se vuelve a buscar y abrir el
  producto (`_buscar_y_abrir_producto()`) antes de la próxima.
- **Exportar Notas:** no hay SAVE, pero tampoco hace falta re-navegar —
  leer una nota abre su vista `#noteeditorview` (para leer el iframe) y
  hay que cerrarla explícitamente con su botón **"Exit"**
  (`cerrar_nota()`, click en `button.tpcancel` dentro de
  `#noteeditorview`, con fallback por texto "EXIT" — mismo botón/criterio
  que `_cerrar_ultimo_tp_dialog()` en Valorización desde Servicio Madre).
  `process_nota()` la cierra sola, en un `finally`, apenas termina de leer
  (haya salido bien o no) — así que en el caso normal no hace falta nada
  más entre notas del mismo grupo. **Bug real encontrado por la usuaria en
  la primera corrida (2026-08-29):** sin este click, la nota quedaba
  abierta y la del producto siguiente se abría encima, sin cerrar la
  anterior — "se abren varias notas juntas". El riesgo residual (si
  `cerrar_nota()` no logra cerrarla) se cubre igual que en Notas SRV: si
  una nota termina en `ERROR` y quedan más del mismo producto, se
  re-sincroniza — primero liviano (`abrir_product_notes()` sola, sin
  repetir la búsqueda del producto) y, si eso también falla, una
  re-sincronización completa (`_buscar_y_abrir_producto()`).
- **En ambos:** si ni la re-sincronización funciona, todas las notas
  restantes de ese grupo quedan en `ERROR` sin seguir intentando sobre un
  producto que no se puede reabrir.

**Notas SRV — encadenamiento por producto confirmado contra Tourplan real
(2026-08-29)** — la usuaria corrió el script tras estos cambios y reportó
que el agrupamiento funciona bien. La re-sincronización tras una nota en
`ERROR` en particular no se reportó como ejercitada en esa corrida (no se
mencionó una nota que fallara en el medio de un grupo), así que sigue
apoyada solo en las simulaciones en Python (los 3 escenarios: nota en
`ERROR` con más notas pendientes, `ERROR` en la última nota del grupo, y
`ERROR` en la re-sincronización misma) hasta confirmarla con un caso real.

**Exportar Notas — encadenamiento agregado, primer intento sin cerrar la
nota (bug encontrado por la usuaria), corregido con el click en "Exit".**
El primer intento de esta sesión re-sincronizaba siempre con
`abrir_product_notes()` entre notas, asumiendo (sin haber visto el HTML
real) que no había forma de cerrar la vista de lectura — la usuaria
corrió el script contra Tourplan real, vio el bug ("se abren varias notas
juntas") y compartió el HTML del botón "Exit" real. Corregido:
`process_nota()` ahora cierra la nota ella misma con ese botón, y la
re-sincronización en MAIN vuelve a ser solo-en-`ERROR`, igual que Notas
SRV. Verificado con una simulación en Python de la lógica de control (7
escenarios: agrupamiento con filas intercaladas; grupo entero sin
errores, sin ningún resync de más; una nota falla en el medio y sí
dispara resync; falla el resync liviano pero el completo la salva; fallan
ambos y el resto del grupo queda en `ERROR` cortando el grupo; falla la
búsqueda del producto y todo el grupo queda en `ERROR`; la última nota
del grupo falla sin disparar resync de más) — **sin correr todavía contra
Tourplan real con el fix del botón "Exit"**.

## Vigencias — categoría "Relevamiento"

Script de relevamiento de vigencias de tarifas, vendorizado desde el repo
[`generico-vs-especificos`](https://github.com/daianalezcano-beep/generico-vs-especificos),
rama `claude/rates-extraction-script-lm4tyn` (`extraccion_vigencias.py`),
con los mismos ajustes de entorno que el resto de los scripts
vendorizados (Chrome multiplataforma vía `common/chrome_bootstrap.py`,
config por variables de entorno) — no se tocó la lógica de negocio, los
selectores ni los nombres de función del script original.

**Fixes aplicados tras la primera corrida real (2026-08-29):**
- El script fallaba al toque con `AttributeError: module 'importlib' has
  no attribute 'util'`. Es el mismo bug ya conocido y corregido en
  `valorizacion_desde_excel.py`: `import importlib` (sin `.util`) no
  garantiza que el submódulo `importlib.util` esté disponible como
  atributo fuera de un proceso donde algo más ya lo haya importado antes
  (como pasaba en Colab, por azar de qué paquetes se cargan antes). Fix:
  `import importlib.util` explícito en el import inicial.
- `crear_driver()` corría con `--headless=new` (heredado tal cual del
  original de Colab, donde no hay pantalla) y sin perfil de Chrome
  aislado — a diferencia del resto de los scripts vendorizados en esta
  app, que sacan `--headless` y usan un `--user-data-dir` temporal para
  correr con ventana visible en la PC de la persona. La usuaria notó que
  no se abría ninguna ventana de Chrome al correrlo (a diferencia de los
  demás scripts). Fix: mismo patrón que `valorizacion_desde_excel.py` —
  sin `--headless`, perfil temporal aislado por corrida.

Ninguno de los dos fixes toca lógica de negocio ni selectores — son el
mismo tipo de ajuste "de entorno" que el resto de la vendorización.
**Confirmado funcionando contra Tourplan real (2026-08-29)** tras
aplicar ambos.

Dado un supplier/location/service type y/o un código puntual (ninguno de
los 4 campos es obligatorio por sí solo, pero se exige que vengan
completos al menos 2 de LOCATION/SUPPLIER/SERVICE TYPE/CODIGO), busca los
product codes correspondientes en Tourplan, entra al módulo RATES de cada
uno y exporta las columnas de la lista de períodos (Rate Period, PC,
Buy/Sell Currency, Sale Period, Rate Status, Rate Text, Rate Name) **sin
abrir ningún período individual** — es un relevamiento de vigencias, no
una lectura de costos. Sirve para detectar tarifas vencidas, sin cargar o
con estado no confirmado.

Es de **solo lectura**: nunca inserta, edita ni guarda nada en Tourplan,
así que no tiene modo lectura/aplicar (como Exportar Notas). Usa Excel
como cola de trabajo: hoja `PRODUCTOS` de entrada (con `ESTADO`/
`OBSERVACIONES`, resumible fila a fila) y hoja `VIGENCIAS` de salida
(append-only), ambas en el mismo archivo — a diferencia de las demás
entradas de la app, el nombre de la hoja de salida queda fijo en
`"VIGENCIAS"` (no es configurable desde la app, que solo pasa el nombre
de la hoja de entrada).

**Sin correr todavía contra Tourplan real desde esta app** — vino de
Colab ya probado ahí (según el repo origen), pero no desde `scripts/
extraccion_vigencias/`. Verificado en este repo solo con `python -m
py_compile` y con un import en limpio de `app.py` (sin correr Selenium).

## URL de Tourplan — producción por default, editable

El campo de URL viene precargado con `https://tourplannx.eurotur.com.ar/tourplannx`
(producción) en todos los scripts, pero sigue siendo editable — correr
contra Test es una decisión deliberada (cambiar el campo a mano), no algo
que pase sin querer por dejar el default de un script suelto en Test.

**Flujo de trabajo para scripts nuevos o cambios grandes:** primero se
iteran y se pulen en Google Colab (ahí sí tiene sentido probar contra Test
y hacer prueba-y-error de selectores) hasta que la lógica de negocio
funciona de punta a punta. Recién cuando el script está sólido se
vendoriza acá, se le aplican los ajustes de entorno para Windows (ver
sección anterior) y se agrega como pestaña nueva. Incluso así, conviene
una vuelta corta de validación en una PC real — no por la lógica de
negocio (esa ya viene probada), sino por posibles bugs específicos de
Windows/la app (justo el tipo de cosas que aparecieron con cada script
nuevo hasta ahora: Chrome, perfiles, encoding, rutas).

## Optimización de navegación (Valorización desde Componente)

Detectado en el uso real: si el Excel tiene 2+ filas para el mismo
componente (Fase 1) o el mismo servicio madre (Fase 2) — típicamente dos
períodos distintos del mismo producto — el script original repetía la
búsqueda completa (Product Search → USED IN → lista de PCMs, o Product
Search → RATES del servicio madre) para cada fila, aunque el componente/
servicio madre ya estuviera identificado.

Ahora `tourplan_valorizacion_pkg_v3.py` agrupa las filas pendientes por
componente (Fase 1) y por servicio madre (Fase 2) antes de procesarlas:

- **Fase 1:** `_buscar_componente_fase1()` hace la búsqueda del producto +
  lista de PCMs **una sola vez por grupo**. Además — segundo nivel de
  optimización, agregado después de una corrida real que mostró el
  problema — `_procesar_pcm_todos_periodos()` **abre cada PCM una sola vez
  por grupo** (no una vez por período): entra al PCM, aplica Change Base
  Date + lee costos para **todos** los períodos del grupo sobre esa misma
  ventana ya abierta (Change Base Date, Markup/Commission y Price Code
  navegan dentro de la ventana ya abierta vía el menú hamburguesa, no hace
  falta cerrar y reabrir entre período y período), y recién ahí cierra la
  ventana. Antes, con N períodos y M PCMs, se abría y cerraba cada PCM N
  veces (uno por período) — reescaneando las ~1900 filas de USED IN cada
  vez; ahora se abre M veces en total (una por PCM), sin importar cuántos
  períodos tenga el grupo.
- **Fase 2:** `_buscar_servicio_madre_fase2()` busca el servicio madre **una
  sola vez por grupo**; `actualizar_rates_servicio_madre()` sigue
  navegando a RATES por cada período (es barato y necesario, porque la
  verificación post-SAVE del período anterior puede dejar al driver
  parado en el detalle de otro período), pero ya no repite la búsqueda del
  producto. Esta fase **no** tuvo el mismo problema reportado para Fase 1 —
  no se le aplicó (todavía) el mismo nivel de optimización por-apertura,
  a confirmar con la usuaria si hace falta.

Si una fila queda sola (sin otras del mismo componente/servicio madre), el
comportamiento es idéntico al de antes — solo cambia cuando hay 2+ filas
agrupables. La lógica de agrupamiento se probó de forma aislada (agrupa y
preserva el orden correctamente), pero el flujo de Selenium en sí **todavía
no se volvió a correr contra Tourplan real con este segundo nivel de
optimización** — la primera versión (agrupar por componente, pero seguir
abriendo cada PCM una vez por período) sí se corrió contra producción real
y fue la que reveló este problema; probar esta versión nueva con un Excel
que tenga 2+ filas del mismo componente/servicio madre antes de confiar en
él para corridas grandes.

## Optimización de navegación (Valorización desde Servicio Madre — Fase 1)

Detectado por la usuaria en el uso real: por cada período cargado en el
Excel de Fase 1, el script volvía a buscar el servicio madre desde cero,
iba a Used In, identificaba el PCM Package Header, lo abría y recién ahí
recalculaba (Change Base Date + Voucher Cost) — aunque el período anterior
hubiera sido del **mismo** servicio madre. Mismo problema ya resuelto antes
para Valorización desde Componente y para la Fase 2 de este mismo script,
pero la Fase 1 de `valorizacion_desde_madre.py` no lo tenía.

La dificultad acá es que cada fila del Excel de Fase 1 no siempre apunta a
un servicio madre puntual: si la columna `SERVICE_CODE` viene vacía,
dispara una búsqueda masiva (supplier+service type+location) que puede
devolver varios códigos — no se sabe qué servicios va a tocar una fila
hasta correr esa búsqueda. Por eso el agrupamiento se hizo en dos pasos:

1. **Resolución** (`_resolver_fila_fase1()`) — por cada fila pendiente,
   resuelve su lista de servicios madre (directo si trae `SERVICE_CODE`, o
   vía búsqueda masiva si no) y clasifica cada uno (Package a procesar /
   Non-Accommodation o categoría desconocida a saltear), sin abrir ningún
   producto todavía. Arma una lista plana de tareas (fila + código).
2. **Agrupamiento** (`_agrupar_tareas_fase1()`) — agrupa esas tareas por
   servicio madre real (location+supplier+código+service type), sin
   importar si el mismo código se descubrió por `SERVICE_CODE` directo en
   una fila o por búsqueda masiva en otra.
3. **Procesamiento por grupo** — `_buscar_servicio_fase1()` busca el
   producto, lee días de operación, va a Used In y busca el PCM Package
   Header **una sola vez por grupo**; `_procesar_pcm_grupo_fase1()` abre
   ese PCM **una sola vez** y, sobre esa misma ventana ya abierta, aplica
   Change Base Date + lee Voucher Cost + Price Code para **todos** los
   períodos del grupo, cerrándolo recién al final.
4. El resultado por fila del Excel (columnas `SERVICIOS PKG PROCESADOS` /
   `SERVICIOS SKIP`) se sigue escribiendo igual que antes — se arma
   juntando lo que aportó cada tarea resuelta a partir de esa fila
   puntual, sin importar en qué grupo terminó procesándose.

Si una fila queda sola (sin otras del mismo servicio madre), el
comportamiento es idéntico al de antes — solo cambia cuando 2+ filas
(directas o vía búsqueda masiva) terminan apuntando al mismo servicio
madre. Verificado con una simulación en Python de la resolución +
agrupamiento + atribución de resultados (mezclando filas con `SERVICE_CODE`
directo y filas de búsqueda masiva apuntando al mismo código, y el caso de
un grupo sin PCM encontrado) — **sin correr todavía contra Tourplan real**.

Deliberadamente **no** se tocó `tourplan_valorizacion_pkg_v3.py`, que ya
tenía este nivel de optimización para su propia Fase 1 desde antes (ver
sección "Optimización de navegación (Valorización desde Componente)"
arriba) — ese script no tiene el problema de la búsqueda masiva por fila,
porque cada fila ya apunta a un componente puntual.

## Optimización de navegación (Valorización desde Servicio Madre — Fase 2)

Detectado en una corrida real: `valorizacion_desde_madre.py` no tenía
NINGÚN agrupamiento en Fase 2 — `actualizar_rates_servicio_madre()` hacía
`buscar_producto()` (la búsqueda completa del servicio madre) en CADA
llamado, sin excepción, así que un servicio madre con 2+ períodos volvía a
buscarse desde cero por cada uno. La usuaria lo notó en vivo: para varios
servicios se creaba/aplicaba el primer período y el script se iba
derecho a buscar el SIGUIENTE SERVICIO en vez de seguir con los demás
períodos del mismo servicio madre.

Fix (mismo patrón que ya tenía Valorización desde Componente): se agregó
`_agrupar_por_servicio_madre()` y `_buscar_servicio_madre_fase2()` —
agrupa las filas pendientes de Fase 2 por servicio madre
(location+service_code+supplier+service_type) y busca el producto **una
sola vez por grupo**; `actualizar_rates_servicio_madre()` ya no hace la
búsqueda, asume que el producto ya está abierto (se le sacó el
`buscar_producto()` interno). La navegación a RATES sigue repitiéndose por
período (es barata y necesaria, la verificación post-SAVE puede dejar al
driver parado en el detalle de otro período).

**Sin correr todavía contra Tourplan real** — probarlo con un Excel que
tenga 2+ períodos para el mismo servicio madre antes de confiar en él para
corridas grandes. Este script sigue sin tener el segundo nivel (abrir/
reutilizar algo una sola vez en vez de por período) porque en Fase 2 no
hay equivalente al "abrir el PCM" de Fase 1 — no se detectó un problema
análogo acá.

## Optimización de navegación (Copy Products) — encadenar copias del mismo origen

Caso real de la usuaria: un Excel de 160 filas que copian el mismo
OPTION_CODE del mismo SUPPLIER_ORIGEN a 160 suppliers destino distintos
(manteniendo location/service type/código iguales, solo cambia el
supplier). El script hacía la búsqueda completa (home→product→abrir modal
→ 4 filtros→resultados→click fila) para **cada una** de las 160 filas,
aunque el origen fuera siempre el mismo. La usuaria notó que, tras guardar
una copia, Tourplan queda posicionado en la opción recién creada — y esa
opción, al tener la misma location+service type+código que el origen, es
un punto de partida tan válido como el origen para la siguiente copia del
lote.

Fix: `_agrupar_por_origen()` agrupa las filas PENDIENTE por origen
(SUPPLIER_ORIGEN+SERVICE_TYPE+OPTION_CODE+LOCATION). Para la primera fila
de cada grupo se hace la búsqueda completa de siempre (`copy_option()`);
para las filas siguientes del mismo grupo, si la fila anterior copió sin
cambiar location/service type/código en el destino
(`_copia_preserva_identidad()` — solo puede haber cambiado el supplier
destino, description o comment), se salta la búsqueda completa y se copia
directo desde donde quedó posicionado el driver (`copy_option_encadenado()`).
Si una fila del grupo falla, o cambia location/service type/código en el
destino, la cadena se corta ahí — la fila siguiente vuelve a buscar desde
cero, nunca se asume una posición sin haberla confirmado.

**Sin correr todavía contra Tourplan real** — la lógica de agrupamiento y
corte de cadena se probó de forma aislada (con búsqueda/copia simuladas:
todas encadenables, falla a mitad de un grupo, corte de cadena por cambio
de campo, y orígenes intercalados en el Excel), pero el flujo de Selenium
en sí todavía no se validó en una corrida real. Probar primero con un lote
chico (5–10 filas) antes de confiar en él para lotes grandes como el de
160 filas que motivó esta optimización.

## Fix — escritura de rates podía ir a la fila de rango de edad equivocada (Fase 2)

Detectado en una corrida real contra producción: al escribir los rates de
un servicio madre con muchos rangos de pax AD (ej. 24 filas), el script
reportaba "Los valores NO quedaron guardados... N/24 rangos difieren tras
recargar" — pero el patrón de los valores (el "leído" de una fila coincidía
con el "esperado" de la fila anterior) mostraba que **no era un problema de
lectura, sino de escritura**: los valores se estaban guardando en la fila
de rango de pax EQUIVOCADA.

Causa: la grilla de RATES usa virtual scroll de Angular (CDK), que reutiliza
un pool fijo de nodos `<tr>`/`<input>` del DOM y los re-vincula a otra fila
lógica a medida que se scrollea. `_escribir_rates()` capturaba la lista de
filas **una sola vez** al principio y después escribía fila por fila
haciendo `scrollIntoView` en cada input — el scroll de una fila podía hacer
que Angular reciclara el nodo ya cacheado de una fila anterior para
representar otra distinta, y el script seguía escribiendo en esa referencia
vieja sin darse cuenta.

Fix: en vez de cachear las filas de antemano, ahora se re-busca en vivo la
fila por su texto de rango (no por posición/índice) justo antes de escribir
cada una — igual que ya se hacía para leer y verificar valores (que siempre
usaron el texto del rango como clave, nunca el índice).

**Importante — dato ya escrito antes de este fix:** como el SAVE se hace
ANTES de la verificación post-guardado, cualquier servicio madre que haya
mostrado este error (o similar) en una corrida con el código viejo puede
tener valores de costo guardados en rangos de pax que no les corresponden,
más allá de lo que diga el Excel. Conviene revisar manualmente en Tourplan
los casos que dieron este error antes de este fix.

**Actualización tras una corrida posterior:** auditando a mano varios casos
de "N/24 rangos difieren", se confirmó que en esos casos el valor "leído"
coincidía EXACTO con el valor VIEJO de esa misma fila (no con el de una
fila vecina) — es decir, esos casos puntuales fueron un fallo genuino de
guardado (SAVE no persistió nada), no una escritura en la fila equivocada.
El fix de arriba (re-buscar la fila por texto en vez de por índice) sigue
siendo válido y necesario — es una causa real y contemplada por el diseño
de la grilla (CDK virtual scroll) — pero **no explica todos los casos de
este error**; queda abierto por qué el SAVE a veces no persiste el valor
aunque el click se ejecute sin problemas aparentes. Pendiente seguir
investigando con capturas de pantalla de una corrida real.

## Fix — falso error "No encontré el botón SAVE" en períodos nuevos (Fase 2)

Detectado auditando a mano un caso de este error: el valor SÍ había quedado
guardado en Tourplan, EXACTO al valor objetivo (241.56), pero el script lo
reportaba como error igual.

Causa: cuando el período no existía todavía, se crea por COPY DATE RANGE
(duplica el período más reciente y le ajusta las fechas). En ese caso el
valor que se escribe en la grilla puede confirmarse solo (sin que aparezca
un botón SAVE habilitado para clickear) — probablemente porque, al ser un
período recién creado, el guardado ocurre en cuanto se pierde el foco del
campo, sin un botón SAVE aparte. El script solo sabía comparar el valor ya
escrito contra el que HABÍA ANTES (`v_viejos`, previo a editar) para decidir
si "no hacía falta guardar"; como el objetivo era distinto del valor
previo, no encontraba el SAVE y tiraba error, sin considerar que el valor
ya podía estar aplicado.

Fix: si no aparece un botón SAVE clickeable y el valor previo SÍ cambió
(no es el caso de "nada que guardar"), ahora se vuelve a leer la grilla en
ese momento y se compara lo que hay AHORA contra el objetivo (con una
tolerancia de 0.30 por las dudas, aunque en el caso auditado el valor
coincidía exacto). Si ya coincide, se toma como guardado correctamente en
vez de marcar error.

**Portado a Valorización desde Servicio Madre:** `valorizacion_desde_madre.py`
tiene una copia casi idéntica de esta misma lógica de escritura de rates
(script vendorizado aparte, no comparte código con el de Componente) — los
2 fixes de esta sección y la anterior se portaron ahí también. Sin correr
todavía contra Tourplan real desde ese script.

## Fix — Valorización desde Servicio Madre: 4 errores de una corrida real (2026-08-25)

Detectados en una corrida real (Excel de 5 períodos × 5 componentes) —
cuatro problemas distintos, cada uno con causa y fix propios:

1. **`abrir_pcm()` podía saltear la fila del PCM en listas grandes de Used
   In** — confirmado con un componente (PSUS) que tiene "cientos" de
   filas en Used In: el escaneo inicial que arma la lista de PCMs
   candidatos usa un scroll de 150px por paso (con acumulación
   progresiva, pensado para listas grandes — ver skill
   `recorriendo-grillas-virtuales-de-tourplan`), pero `abrir_pcm()`
   (que vuelve a buscar/clickear ese mismo PCM más tarde) usaba 1500px
   por paso con un chequeo puntual en cada posición — 10x más grande.
   En un Used In con cientos de filas, ese salto podía dejar la fila del
   PCM en el "hueco" entre lo renderizado en una posición y la
   siguiente (virtual scroll: Angular solo mantiene en el DOM lo cercano
   al viewport), y el PCM nunca se encontraba pese a existir — mismo PCM
   que sí se había detectado bien en el escaneo inicial. Fix: `step` bajado
   a 150 (igual que el escaneo inicial) y `max_steps` subido en la misma
   proporción para cubrir la misma distancia total. Portado también a
   `tourplan_valorizacion_pkg_v3.py`, que tenía el mismo código.
2. **El campo TO del diálogo COPY DATE RANGE podía quedar vacío sin que
   el script lo notara** — confirmado con captura de pantalla: el campo
   TRAVEL DATE FROM se completaba bien pero TRAVEL DATE TO quedaba vacío
   (borde rojo, campo requerido) pese a que `_fill_dialog_date()` no
   tiraba ninguna excepción — el setter sintético (`set_val` + evento
   `change`) a veces no "pega" en ese campo puntual, y el código seguía
   de largo clickeando OK con el dialog a medio llenar. Fix:
   `_fill_dialog_date()` ahora relee el valor tras escribirlo; si no
   coincide, reintenta con foco + Ctrl+A + Delete + `send_keys` (eventos
   de teclado nativos); si sigue sin coincidir, lanza una excepción para
   que el caller haga el fallback a INSERT que ya existía, en vez de
   confirmar un dialog incompleto. Portado a los 3 scripts que comparten
   este código (`valorizacion_desde_madre.py`,
   `tourplan_valorizacion_pkg_v3.py`, `valorizacion_desde_excel.py`).
3. **El fallback a INSERT terminaba leyendo la grilla de rates en vez de
   un dialog nuevo — causa confirmada con captura de pantalla.** Tras
   cancelar el dialog COPY DATE RANGE por el bug del punto 2, el código
   buscaba un botón `tp-button.cancel > button` o `tp-button.close >
   button` para cerrarlo — pero el botón real del dialog (confirmado por
   captura) es **"EXIT"**, no "cancel"/"close": el click no encontraba
   nada y el dialog quedaba abierto en silencio. Con ese dialog viejo
   todavía abierto, el siguiente click en "INSERT" terminaba siendo en
   realidad el botón **"INSERT RATE SET" del propio dialog** (agregar una
   fila de rate DENTRO del período actual, no un período nuevo) — que
   abre otro dialog ("Insert Rate") apilado encima, con sus propios
   TRAVEL DATE FROM/TO (TO también vacío, mismo bug del punto 2). Por eso
   `_insertar_rate_period()` terminaba leyendo los inputs numéricos de la
   grilla de rates de atrás en vez de fechas. Fix: el click para cerrar
   ahora también prueba `tp-button.exit > button` y, si no encuentra nada
   por clase, busca por texto ("EXIT"/"CANCEL"/"CLOSE"). Además, si tras
   varios reintentos el dialog sigue sin cerrarse, la fila se aborta con
   un error explícito en vez de seguir a un INSERT que iba a terminar mal
   igual. Portado a los mismos 3 scripts del punto 2.
4. **`UnboundLocalError` en la verificación post-SAVE** — la variable
   `periodo_encontrado` solo se asignaba en un camino del código; si COPY
   DATE RANGE dejaba al driver directo en el detalle del período nuevo y
   la verificación post-SAVE no encontraba el período al recargar RATES,
   el código intentaba usar una variable nunca asignada. Fix: default
   asignado antes del `if`. Portado a `tourplan_valorizacion_pkg_v3.py`
   también (mismo bug); `valorizacion_desde_excel.py` ya lo tenía bien.

**Sin correr todavía contra Tourplan real** — los 4 fixes se armaron a
partir del Excel de resultados y capturas de pantalla de la corrida real
que los reveló, pero ninguno se validó todavía en una corrida nueva.
Probar con un lote chico (ideal: repetir el mismo Excel que falló) antes
de confiar en esto para lotes grandes.

## Optimización — Ordenar Used In por Date para evitar el scroll completo

En la grilla **Used In** (donde se busca el PCM Package Header de cada
servicio), la usuaria observó que el scroll para encontrarlo se veía
"raro": escrollea rápido hasta el final, se pasa de largo el Package
Header, y recién en un segundo pase más lento lo encuentra — comportamiento
esperable de un virtual scroll de Angular (solo lo cercano al viewport
está en el DOM en cada momento; ver skill
`recorriendo-grillas-virtuales-de-tourplan`), no un bug.

La usuaria propuso una mejora: los PCM Package Header **siempre traen la
columna "Date" en blanco** (confirmado con captura real de un Used In
ordenado así, con el Package Header primero). Clickeando ese encabezado
para ordenar ascendente, esas filas quedan al principio de la lista y
casi no hace falta escrollear para encontrarlas. Se implementó en **dos
puntos** de `valorizacion_desde_madre.py` y `tourplan_valorizacion_pkg_v3.py`,
ambos como **paso 0, best-effort y verificado** — si falla en cualquier
punto, cae exactamente al procedimiento anterior sin cambios:

1. **`abrir_pcm()`** — antes de escrollear para clickear la fila del PCM
   ya identificado, ordena por Date (`_ordenar_used_in_por_date()`). Si el
   Package Header quedó primero, el scroll de siempre lo encuentra casi
   sin moverse; si el ordenamiento no se pudo verificar (no encontró el
   encabezado, o la primera fila visible no cambió tras el click), el
   scroll sigue funcionando exactamente igual que antes.
2. **`leer_pcm_list()` / `leer_pcm_list_package_header()`** (la lectura
   inicial de TODOS los PCMs candidatos del Used In) — mismo paso 0:
   ordena por Date y revisa las filas iniciales con Date en blanco. Si
   alguna cumple los filtros ya existentes de PCM válido (tipo Package
   Header/Package/PCM Service, nombre "PKG-", no eliminado, rango de
   fechas ≤365 días — la misma función de filtro que usa el camino lento,
   compartida entre ambos para que den siempre el mismo resultado), la usa
   como resultado y **se salta por completo el escaneo con scroll de
   miles de filas**. Si el ordenamiento falla, no hay ninguna fila con
   Date en blanco, o las que hay no pasan los filtros (por ejemplo si
   otro tipo de fila también tiene la fecha vacía), cae sin cambios al
   procedimiento de siempre: ordenar por Type + escrollear toda la lista.

**Sin correr todavía contra Tourplan real** — la lógica se verificó con
simulaciones en Python contra distintos escenarios (fila válida en blanco,
fila en blanco que no es un PCM válido, columna "Date" no reconocida), pero
falta confirmar en una corrida real que el ordenamiento se comporta igual
que en la captura de pantalla que motivó el cambio.

**Ajuste posterior — SOLO en `valorizacion_desde_madre.py`:** la usuaria
observó en una corrida real que, aunque el orden por Date sí dejaba el
Package Header como primera fila, `abrir_pcm()` igual escrolleaba para
"buscarlo" — trabajo de más, porque este script únicamente procesa el PCM
tipo Package Header (`leer_pcm_list_package_header()` ya filtra por ese
tipo, y el llamador siempre toma el primero de la lista). Con el orden por
Date ya confirmado, no hace falta escanear nada: se clickea directo la
primera fila (scroll 0). Si el ordenamiento no se pudo verificar, o pese a
estar confirmado la primera fila no fuera la esperada (caso no esperado),
cae al escaneo con scroll de siempre, sin cambios. Deliberadamente **no**
se tocó `tourplan_valorizacion_pkg_v3.py` — a diferencia de Valorización
desde Servicio Madre, ese script también puede procesar PCMs tipo "PCM
Service" (no solo Package Header), así que no vale la misma garantía de
"siempre el primero". Sin correr todavía contra Tourplan real.

## Fix — `ir_a_used_in()` no encontraba "Used In" tras navegar ahí una vez (Valorización desde Servicio Madre)

Detectado en una corrida real (mismo caso que el diagnóstico agregado más
abajo, PSUS): la Fase 1 llama a `ir_a_used_in()` **dos veces** para el
mismo producto — una desde `_buscar_servicio_fase1()` (para armar la
lista de PCMs, solo lectura) y otra desde `_procesar_pcm_grupo_fase1()`
(justo antes de `abrir_pcm()`, para clickear y abrir el PCM ya
identificado). **Esta doble visita no la introdujo la optimización de
Fase 1 de esta sesión** — ya existía en el código original
(`procesar_fila_fase1`, con el comentario `# Re-navegar a USED IN antes
de abrir PCM`); el agrupamiento solo la repartió en dos funciones sin
cambiar la secuencia. Es una limpieza defensiva ya usada también en
`tourplan_valorizacion_pkg_v3.py` (comentario ahí: "el orden del listado
puede cambiar tras Change Base Date en PCMs anteriores del mismo grupo")
para garantizar que `abrir_pcm()` arranque siempre desde una vista
fresca. La segunda llamada falló con un timeout sin mensaje útil ("Re-nav
USED IN falló") al buscar "Used In" dentro del submenú ya expandido de
UTILITIES — un bug latente en ese patrón, no algo nuevo que haya
aparecido con el agrupamiento; simplemente no se había visto disparado en
una corrida real hasta ahora.

Causa confirmada comparando los `Menu items` que imprime `hamburger()` en
cada llamado: el menú de Tourplan es **contextual**. Desde la vista neutra
del producto (recién abierto), "USED IN" aparece anidado dentro de
UTILITIES — como siempre asumió el código. Pero una vez que ya se navegó a
Used In antes en la misma sesión de producto, Tourplan lo muestra como
**ítem de primer nivel directo**, sin UTILITIES en el medio — por eso
buscarlo adentro del submenú expandido de UTILITIES nunca lo encontraba.

Fix: `ir_a_used_in()` ahora prueba primero si "Used In" ya está visible
como ítem de primer nivel (sin expandir nada) y lo clickea directo; solo
si no lo encuentra ahí, cae al camino de siempre (UTILITIES → Used In
anidado), sin cambios. La detección de visibilidad usa el mismo criterio
ya usado en otros lados del script (`offsetWidth`/`offsetHeight`/
`getClientRects().length`), para no clickear por accidente un ítem "Used
In" que exista en el DOM pero esté oculto dentro de un submenú colapsado.

Importante: esto no rompió nada — la red de seguridad de la optimización
de Date-sort en `abrir_pcm()` (ver sección de arriba) detectó que el
orden por Date no daba el resultado esperado y cayó al escaneo con scroll
completo, encontrando igual el PCM correcto (más lento, pero sin abrir
nada equivocado). Este fix ataca la causa real para que ni siquiera haga
falta ese fallback. **Sin correr todavía contra Tourplan real.**

## Simplificación — fusionar identificar y abrir el PCM en una sola visita a Used In (Valorización desde Servicio Madre)

La usuaria preguntó por qué hacía falta armar la lista de PCMs primero
(`leer_pcm_list_package_header()`, dentro de `_buscar_servicio_fase1()`) y
recién después reabrir Used In para abrir el PCM (`abrir_pcm()`, dentro de
`_procesar_pcm_grupo_fase1()`), si cada servicio madre está linkeado a un
único PCM Package Header. Respuesta corta: no hacía falta — esa segunda
visita a Used In era, de hecho, la causa raíz de los dos bugs anteriores
(el menú contextual y el segundo click que invertía el orden). En vez de
solo evitarlos defensivamente, se eliminó la visita redundante:

`leer_pcm_list_package_header()` (llamada desde `_buscar_servicio_fase1()`)
ya deja la página parada en Used In con scroll en 0 al terminar — tanto en
su pasada rápida (sin scroll, orden por Date) como en su fallback con
scroll completo (que también resetea el scroll a 0 antes de retornar). Por
eso `_procesar_pcm_grupo_fase1()` ya no vuelve a llamar a `ir_a_used_in()`
antes de `abrir_pcm()` — lo llama directo, sobre la misma vista en la que
ya está parada. `abrir_pcm()` reordena por Date igual que antes (con el
fix que no clickea de nuevo si ya está ordenado — esta vez debería
detectar que sí lo está y no clickear nada) y busca la fila ahí mismo.

Como red de seguridad (por si algo más cambiara el estado de la página
entre medio, o el intento directo fallara por cualquier otro motivo), si
`abrir_pcm()` falla sin haber re-navegado, se reintenta **una vez** con la
re-navegación explícita de siempre (`ir_a_used_in()` + `abrir_pcm()`) antes
de recién ahí marcar error — mismo comportamiento de fallback que existía
antes, solo que ahora es el segundo intento en vez del único.

**Sin correr todavía contra Tourplan real.**

## Fix — el segundo click en "Date" invertía el orden de Used In en vez de mantenerlo (Valorización desde Servicio Madre)

Detectado en una corrida real, viendo la ejecución en vivo: después de
confirmar que el fix de `ir_a_used_in()` (menú contextual) funcionaba
bien — la usuaria vio en el log `'Used In' encontrado como ítem de
primer nivel — click directo` —, igual volvió a pasar que `abrir_pcm()`
no encontraba el Package Header en la primera fila y caía al scroll
completo. Esta vez el fallback SÍ encontró `PKG-BUE-EX-1STA01-PSUS` (en
la posición 22800) — lo cual confirma que la tabla era la correcta, el
problema era solo el **orden**.

Causa: `_ordenar_used_in_por_date()` se llama **dos veces** para el mismo
producto — una en `leer_pcm_list_package_header()` (arma la lista) y otra
en `abrir_pcm()` (para reabrir el PCM). Si la grilla Used In ya estaba
ordenada ascendente por Date desde el primer llamado (el componente
Angular puede retener su estado entre navegaciones en vez de resetearlo),
el segundo click sobre el mismo encabezado la **invierte** a descendente
— mismo comportamiento ya conocido para el header Type en otros scripts
de esta app (1er click ordena, 2do invierte). Con la grilla invertida, el
Package Header (Date en blanco) deja de estar primero — apareció en su
lugar un PCM totalmente distinto con la fecha más lejana
(`PCM Service EAGP252846 - Riviera 2028`, 21/Mar/2028).

Fix: `_ordenar_used_in_por_date()` ahora verifica el estado de la
primera fila **antes** de clickear — si la columna Date ya está en
blanco ahí (ya ordenada como se quiere), no vuelve a clickear el
encabezado. Solo clickea cuando hace falta. Beneficia a los dos llamados
(la lista y la reapertura del PCM) sin necesidad de distinguir cuál es
cuál. **Sin confirmar todavía en una corrida real con este fix.**

## Fix confirmado — "No encontré campo FROM/TO" en COPY DATE RANGE: el diálogo correcto es el ÚLTIMO `<tp-dialog>`, no el segundo (Valorización desde Servicio Madre)

Detectado en una corrida real: al valorizar varios períodos consecutivos
del mismo servicio madre (PSUS), el primer período se armó y valorizó
bien, pero desde el segundo en adelante `_copiar_ultimo_period()` fallaba
con `⚠ No encontré campo FROM en dialog COPY`, después no lograba cerrar
el diálogo (`El dialog COPY DATE RANGE no se cerró tras cancelar`), y en
períodos posteriores aparecía además `No encontré el botón SAVE de la
grilla de rates`.

La usuaria compartió los 4 dumps de HTML reales guardados por el
diagnóstico agregado en la sesión anterior, que confirmaron la causa
exacta (no una hipótesis):

- El diálogo **"Copy Rate" no es un componente propio** — es el mismo
  `tp-insert-rate` / `#insertrateview` que ya se usaba para el fallback a
  INSERT; Tourplan le pone un título distinto ("Copy Rate" vs "Insert
  Rate") según el contexto, pero la estructura interna (Price Code +
  Travel Date From/To en `li:nth-of-type(2)`/`(3)` + Based On) es idéntica.
- **Angular nunca saca del DOM un `<tp-dialog>` ya cerrado** — solo deja
  de mostrarlo. Cada período procesado (cada vez que este flujo cae al
  fallback a INSERT) deja un `<tp-dialog>` más acumulado: los 4 dumps
  mostraron 3, 2, 6 y hasta 8 diálogos en la misma corrida, nunca 2. El
  selector `tp-dialog:nth-of-type(2)` agarraba siempre un diálogo viejo y
  abandonado (normalmente uno de "Product Costs", con una estructura de
  formulario totalmente distinta) — por eso nunca encontraba los campos
  FROM/TO ahí. El diálogo recién abierto es siempre el **último** del
  documento.
- Esto además rompía el intento de cancelar: la búsqueda del botón EXIT
  usaba `document.querySelector(...)` **sin acotar a ningún diálogo**, así
  que encontraba y clickeaba el botón Exit del diálogo viejo (que
  visualmente parecía válido — sin `display:none` cerca), sin ningún
  efecto sobre el diálogo real todavía abierto.
- El error posterior de "No encontré el botón SAVE" en períodos más
  avanzados es, con alta probabilidad, una consecuencia en cascada de lo
  anterior: como el período nunca se creaba bien, no había nada real que
  guardar en la grilla.

**Fix:** se reemplazó `tp-dialog:nth-of-type(2)` por `tp-dialog:last-of-type`
en los 3 puntos que lo usaban (`_fill_dialog_date()` para FROM y TO, y las
dos verificaciones de `_cancelar_dialog_y_hacer_insert()`), y se acotó la
búsqueda del botón EXIT al último diálogo en vez de buscarlo en todo el
documento. El diagnóstico agregado la sesión anterior (`dump()` + conteo
de `<tp-dialog>`) queda como red de verificación para la próxima corrida.
**Sin volver a correr todavía contra Tourplan real con este fix.**

## Fix — Falso positivo en `_fill_dialog_date()`: campo TO reformateado se tomaba como falla

Detectado en una corrida real de Valorización desde Excel (fila AEEZMV,
PC=TR): el campo **FROM** del dialog COPY DATE RANGE se completaba y
verificaba bien (`01/11/26`), pero el campo **TO** disparaba la alerta de
reintento y terminaba abortando el COPY pese a estar bien completado:

```
⚠ TO quedó '31/Mar/2027' tras set_val — reintentando con send_keys
TO: 31/Mar/2027
⚠ No encontré campo TO en dialog COPY: TO quedó '31/Mar/2027' en vez de '31/03/27' tras 2 intentos → cancela y hace INSERT
```

Causa: `_fill_dialog_date()` (fix anterior, ver sección "4 errores de una
corrida real") comparaba el valor leído del campo contra el string exacto
que se había escrito (`"31/03/27"`). El campo TO, a diferencia de FROM,
re-formatea el valor apenas lo acepta como fecha válida (a `"31/Mar/2027"`,
formato distinto pero **la misma fecha**) — la comparación de texto
literal tomaba ese reformateo como si el campo hubiera quedado vacío/mal,
cuando en realidad COPY DATE RANGE había funcionado. Eso disparaba el
fallback a INSERT innecesariamente, que además terminó creando el período
con la fecha mal (`01/Nov/2026–01/Nov/2026` en vez del rango pedido del
Excel) por el problema ya conocido de asignación de Price Code en el
diálogo INSERT.

Fix: la comparación ahora es **por fecha**, no por texto — si el string
exacto no coincide, se parsean ambos valores con `parsear_fecha()` (que ya
entiende tanto `DD/MM/YY` como `DD/Mon/YYYY`) y se los compara como fechas.
Solo se reintenta/aborta si de verdad no son la misma fecha (campo vacío,
fecha incorrecta, texto no parseable). Portado a los 3 scripts que
comparten este código: `valorizacion_desde_excel.py`,
`valorizacion_desde_madre.py`, `tourplan_valorizacion_pkg_v3.py`.

**Sin correr todavía contra Tourplan real** — el fix se armó a partir del
log de la corrida real que lo reveló y se verificó con una simulación en
Python del caso real (`31/Mar/2027` vs `31/03/27`) y de los casos de falla
genuina (campo vacío, fecha incorrecta), pero falta confirmar que el COPY
DATE RANGE completo funciona en una corrida nueva.

## Fix — Copy y Linkeo PCM buscaba el producto equivocado al linkear

Detectado en una corrida real: **todas** las filas terminaban en error, la
mayoría con "No encontré resultados para .../CODIGO_NUEVO en PRODUCT
SEARCH". `linkear_pcm_a_servicio()` (el paso final, que linkea el PCM ya
copiado/renombrado a un producto) buscaba en Tourplan usando **CODIGO
NUEVO** — pero CODIGO NUEVO no es un product code que exista como tal en
Tourplan: es solo el identificador que se usa para renombrar el PCM copiado
(va en su nombre, ver `calcular_nombre_pcm_nuevo`). El producto al que hay
que linkear el PCM copiado es el servicio madre **original** — el mismo
que se buscó al principio por su **SERVICE CODE** (columna "SERVICE CODE"
del Excel, no "CODIGO NUEVO"). Fix: `procesar_fila()` ahora le pasa el
SERVICE CODE original a `linkear_pcm_a_servicio()` en vez de CODIGO NUEVO.

De paso, se corrigió un problema de auditoría relacionado: si el PCM ya se
había copiado/renombrado en Tourplan con éxito pero fallaba el paso de
linkeo (como en esta misma corrida), el Excel de salida no registraba el
nombre del PCM ya creado — quedaba en error sin ese dato, dificultando la
corrección manual. Ahora `escribir_resultado()` en el camino de error
también incluye el PCM original/copiado si ya se conocían para cuando
ocurrió la falla.

**Sin correr todavía contra Tourplan real** — la próxima corrida con este
fix aplicado va a confirmar si el linkeo ahora encuentra el producto
correctamente.

## Seguridad — qué tener en cuenta

- **La app solo escucha en la propia PC** (`--server.address=localhost` en
  `run_app.bat`). Sin esto, Streamlit escucha en todas las interfaces de red
  por default — cualquiera en la misma red de oficina podría haber abierto
  la app en su navegador, ya que no tiene login propio.
- **Credenciales:** viajan solo en el entorno del subproceso, nunca se
  escriben a disco ni se loguean. Esto es mucho más seguro que hardcodearlas
  en un archivo, pero no es una bóveda: alguien con herramientas de
  inspección de procesos y permisos suficientes en la MISMA PC, durante la
  corrida, técnicamente podría leerlas. No son recuperables de forma remota.
- **Sin login propio:** cualquiera con acceso a la sesión de Windows de esa
  PC puede abrir la app y correr los scripts contra Tourplan. Igual que
  pasaría si alguien dejara una notebook de Colab abierta y logueada.
- **Archivos temporales:** el Excel subido y las capturas de pantalla de
  cada corrida quedan en una carpeta temporal del sistema — no se borran
  solos. Con el tiempo se acumulan datos de negocio ahí; conviene una
  limpieza manual periódica si eso importa.
- **`--no-sandbox` en Chrome:** heredado de los scripts originales de Colab
  (necesario ahí porque corre como root en un contenedor). En una PC Windows
  normal no hace falta y reduce una capa de aislamiento de Chrome. Pendiente
  de sacar — no se hizo todavía para no tener que re-validar los scripts que
  ya andan.
- **Nada de esto es exclusivo de esta app:** el riesgo de fondo (automatizar
  cambios reales en producción con Selenium) ya existía con los scripts de
  Colab — la app no lo aumenta, solo lo hace más visible (con avisos de modo
  y de producción en la UI).

## Próximos pasos

- [x] Validar Copy Products contra Tourplan real (Windows).
- [x] Validar Rename Products contra Tourplan real.
- [x] Validar Flag as Deleted contra Tourplan real.
- [x] Validar Copy PCM Linkeo contra Tourplan real.
- [x] Validar Valorización desde Componente contra Tourplan real.
- [x] Validar Valorización desde Servicio Madre contra Tourplan real.
- [x] Los 6 scripts en alcance ya están cargados y validados en la app.
- [x] Validar Flag as Deleted v1.1 contra Tourplan real — confirmado
      funcionando (2026-07-24).
- [ ] Validar Copy Products (versión nueva, reemplaza a Copy Products y
      Rename Products) contra Tourplan real (Windows) — vendorizado en
      `scripts/copy_products/`, probado antes en Colab pero no desde esta
      app.
- [ ] Validar el encadenado de copias del mismo origen en Copy Products
      (ver sección "Optimización de navegación (Copy Products)" arriba) —
      lógica de agrupamiento/corte de cadena probada de forma aislada,
      pero el flujo de Selenium encadenado todavía no se corrió contra
      Tourplan real. Probar con un lote chico antes de un lote grande.
- [ ] Validar Modificar Description y Comment contra Tourplan real
      (Windows) — vendorizado en `scripts/modificar_description_comment/`.
- [ ] Validar Valorización desde Excel contra Tourplan real (Windows) —
      vendorizado en `scripts/valorizacion_excel/`.
- [ ] Re-validar Valorización desde Componente contra Tourplan real con un
      Excel de 2+ filas del mismo componente/servicio madre — se cambió el
      flujo de navegación para no repetir búsquedas Y para abrir cada PCM
      una sola vez por grupo en vez de una vez por período (ver sección
      "Optimización de navegación" arriba, segundo nivel).
- [x] Portar la optimización de navegación de Fase 2 a Valorización desde
      Servicio Madre (agrupar por servicio madre para no repetir la
      búsqueda por cada período) — hecho, sin correr todavía contra
      Tourplan real (ver sección "Optimización de navegación —
      Valorización desde Servicio Madre" arriba).
- [x] Portar el nivel de optimización de Fase 1 (agrupar por servicio
      madre en la lectura de costos) a Valorización desde Servicio Madre —
      hecho con un agrupamiento en dos pasos (resolver filas a códigos,
      después agrupar) por la búsqueda masiva propia de la Fase 1 de este
      script (ver sección "Optimización de navegación — Fase 1" arriba).
- [ ] Validar contra Tourplan real la optimización de Fase 1 recién
      agregada — verificada por simulación en Python (resolución +
      agrupamiento + atribución de resultados por fila), no corrida
      todavía. Probar con un Excel de Fase 1 que tenga 2+ filas (mezclando
      `SERVICE_CODE` directo y búsqueda masiva si es posible) apuntando al
      mismo servicio madre.
- [ ] Correr Valorización desde Servicio Madre contra Tourplan real tras
      portarle los 2 fixes de escritura de rates (re-búsqueda de fila por
      texto + falso error de SAVE en períodos nuevos) y la optimización de
      navegación de Fase 2 — todo portado pero sin probar todavía desde
      este script.
- [ ] Seguir investigando el error "N/24 rangos difieren tras recargar"
      (Fase 2) — sigue sin confirmarse si es un fallo real de guardado o
      un tercer falso positivo (ver sección de fixes arriba).
- [ ] Validar contra Tourplan real la optimización de ordenar Used In por
      Date (paso 0 en `abrir_pcm()` y en `leer_pcm_list()`/
      `leer_pcm_list_package_header()`, ver sección "Optimización —
      Ordenar Used In por Date" arriba) — verificada por simulación en
      Python, no corrida todavía.
- [ ] Validar contra Tourplan real el fix de comparación por fecha (no por
      texto) en `_fill_dialog_date()` (ver sección "Falso positivo en
      `_fill_dialog_date()`" arriba) — se confirmó que el bug anterior
      quedó resuelto en la simulación con los valores reales de la corrida
      que falló (AEEZMV), pero falta correr un lote nuevo.
- [ ] Validar Notas SRV contra Tourplan real desde esta app (Windows) —
      vendorizado en `scripts/notas_srv/`; en el repo origen el INSERT ya
      está confirmado en una corrida real de 305 filas, pero la lectura de
      notas existentes (usada para detectar "ya existe") se reescribió
      recién en la última sesión (v1.12) y todavía no se re-probó, y el
      flujo de EDITAR una nota existente (`EDICION="SI"`) nunca se probó.
- [ ] Validar Exportar Notas contra Tourplan real desde esta app (Windows)
      — vendorizado en `scripts/exportar_notas/`; la lectura de notas
      existentes se reescribió recién (v1.3) tras un diagnóstico por dump
      real de DOM y todavía no se re-probó contra Tourplan.
- [ ] Validar contra Tourplan real los 15 códigos Plain Text sumados a
      `CODIGOS_NOTA_VALIDOS` en esta sesión (`DRT`, `LWE`/`LWF`/`LWG`/`LWI`/
      `LWS`, `PCR`, `REM`, `REO`, `SC2`, `TAL`/`TES`/`TFR`/`TIN`/`TIT`) —
      confirmados como Product/Plain Text por captura real del catálogo,
      pero ninguno se probó todavía en una corrida real de Notas SRV
      (solo los 5 "Nota SRV" originales y los 5 "Descriptivo" tienen
      corridas reales o casi).
- [ ] Si en algún momento hace falta soportar códigos de nota **HTML**
      (`OBS`, `NAP`, `SDV`, los `HY*`, los `TB*`, etc. — mismo catálogo,
      quedaron deliberadamente afuera de `CODIGOS_NOTA_VALIDOS`), primero
      confirmar si el mecanismo de escritura actual (bypass directo del
      `<body contenteditable>` del iframe) funciona igual contra un editor
      HTML, o si hace falta un camino distinto — no asumido, no probado.
- [x] Validar contra Tourplan real el encadenamiento por producto de
      Notas SRV — confirmado funcionando por la usuaria (2026-08-29).
- [ ] Validar contra Tourplan real la re-sincronización tras una nota en
      `ERROR` dentro de un grupo encadenado de **Notas SRV** (ver
      "Re-sincronización..." en la sección "Notas SRV / Exportar Notas"
      arriba) — el agrupamiento en sí ya se confirmó (ver ítem arriba),
      pero la corrida real no reportó una nota fallida en el medio de un
      grupo, así que este camino sigue apoyado solo en simulaciones en
      Python. Probar con un Excel que tenga 2+ notas del mismo producto,
      idealmente forzando que una falle a propósito (ej. un `Codigo_Nota`
      inválido en el medio) para confirmar que las siguientes del mismo
      producto igual se procesan.
- [ ] Validar contra Tourplan real el encadenamiento por producto de
      **Exportar Notas** con el fix del botón "Exit" (`cerrar_nota()` — ver
      sección "Notas SRV / Exportar Notas" arriba). La primera corrida
      real (2026-08-29) mostró el bug que el fix corrige: sin cerrar la
      nota, se abrían varias notas juntas al pasar a la siguiente del
      mismo producto. El fix en sí (click en "Exit" tras cada lectura)
      todavía no se re-probó — verificado solo con una simulación en
      Python de la lógica de control (7 escenarios). Probar con un Excel
      que tenga 2+ notas del mismo producto y confirmar en las capturas
      que cada nota se cierra antes de abrir la siguiente.
- [ ] Validar contra Tourplan real el fix del segundo click en "Date"
      (ver sección "Fix — el segundo click en 'Date' invertía el orden..."
      arriba) — en la corrida real que reveló esto, `abrir_pcm()` seguía
      escrolleando de más pese al fix de `ir_a_used_in()` (que sí
      funcionó), porque el segundo click sobre "Date" invertía el orden
      en vez de mantenerlo. Confirmar que ahora `abrir_pcm()` abre el PCM
      sin escrollear cuando la grilla ya está bien ordenada desde la
      lectura inicial.
- [ ] Validar contra Tourplan real la fusión de identificar + abrir el
      PCM en Fase 1 (ver sección "Simplificación — fusionar identificar y
      abrir el PCM..." arriba) — `_procesar_pcm_grupo_fase1()` ya no
      vuelve a navegar a Used In antes de `abrir_pcm()`, confía en que
      `_buscar_servicio_fase1()` la dejó ya parada ahí. Confirmar que el
      camino directo funciona (sin el mensaje de reintento con
      re-navegación) y que, si hiciera falta, el reintento de respaldo
      también funciona.
- [ ] Validar contra Tourplan real el fix de `tp-dialog:last-of-type` en
      COPY DATE RANGE (ver sección "Fix confirmado — 'No encontré campo
      FROM/TO'..." arriba) — la causa se confirmó con dumps de HTML reales
      de una corrida (PSUS), pero el fix en sí todavía no se corrió. Repetir
      con un Excel de 2+ períodos del mismo servicio madre, idealmente
      alguno que necesite pasar por el fallback a INSERT más de una vez,
      para confirmar que ya no se acumulan diálogos que rompan el
      siguiente período.
- [ ] Confirmar si el error "No encontré el botón SAVE de la grilla de
      rates" (visto en períodos avanzados de la misma corrida de PSUS)
      desaparece solo al corregir el bug de arriba, o si es un problema
      independiente — no se tocó ese código todavía porque no había
      evidencia de que fuera un bug propio y no una consecuencia en
      cascada.
- [ ] Validar contra Tourplan real el fix de `ir_a_used_in()` (menú
      contextual — "Used In" de primer nivel vs anidado en UTILITIES, ver
      sección "Fix — `ir_a_used_in()` no encontraba..." arriba) — detectado
      y diagnosticado con el log/capturas de una corrida real, pero el fix
      en sí todavía no se corrió.
- [x] Validar Vigencias contra Tourplan real desde esta app (Windows) —
      vendorizado en `scripts/extraccion_vigencias/`, categoría
      "Relevamiento". Las dos fallas de las primeras corridas
      (`AttributeError` de `importlib.util`, y Chrome en modo `--headless`
      sin abrir ventana — ver sección "Vigencias" arriba) ya se
      corrigieron, y la usuaria confirmó que ahora funciona (2026-08-29).
- [ ] Revisar "FECHA RECALCULATE PCM" (`valorizacion_desde_madre.py`) /
      "RATE FROM D" (`tourplan_valorizacion_pkg_v3.py`) cuando dos o más
      filas del mismo servicio madre piden la MISMA fecha de recalculate:
      hoy no se detecta esa coincidencia para reusar el recálculo ya
      hecho. `cambiar_fecha_pcm()` vuelve a correr el flujo completo
      (hamburger → ITINERARY → CHANGE BASE DATE → diálogo de recálculo)
      en cada fila, y como el "avanzar un día más" (para evitar el no-op
      de Tourplan cuando la fecha no cambia) sigue activo, la segunda
      fila y las siguientes terminan recalculadas en el día hábil
      SIGUIENTE al pedido, no en la fecha exacta cargada. Comportamiento
      deseado (pendiente de implementar): si dos o más filas del mismo
      grupo piden la misma fecha de recalculate, directamente reusar el
      resultado ya generado por el primer Change Base Date del grupo
      (misma `nueva_fecha`/costos leídos) para esas filas — sin volver a
      abrir el diálogo de recálculo NI forzar el avance al día siguiente,
      que solo debe aplicar cuando la fecha pedida es distinta de la ya
      aplicada.
- [ ] Validar contra Tourplan real el corte temprano por Non Accommodation
      en Valorización desde Servicio Madre (`_buscar_servicio_fase1()` /
      `leer_service_category()`, agregado en v2.8) — la grilla de
      resultados de Product Search no trae una columna Category poblada
      (confirmado por la usuaria con captura real), así que ahora el
      script lee el dropdown "Service Category" directo en la pantalla
      donde ya queda el producto al seleccionarlo desde la grilla de
      resultados (confirmado por la usuaria que no hace falta navegar a
      ningún lado), ANTES de ir a Used In — si da Non Accommodation,
      corta ahí sin escanear la grilla virtual de Used In (hasta ~1900
      filas). El selector del dropdown (basado en el HTML del campo que
      pasó la usuaria: `label[for="serviceCategory"]`) todavía no se
      probó contra Tourplan real. Probar con un supplier que tenga
      servicios Package y Non Accommodation mezclados (confirmado que
      existen) y verificar que los Non Accommodation se saltean sin
      pasar por Used In, y que los Package siguen procesándose igual que
      antes.
- [ ] Validar contra Tourplan real el fix del checkbox "Update Exchange
      Rates" en el diálogo de recálculo del PCM (`cambiar_fecha_pcm()`,
      `valorizacion_desde_madre.py` v2.11 / `tourplan_valorizacion_pkg_v3.py`
      v3.45) — la usuaria confirmó en corrida real que el checkbox
      quedaba destildado pese al click. Causa confirmada con una
      grabación nueva de Chrome DevTools: la grabación original clickeaba
      el checkbox DOS veces (su `<label>` de texto y el `<label>` de su
      propio widget — ambos hijos del mismo `#exchange-rate-panel`, mismo
      checkbox), tildándolo y destildándolo de nuevo. Ahora se clickea
      `#updateexchangerates` (el `<input>`) UNA sola vez, con
      verificación y reintentos. El fix en sí todavía no se re-probó —
      confirmar en una corrida real que "Update Exchange Rates" termina
      tildado al cerrar el diálogo de recálculo.
- [ ] Sacar `--no-sandbox` de Chrome en los 6 scripts (innecesario en
      Windows, pospuesto para no forzar una re-validación — ver sección
      Seguridad).
- [ ] Empaquetar como ejecutable de Windows (evaluando PyInstaller) para que
      el equipo no técnico abra la app con doble clic sin instalar Python.

## Estructura

```
app.py                          # Streamlit — sidebar por categoría, un script por vez
common/chrome_bootstrap.py      # Detección/instalación de Chrome, multiplataforma
scripts/copy_products/          # Copy Products (nuevo) — Productos
scripts/flag_as_deleted/        # Flag as Deleted — Productos
scripts/modificar_description_comment/  # Modificar Description y Comment — Productos
scripts/copy_pcm_linkeo/        # Copy y Linkeo de PCM — PCM
scripts/valorizacion_pkg/       # Valorización desde Componente — Valorización
scripts/valorizacion_madre/     # Valorización desde Servicio Madre — Valorización
scripts/valorizacion_excel/     # Valorización desde Excel — Valorización
scripts/notas_srv/              # Notas SRV (insertar/editar) — Notas
scripts/exportar_notas/         # Exportar Notas (solo lectura) — Notas
scripts/extraccion_vigencias/   # Vigencias (solo lectura) — Relevamiento
requirements.txt
run_app.bat                     # Launcher para Windows
```
