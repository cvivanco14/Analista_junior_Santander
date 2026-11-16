# IMCE: proceso mensual

## Minuta IMCE (`run_minuta_imce.py`, `Code/`, `Inputs/`, `Outputs/`)

Presentación mensual de 16 láminas en PPTX (gráficos nativos editables, con sus datos en "Editar
datos") y PDF, con `formato_graficos.py`. Misma estructura que la Minuta de Remuneraciones:

| Carpeta / archivo | Contenido |
|---|---|
| `run_minuta_imce.py` | Lanzador: corre `Code/minuta_imce.py` (F5 en Spyder). `SIN_DESCARGA = True` usa el `Inputs/imce.csv` existente. |
| `Code/` | `minuta_imce.py` (arma la presentación y antes corre `datos_imce.py`) y `datos_imce.py` (descarga y Excel de bases). |
| `Inputs/` | `imce.csv` (series de la API del BCCh), `Base_IMCE_plantilla.xlsx` (planilla IMCE con hojas `IMCE` y `Grafs`) y `hacienda_imce_sa.csv` (referencia). |
| `Outputs/` | `Base_IMCE.xlsx` y `Presentación_IMCE_<mes>_<año>.pptx`, `.pdf` y `_datos.xlsx` (datos de cada gráfico). Al empezar, las presentaciones anteriores pasan a `Outputs/<año>/` con la fecha al inicio del nombre; `Base_IMCE.xlsx` no se mueve. |

```bash
python imce/run_minuta_imce.py
```

**Láminas.** Portada; resumen y resumen sin minería (con gráfico por mes calendario); principales
resultados (2); y luego, cada lámina de gráficos con una viñeta por gráfico con el dato del mes y su
diferencia con el mes anterior (por sobre / por debajo), sin el número al final de las líneas:

| Lámina | Gráficos |
|---|---|
| 6 | IMCE total y sin minería |
| 7 | IMCE por sectores (comercio, industria y construcción) |
| 8 | Expectativas de salarios, costos y precios por sector |
| 9 | Promedio simple de los sectores: salarios (construcción e industria), costos y precios |
| 10 | Expectativas de empleo por sector |
| 11 | Promedio simple de expectativas de empleo |
| 12 | Situación actual de inventarios y demanda (desest.) por sector |
| 13 | Promedio simple de inventarios (comercio e industria) y de demanda (desest.) |
| 14 | Expectativas de inversión, ventas y producción |
| 15 y 16 | Índices compuestos OCSI/PCSI y CFEI/OFEI con contribuciones por sector |

**Pesos de los índices compuestos.** Participación en el PIB 2018 (BCCh, cuentas nacionales
referencia 2018): comercio 9,2%, construcción 6,5% e industria 9,6%, normalizada entre los tres
sectores (36,4%, 25,7% y 37,9%). La planilla de Hacienda usaba "Comercio, restaurantes y hoteles"
(11,4%), que daba 41,4% / 23,6% / 35,0%; pero el IMCE comercio no encuesta restaurantes ni hoteles.
Referencia: el IMCE sin minería que publica ICARE pondera fijo comercio 31,1%, construcción 22,7% e
industria 46,2% (se verifica exacto en toda la serie). Los pesos están en `PIB_2018` de `datos_imce.py`.

- `Base_IMCE.xlsx` se actualiza en cada corrida. Si no existe, se crea copiando
  `Inputs/Base_IMCE_plantilla.xlsx` (o, si falta, la ubicación anterior `IMCE\Base\IMCE.xlsx`).
- **No se sobrescribe nunca:** no se borra el archivo ni se limpia ninguna hoja. En cada hoja el
  script solo escribe sus columnas (las ubica por su nombre en la fila 3; si faltan, las agrega al
  final) y agrega los meses nuevos al final, con el formato de la fila de arriba. Columnas, fórmulas,
  hojas y gráficos propios se conservan; los gráficos (hoja `Grafs` y el de la hoja `IMCE`) se
  extienden hasta el último mes.
- Hojas que agrega: `SA` (desestacionalizadas con STL), `Pesos` (se crea una vez; si cambias los
  pesos, se respetan; los pesos anteriores de Hacienda se reemplazan solos), `Compuestos` (OCSI/PCSI/CFEI/OFEI con fórmulas sobre `SA` y `Pesos`),
  `Promedios` (promedio simple de los sectores de precios, costos, demanda, demanda desestacionalizada,
  salarios, empleo e inventarios, con fórmulas `=AVERAGE(...)` sobre `IMCE` y `SA`) y `Léeme`.
- Si el archivo se guardó con Excel y tiene gráficos, se actualiza con Excel mediante `xlwings`; si
  `xlwings` no está, se actualiza con openpyxl (conserva los gráficos, pero puede simplificar parte de
  su formato) y deja un respaldo en `Outputs/respaldo/`.
- `minuta_imce.py` corre primero `datos_imce.py`. Con `--sin-descarga` usa el `Inputs/imce.csv` existente
  (en Spyder: `runfile('minuta_imce.py', args='--sin-descarga')`). `datos_imce.py` se puede correr
  solo para actualizar únicamente los datos y el Excel.
- `datos_imce.py --sin-descarga` rehace `Base_IMCE.xlsx` con el `Inputs/imce.csv` existente.
- `formato_graficos.py` se busca en `00_Administrativos\Manual de Estilo` (o en `Administrativos`, el
  nombre anterior): subiendo desde `Code`, en OneDrive/SharePoint y en la carpeta Santander de iCloud Drive (o `FORMATO_GRAFICOS_DIR`). Usa la versión con portada roja y pie de lámina.
- Las credenciales del BCCh las maneja `formato_graficos.py` (variable de entorno o Llavero; ver README
  principal). No hay credenciales en el código.

## Chartbook (`imce_principal.py`)

Script anterior: `imce_principal.py` actualiza `IMCE.xlsx` desde la API del BCCh y genera el chartbook en PPTX
(gráficos nativos editables) y PDF. No define formato propio: usa `formato_graficos.py`
(en SharePoint, `Administrativos\Manual de Estilo`), igual que el chartpack de Expectativas.

## Carpetas en SharePoint (sirven para cualquier usuario)

| Qué | Dónde |
|---|---|
| Este script | `...\Estudios - Documentos\General\Expectativas y Encuestas\IMCE\Code` |
| `formato_graficos.py` + `assets\` | `...\Estudios - Documentos\General\Administrativos\Manual de Estilo` |

- El script ubica el Manual de Estilo subiendo desde su propia carpeta; si no lo encuentra,
  lo busca en la carpeta sincronizada `Santander Office 365` del usuario que lo corre. No usa el
  nombre de usuario (`C:\Users\<usuario>`), así que funciona para cualquiera.
  `FORMATO_GRAFICOS_DIR` permite forzar otra carpeta.
No hay que editar nada:

- Credenciales desde variables de entorno (`BDE_TOKEN`/`BCCH_API_token`, o `BCCH_USER` y
  `BCCH_PASS`) o desde el Llavero de macOS (ver README principal); nunca escritas en el código.
- Funciona en Windows y macOS: en Mac encuentra SharePoint en `~/Library/CloudStorage` y exporta
  el PDF con PowerPoint para Mac (o LibreOffice).
- Estructura del proyecto: `IMCE\Code` (scripts), `IMCE\Base` (`IMCE.xlsx`) e `IMCE\Output`
  (chartbook PPTX y PDF; se crea si no existe).
- Busca `IMCE.xlsx` primero en `IMCE\Base\`, luego en el resto de la carpeta `IMCE` y por último
  en la ubicación anterior
  (`.../Bases - Chartpack/Bases/Local/IMCE/Base/IMCE.xlsx`). `IMCE_XLSX` permite forzar otra ruta.
- El PDF lo exporta PowerPoint en Windows (con `pywin32`) o LibreOffice si está instalado.
- Antes de modificar `IMCE.xlsx` prueba la credencial; si el Banco Central la rechaza
  ("Invalid username or password"), se detiene e indica qué variable usó y cómo corregirla
  (la causa más común es guardar el token con comillas dobles en PowerShell).

```powershell
python imce/imce_principal.py
```

## Si aparece "Invalid username or password"

El script indica de dónde leyó el token. Si dice "de la variable BDE_TOKEN" y el largo no es ~60
o no empieza con `$2a$`, esa variable tiene un valor equivocado (típico de guardarla con comillas
dobles, que borran las partes con `$`). Si hay un token válido en el Llavero, el script lo usa
igual y avisa; para que no vuelva a interferir, borra la variable:

- **macOS**, en la Terminal:
  ```bash
  grep -n BDE_TOKEN ~/.zshrc ~/.zprofile ~/.bash_profile 2>/dev/null   # borra la línea que aparezca
  launchctl unsetenv BDE_TOKEN
  conda env config vars unset BDE_TOKEN
  ```
  y cierra y abre Spyder.
- **Windows**, en Anaconda Prompt: `reg delete HKCU\Environment /v BDE_TOKEN /f`, y cierra y abre Spyder.
