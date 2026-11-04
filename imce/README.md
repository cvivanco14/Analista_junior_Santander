# IMCE: proceso mensual

## Minuta IMCE (`Code/`, `Inputs/`, `Outputs/`)

Minuta mensual de 13 láminas (resumen con textos automáticos, resultados, gráficos por sector e
índices compuestos OCSI/PCSI/CFEI/OFEI), en PPTX con gráficos nativos editables y en PDF, con
`formato_graficos.py`. Misma estructura que el chartpack de Expectativas:

| Carpeta | Contenido |
|---|---|
| `Code/` | `datos_imce.py` (descarga y Excel de bases) y `minuta_imce.py` (arma la minuta). |
| `Inputs/` | `imce.csv` (series descargadas de la API del BCCh) y `hacienda_imce_sa.csv` (compuestos de Hacienda, referencia). |
| `Outputs/` | `IMCE.xlsx` (bases y cálculos) y `AAAAMMDD_Minuta_IMCE.pptx`, `.pdf` y `_datos.xlsx` (datos de cada gráfico). |

```bash
cd imce/Code
python datos_imce.py      # descarga, actualiza Inputs/imce.csv y Outputs/IMCE.xlsx
python minuta_imce.py     # genera la minuta en Outputs
```

- `IMCE.xlsx`: usa el que ya esté en `Outputs`; si no, el de la ubicación anterior
  (`IMCE\Base\IMCE.xlsx` o `Bases - Chartpack/.../IMCE/Base/IMCE.xlsx`), que copia a `Outputs`; si
  no hay ninguno, crea uno nuevo. `IMCE_XLSX` permite forzar otro archivo.
- En la hoja `IMCE` (fila 3 = nombres, como la planilla de siempre) solo escribe las columnas de las
  series: agrega meses y series nuevas y conserva columnas y fórmulas propias. Las hojas `Léeme`,
  `SA` (desestacionalizadas con STL), `Pesos` y `Compuestos` (con fórmulas de Excel) se reescriben en
  cada corrida; los cálculos y gráficos propios van en otras hojas.
- Si `IMCE.xlsx` tiene gráficos, lo actualiza con Excel mediante `xlwings` (openpyxl los borraría);
  si `xlwings` no está instalado, no toca el archivo y avisa.
- `datos_imce.py --sin-descarga` rehace `IMCE.xlsx` con el `Inputs/imce.csv` existente.
- El token del BCCh lo lee `formato_graficos.py` (variable de entorno o Llavero; ver README
  principal). No hay credenciales en el código.

## Chartbook (`imce_principal.py`)

Script anterior: `imce_principal.py` actualiza `IMCE.xlsx` desde la API del BCCh y genera el chartbook en PPTX
(gráficos nativos editables) y PDF. No define formato propio: usa `formato_graficos.py`
(en este repositorio, [`../formato/`](../formato/)), igual que el chartpack de Expectativas.

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
