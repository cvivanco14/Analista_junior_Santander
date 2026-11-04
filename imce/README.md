# IMCE: proceso mensual

`imce_principal.py` actualiza `IMCE.xlsx` desde la API del BCCh y genera el chartbook en PPTX
(gráficos nativos editables) y PDF. No define formato propio: usa `formato_graficos.py`
(en este repositorio, [`../formato/`](../formato/)), igual que el chartpack de Expectativas.

## Carpetas en SharePoint (sirven para cualquier usuario)

| Qué | Dónde |
|---|---|
| Este script | `...\Estudios - Documentos\General\Expectativas y Encuestas\IMCE\Code` |
| `formato_graficos.py` | `...\Estudios - Documentos\General\Administrativos\Manual de Estilo` |

- El script ubica el Manual de Estilo subiendo desde su propia carpeta; si no lo encuentra,
  lo busca en la carpeta sincronizada `Santander Office 365` del usuario que lo corre o, en Mac,
  en la carpeta `Santander` de iCloud Drive. No usa el
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
