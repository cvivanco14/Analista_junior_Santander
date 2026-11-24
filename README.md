# Analista junior Santander

Herramientas de la Gerencia de Estudios (Santander Chile). Cada carpeta es un proyecto
independiente con su propio README:

| Carpeta | Qué hace |
|---|---|
| [`formato/`](formato/) | `formato_graficos.py`, versión vigente del módulo común de datos y formato (gráficos nativos, plantilla de presentación, acceso al BDE). En SharePoint debe haber una copia idéntica en `00_Administrativos\Manual de Estilo`. |
| [`imce/`](imce/) | Minuta mensual del IMCE (`run_minuta_imce.py`; PPTX/PDF con gráficos nativos), con `Code/`, `Inputs/` y `Outputs/` (incluye el Excel de bases `Base_IMCE.xlsx`); y el chartbook anterior `imce_principal.py`. |
| [`minuta_ipc/`](minuta_ipc/) | Minuta IPC mensual en PDF (agregados con y sin volátiles, CBA, difusión y anexo de subclases), con `Code/`, `Inputs/` y `Outputs/`. |
| [`remuneraciones/`](remuneraciones/) | Minuta de Remuneraciones (INE), un solo script: `minuta_remuneraciones.py` descarga del INE (series originales y desestacionalizadas), deja las bases en `Inputs/` y arma el documento PDF, la presentación PPTX/PDF y `Outputs/IRL_Base.xlsx` (IR, ICL, IR real y masa salarial real). |
| [`chartpack_expectativas/`](chartpack_expectativas/) | Chartpack de Expectativas (PPTX/PDF con gráficos nativos): IMCE, EEE/EOF, compensaciones, EDEP, tipo de cambio, SPC, IPEC y más, con `Code/` (`encuestas_bcch.py` y `build_chartpack.py`), `Inputs/` y `Outputs/`; la base `Chartpack_Expectativas.xlsx` se actualiza sin perder fórmulas ni gráficos. |
| [`proyeccion_ipc/`](proyeccion_ipc/) | Proyección del IPC 2026-2027: un solo script (`Code/proyeccion_ipc.py`) que descarga INE y BDE, corre el modelo por bloques (Carlomagno) y el de 172 subclases, el nowcast de combustibles (ENAP) y la evaluación fuera de muestra, y genera `Outputs/IPC_Chile_Proyecciones.xlsx` (con fórmulas) y un reporte HTML. |

El formato común (`formato_graficos.py`) se usa desde SharePoint, en
`Estudios - Documentos\General\00_Administrativos\Manual de Estilo`; la versión vigente está en
[`formato/`](formato/) y esa copia de SharePoint debe reemplazarse cada vez que cambie. Los scripts lo
encuentran solos (o con la variable `FORMATO_GRAFICOS_DIR`) y muestran qué copia usan.

```bash
python -m pip install -r requirements.txt
```

Credenciales del BDE, una vez por computador; nunca se escriben en el código. Hay dos formas de acceso:
el **API token** (`BDE_TOKEN`, recomendado) o el **usuario y la clave** de la cuenta del BDE
(`BDE_USER` y `BDE_PASS`, siempre juntos). `formato_graficos.py` prueba primero el token y, si falta o
el BCCh lo rechaza, el usuario y la clave; si no hay nada guardado, los pide en pantalla.

**Windows con Visual Studio Code** (se guarda en el Administrador de credenciales de Windows):

1. Abrir la carpeta en VS Code y elegir el Python con que se corren los scripts:
   `Ctrl+Shift+P` → *Python: Select Interpreter*.
2. Abrir una terminal (*Terminal → Nuevo terminal*) y, una sola vez:
   ```powershell
   python -m pip install keyring
   python -c "import keyring, getpass; keyring.set_password('estudios_santander', 'BDE_TOKEN', getpass.getpass('Token BCCh: '))"
   ```
   El segundo comando pide el token sin mostrarlo: pegarlo con `Ctrl+V` y Enter. No hay que reiniciar
   VS Code. Para usar usuario y clave, lo mismo con `'BDE_USER'` (con `input(...)`) y `'BDE_PASS'`.
3. Comprobar: `python -c "import keyring; print(keyring.get_password('estudios_santander', 'BDE_TOKEN')[:4])"`
   debe mostrar `$2a$`.

Alternativa en Windows: `setx BDE_TOKEN 'tu_token'` en la terminal de VS Code (PowerShell, comillas
simples) o `setx BDE_TOKEN "tu_token"` en Anaconda Prompt, y luego cerrar y abrir VS Code por completo.

**macOS** (consola de Spyder; queda en el Llavero):
`import keyring, getpass; keyring.set_password('estudios_santander', 'BDE_TOKEN', getpass.getpass('Token BCCh: '))`
