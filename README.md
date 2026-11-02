# Analista junior Santander

Herramientas de la Gerencia de Estudios (Santander Chile). Cada carpeta es un proyecto
independiente con su propio README:

| Carpeta | Qué hace |
|---|---|
| [`formato/`](formato/) | `formato_graficos.py`: datos (API BDE, FRED, FMI, GSCPI), transformaciones con nota metodológica y formato Santander de gráficos nativos de PowerPoint; portada y logos. En SharePoint vive en `Administrativos\Manual de Estilo`; lo usan el chartpack y el IMCE. |
| [`chartpack_expectativas/`](chartpack_expectativas/) | Chartpack de Expectativas (PPTX/PDF con gráficos nativos editables), con `Code/`, `Inputs/` (EEE y EOF) y `Outputs/`. |
| [`imce/`](imce/) | Proceso mensual del IMCE: actualiza `IMCE.xlsx` y genera el chartbook en PPTX y PDF. |

```bash
python -m pip install -r requirements.txt
```

Token del BCCh (API Key Token), una vez por computador; nunca se escribe en el código:

- **macOS** (consola de Spyder; queda en el Llavero):
  `import keyring, getpass; keyring.set_password('estudios_santander', 'BDE_TOKEN', getpass.getpass('Token BCCh: '))`
- **Windows**: `setx BDE_TOKEN "tu_token"` en Anaconda Prompt/cmd (en PowerShell, con comillas
  simples), y luego cerrar y abrir Spyder / VS Code. También se acepta la variable `BCCH_API_token`.
