# Analista junior Santander

Herramientas de la Gerencia de Estudios (Santander Chile). Cada carpeta es un proyecto
independiente con su propio README:

| Carpeta | Qué hace |
|---|---|
| [`imce/`](imce/) | Minuta mensual del IMCE (PPTX/PDF con gráficos nativos), con `Code/`, `Inputs/` y `Outputs/` (incluye el Excel de bases `Base_IMCE.xlsx`); y el chartbook anterior `imce_principal.py`. |
| [`minuta_ipc/`](minuta_ipc/) | Minuta IPC mensual en PDF (agregados con y sin volátiles, CBA, difusión y anexo de subclases), con `Code/`, `Inputs/` y `Outputs/`. |
| [`remuneraciones/`](remuneraciones/) | Minuta de Remuneraciones (INE), un solo script: `minuta_remuneraciones.py` descarga del INE (series originales y desestacionalizadas), deja las bases en `Inputs/` y arma el documento PDF, la presentación PPTX/PDF y `Outputs/IRL_Base.xlsx` (IR, ICL, IR real y masa salarial real). |
| [`proyeccion_ipc/`](proyeccion_ipc/) | Proyección del IPC 2026-2027: un solo script (`Code/proyeccion_ipc.py`) que descarga INE y BDE, corre el modelo por bloques (Carlomagno) y el de 172 subclases, el nowcast de combustibles (ENAP) y la evaluación fuera de muestra, y genera `Outputs/IPC_Chile_Proyecciones.xlsx` (con fórmulas) y un reporte HTML. |

El formato común (`formato_graficos.py` y su carpeta `assets`) vive en SharePoint, en
`Estudios - Documentos\General\Administrativos\Manual de Estilo`; los scripts lo encuentran solos
(o con la variable `FORMATO_GRAFICOS_DIR`).

```bash
python -m pip install -r requirements.txt
```

Token del BCCh (API Key Token), una vez por computador; nunca se escribe en el código:

- **macOS** (consola de Spyder; queda en el Llavero):
  `import keyring, getpass; keyring.set_password('estudios_santander', 'BDE_TOKEN', getpass.getpass('Token BCCh: '))`
- **Windows**: `setx BDE_TOKEN "tu_token"` en Anaconda Prompt/cmd (en PowerShell, con comillas
  simples), y luego cerrar y abrir Spyder / VS Code. También se acepta la variable `BCCH_API_token`.
