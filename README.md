# Analista junior Santander

Herramientas de la Gerencia de Estudios (Santander Chile). Cada carpeta es un proyecto
independiente con su propio README:

| Carpeta | Qué hace |
|---|---|
| [`formato/`](formato/) | `formato_graficos.py`: datos (API BDE, FRED, FMI, GSCPI), transformaciones con nota metodológica y formato Santander de gráficos nativos de PowerPoint; portada y logos. Lo usan el chartpack y el IMCE. |
| [`chartpack_expectativas/`](chartpack_expectativas/) | Chartpack de Expectativas (PPTX/PDF con gráficos nativos editables). |
| [`encuestas_eee_eof/`](encuestas_eee_eof/) | Descarga de planillas de la EEE y la EOF del Banco Central y series de dispersión (deciles EEE, umbrales EOF). |
| [`imce/`](imce/) | Proceso mensual del IMCE: actualiza `IMCE.xlsx` y genera el chartbook en PPTX y PDF. |

```bash
python -m pip install -r requirements.txt
setx BDE_TOKEN 'tu_token'   # una vez por computador: API Key Token del BCCh (también acepta BCCH_API_token)
```

Las credenciales se leen siempre de variables de entorno; nunca se escriben en el código.
