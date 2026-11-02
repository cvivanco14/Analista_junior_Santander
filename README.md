# Analista junior Santander

Herramientas de la Gerencia de Estudios (Santander Chile). Cada carpeta es un proyecto
independiente con su propio README:

| Carpeta | Qué hace |
|---|---|
| [`chartpack_expectativas/`](chartpack_expectativas/) | Chartpack de Expectativas (PPTX/PDF con gráficos nativos editables) y el módulo de formato `formato_graficos.py`. |
| [`encuestas_eee_eof/`](encuestas_eee_eof/) | Descarga de planillas de la EEE y la EOF del Banco Central y series de dispersión (deciles EEE, umbrales EOF). |
| [`imce/`](imce/) | Proceso mensual del IMCE: actualiza `IMCE.xlsx` y genera el chartbook PDF. |

```bash
python -m pip install -r requirements.txt
setx BDE_TOKEN 'tu_token'   # una vez por computador: API Key Token del BCCh (también acepta BCCH_API_token)
```

Las credenciales se leen siempre de variables de entorno; nunca se escriben en el código.
