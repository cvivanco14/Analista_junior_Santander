# Chartpack_santander

## EEE: dispersión de expectativas de inflación

Replica los gráficos "EEE: Inflación en 11 meses" y "EEE: Inflación en 23 meses"
(mediana, decil 1, decil 9 y barras decil 9 − decil 1) con las planillas mensuales de la
[Encuesta de Expectativas Económicas](https://www.bcentral.cl/encuestas-economicas/resultados-anteriores-eee)
del Banco Central de Chile, desde enero de 2014.

```bash
pip install -r requirements.txt
python descargar_bcch.py eee   # baja las planillas a eee/raw/ (usa Chromium vía Playwright)
python eee/graficos.py    # genera eee/output/*.png y eee_inflacion_deciles.csv
```

- `descargar_bcch.py`: el sitio está protegido por Incapsula, así que las planillas se bajan
  con un navegador headless. Solo descarga las que faltan en `eee/raw/`.
- Si un mes aún no está en la página, deja el Excel ("Datos mensuales <mes> <año>.xlsx")
  en `eee/raw/`, o pasa otra carpeta: `python eee/graficos.py ruta/a/planillas`.
- `eee/output/eee_inflacion_deciles.csv`: serie mensual por horizonte (11m, 23m) con
  mediana, decil 1, decil 9, rango 9−1 y archivo de origen.
