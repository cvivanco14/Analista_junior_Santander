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

## EOF: operadores con inflación a 2 años sobre/bajo umbrales

Con la hoja "Distribución" de cada [Encuesta de Operadores Financieros](https://www.bcentral.cl/encuestas-economicas/resultados-anteriores-eof)
(pregunta de inflación en los siguientes 12 meses, 13 al 24) cuenta cuántos operadores
responden **sobre 3,0%**, **sobre 3,2%** y **bajo 2,8%** (desigualdades estrictas), desde agosto de 2013.

```bash
python descargar_bcch.py eof   # baja las planillas a eof/raw/
python eof/graficos.py         # genera eof/output/*.png y eof_inflacion_2a_umbrales.csv
```

- Los gráficos muestran un dato por mes: el promedio de los porcentajes de las encuestas que
  cerraron ese mes (las dos quincenales hasta 2018; luego Pre RPM y Post RPM). Noviembre de
  2021 queda vacío porque no hubo encuesta que cerrara ese mes.
- `eof/output/eof_inflacion_2a_umbrales_mensual.csv`: promedio mensual y n° de encuestas promediadas.
- `eof/output/eof_inflacion_2a_umbrales.csv`: detalle por encuesta (fecha de cierre, tipo,
  total de respuestas, conteos y porcentajes por umbral, archivo de origen).
