# Chartpack de Expectativas

Chartpack de Expectativas de Estudios Económicos (14 láminas): IMCE, expectativas de inflación
EEE/EOF, dispersión de la EEE (deciles), EOF con inflación a 2 años sobre/bajo umbrales,
compensación inflacionaria, EDEP, costos de transporte, TCN/TCM, tasas SPC, EEE de actividad,
IPEC (índice de percepción de la economía y sus componentes), expectativas de consumo de artículos
para el hogar (IPEC) e IMCE vs Imacec no minero.

## Carpetas

| Carpeta | Contenido |
|---|---|
| (raíz) | `run_chartpack_expectativas.py`: corre los dos scripts de `Code/` en orden. |
| `Code/` | `encuestas_bcch.py` (descarga y procesa la EEE y la EOF) y `build_chartpack.py` (arma el chartpack). |
| `Inputs/eee/` | Planillas de la EEE (`raw/`) y el Excel resultante `eee_inflacion_deciles.xlsx`. |
| `Inputs/eof/` | Planillas de la EOF (`raw/`) y el Excel resultante `eof_inflacion_2a_umbrales.xlsx` (hojas `mensual` y `por_encuesta`). |
| `Outputs/` | `AAAAMMDD_Chartpack_Expectativas.pptx` y `.pdf`; en macOS, también `Chartpack_Expectativas.xlsx`. |

**Base de datos `Chartpack_Expectativas.xlsx`** (una hoja por gráfico, siempre con ese nombre):

- En Windows con OneDrive/SharePoint se guarda en
  `Estudios - Documentos\General\10_Bases y Analisis\Local\Expectativas y Encuestas` (la carpeta
  `Local\Expectativas y Encuestas` se crea si no existe). En macOS (o si no se encuentra
  `10_Bases y Analisis`) queda en `Outputs/`.
- Si el archivo ya existe, se abre y solo se reescriben las columnas de cada hoja del chartpack (desde
  A1); hojas, columnas, fórmulas y gráficos propios se conservan. Si lo guardó Excel y tiene
  gráficos, se actualiza con Excel (`xlwings`); si Excel no está disponible, con openpyxl, dejando un
  respaldo en `Inputs/respaldo/` (openpyxl conserva fórmulas, pero puede perder gráficos).
- Los nombres de las hojas son fijos (por ejemplo `IPEC`, `IMCE precios`), para que las fórmulas que
  las usan sigan funcionando aunque cambie el orden de las láminas.

El formato (`formato_graficos.py`) está en `Estudios - Documentos\General\00_Administrativos\Manual de Estilo`.
Los scripts lo encuentran subiendo desde su propia carpeta, en la carpeta sincronizada de SharePoint
del usuario (Windows o macOS) o en la carpeta `Santander` de iCloud Drive (macOS); no usan el nombre
de usuario. También se puede indicar con la variable `FORMATO_GRAFICOS_DIR`.

## Uso (no hay que editar nada)

Lo más simple: abrir `run_chartpack_expectativas.py` (en la carpeta Expectativas, junto a `Code/`,
`Inputs/` y `Outputs/`) en Spyder o VS Code y presionar F5 / Run. Corre los dos scripts en orden:
primero `encuestas_bcch.py` y después `build_chartpack.py`. Con `SIN_DESCARGA_ENCUESTAS = True` no
descarga planillas EEE/EOF. Si las encuestas fallan, el chartpack se arma igual con los Excel que ya
hay en `Inputs`.

También se pueden correr por separado:

```bash
cd chartpack_expectativas/Code
python encuestas_bcch.py      # baja las planillas EEE/EOF que falten y actualiza Inputs
python build_chartpack.py     # genera el chartpack en Outputs y actualiza Chartpack_Expectativas.xlsx
```

- Credenciales del BDE: ver el README principal (Windows con VS Code o macOS).
- `encuestas_bcch.py --sin-descarga` reprocesa solo las planillas ya descargadas;
  `--solo eee` o `--solo eof` procesa una encuesta.
- Las planillas se descargan con Chromium (Playwright) porque el sitio del Banco Central bloquea
  descargas directas: instalar una vez con `python -m playwright install chromium` (en Windows, si
  falta, usa Microsoft Edge).
- EEE: mediana, decil 1 y decil 9 de inflación en 11 y 23 meses, desde 2014. Reconoce los nombres
  de archivo antiguos y los nuevos (`datos-mensuales-<mes>-<año>-xlsx`, desde febrero de 2026).
- EOF: % de operadores con inflación a 13–24 meses sobre 3,0%, sobre 3,2% y bajo 2,8%, desde
  agosto de 2013; un dato por mes (promedio de las encuestas que cerraron ese mes; noviembre de
  2021 queda vacío).
- IPEC (GfK Adimark, publicado en la BDE): índice de percepción de la economía y sus componentes
  situación económica personal actual, situación económica del país a 12 meses y expectativas de
  consumo de artículos para el hogar (`F089.IPE.IND.70M2`, `72M2`, `74M2` y `75M2`).
- `build_chartpack.py` descarga las demás series de la API del BCCh y exporta el PDF con PowerPoint
  (Windows/Mac) o LibreOffice.
- Las planillas raw (~260 MB) no se suben a GitHub; en SharePoint quedan en `Inputs/*/raw/`.
