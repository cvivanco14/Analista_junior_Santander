# Chartpack de Expectativas

Chartpack de Expectativas de Estudios Económicos (12 láminas): IMCE, expectativas de inflación
EEE/EOF, dispersión de la EEE (deciles), EOF con inflación a 2 años sobre/bajo umbrales,
compensación inflacionaria, EDEP, costos de transporte, TCN/TCM, tasas SPC, EEE de actividad e
IMCE vs Imacec no minero.

## Carpetas

| Carpeta | Contenido |
|---|---|
| `Code/` | `encuestas_bcch.py` (descarga y procesa la EEE y la EOF) y `build_chartpack.py` (arma el chartpack). |
| `Inputs/eee/` | Planillas de la EEE (`raw/`) y el Excel resultante `eee_inflacion_deciles.xlsx`. |
| `Inputs/eof/` | Planillas de la EOF (`raw/`) y el Excel resultante `eof_inflacion_2a_umbrales.xlsx` (hojas `mensual` y `por_encuesta`). |
| `Outputs/` | `AAAAMMDD_Chartpack_Expectativas.pptx`, `.pdf` y `_datos.xlsx` (datos de cada gráfico). |

El formato (`formato_graficos.py`) está en
`Estudios - Documentos\General\Administrativos\Manual de Estilo` (en este repositorio,
[`../formato/`](../formato/)). Los scripts lo encuentran subiendo desde su propia carpeta o en la
carpeta sincronizada de SharePoint del usuario, en Windows o macOS; no usan el nombre de usuario.

## Uso (no hay que editar nada)

```bash
cd chartpack_expectativas/Code
python encuestas_bcch.py      # baja las planillas EEE/EOF que falten y actualiza Inputs
python build_chartpack.py     # genera el chartpack en Outputs
```

- `encuestas_bcch.py --sin-descarga` reprocesa solo las planillas ya descargadas;
  `--solo eee` o `--solo eof` procesa una encuesta.
- Las planillas se descargan con Chromium (Playwright) porque el sitio del Banco Central bloquea
  descargas directas: instalar una vez con `python -m playwright install chromium`.
- EEE: mediana, decil 1 y decil 9 de inflación en 11 y 23 meses, desde 2014. Reconoce los nombres
  de archivo antiguos y los nuevos (`datos-mensuales-<mes>-<año>-xlsx`, desde febrero de 2026).
- EOF: % de operadores con inflación a 13–24 meses sobre 3,0%, sobre 3,2% y bajo 2,8%, desde
  agosto de 2013; un dato por mes (promedio de las encuestas que cerraron ese mes; noviembre de
  2021 queda vacío).
- `build_chartpack.py` descarga las demás series de la API del BCCh (token, ver README principal) y
  exporta el PDF con PowerPoint (Windows/Mac) o LibreOffice.
- Las planillas raw (~260 MB) no se suben a GitHub; en SharePoint quedan en `Inputs/*/raw/`.
