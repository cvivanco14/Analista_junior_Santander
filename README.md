# Chartpack_santander

## Chartpack de Expectativas (`expectativas/`)

Actualiza el chartpack de Expectativas de Estudios Económicos (10 láminas: IMCE, expectativas de
inflación EEE/EOF, compensación inflacionaria, EDEP, costos de transporte, TCN/TCM, tasas SPC, EEE de
actividad e IMCE vs Imacec no minero) con el formato de la versión de abril 2025.

Dos scripts:

- `expectativas/macro.py`: módulo de datos y formato de la Gerencia de Estudios (API BDE, FRED,
  FMI, GSCPI; transformaciones que devuelven su nota metodológica; estilo Santander; gráficos
  nativos de PowerPoint editables; portada y láminas).
- `expectativas/build_chartpack.py`: contenido del chartpack (series, transformaciones, láminas);
  llama a `macro`.

```bash
python -m pip install bcchapi pandas openpyxl requests python-pptx matplotlib
setx BDE_TOKEN 'tu_token'               # API Key Token del BCCh (también acepta BCCH_API_token)
cd expectativas
python build_chartpack.py               # output/AAAAMMDD_Chartpack_Expectativas.{pptx,pdf} + _datos.xlsx
```

El PDF se genera solo si LibreOffice (`soffice`) está instalado. Las tipografías Santander se usan si
están instaladas; si no, Arial.
