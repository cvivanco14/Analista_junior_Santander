# Chartpack_santander

## Chartpack de Expectativas (`expectativas/`)

Actualiza el chartpack de Expectativas de Estudios Económicos (10 láminas: IMCE, expectativas de
inflación EEE/EOF, compensación inflacionaria, EDEP, costos de transporte, TCN/TCM, tasas SPC, EEE de
actividad e IMCE vs Imacec no minero) con el formato de la versión de abril 2025.

```bash
pip install pandas requests matplotlib python-pptx openpyxl
export BCCH_USER=... BCCH_PASS=...      # credenciales API BDE Banco Central
python expectativas/fetch_data.py       # descarga series a expectativas/data/
python expectativas/build_chartpack.py  # genera expectativas/output/AAAAMMDD_Chartpack_Expectativas.{pptx,pdf}
```

La conversión a PDF usa LibreOffice (`soffice`). Fuentes: Banco Central de Chile (API BDE) y
Reserva Federal de Nueva York (GSCPI, descarga directa).
