# Chartpack de Expectativas

Chartpack de Expectativas de Estudios Económicos (12 láminas): IMCE, expectativas de inflación
EEE/EOF, dispersión de la EEE (deciles), EOF con inflación a 2 años sobre/bajo umbrales,
compensación inflacionaria, EDEP, costos de transporte, TCN/TCM, tasas SPC, EEE de actividad e
IMCE vs Imacec no minero.

- `formato_graficos.py`: datos (API BDE, FRED, FMI, GSCPI), transformaciones que devuelven su
  nota metodológica y formato Santander con gráficos nativos editables de PowerPoint (líneas de
  2,25 pt mensuales y 1,75 pt diarias, fechas horizontales, números sin ceros sobrantes, último
  dato etiquetado —con el día en series diarias—, descriptor de la serie en 11 y fuente en
  Santander Text Light 9).
- `build_chartpack.py`: contenido del chartpack; llama a `formato_graficos`. Las láminas de
  dispersión EEE y umbrales EOF leen las series de [`../encuestas_eee_eof/`](../encuestas_eee_eof/)
  (actualizarlas primero).
- `assets/`: portada y logos. `output/`: PPTX, PDF y Excel con los datos de cada gráfico.

```bash
python chartpack_expectativas/build_chartpack.py   # output/AAAAMMDD_Chartpack_Expectativas.{pptx,pdf} + _datos.xlsx
```

El PDF se genera solo si LibreOffice (`soffice`) está instalado. Las tipografías Santander se usan si
están instaladas; si no, Arial.
