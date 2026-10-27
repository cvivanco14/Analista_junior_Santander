# Chartpack_santander

## Proyección del IPC por bloques (método Carlomagno)

| Archivo | Contenido |
|---|---|
| `output/Proyeccion_IPC_Chile.html` | Reporte con metodología, proyección, evaluación y comparación con el mercado. Usa el formato de `Apuntes_IPC_Chile.html`. |
| `output/Base_Proyeccion_IPC.xlsx` | Base con todo lo utilizado: proyecciones, incidencias, modelos, evaluación, supuestos y series de la BDE. |
| `data/bde/` | Variables macro descargadas de la BDE del Banco Central (`bde_mensual.csv`, `bde_diario.csv`, `catalogo_series.csv`). |
| `data/insumos/` | Insumos originales (Excel automatizado, canasta 2023, matriz de productos, difusión, etc.). |
| `scripts/` | Pipeline reproducible, scripts 01 a 05. |

```bash
pip install -r requirements.txt
export BCCH_USER=...  BCCH_PASS=...
python scripts/01_descarga_bde.py && python scripts/02_base.py && python scripts/03_modelo.py \
  && python scripts/04_resultados.py && python scripts/05_reporte_html.py
```
