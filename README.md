# Chartpack_santander

## Proyección del IPC por bloques (método Carlomagno)

| Archivo | Contenido |
|---|---|
| `output/Proyeccion_IPC_Chile.html` | Reporte con metodología, proyección, evaluación y comparación con el mercado. Usa el formato de `Apuntes_IPC_Chile.html`. |
| `output/Base_Proyeccion_IPC.xlsx` | Base con todo lo utilizado: proyecciones, incidencias, modelos, evaluación, supuestos y series de la BDE. |
| `data/bde/` | Variables macro descargadas de la BDE del Banco Central (`bde_mensual.csv`, `bde_diario.csv`, `catalogo_series.csv`). |
| `data/insumos/` | Insumos originales (Excel automatizado, canasta 2023, matriz de productos, difusión, etc.). |
| `data/nowcast/` | Insumos del nowcasting: informes semanales de ENAP (con la URL de cada dato), precios ancla CNE y canasta de supermercados. |
| `output/nowcast/` | Nowcast de combustibles, backtest y puente al IPC (también en las hojas `NC_*` del Excel). |
| `scripts/` | Pipeline reproducible, scripts 01 a 07. |

```bash
pip install -r requirements.txt
export BCCH_USER=...  BCCH_PASS=...
python scripts/01_descarga_bde.py && python scripts/02_base.py && python scripts/03_modelo.py \
  && python scripts/04_resultados.py \
  && python scripts/06_scraper_supermercados.py && python scripts/07_nowcast.py && python scripts/05_reporte_html.py
```

El scraper de supermercados (`06`) necesita que el entorno permita jumbo.cl, santaisabel.cl y unimarc.cl, y debe correr a diario.
