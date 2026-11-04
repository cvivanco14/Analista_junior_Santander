# Formato de gráficos Santander

`formato_graficos.py` es el módulo común de datos y formato de la Gerencia de Estudios. Lo usan
[`../chartpack_expectativas/`](../chartpack_expectativas/) y [`../imce/`](../imce/). En SharePoint está en `Estudios - Documentos\General\Administrativos\Manual de Estilo`.

- Datos: API BDE del BCCh (`bde`, `buscar_bde`), FRED, FMI (WEO) y GSCPI de la Fed de Nueva York.
- Transformaciones explícitas que devuelven su nota metodológica (`var_anual`, `media_movil`,
  `base100`, `promedio`).
- Barras nativas editables (`agregar_barras`) con el formato de la plantilla: columnas o barras
  horizontales, agrupadas o apiladas, valor sobre cada barra en el color de su serie (sin ceros
  sobrantes), sin eje de valores y con líneas guía suaves, leyenda abajo; colores en orden #EC0000,
  #000000, #4DC4CA, #FFAE42 y #B2B2B2 (`PALETA_BARRAS`).
- Leyenda de los gráficos de líneas (`leyenda="auto"`, por defecto): dentro del gráfico, en el
  rincón donde no tapa series, etiquetas ni el descriptor de unidad; si no hay espacio, arriba.
- Gráficos nativos editables de PowerPoint (`agregar_grafico`, `grafico_lineas`): líneas de
  2,25 pt mensuales y 1,75 pt diarias, fechas horizontales, números sin ceros sobrantes, último
  dato etiquetado (con el día en series diarias), descriptor de la serie en 11, nota y fuente en
  Santander Text Light 9, eje derecho y columnas.
- Presentación: `nueva_presentacion`, `portada`, `lamina`, `pie` y `exportar_pdf`.
- Plantilla de presentación (minuta mensual de Estudios): portada roja y láminas con franja roja,
  título rojo, línea gris inferior, pie "documento | autor" y número de página.

Para usarlo desde otra carpeta del repositorio:

```python
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "formato"))
import formato_graficos as fg
```
