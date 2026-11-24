# Formato de gráficos Santander

`formato_graficos.py` es el módulo común de datos y formato de la Gerencia de Estudios. Lo usan
el [chartpack de Expectativas](../chartpack_expectativas/), la [Minuta IMCE](../imce/) y la
[Minuta de Remuneraciones](../remuneraciones/).

**Esta carpeta guarda la versión vigente.** En SharePoint debe haber una copia idéntica en
`Estudios - Documentos\General\00_Administrativos\Manual de Estilo\formato_graficos.py`: cada vez que
cambie aquí, reemplazar esa copia (y borrar las copias antiguas, por ejemplo en `Administrativos\`).
La constante `VERSION` indica la fecha de la versión; `build_chartpack.py` la muestra al partir
(`Formato: <ruta> (versión ...)`), así se ve qué copia se está usando.

Contenido:

- Datos: API BDE del BCCh (`bde`, `buscar_bde`), FRED, FMI (WEO) y GSCPI de la Fed de Nueva York.
  Acceso al BDE con API token (`BDE_TOKEN`) o con usuario y clave (`BDE_USER` y `BDE_PASS`); ver el
  docstring del módulo y el README principal para guardarlos en Windows (VS Code) o macOS.
- Transformaciones explícitas que devuelven su nota metodológica (`var_anual`, `media_movil`,
  `base100`, `promedio`).
- Gráficos de líneas nativos y editables (`agregar_grafico`, `grafico_lineas`): líneas de 2,25 pt
  mensuales y 1,75 pt diarias, fechas horizontales, números sin ceros sobrantes, último dato
  etiquetado (con el día en series diarias), descriptor de la serie en 11, nota y fuente en
  Santander Text Light 9, eje derecho, columnas, leyenda automática (`leyenda="auto"`) y rótulos
  "▲ optimista / ▼ pesimista" para índices de difusión (`zonas=`).
- Barras nativas (`agregar_barras`) con el formato de la plantilla y `PALETA_BARRAS`.
- Presentación con la plantilla de la minuta de Estudios: `nueva_presentacion`, `portada`, `lamina`,
  `pie` y `exportar_pdf`.

Para usarlo desde otra carpeta del repositorio (los scripts ya lo buscan solos):

```python
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "formato"))
import formato_graficos as fg
```
