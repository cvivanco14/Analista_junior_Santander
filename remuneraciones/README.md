# Minuta de Remuneraciones (INE)

Descarga las cifras de empleo y remuneraciones del INE, calcula la masa salarial real y arma la
minuta en PPTX (gráficos nativos editables) y PDF, con `formato_graficos.py`. Misma estructura que la
minuta IMCE: un script de datos y uno de minuta.

```
masa salarial real = ocupados (desest. INE) × horas habituales semanales (desest. STL) × IR real (desest. STL)
```

| Carpeta | Contenido |
|---|---|
| `Code/` | `datos_remuneraciones.py` (descarga del INE, cálculos y Excel de bases) y `minuta_remuneraciones.py` (minuta PPTX/PDF y Excel de los gráficos). |
| `Inputs/` | Archivos descargados del INE (`ajuste_estacional.xlsx`, `horas.xlsx`, `tabulado_ir_real_empalmado.xlsx`, etc.). |
| `Outputs/` | Bases: `Masa_salarial_Base.xlsx` y `Cifras_empleo.xlsx`. Minuta: documento `Minuta_Remuneraciones_<mes>_<año>.pdf`, láminas `AAAAMMDD_Minuta_Remuneraciones.pptx` y `.pdf`, `_datos.xlsx` y `Masa_salarial.png`. |

```bash
cd remuneraciones/Code
python datos_remuneraciones.py      # descarga del INE y actualiza los Excel de bases
python minuta_remuneraciones.py     # arma la minuta en Outputs
```

En Spyder: abre y corre (F5) primero `datos_remuneraciones.py` y después `minuta_remuneraciones.py`.
`datos_remuneraciones.py --sin-descarga` rehace las bases con los archivos que ya están en `Inputs`.

## Minuta (`minuta_remuneraciones.py`)

### Documento `Minuta_Remuneraciones_<mes>_<año>.pdf`

Minuta escrita (como la de Coordinación Macroeconómica) con colores Santander, generada con `reportlab`:

- Título con el resultado del mes (salarios reales m/m sa y a/a), fecha y autor.
- **Principales resultados** (textos automáticos): IR nominal m/m (sa) frente al mes anterior, a/a, la
  variación m/m de la serie original frente al promedio 2010-2019 del mismo mes, e ICL; salarios reales
  m/m (sa) y a/a, inflación implícita en el IR real y meses consecutivos de alzas (o caídas) anuales;
  masa salarial real con sus contribuciones.
- **Cuadro 1**: variaciones m/m (sa) y a/a de IR nominal, ICL nominal e IR real (mes anterior y actual).
- **Figuras**: IR real desestacionalizado; variación anual de IR, ICL e IR real; y variación m/m (sa)
  del IR nominal y real por mes calendario, con la banda mín-máx y el promedio de 2010 a tres años
  antes, los dos años anteriores en líneas y el año en curso en rombos.
- "sa" usa la serie desestacionalizada del INE si el cuadro del INE la trae; si no, la de STL (la nota
  del cuadro dice cuál se usó).

### Láminas `AAAAMMDD_Minuta_Remuneraciones.pptx`

| Lámina | Contenido |
|---|---|
| Portada | Mes del último dato del IR y fecha. |
| Resumen | Textos automáticos: IR nominal e ICL (a/a), IR real (a/a y promedio del año) y masa salarial real con sus contribuciones; gráfico del IR nominal y real (a/a). |
| Remuneraciones y costo laboral | IR e ICL nominales (a/a) e IR real en nivel (original y desestacionalizado). |
| Masa salarial real | Variación anual con contribuciones apiladas de ocupados, horas y salario real, y nivel con y sin horas. |

- `_datos.xlsx` tiene una hoja con los datos de cada gráfico de las láminas y del documento (`Doc ...`).
- `formato_graficos.py` se busca como en la minuta IMCE: subiendo desde `Code`, en OneDrive/SharePoint
  y en la carpeta Santander de iCloud Drive (o en `FORMATO_GRAFICOS_DIR`).
- El autor (portada, pie de las láminas y documento) se cambia en `AUTOR`, al inicio del script.

## Bases (`datos_remuneraciones.py`)

- **`Masa_salarial_Base.xlsx`**:
  - *Masa salarial*: por mes, ocupados SA (miles), horas habituales SA, IR real SA (2023 = 100), masa
    salarial real (índice 2023 = 100) y su variación a/a, contribuciones de ocupados, horas y salario
    real (descomposición logarítmica; suman la variación total) y la versión sin horas (ocupados × IR
    real).
  - *Gráficos*: gráficos de Excel (editables) desde 2015.
  - *Léeme*: definiciones y fuentes.
- **`Cifras_empleo.xlsx`**: indicadores principales, categoría, informalidad, ajuste estacional,
  comparación ENE-RRAA, horas, rama y remuneraciones (IR, ICL e IR real, originales y desestacionalizados).

`Periodo` es el último mes del trimestre móvil de la ENE (`Nov - Ene` de 2015 = enero 2016); el IR es
mensual y se cruza con ese mes. Todo se reescribe en cada corrida.

## Cambios respecto del notebook "Series empleo3.ipynb"

- Rutas relativas al proyecto (`Inputs`/`Outputs`), sin mapeo de usuarios: funciona en Windows y macOS.
- La masa salarial se calcula con los datos de la misma corrida; antes se leía de `cifras_empleo.xlsx`
  de la corrida anterior, así que quedaba un mes atrasada.
- Las columnas se buscan por nombre y no por posición (`usecols=[0, 5]`, etc.).
- No usa `xlwings`; escribe con openpyxl. No toca el `cifras_empleo.xlsx` de SharePoint.
- Si una descarga falla o el INE responde con una página de error, conserva el archivo anterior. Si
  cambió el parámetro `sfvrsn` de la URL, reintenta sin él.
