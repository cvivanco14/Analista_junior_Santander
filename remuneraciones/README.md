# Minuta de Remuneraciones (INE)

Descarga las cifras de remuneraciones y empleo del INE, calcula la masa salarial real y arma la minuta:
documento PDF, presentación PPTX/PDF (con `formato_graficos.py`) e `IRL_Base.xlsx`. Misma estructura
que la minuta IMCE, todo en un solo script: `minuta_remuneraciones.py`.

```
masa salarial real = ocupados (sa INE) × horas habituales semanales (sa STL) × IR real (sa INE)
```

| Carpeta / archivo | Contenido |
|---|---|
| `run_minuta_remuneraciones.py` | Lanzador: corre `Code/minuta_remuneraciones.py` (F5 en Spyder). |
| `Code/` | `minuta_remuneraciones.py`: descarga del INE, cálculos, bases, documento, presentación e `IRL_Base.xlsx`. |
| `Inputs/` | Archivos del INE, bases intermedias (`Masa_salarial_Base.xlsx`, `Cifras_empleo.xlsx`) y `Figuras/` (imágenes de las figuras). |
| `Outputs/` | `IRL_Base.xlsx`, documento `Minuta_Remuneraciones_<mes>_<año>.pdf` y presentación `Presentación_Remuneraciones_<mes>_<año>.pptx` / `.pdf`. En `Outputs/<año>/` quedan las minutas y presentaciones anteriores, con la fecha al inicio del nombre. |

En Spyder: abre `run_minuta_remuneraciones.py` (en la carpeta `remuneraciones`) y presiona F5. Para
usar los archivos del INE que ya están en `Inputs`, sin descargar, pon `SIN_DESCARGA = True` al inicio de
ese archivo. Desde la terminal:

```bash
python remuneraciones/run_minuta_remuneraciones.py
```

## Datos que se descargan del INE

Solo lo que usa la minuta (7 archivos, a `Inputs`):

| Archivo | Uso |
|---|---|
| `tabulado_ir_empalmado.xlsx`, `tabulado_icl_empalmado.xlsx`, `tabulado_ir_real_empalmado.xlsx` | IR, ICL e IR real, series originales (nsa) |
| `tabulado_ir_icl_nominal_general_empalmado_desestacionalizado.xlsx` | IR e ICL desestacionalizados por el INE (sa) |
| `tabulado_ir_real_general_empalmado_desestacionalizado.xlsx` | IR real desestacionalizado por el INE (sa) |
| `ajuste_estacional.xlsx` | Personas ocupadas desestacionalizadas por el INE |
| `horas.xlsx` | Horas habitualmente trabajadas (el INE no las publica desestacionalizadas: se desestacionalizan con STL) |

Los archivos desestacionalizados del INE traen una hoja por publicación (`DDMMAAAA`); se usa la más
reciente. Si alguno no se puede leer, el script usa STL y lo avisa en la consola, en la nota del
Cuadro 1 y en la hoja `Léeme` de `IRL_Base.xlsx`. Las URLs están en `URLS`, al inicio de
`minuta_remuneraciones.py`.

## Documento `Minuta_Remuneraciones_<mes>_<año>.pdf`

- Título con el resultado del mes (salarios reales m/m sa y a/a), fecha y autor (Gerencia de Estudios).
- **Principales resultados**: IR nominal, ICL e IR real en m/m (sa) y a/a, y su variación m/m en la
  serie original (nsa) frente al promedio 2010-año actual del mismo mes y su rango histórico (para ver
  si el dato está en su normalidad); para el IR real, la variación nsa m/m y la a/a se comparan con las
  del mes anterior; inflación implícita y meses consecutivos de alzas; masa salarial real con sus
  contribuciones.
- **Cuadro 1**: m/m (sa), m/m (nsa) y a/a de IR nominal, ICL nominal e IR real (mes anterior y actual),
  con una fila de promedio histórico del mismo mes (2010-año actual) bajo cada índice.
- **Figuras** (también en `Inputs/Figuras`):

| Figura | Contenido |
|---|---|
| 1 | IR real desestacionalizado (nivel) |
| 2 | Variación anual de IR, ICL e IR real, con su promedio histórico (líneas punteadas) |
| 3 | Masa salarial real a/a con contribuciones de ocupados, horas y salario real, y su promedio histórico |
| 4 | Masa salarial real (ocupados × horas × salario real) y horas × salario real, en nivel |
| 5 y 6 | IR nominal m/m por mes calendario, sa y nsa: banda mín-máx y promedio 2010 a tres años antes, dos años anteriores y año en curso |
| 7 y 8 | Lo mismo para el IR real, sa y nsa |
| 9 | Velocidad de los salarios: IR nominal e IR real |
| 10 | Velocidad de la masa salarial real y de horas × salario real |

Velocidad: `TM_t / TM_t-1 × 100 − 100`, donde `TM_t` es el promedio de los índices desestacionalizados de
los últimos 3 meses (trimestre móvil) y `TM_t-1` el del mes anterior; con su promedio histórico en líneas
punteadas. Los gráficos parten cuatro años atrás para que
el Covid no aplaste la escala.

## Presentación `Presentación_Remuneraciones_<mes>_<año>.pptx`

Los mismos contenidos del documento: portada, principales resultados, Cuadro 1 (tabla editable) y las
diez figuras, dos por lámina, como **gráficos nativos de PowerPoint** (editables, con sus datos en
"Editar datos"): los mismos datos, colores, ejes y promedios que las figuras del documento. `formato_graficos.py` se busca como en la minuta IMCE
(o en `FORMATO_GRAFICOS_DIR`); si no se encuentra, se generan igual el documento y `IRL_Base.xlsx`.

## Archivo de minutas anteriores

Al empezar cada corrida, los PDF y PPTX que estén sueltos en `Outputs` se mueven a `Outputs/<año>/` con la
fecha al inicio del nombre (`AAAAMMDD_<nombre>`): la que el nombre ya traiga o, si no tiene, la de su
última modificación. Así en `Outputs` quedan solo la minuta y la presentación del mes, más
`IRL_Base.xlsx`, que no se mueve.

## `Outputs/IRL_Base.xlsx`

- **Datos**: toda la historia disponible: IR, ICL e IR real (nsa y sa), sus variaciones m/m (nsa y sa) y
  a/a, la masa salarial (componentes, índice, a/a, contribuciones y horas × salario real) y, en las
  últimas columnas, las velocidades.
- **Cuadro 1** y **Figura 1** a **Figura 10**: los datos de cada cuadro y figura.
- **Léeme**: contenido y origen de cada serie desestacionalizada.

El script **solo escribe en sus columnas** (desde A1 de cada hoja): las demás columnas, hojas, fórmulas
y gráficos que agregues se conservan. Si el archivo lo guardó Excel y tiene gráficos o tablas, se
actualiza con Excel (`xlwings`); si no hay Excel disponible, con openpyxl y un respaldo en
`Inputs/respaldo`.

## Bases intermedias (en `Inputs`)

- **`Masa_salarial_Base.xlsx`**: masa salarial real (componentes, índice 2023 = 100, a/a,
  contribuciones y horas × salario real), con gráficos de Excel y hoja Léeme.
- **`Cifras_empleo.xlsx`**: hojas ajuste_estacional, horas y remuneraciones (nsa y sa).

`Periodo` es el último mes del trimestre móvil de la ENE (`Nov - Ene` de 2015 = enero 2016); el IR es
mensual y se cruza con ese mes.
