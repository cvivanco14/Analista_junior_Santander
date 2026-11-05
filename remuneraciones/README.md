# Minuta de Remuneraciones (INE)

Descarga las cifras de remuneraciones y empleo del INE, calcula la masa salarial real y arma la minuta:
documento PDF, presentación PPTX/PDF (con `formato_graficos.py`) e `IRL_Base.xlsx`. Misma estructura
que la minuta IMCE: un script de datos y uno de minuta.

```
masa salarial real = ocupados (sa INE) × horas habituales semanales (sa) × IR real (sa INE)
```

| Carpeta | Contenido |
|---|---|
| `Code/` | `datos_remuneraciones.py` (descarga del INE, cálculos y bases) y `minuta_remuneraciones.py` (documento, presentación e `IRL_Base.xlsx`). |
| `Inputs/` | Archivos del INE, bases intermedias (`Masa_salarial_Base.xlsx`, `Cifras_empleo.xlsx`) y `Figuras/` (imágenes de las figuras). |
| `Outputs/` | `IRL_Base.xlsx`, documento `Minuta_Remuneraciones_<mes>_<año>.pdf` y presentación `AAAAMMDD_Minuta_Remuneraciones.pptx` / `.pdf`. |

```bash
cd remuneraciones/Code
python datos_remuneraciones.py      # descarga del INE y actualiza las bases en Inputs
python minuta_remuneraciones.py     # arma la minuta en Outputs
```

En Spyder: corre (F5) primero `datos_remuneraciones.py` y después `minuta_remuneraciones.py`.
`datos_remuneraciones.py --sin-descarga` rehace las bases con los archivos que ya están en `Inputs`.

## Series desestacionalizadas (sa)

Se usan las que publica el INE, no una desestacionalización propia:

- **IR, ICL e IR real**: el script busca la serie desestacionalizada en todas las hojas de los tabulados
  del INE (columna u hoja con "desestacionalizada"). Si el INE la publica en un archivo aparte, agrega
  su URL en `URLS_SA` (al inicio de `datos_remuneraciones.py`) y se descarga y usa.
- **Ocupados**: `ajuste_estacional.xlsx` del INE. **Horas**: la de `ajuste_estacional.xlsx` si la trae;
  si no, STL sobre `horas.xlsx`.
- Si no encuentra la serie del INE, usa STL y avisa en la consola, en la nota del Cuadro 1 y en la
  hoja `Léeme` de `IRL_Base.xlsx` (que dice de dónde salió cada serie).

## Documento `Minuta_Remuneraciones_<mes>_<año>.pdf`

- Título con el resultado del mes (salarios reales m/m sa y a/a), fecha y autor (Gerencia de Estudios).
- **Principales resultados**: IR nominal, ICL e IR real en m/m (sa) y a/a, y su variación m/m en la
  serie original (nsa) frente al promedio 2010-año actual del mismo mes y su rango histórico (para ver
  si el dato está en su normalidad); inflación implícita y meses consecutivos de alzas; masa salarial
  real con sus contribuciones.
- **Cuadro 1**: m/m (sa), m/m (nsa) y a/a de IR nominal, ICL nominal e IR real (mes anterior y actual),
  con una fila de promedio histórico del mismo mes (2010-año actual) bajo cada índice.
- **Figuras** (también en `Inputs/Figuras`):

| Figura | Contenido |
|---|---|
| 1 | IR real desestacionalizado (nivel) |
| 2 | Variación anual de IR, ICL e IR real, con su promedio histórico (líneas punteadas) |
| 3 y 4 | IR nominal m/m por mes calendario, sa y nsa: banda mín-máx y promedio 2010 a tres años antes, dos años anteriores y año en curso |
| 5 y 6 | Lo mismo para el IR real, sa y nsa |
| 7 | Masa salarial real a/a con contribuciones de ocupados, horas y salario real, y su promedio histórico |
| 8 | Masa salarial real (ocupados × horas × salario real) y horas × salario real, en nivel |

## Presentación `AAAAMMDD_Minuta_Remuneraciones.pptx`

Los mismos contenidos del documento: portada, principales resultados, Cuadro 1 (tabla editable) y las
ocho figuras (las mismas imágenes), dos por lámina. `formato_graficos.py` se busca como en la minuta IMCE
(o en `FORMATO_GRAFICOS_DIR`); si no se encuentra, se generan igual el documento y `IRL_Base.xlsx`.

## `Outputs/IRL_Base.xlsx`

- **Datos**: toda la historia disponible: IR, ICL e IR real (nsa y sa), sus variaciones m/m (nsa y sa) y
  a/a, y la masa salarial (componentes, índice, a/a, contribuciones y horas × salario real).
- **Cuadro 1** y **Figura 1** a **Figura 8**: los datos de cada cuadro y figura.
- **Léeme**: contenido y origen de cada serie desestacionalizada.

El script **solo escribe en sus columnas** (desde A1 de cada hoja): las demás columnas, hojas, fórmulas
y gráficos que agregues se conservan. Si el archivo lo guardó Excel y tiene gráficos o tablas, se
actualiza con Excel (`xlwings`); si no hay Excel disponible, con openpyxl y un respaldo en
`Inputs/respaldo`.

## Bases (`datos_remuneraciones.py`, en `Inputs`)

- **`Masa_salarial_Base.xlsx`**: masa salarial real (componentes, índice 2023 = 100, a/a,
  contribuciones y horas × salario real), con gráficos de Excel y hoja Léeme.
- **`Cifras_empleo.xlsx`**: indicadores principales, categoría, informalidad, ajuste estacional,
  comparación ENE-RRAA, horas, rama y remuneraciones (nsa y sa).

`Periodo` es el último mes del trimestre móvil de la ENE (`Nov - Ene` de 2015 = enero 2016); el IR es
mensual y se cruza con ese mes.
