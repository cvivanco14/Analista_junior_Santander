# Minuta de Remuneraciones (INE)

Descarga las cifras de remuneraciones y empleo del INE, calcula la masa salarial real y arma la minuta:
documento PDF, presentación PPTX/PDF (con `formato_graficos.py`) e `IRL_Base.xlsx`. Misma estructura
que la minuta IMCE, todo en un solo script: `minuta_remuneraciones.py`.

```
masa salarial real = ocupados (sa INE) × horas habituales semanales (sa STL) × IR real (sa INE)
```

| Carpeta | Contenido |
|---|---|
| `Code/` | `minuta_remuneraciones.py`: descarga del INE, cálculos, bases, documento, presentación e `IRL_Base.xlsx`. |
| `Inputs/` | Archivos del INE, bases intermedias (`Masa_salarial_Base.xlsx`, `Cifras_empleo.xlsx`) y `Figuras/` (imágenes de las figuras). |
| `Outputs/` | `IRL_Base.xlsx`, documento `Minuta_Remuneraciones_<mes>_<año>.pdf` y presentación `AAAAMMDD_Minuta_Remuneraciones.pptx` / `.pdf`. |

```bash
cd remuneraciones/Code
python minuta_remuneraciones.py                 # descarga del INE y arma todo
python minuta_remuneraciones.py --sin-descarga  # usa los archivos que ya están en Inputs
```

En Spyder: abre `minuta_remuneraciones.py` y córrelo con F5.

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
  si el dato está en su normalidad); inflación implícita y meses consecutivos de alzas; masa salarial
  real con sus contribuciones.
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

## Bases intermedias (en `Inputs`)

- **`Masa_salarial_Base.xlsx`**: masa salarial real (componentes, índice 2023 = 100, a/a,
  contribuciones y horas × salario real), con gráficos de Excel y hoja Léeme.
- **`Cifras_empleo.xlsx`**: hojas ajuste_estacional, horas y remuneraciones (nsa y sa).

`Periodo` es el último mes del trimestre móvil de la ENE (`Nov - Ene` de 2015 = enero 2016); el IR es
mensual y se cruza con ese mes.
