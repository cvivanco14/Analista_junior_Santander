# Masa salarial real (INE)

Descarga las cifras de empleo y remuneraciones del INE, calcula la masa salarial real y deja sus
series y gráficos en `Outputs/Masa_salarial_Base.xlsx`.

```
masa salarial real = ocupados (desest. INE) × horas habituales semanales (desest. STL) × IR real (desest. STL)
```

| Carpeta | Contenido |
|---|---|
| `Code/` | `datos_empleo.py` (descarga del INE y hojas de empleo) y `masa_salarial.py` (cálculo, Excel y gráfico). |
| `Inputs/` | Archivos descargados del INE (`ajuste_estacional.xlsx`, `horas.xlsx`, `tabulado_ir_real_empalmado.xlsx`, etc.). |
| `Outputs/` | `Masa_salarial_Base.xlsx`, `Masa_salarial.png` y `Cifras_empleo.xlsx`. |

```bash
cd masa_salarial/Code
python masa_salarial.py                 # descarga del INE y genera todo en Outputs
python masa_salarial.py --sin-descarga  # usa los archivos que ya están en Inputs
python datos_empleo.py                  # solo las cifras de empleo (Outputs/Cifras_empleo.xlsx)
```

En Spyder: abre `Code/masa_salarial.py` y córrelo con F5.

## `Masa_salarial_Base.xlsx`

- **Masa salarial**: por mes, ocupados SA (miles), horas habituales SA, IR real SA (2023 = 100),
  masa salarial real (índice 2023 = 100) y su variación a/a, contribuciones a esa variación de
  ocupados, horas y salario real (descomposición logarítmica; suman la variación total) y la versión
  sin horas (ocupados × IR real).
- **Gráficos**: gráficos de Excel (editables) desde 2015: variación a/a con contribuciones apiladas,
  y el nivel de las dos medidas.
- **Léeme**: definiciones y fuentes.

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
