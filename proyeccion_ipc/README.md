# Proyección del IPC (modelos, nowcast, Excel y reporte)

Descarga el IPC base 2023 del INE (283 productos y 172 subclases) y las variables macro y de expectativas de la BDE,
proyecta la inflación hasta diciembre de 2027 con dos modelos y un nowcast de combustibles, los evalúa fuera de
muestra y genera un Excel autocontenido con fórmulas y un reporte HTML (formato de `Apuntes_IPC_Chile.html`).

| Carpeta | Contenido |
|---|---|
| `Code/` | `proyeccion_ipc.py`: un solo script con todo el flujo (descargas, base, modelos, nowcast, evaluación, Excel y reporte). |
| `Inputs/` | Insumos fijos: `IPC_2023_Automatizado_v2.xlsx` (historia 2018-2022 de productos y 1998-2017 de bloques), `matriz_productos.csv` (clasificación sin volátiles / volátiles), `difusion_historica.csv` y `Apuntes_IPC_Chile.html` (plantilla del reporte). |
| `Inputs/INE/` | Excel del INE (`ipc_base2023.xlsx`, `ipc_referencial_2023.xlsx`, `analiticos_base2023.xlsx`); se reemplazan en cada corrida. |
| `Inputs/BDE/` | Series de la BDE (`bde_mensual.csv`, `bde_diario.csv`, `catalogo_bde.csv`); se reemplazan en cada corrida. |
| `Outputs/` | `IPC_Chile_Proyecciones.xlsx`, `Proyeccion_IPC_Chile_modelos.html` y `tablas/` (CSV). |

```bash
cd proyeccion_ipc/Code
python proyeccion_ipc.py                 # descarga BDE e INE, modela y genera las salidas (~2 min)
python proyeccion_ipc.py --sin-descarga  # usa lo ya descargado en Inputs/INE e Inputs/BDE (~30 s)
python proyeccion_ipc.py --sin-reporte   # sin el HTML (no requiere pandoc)
```

En Spyder: abrir `proyeccion_ipc.py` y correr con F5. Las rutas se arman desde la ubicación del script
(`Path(__file__)`), así que funciona en cualquier computador.

## Credenciales del BCCh

Igual que en la minuta IMCE: una vez por computador y nunca dentro del script. API Key Token:

- macOS (consola de Spyder; queda en el Llavero):
  `import keyring, getpass; keyring.set_password('estudios_santander', 'BDE_TOKEN', getpass.getpass('Token BCCh: '))`
- Windows, Anaconda Prompt o cmd: `setx BDE_TOKEN "tu_token"` (PowerShell: `setx BDE_TOKEN 'tu_token'`); después cerrar y
  abrir Spyder / VS Code.

También acepta la variable `BCCH_API_token`, o usuario y clave de la API en `BCCH_USER` y `BCCH_PASS`. Si el BCCh
rechaza la credencial, el script dice de dónde la tomó y cómo corregirla, sin mostrarla.

## Modelos

- **Modelo A (agregados, recomendado como central)**: 11 bloques de la partición sin volátiles / volátiles
  (Carlomagno, Fornero y Sansone), con cuatro ecuaciones por bloque ancladas a las expectativas: estacional, AR, con
  impulsores macro, y con corrección de error frente a las tendencias comunes detectadas por pares (Carlomagno y
  Espasa). Se combinan por el error fuera de muestra.
- **Modelo B (subclases)**: una ecuación por cada una de las 172 subclases, con corrección de error frente a su bloque
  cuando el test lo acepta. Es la base editable del Excel.
- **Nowcast**: precio diario de gasolina y diésel reconstruido con los informes semanales de ENAP (tabla
  `ENAP_INFORMES` al inicio del script; agregar cada informe nuevo) y escenarios para los próximos ajustes del MEPCO
  (`ENAP_ESCENARIOS`).

## `Outputs/IPC_Chile_Proyecciones.xlsx`

Hojas agregadas al estilo del Excel automatizado:

- `Resumen_Proyecciones`: var. mensual y en 12 meses 2025-2027 de las 20 clases del BCCh (IPC, sin volátiles,
  bienes y servicios SV con sus subclases, volátiles con energía, electricidad, combustibles, alimentos, FyV y resto),
  más alimentos, energía y SAE del INE; memo con el Modelo A y la EEE, y dic/dic y promedios anuales.
- `Perfiles_Analiticos`: m/m por mes calendario 2010-2027 de cada clase, con promedio 2010-2019 y 2013-2019,
  máximo, mínimo y rango. Hasta 2023 usa las series empalmadas de la hoja `Historia` del Excel automatizado; desde
  2024 se enlaza a `Var_mm`, así que la proyección se actualiza sola.
- `Productos`: los 283 productos con su historia y la proyección = índice del último mes × variación de su
  subclase (hoja `Indices`); la última fila recalcula el IPC desde los productos como control.

Libro autocontenido con fórmulas: `Matriz` (283 productos) → `Subclases` (pesos con SUMIF) → `Indices`
(historia + proyección = índice anterior × (1 + `Proy_mm`)) → `Agregados`, `Var_mm`, `Var_12m`, `Incidencias`,
`Difusion` y `Resumen`. Las celdas amarillas de `Proy_mm` y `Nowcast_ENAP` son supuestos editables; gasolina y diésel
de nov-dic están enlazados al nowcast. Si hay LibreOffice instalado, el script deja el libro recalculado; Excel lo
recalcula al abrirlo.
