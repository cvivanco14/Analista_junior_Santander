# Minuta IPC

Descarga los índices del INE (canasta 2023=100), calcula los agregados con la matriz de volátiles
del Banco Central y genera la minuta en PDF (destacados, cuadros, gráficos, difusión, CBA y anexo
de 172 subclases).

| Carpeta | Contenido |
|---|---|
| `Code/` | `minuta_ipc.py` (arma la minuta), `ipc.py` (datos y cálculos), `graficos_ipc.py` (gráficos) y `anexo_subclases.py` (anexo). |
| `Inputs/` | Insumos fijos: `matriz_productos.csv`, `historico_bcch_indices.csv`, `pesos_canasta_2018.csv`, `difusion_historica.csv`, `ipc_base2018_alimentos.csv` y `subclases_bcch_1989_2023.csv`. |
| `Inputs/INE/` | Excel descargados del INE (`ipc-base-2023.xlsx`, `ipc-referencial-2023.xlsx`); se reemplazan en cada corrida. |
| `Outputs/` | `Minuta IPC <mes> <año>.pdf`, `IPC_Base.xlsx` (datos de cada gráfico y cuadro), `IPC_Base_Subclases.xlsx` (datos del anexo) y `figuras/` (PNG que se regeneran en cada corrida). |

```bash
cd minuta_ipc/Code
python minuta_ipc.py                 # descarga del INE y genera la minuta del último mes
python minuta_ipc.py --sin-descarga  # usa los Excel ya guardados en Inputs/INE/
python minuta_ipc.py --mes 2026-07   # minuta de un mes anterior
python minuta_ipc.py --sin-anexo     # sin el anexo de subclases (más rápido)
```

- Las rutas se arman desde la ubicación del script (`Path(__file__)`): funciona en cualquier
  computador y también con F5 en Spyder / VS Code.
- Antes de cada minuta, completar `EXPECTATIVAS` (y `COMENTARIO` si corresponde) al inicio de
  `minuta_ipc.py`.
- Si falta algún insumo en `Inputs/`, el script lo avisa con la lista de archivos.
- `IPC_Base.xlsx` se reescribe en cada corrida: hoja `Índice`, `Cuadro 1` y `Cuadro 2` (tal como
  salen en la minuta) y una hoja por gráfico (`G1 Inflación total`, `G2 Agregados mes`, ...), con
  el mismo número que el gráfico en el PDF. Si el archivo está abierto en Excel, avisa y no lo sobrescribe.
- `IPC_Base_Subclases.xlsx` trae los datos de los cuatro gráficos del anexo para las 172 subclases
  (una columna por subclase, en el orden del IPC): `Resumen` (peso, m/m, incidencia, a/a y
  clasificación analítica del mes), `Perfil prom 2013-19` y `Perfil de 2013-19` (promedio y desviación
  estándar por mes calendario), `Perfil <año>` de los últimos años, `Var m-m` y `Var a-a` de los
  últimos 5 años e `Índice 2023=100` desde 2010. Con `--sin-anexo` no se genera.
