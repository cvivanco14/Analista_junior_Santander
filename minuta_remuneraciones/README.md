# Minuta Remuneraciones

Minuta mensual en PDF con los índices de remuneraciones del INE (IR, ICL e IR real, base anual
2023=100) y las horas ordinarias y extraordinarias, con textos automáticos a partir del último dato
y en colores Santander (títulos, cuadro y pie en rojo Santander, logo si se encuentra el Manual de
Estilo). Misma estructura que el chartpack de Expectativas:

| Carpeta | Contenido |
|---|---|
| `Code/` | `datos_remuneraciones.py` (descarga y lectura de los tabulados del INE) y `minuta_remuneraciones.py` (arma la minuta). |
| `Inputs/` | Tabulados del INE (`tabulado_*.xlsx`), los deja `datos_remuneraciones.py`. |
| `Outputs/` | `AAAAMMDD_Minuta_Remuneraciones.pdf` y `figuras/` (PNG de las cuatro figuras). |

## Uso (no hay que editar nada)

```bash
cd minuta_remuneraciones/Code
python datos_remuneraciones.py               # descarga los tabulados del INE a Inputs
python minuta_remuneraciones.py              # genera la minuta del último mes en Outputs
python minuta_remuneraciones.py --mes 2026-06  # minuta de un mes anterior
```

- Los tabulados se reemplazan en cada descarga (el INE los publica siempre con el mismo nombre). Si
  alguno falla, se usa el que ya estaba en `Inputs`.
- El logo (`assets/logo.png`) se busca como `formato_graficos.py` en el chartpack: variable
  `FORMATO_GRAFICOS_DIR`, subiendo desde `Code` hasta `Administrativos\Manual de Estilo`, o en la
  carpeta sincronizada de SharePoint. Si no lo encuentra, la minuta sale sin logo.
- Requiere `pandas`, `openpyxl`, `requests`, `matplotlib` y `reportlab`.
