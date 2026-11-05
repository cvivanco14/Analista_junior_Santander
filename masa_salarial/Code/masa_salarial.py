# -*- coding: utf-8 -*-
r"""Masa salarial real: ocupados × horas × remuneración real (INE).

    masa salarial real = ocupados (desestacionalizados por el INE)
                         × horas habitualmente trabajadas a la semana (desestacionalizadas con STL)
                         × índice de remuneraciones real (desestacionalizado con STL)

Igual que la celda "9) Masa salarial" del notebook "Series empleo3.ipynb", pero con las series
calculadas en la misma corrida (antes se leían de cifras_empleo.xlsx de la corrida anterior) y
buscando las columnas por nombre (antes por posición).

Salidas en Outputs\:
    Masa_salarial_Base.xlsx   hojas 'Masa salarial' (series), 'Gráficos' (gráficos de Excel) y 'Léeme'
    Masa_salarial.png         gráfico: variación a/a y contribuciones de ocupados, horas y salario real
    Cifras_empleo.xlsx        todas las hojas de empleo del INE (las de datos_empleo.py)

Uso (no hay que editar nada):
    python masa_salarial.py                 # descarga del INE y genera las salidas
    python masa_salarial.py --sin-descarga  # usa los archivos que ya están en Inputs
"""
import argparse
import sys

import numpy as np
import pandas as pd

import datos_empleo as de

OUT = de.OUT
EXCEL = OUT / "Masa_salarial_Base.xlsx"
PNG = OUT / "Masa_salarial.png"
DESDE_GRAFICO = "2015-01-01"
ANIO_BASE = 2023          # base de los índices (como el IR del INE: base anual 2023 = 100)
FUENTE = "Fuente: INE (Encuesta Nacional de Empleo e Índice de Remuneraciones)."

OCUPADOS = "Ocupados SA (miles)"
HORAS = "Horas habituales semanales SA"
IR_REAL = f"IR real SA ({ANIO_BASE}=100)"
MASA = f"Masa salarial real ({ANIO_BASE}=100)"
MASA_AA = "Masa salarial real a/a (%)"
CONTRIB = {OCUPADOS: "Contribución ocupados (pp)", HORAS: "Contribución horas (pp)",
           IR_REAL: "Contribución salario real (pp)"}
MASA_SH = f"Masa salarial real sin horas ({ANIO_BASE}=100)"
MASA_SH_AA = "Masa salarial real sin horas a/a (%)"

# colores Santander (rojo, gris oscuro, gris claro) y negro para el total
ROJO, GRIS, GRIS_CLARO, NEGRO = "EC0000", "6F7779", "BFC5C7", "1D1D1B"


# ─────────────────────────────────────────────────────────────────────────────
# Series
# ─────────────────────────────────────────────────────────────────────────────

def _columna(df, *palabras, excluir=("tasa", "variación", "variacion"), respaldo=None, hoja=""):
    """Primera columna cuyo nombre contiene todas las 'palabras' (sin mayúsculas) y ninguna de 'excluir'."""
    for c in df.columns:
        n = c.lower()
        if all(p in n for p in palabras) and not any(e in n for e in excluir):
            return c
    if respaldo is not None and respaldo < len(df.columns):
        c = df.columns[respaldo]
        print(f"  ⚠️ En '{hoja}' no encontré {palabras}; uso la columna {respaldo}: '{c}'")
        return c
    raise KeyError(f"En '{hoja}' no encontré una columna con {palabras}. Columnas: {list(df.columns)}")


def _indice(s, anio=ANIO_BASE):
    return 100 * s / s[s.index.year == anio].mean()


def _aa(s):
    return 100 * (s / s.shift(12) - 1)


def masa_salarial(datos):
    """DataFrame mensual (índice = Periodo) con componentes, masa salarial, variación a/a y contribuciones."""
    ae = datos["ajuste_estacional"].set_index("Periodo")
    hr = datos["horas"].set_index("Periodo")
    rem = datos["remuneraciones"].set_index("Periodo")
    # la columna 5 es la que usaba el notebook (usecols=[0, 5] de la hoja ajuste_estacional)
    col_ocu = _columna(ae, "ocupad", respaldo=4, hoja="ajuste_estacional")
    col_hrs = de.COL_HORAS_SA if de.COL_HORAS_SA in hr else _columna(hr, "desest", hoja="horas")
    col_ir = "IR Real (Desest.)" if "IR Real (Desest.)" in rem else _columna(rem, "real", "desest",
                                                                                 hoja="remuneraciones")
    print(f"  Ocupados: '{col_ocu}'\n  Horas:    '{col_hrs}'\n  IR real:  '{col_ir}'")

    df = pd.concat({OCUPADOS: ae[col_ocu], HORAS: hr[col_hrs], IR_REAL: rem[col_ir]}, axis=1, join="inner")
    df = df.apply(pd.to_numeric, errors="coerce").dropna().sort_index()
    df.index.name = "Periodo"
    df[IR_REAL] = _indice(df[IR_REAL])

    df[MASA] = _indice(df[OCUPADOS] * df[HORAS] * df[IR_REAL])
    df[MASA_AA] = _aa(df[MASA])
    # contribuciones a la variación a/a (descomposición logarítmica: suman exactamente la variación)
    log_total = np.log(df[MASA] / df[MASA].shift(12))
    escala = (df[MASA_AA] / log_total).where(log_total.abs() > 1e-12, 100.0)
    for comp, nombre in CONTRIB.items():
        df[nombre] = np.log(df[comp] / df[comp].shift(12)) * escala
    df[MASA_SH] = _indice(df[OCUPADOS] * df[IR_REAL])
    df[MASA_SH_AA] = _aa(df[MASA_SH])
    return df


# ─────────────────────────────────────────────────────────────────────────────
# Gráfico PNG
# ─────────────────────────────────────────────────────────────────────────────

def grafico_png(df, ruta=PNG, desde=DESDE_GRAFICO):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.dates as mdates
    import matplotlib.pyplot as plt

    d = df[df.index >= desde].dropna(subset=[MASA_AA])
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5.2), dpi=150)
    colores = ["#" + c for c in (ROJO, GRIS, GRIS_CLARO)]

    # 1) variación a/a con contribuciones apiladas
    pos = np.zeros(len(d))
    neg = np.zeros(len(d))
    for (comp, nombre), color in zip(CONTRIB.items(), colores):
        v = d[nombre].to_numpy()
        base = np.where(v >= 0, pos, neg)
        ax1.bar(d.index, v, bottom=base, width=25, color=color, label=nombre.replace(" (pp)", ""))
        pos, neg = pos + np.clip(v, 0, None), neg + np.clip(v, None, 0)
    ax1.plot(d.index, d[MASA_AA], color="#" + NEGRO, lw=1.8, label="Masa salarial real")
    ax1.axhline(0, color="#" + NEGRO, lw=0.6)
    ultimo = d.index[-1]
    ax1.annotate(f"{d[MASA_AA].iloc[-1]:.1f}".replace(".", ","), (ultimo, d[MASA_AA].iloc[-1]),
                 xytext=(6, 0), textcoords="offset points", va="center", fontsize=9, fontweight="bold")
    ax1.set_title("Masa salarial real: variación anual y contribuciones", loc="left", fontsize=12,
                  fontweight="bold")
    ax1.set_ylabel("(Var. a/a %, contribuciones en pp)")

    # 2) nivel (índice)
    ax2.plot(d.index, d[MASA], color="#" + ROJO, lw=2, label="Ocupados × horas × salario real")
    ax2.plot(d.index, d[MASA_SH], color="#" + GRIS, lw=1.6, label="Ocupados × salario real")
    ax2.axhline(100, color="#" + GRIS_CLARO, lw=0.8, ls="--")
    ax2.set_title("Masa salarial real: nivel", loc="left", fontsize=12, fontweight="bold")
    ax2.set_ylabel(f"(Índice {ANIO_BASE} = 100, desestacionalizado)")

    for ax in (ax1, ax2):
        ax.xaxis.set_major_locator(mdates.YearLocator(2 if len(d) > 100 else 1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
        ax.spines[["top", "right"]].set_visible(False)
        ax.grid(axis="y", color="#E5E5E5", lw=0.6)
        ax.set_axisbelow(True)
        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=2, frameon=False, fontsize=8.5)
    mes = ultimo.strftime("%m-%Y")
    fig.text(0.01, 0.005, f"{FUENTE} Último dato: trimestre móvil terminado en {mes}. Ocupados desestacionalizados "
             "por el INE; horas e IR real desestacionalizados con STL.", fontsize=7.5, color="#" + GRIS)
    fig.tight_layout(rect=(0, 0.03, 1, 1))
    fig.savefig(ruta)
    plt.close(fig)
    return ruta


# ─────────────────────────────────────────────────────────────────────────────
# Excel
# ─────────────────────────────────────────────────────────────────────────────

def guardar_base(df, ruta=EXCEL, desde=DESDE_GRAFICO):
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.chart.series import SeriesLabel
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    tabla = df.reset_index()
    leeme = pd.DataFrame({"Masa salarial real": [
        "Masa salarial real = ocupados × horas habituales semanales × índice de remuneraciones real.",
        "Ocupados: ENE, cifras desestacionalizadas por el INE (archivo ajuste_estacional.xlsx), en miles.",
        "Horas: promedio de horas habitualmente trabajadas a la semana (horas.xlsx), desestacionalizado con STL.",
        "IR real: índice de remuneraciones real empalmado (base anual 2023 = 100), desestacionalizado con STL.",
        "Periodo: último mes del trimestre móvil de la ENE (el IR es mensual y se cruza con ese mes).",
        f"Índices: promedio {ANIO_BASE} = 100. Variación a/a: respecto del mismo mes del año anterior.",
        "Contribuciones: descomposición logarítmica de la variación a/a; las tres suman la variación total.",
        "Sin horas: ocupados × IR real, la medida más usada cuando no se ajusta por jornada.",
        FUENTE,
        "Se reescribe en cada corrida con masa_salarial.py (proyecto masa_salarial del repositorio).",
    ]})
    try:
        with pd.ExcelWriter(ruta, engine="openpyxl") as xl:
            tabla.to_excel(xl, sheet_name="Masa salarial", index=False)
            leeme.to_excel(xl, sheet_name="Léeme", index=False)
            ws = xl.sheets["Masa salarial"]
            for celda in ws[1]:
                celda.font = Font(bold=True)
                celda.fill = PatternFill("solid", start_color="CCEAFF")
                celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.row_dimensions[1].height = 45
            ws.freeze_panes = "B2"
            ws.column_dimensions["A"].width = 11
            for j in range(2, ws.max_column + 1):
                ws.column_dimensions[get_column_letter(j)].width = 15
            formatos = {OCUPADOS: "#,##0.0", HORAS: "0.00"}
            for j, col in enumerate(tabla.columns, start=1):
                fmt = "mmm-yyyy" if j == 1 else formatos.get(col, "0.0")
                for (c,) in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                    c.number_format = fmt
            xl.sheets["Léeme"].column_dimensions["A"].width = 110

            # gráficos de Excel (nativos), desde DESDE_GRAFICO hasta el último mes
            gs = xl.book.create_sheet("Gráficos", 1)
            ini = 2 + int((tabla["Periodo"] < pd.Timestamp(desde)).sum())
            fin = len(tabla) + 1
            col = {c: j for j, c in enumerate(tabla.columns, start=1)}
            fechas = Reference(ws, min_col=1, min_row=ini, max_row=fin)

            barras = BarChart()
            barras.type, barras.grouping, barras.overlap, barras.gapWidth = "col", "stacked", 100, 30
            for nombre, color in zip(CONTRIB.values(), (ROJO, GRIS, GRIS_CLARO)):
                barras.add_data(Reference(ws, min_col=col[nombre], min_row=ini, max_row=fin),
                                titles_from_data=False)
                s = barras.series[-1]
                s.tx = SeriesLabel(v=nombre.replace(" (pp)", ""))
                s.graphicalProperties = GraphicalProperties(solidFill=color)
                s.graphicalProperties.line.noFill = True
            linea = LineChart()
            linea.add_data(Reference(ws, min_col=col[MASA_AA], min_row=ini, max_row=fin), titles_from_data=False)
            linea.series[0].tx = SeriesLabel(v="Masa salarial real")
            linea.series[0].graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=NEGRO, w=22000))
            linea.series[0].smooth = False
            barras.set_categories(fechas)
            barras += linea
            barras.title = "Masa salarial real: variación anual y contribuciones"
            barras.y_axis.title = "Var. a/a %, contribuciones en pp"
            barras.x_axis.number_format = "yyyy"
            barras.x_axis.tickLblSkip = 12
            barras.x_axis.tickMarkSkip = 12
            barras.x_axis.tickLblPos = "low"
            barras.x_axis.delete = barras.y_axis.delete = False
            barras.legend.position = "b"
            barras.height, barras.width = 10, 24
            gs.add_chart(barras, "A1")

            nivel = LineChart()
            for nombre, color, ancho in ((MASA, ROJO, 25000), (MASA_SH, GRIS, 19000)):
                nivel.add_data(Reference(ws, min_col=col[nombre], min_row=ini, max_row=fin), titles_from_data=False)
                s = nivel.series[-1]
                s.tx = SeriesLabel(v=nombre)
                s.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=color, w=ancho))
                s.smooth = False
            nivel.set_categories(fechas)
            tramo = tabla.iloc[ini - 2:][[MASA, MASA_SH]]
            nivel.y_axis.scaling.min = 10 * np.floor(np.nanmin(tramo.to_numpy()) / 10)
            nivel.y_axis.scaling.max = 10 * np.ceil(np.nanmax(tramo.to_numpy()) / 10)
            nivel.title = "Masa salarial real: nivel"
            nivel.y_axis.title = f"Índice {ANIO_BASE} = 100, desestacionalizado"
            nivel.x_axis.number_format = "yyyy"
            nivel.x_axis.tickLblSkip = 12
            nivel.x_axis.tickMarkSkip = 12
            nivel.x_axis.delete = nivel.y_axis.delete = False
            nivel.legend.position = "b"
            nivel.height, nivel.width = 10, 24
            gs.add_chart(nivel, "A22")
            gs["A43"] = FUENTE
    except PermissionError:
        sys.exit(f"No pude guardar {ruta}: está abierto. Ciérralo y vuelve a correr.")
    return ruta


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--sin-descarga", action="store_true", help="usa los archivos que ya están en Inputs")
    args = p.parse_args()
    print("=" * 60 + "\n  Masa salarial real (INE)\n" + "=" * 60)
    datos = de.cargar(descarga=not args.sin_descarga)
    faltan = [h for h in ("ajuste_estacional", "horas", "remuneraciones") if h not in datos]
    if faltan:
        sys.exit(f"Faltan datos para la masa salarial: {', '.join(faltan)}")
    OUT.mkdir(parents=True, exist_ok=True)
    print(f"  ✅ Cifras de empleo: {de.guardar_excel(datos, de.EXCEL_CIFRAS)}")
    ms = masa_salarial(datos)
    t = ms.index[-1]
    print(f"  Último dato: {t:%m-%Y}  masa salarial real {ms[MASA_AA].iloc[-1]:+.1f}% a/a "
          f"(ocupados {ms[CONTRIB[OCUPADOS]].iloc[-1]:+.1f} pp, horas {ms[CONTRIB[HORAS]].iloc[-1]:+.1f} pp, "
          f"salario real {ms[CONTRIB[IR_REAL]].iloc[-1]:+.1f} pp)")
    print(f"  ✅ Excel: {guardar_base(ms)}")
    print(f"  ✅ Gráfico: {grafico_png(ms)}")
