"""Gráficos de dispersión de expectativas de inflación (EEE, Banco Central de Chile).

Lee las planillas mensuales de la EEE (eee/raw/ y cualquier carpeta extra que se
pase como argumento), extrae Mediana, Decil 1 y Decil 9 de inflación en 11 y 23
meses, guarda la serie en eee/output/eee_inflacion_deciles.csv y genera:

    eee/output/eee_inflacion_11m.png
    eee/output/eee_inflacion_23m.png

Uso:
    python eee/graficos.py [carpeta_extra ...]
"""
import re
import sys
import warnings
from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import pandas as pd

BASE = Path(__file__).resolve().parent
RAW = BASE / "raw"
OUT = BASE / "output"
DESDE = "2014-01-01"

MESES = {m: i for i, m in enumerate(
    ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
     "septiembre", "octubre", "noviembre", "diciembre"], start=1)}

HORIZONTES = {"11m": r"en 11 meses", "23m": r"en 23 meses"}

COLORES = {"d9": "#E3262B", "mediana": "#8DC63F", "d1": "#1D2A6B", "rango": "#E8C766", "titulo": "#C2A15A"}


def fecha_encuesta(path):
    """Mes de la encuesta a partir del nombre del archivo."""
    n = path.stem.lower()
    m = re.search(r"datos[ _+]mensuales[ _+](\w+)[ _+](\d{4})", n)
    if m:
        return pd.Timestamp(int(m.group(2)), MESES[m.group(1)], 1)
    m = re.search(r"resultado_?(\d{1,2})(\d{4})$", n)
    if m:
        return pd.Timestamp(int(m.group(2)), int(m.group(1)), 1)
    return None


def leer_planilla(path):
    """Devuelve {horizonte: (mediana, decil1, decil9)} de la hoja resumen."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df = pd.read_excel(path, sheet_name=0, header=None)
    res = {}
    for _, fila in df.iterrows():
        celdas = list(fila)
        for i, c in enumerate(celdas):
            if not isinstance(c, str):
                continue
            etiqueta = c.replace("\xa0", " ").lower()
            for h, patron in HORIZONTES.items():
                # La primera aparición corresponde a la sección "Inflación" (IPC total).
                if h in res or not re.search(patron, etiqueta):
                    continue
                nums = [v for v in celdas[i + 1:] if isinstance(v, (int, float)) and not pd.isna(v)]
                if len(nums) >= 3:
                    res[h] = tuple(float(v) for v in nums[:3])
    return res


def construir_serie(carpetas):
    filas = {}
    for carpeta in carpetas:
        for path in sorted(Path(carpeta).glob("*.xls*")):
            fecha = fecha_encuesta(path)
            if fecha is None:
                continue
            datos = leer_planilla(path)
            for h, (med, d1, d9) in datos.items():
                if not d1 <= med <= d9:
                    print(f"aviso: orden inesperado en {path.name} {h}: {med}, {d1}, {d9}")
                filas[(fecha, h)] = {"fecha": fecha, "horizonte": h, "mediana": med,
                                     "decil_1": d1, "decil_9": d9, "archivo": path.name}
            if len(datos) < len(HORIZONTES):
                print(f"aviso: {path.name} sin datos para {set(HORIZONTES) - set(datos)}")
    serie = pd.DataFrame(filas.values()).sort_values(["horizonte", "fecha"])
    serie = serie[serie["fecha"] >= DESDE]
    serie[["mediana", "decil_1", "decil_9"]] = serie[["mediana", "decil_1", "decil_9"]].round(2)
    serie["rango_9_1"] = (serie["decil_9"] - serie["decil_1"]).round(2)
    return serie


def graficar(serie, horizonte, titulo, archivo):
    s = serie[serie["horizonte"] == horizonte].set_index("fecha")
    tope = max(6, float(s["decil_9"].max()))
    tope = 2 * int(tope / 2 + 1) if tope > 6 else 6

    fig, ax = plt.subplots(figsize=(6.4, 5.2), dpi=200)
    ax2 = ax.twinx()
    ax2.bar(s.index, s["rango_9_1"], width=25, color=COLORES["rango"], label="Decil 9 - 1", zorder=1)
    ax.plot(s.index, s["decil_9"], color=COLORES["d9"], lw=1.6, label="Decil 9", zorder=3)
    ax.plot(s.index, s["mediana"], color=COLORES["mediana"], lw=1.6, label="Mediana", zorder=3)
    ax.plot(s.index, s["decil_1"], color=COLORES["d1"], lw=1.6, label="Decil 1", zorder=3)

    # Líneas por encima de las barras.
    ax.set_zorder(ax2.get_zorder() + 1)
    ax.patch.set_visible(False)

    for a in (ax, ax2):
        a.set_ylim(0, tope)
        a.tick_params(labelsize=9, length=0)
        for lado in ("top", "left", "right"):
            a.spines[lado].set_visible(False)
        a.spines["bottom"].set_color("#7F7F7F")
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%y"))
    ax.set_xlim(s.index.min() - pd.Timedelta(days=20), s.index.max() + pd.Timedelta(days=20))

    fig.suptitle(titulo, color=COLORES["titulo"], fontsize=14, y=0.97)
    ax.set_title("(variación anual, porcentaje)", fontsize=10.5, pad=10)
    lineas = ax2.get_legend_handles_labels()[0] + ax.get_legend_handles_labels()[0]
    orden = [0, 2, 3, 1]  # Decil 9 - 1, Mediana, Decil 1, Decil 9
    fig.legend([lineas[i] for i in orden], ["Decil 9 - 1", "Mediana", "Decil 1", "Decil 9"],
               loc="lower center", ncol=4, frameon=False, fontsize=9)
    ultima = s.index.max()
    fig.text(0.01, 0.01, f"Fuente: Banco Central de Chile, EEE. Último dato: {ultima:%m-%Y}.",
             fontsize=7, color="#595959")
    fig.tight_layout(rect=(0, 0.06, 1, 0.95))
    fig.savefig(archivo)
    plt.close(fig)


def main():
    carpetas = [RAW] + [Path(c) for c in sys.argv[1:]]
    OUT.mkdir(exist_ok=True)
    serie = construir_serie(carpetas)
    serie.to_csv(OUT / "eee_inflacion_deciles.csv", index=False, date_format="%Y-%m")
    graficar(serie, "11m", "EEE: Inflación en 11 meses", OUT / "eee_inflacion_11m.png")
    graficar(serie, "23m", "EEE: Inflación en 23 meses", OUT / "eee_inflacion_23m.png")
    print(serie.groupby("horizonte")["fecha"].agg(["min", "max", "count"]))


if __name__ == "__main__":
    main()
