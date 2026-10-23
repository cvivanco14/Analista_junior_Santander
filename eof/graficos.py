"""EOF: porcentaje de operadores con inflación a 2 años sobre/bajo umbrales.

Lee la hoja "Distribución" de cada planilla de la Encuesta de Operadores
Financieros (eof/raw/ y carpetas extra pasadas como argumento) y, para la
pregunta de inflación en los siguientes 12 meses (13 al 24), cuenta cuántos
operadores responden:

    > 3,0%   > 3,2%   < 2,8%

Guarda la serie en eof/output/eof_inflacion_2a_umbrales.csv y un gráfico por umbral:

    eof/output/eof_sobre_3_0.png
    eof/output/eof_sobre_3_2.png
    eof/output/eof_bajo_2_8.png

Uso:
    python eof/graficos.py [carpeta_extra ...]
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

MESES = {m: i for i, m in enumerate(
    ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
     "septiembre", "octubre", "noviembre", "diciembre"], start=1)}

# "Siguientes 12 meses (13 al 24)" desde 2018; "IPC (13 a 24 Meses)" antes.
PREGUNTA = re.compile(r"13\s*al?\s*24", re.I)

UMBRALES = [  # (columna, título, función sobre el nivel en %, archivo)
    ("sobre_3_0", "sobre 3,0%", lambda x: x > 3.0, "eof_sobre_3_0.png"),
    ("sobre_3_2", "sobre 3,2%", lambda x: x > 3.2, "eof_sobre_3_2.png"),
    ("bajo_2_8", "bajo 2,8%", lambda x: x < 2.8, "eof_bajo_2_8.png"),
]

COLORES = {"serie": "#1D2A6B", "relleno": "#E8C766", "titulo": "#C2A15A", "ref": "#7F7F7F"}


def es_numero(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not pd.isna(v)


def texto(v):
    return str(v).replace("\xa0", " ").strip().lower() if isinstance(v, str) else ""


def tipo_encuesta(nombre):
    n = nombre.lower()
    if "post" in n:
        return "Post RPM"
    if "pre" in n:
        return "Pre RPM"
    return "Quincenal"


def fecha_desde_nombre(nombre):
    n = nombre.lower().replace("_", " ")
    m = re.search(r"(\d{4})(\d{2})q([12])", n)  # EOFRESULTADOS_201601q1
    if m:
        return pd.Timestamp(int(m.group(1)), int(m.group(2)), 8 if m.group(3) == "1" else 22)
    m = re.search(r"rpm (\d{4})(\d{2})\b", n)  # EOF_RESULTADOS_Pre_RPM_201905
    if m:
        return pd.Timestamp(int(m.group(1)), int(m.group(2)), 15)
    m = re.search(r"(" + "|".join(MESES) + r") (\d{4})", n)  # ... Pre RPM Octubre 2026
    if m:
        return pd.Timestamp(int(m.group(2)), MESES[m.group(1)], 15)
    m = re.search(r"(\d{2})(\d{2})(\d{4})", n)  # EOFRESULTADOS22122015
    if m:
        return pd.Timestamp(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    return None


def fecha_encuesta(hojas, nombre):
    """Fecha de cierre: la última fecha del encabezado de la hoja resumen."""
    resumen = next(iter(hojas.values()))
    fechas = [v for v in resumen.iloc[:7].to_numpy().ravel() if isinstance(v, (pd.Timestamp,))
              or hasattr(v, "year") and hasattr(v, "month") and not isinstance(v, str)]
    fechas = [pd.Timestamp(f) for f in fechas if 2000 < pd.Timestamp(f).year < 2100]
    return max(fechas) if fechas else fecha_desde_nombre(nombre)


def leer_distribucion(df):
    """Devuelve [(nivel en %, n° respuestas)] de la pregunta 13 al 24 meses."""
    for (r, c), v in df.stack().items():
        if PREGUNTA.search(texto(v)):
            break
    else:
        return []
    if texto(df.iat[r + 1, c]).startswith("nivel"):
        # Layout por pares de columnas: Nivel | N° R bajo cada pregunta.
        col_nivel, col_n, fila = c, c + 1, r + 2
    else:
        # Layout con una columna "Nivel" compartida a la izquierda de cada bloque.
        col_nivel = max(j for j in range(c) if texto(df.iat[r, j]).startswith("nivel"))
        col_n, fila = c, r + 1
    dist = []
    for i in range(fila, len(df)):
        if any("resp" in texto(x) for x in df.iloc[i, :col_nivel + 1]):
            break
        nivel, n = df.iat[i, col_nivel], df.iat[i, col_n]
        if es_numero(nivel) and es_numero(n) and n > 0:
            dist.append((round(nivel * 100, 3), int(n)))
    return dist


def construir_serie(carpetas):
    filas = []
    for carpeta in carpetas:
        for path in sorted(Path(carpeta).glob("*.xls*")):
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                hojas = pd.read_excel(path, sheet_name=None, header=None)
            hoja = next((h for nombre, h in hojas.items() if "distrib" in nombre.lower()), None)
            dist = leer_distribucion(hoja) if hoja is not None else []
            if not dist:
                print(f"aviso: {path.name} sin distribución 13-24 meses")
                continue
            fila = {"fecha": fecha_encuesta(hojas, path.name), "tipo": tipo_encuesta(path.name),
                    "n": sum(n for _, n in dist), "archivo": path.name}
            for col, _, cond, _ in UMBRALES:
                fila[col] = sum(n for nivel, n in dist if cond(nivel))
            filas.append(fila)
    serie = (pd.DataFrame(filas).sort_values("fecha")
             .drop_duplicates(["fecha", "tipo"], keep="last").reset_index(drop=True))
    for col, *_ in UMBRALES:
        serie[f"pct_{col}"] = (100 * serie[col] / serie["n"]).round(1)
    return serie


def graficar(serie, col, etiqueta, archivo):
    s = serie.set_index("fecha")
    pct = s[f"pct_{col}"]

    fig, ax = plt.subplots(figsize=(6.4, 5.2), dpi=200)
    ax.fill_between(pct.index, pct, color=COLORES["relleno"], alpha=0.55, lw=0, step="mid")
    ax.plot(pct.index, pct, color=COLORES["serie"], lw=1.2, drawstyle="steps-mid")
    ax.axhline(50, color=COLORES["ref"], lw=0.8, ls="--")

    ult = s.iloc[-1]
    ax.annotate(f"{int(ult[col])}/{int(ult['n'])}\n({ult[f'pct_{col}']:.0f}%)",
                xy=(pct.index[-1], pct.iloc[-1]), xytext=(-8, 18), textcoords="offset points",
                ha="right", va="bottom", fontsize=9, color=COLORES["serie"], fontweight="bold",
                bbox={"boxstyle": "round,pad=0.25", "fc": "white", "ec": "none", "alpha": 0.9})

    ax.set_ylim(0, 100)
    ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}")
    ax.tick_params(labelsize=9, length=0)
    for lado in ("top", "right"):
        ax.spines[lado].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color("#7F7F7F")
    ax.grid(axis="y", color="#E5E5E5", lw=0.6)
    ax.set_axisbelow(True)
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%y"))
    ax.set_xlim(pct.index.min() - pd.Timedelta(days=30), pct.index.max() + pd.Timedelta(days=30))

    fig.text(0.5, 0.965, f"EOF: Inflación a 2 años {etiqueta}", color=COLORES["titulo"],
             fontsize=14, ha="center", va="top")
    fig.text(0.5, 0.915, "(porcentaje de operadores, inflación 13 a 24 meses adelante)",
             fontsize=10.5, ha="center", va="top")
    fig.text(0.01, 0.01, f"Fuente: Banco Central de Chile, EOF. Último dato: {ult['tipo']} "
             f"{pct.index[-1]:%d-%m-%Y}.", fontsize=7, color="#595959")
    fig.tight_layout(rect=(0, 0.03, 1, 0.87))
    fig.savefig(archivo)
    plt.close(fig)


def main():
    carpetas = [RAW] + [Path(c) for c in sys.argv[1:]]
    OUT.mkdir(exist_ok=True)
    serie = construir_serie(carpetas)
    serie.to_csv(OUT / "eof_inflacion_2a_umbrales.csv", index=False, date_format="%Y-%m-%d")
    for col, etiqueta, _, archivo in UMBRALES:
        graficar(serie, col, etiqueta, OUT / archivo)
    print(serie.tail(3).to_string())
    print(f"{len(serie)} encuestas, {serie['fecha'].min():%Y-%m-%d} a {serie['fecha'].max():%Y-%m-%d}")


if __name__ == "__main__":
    main()
