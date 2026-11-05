"""Gráficos de la minuta IPC (matplotlib)."""

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import matplotlib.ticker as mticker  # noqa: E402
import pandas as pd  # noqa: E402

AZUL, ROJO, NEGRO, GRIS, VERDE = "#1F3FBF", "#D62728", "black", "#CBCCCB", "darkgreen"
COLORES = ["#1F3FBF", "#D62728", "#7F7F7F", "#F2A93B", "#2CA02C", "#17BECF"]
MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]

plt.rcParams.update({
    "font.family": "Helvetica", "font.size": 11,
    "axes.spines.top": False, "axes.spines.right": False, "axes.spines.left": False,
    "axes.edgecolor": "#AFB3B3", "axes.grid": True, "axes.grid.axis": "y",
    "grid.color": "#AFB3B3", "grid.linewidth": 0.6, "ytick.left": False,
    "legend.frameon": False, "savefig.dpi": 200,
})


def _num(dec=1):
    def f(x, _p=None):
        return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return mticker.FuncFormatter(f)


def _figura(titulo, subtitulo, nota):
    fig, ax = plt.subplots(figsize=(6.6, 4.6))
    fig.subplots_adjust(top=0.78, bottom=0.15, left=0.11, right=0.89)
    fig.text(0.02, 0.965, titulo, fontsize=13, fontweight="bold", va="top")
    fig.text(0.02, 0.905, subtitulo, fontsize=11, va="top")
    fig.text(0.02, 0.015, nota, fontsize=9, va="bottom")
    return fig, ax


def _leyenda(fig, handles, ncol=None):
    fig.legend(handles=handles, loc="upper center", bbox_to_anchor=(0.5, 0.865),
               ncol=ncol or len(handles), fontsize=9.5, handlelength=1.2, columnspacing=1.0)


def _eje_tiempo(ax, periodos, linea_canasta=None):
    x = range(len(periodos))
    marcas = [i for i, p in enumerate(periodos) if p.month == 1]
    ax.set_xticks(marcas, [f"{MESES[0]}-{str(periodos[i].year)[2:]}" for i in marcas], fontsize=9.5)
    ax.set_xlim(-1, len(periodos))
    if linea_canasta is not None and linea_canasta in list(periodos):
        ax.axvline(list(periodos).index(linea_canasta) + 0.5, color=NEGRO, ls=":", lw=1.2)
    return x


def _guardar(fig, ruta):
    fig.savefig(ruta)
    plt.close(fig)
    return ruta


def lineas_barras(ruta, titulo, subtitulo, lineas: dict, barras: pd.Series = None, nota="Fuente: INE.",
                  linea_canasta=None, dec=1, etiqueta_barras="Var. % m/m (der.)"):
    """Líneas en el eje izquierdo y, opcionalmente, barras en el eje derecho."""
    fig, ax = _figura(titulo, subtitulo, nota)
    periodos = next(iter(lineas.values())).index
    x = _eje_tiempo(ax, periodos, linea_canasta)
    handles = []
    if barras is not None:
        ax2 = ax.twinx()
        handles.append(ax2.bar(x, barras.reindex(periodos).values, color=GRIS, width=0.8,
                               label=etiqueta_barras))
        ax2.grid(False)
        ax2.yaxis.set_major_formatter(_num(1))
        for lado in ("top", "left", "right"):
            ax2.spines[lado].set_visible(False)
        ax2.tick_params(axis="y", length=0)
        ax.set_zorder(ax2.get_zorder() + 1)
        ax.patch.set_visible(False)
    for (etq, s), color in zip(lineas.items(), [AZUL, ROJO, NEGRO, VERDE]):
        handles += ax.plot(x, s.reindex(periodos).values, color=color, lw=2, label=etq)
    ax.yaxis.set_major_formatter(_num(dec))
    ax.margins(y=0.08)
    _leyenda(fig, handles)
    return _guardar(fig, ruta)


def barras_mes(ruta, titulo, subtitulo, variacion: pd.Series, incid: pd.Series, nota="Fuente: INE."):
    """Variación m/m e incidencia por agregado para el mes."""
    fig, ax = _figura(titulo, subtitulo, nota)
    n = range(len(variacion))
    h1 = ax.bar([i - 0.2 for i in n], variacion.values, width=0.4, color=AZUL, label="Var. % m/m")
    h2 = ax.bar([i + 0.2 for i in n], incid.values, width=0.4, color=GRIS, label="Incidencia (pp.)")
    ax.axhline(0, color="#7F7F7F", lw=0.8)
    ax.set_xticks(list(n), [s.replace(" ", "\n", 1) for s in variacion.index], fontsize=9)
    ax.yaxis.set_major_formatter(_num(1))
    ax.margins(y=0.12)
    _leyenda(fig, [h1, h2])
    return _guardar(fig, ruta)


def incidencias_apiladas(ruta, titulo, subtitulo, total: pd.Series, componentes: dict,
                         nota="Fuente: INE y Banco Central de Chile.", linea_canasta=None):
    fig, ax = _figura(titulo, subtitulo, nota)
    periodos = total.index
    x = list(_eje_tiempo(ax, periodos, linea_canasta))
    pos = [0.0] * len(x)
    neg = [0.0] * len(x)
    handles = []
    for (etq, s), color in zip(componentes.items(), COLORES):
        v = s.reindex(periodos).fillna(0).values
        base = [pos[i] if v[i] >= 0 else neg[i] for i in range(len(v))]
        handles.append(ax.bar(x, v, bottom=base, width=0.85, color=color, label=etq))
        for i in range(len(v)):
            if v[i] >= 0:
                pos[i] += v[i]
            else:
                neg[i] += v[i]
    handles += ax.plot(x, total.values, color=NEGRO, lw=2, label="IPC total")
    ax.yaxis.set_major_formatter(_num(1))
    _leyenda(fig, handles, ncol=3)
    fig.subplots_adjust(top=0.74)
    return _guardar(fig, ruta)


def difusion(ruta, titulo, hist: pd.DataFrame, previo: pd.Series, actual: pd.Series,
             anio_actual: int, nota="Fuente: INE."):
    """Banda mín–máx y promedio 2019-2022, año anterior (línea) y año actual (puntos)."""
    fig, ax = _figura(titulo, "(% de productos con var. mensual positiva)", nota)
    x = range(1, 13)
    banda = ax.fill_between(x, hist["min"], hist["max"], color=GRIS, lw=0, label="Mín-Máx 2019-22")
    handles = [banda]
    handles += ax.plot(x, hist["prom"], color=AZUL, lw=2, label="Prom. 2019-22")
    handles += ax.plot(x, previo.reindex(x).values, color=ROJO, lw=2, label=str(anio_actual - 1))
    handles += ax.plot(x, actual.reindex(x).values, ls="none", marker="D", ms=6, color=VERDE,
                       label=str(anio_actual))
    ax.set_xticks(list(x), MESES, fontsize=10)
    ax.yaxis.set_major_formatter(_num(0))
    ax.margins(y=0.1)
    _leyenda(fig, handles)
    return _guardar(fig, ruta)
