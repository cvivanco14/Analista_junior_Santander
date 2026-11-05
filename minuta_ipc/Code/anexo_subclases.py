"""Anexo de la minuta IPC: cuatro gráficos por cada una de las 172 subclases del IPC (canasta 2023).

Por subclase: perfil de la variación mensual por mes calendario (promedio 2013-2019 ± 1 d.e. y los
últimos años), variación mensual y anual de los últimos cinco años, y nivel del índice (2023=100).

Datos:
  * Inputs/subclases_bcch_1989_2023.csv: empalme de subclases del Banco Central (1989-2023, promedio
    2023=100), con las mismas 172 subclases, en el mismo orden y con los mismos nombres que el INE base 2023.
  * Desde enero de 2024: índices por subclase del INE base 2023 (los mismos Excel de la minuta),
    encadenados en diciembre de 2023 (el factor de enlace es 1,0000 dentro del redondeo).
  * Clasificación analítica (volátiles / sin volátiles): Inputs/matriz_productos.csv, por mayoría de
    ponderación de los productos de cada subclase.
"""

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402
from PIL import Image  # noqa: E402

import graficos_ipc as g  # noqa: E402  (estilo y formato de números de la minuta)
import ipc  # noqa: E402

PERIODO_PERFIL = (2013, 2019)
ANIOS_PERFIL = 3          # años completos previos al actual que se muestran en el perfil
ANIOS_RECIENTES = 5       # ventana de los gráficos m/m y a/a
INICIO_NIVELES = "2010-01"
COLORES_ANIOS = ["#7F7F7F", "#2CA02C", "#F2A93B"]  # años previos (de más antiguo a más reciente)


def cargar_subclases(datos: Path, insumos: Path) -> dict:
    """Índices por subclase empalmados, metadatos (peso, nombre, ruta analítica)."""
    ine = pd.concat([ipc._leer_ine(datos / "ipc-referencial-2023.xlsx"), ipc._leer_ine(datos / "ipc-base-2023.xlsx")])
    ine = ine.drop_duplicates(["periodo", "div", "grp", "clase", "subclase", "prod", "glosa"], keep="last")
    sc = ine[(ine["subclase"] > 0) & (ine["prod"] == 0)].copy()
    sc["clave"] = list(zip(sc["div"], sc["grp"], sc["clase"], sc["subclase"]))
    orden = sc[sc.periodo == sc.periodo.max()].reset_index(drop=True)
    orden["codigo"] = orden.index + 1
    sc["codigo"] = sc["clave"].map(dict(zip(orden.clave, orden.codigo)))
    nuevo = sc.pivot_table(index="periodo", columns="codigo", values="indice").sort_index()

    hist = pd.read_csv(insumos / "subclases_bcch_1989_2023.csv", index_col=0)
    hist.index = pd.PeriodIndex(hist.index, freq="M")
    hist.columns = hist.columns.astype(int)
    if list(hist.columns) != list(nuevo.columns):
        raise RuntimeError("Las subclases del INE no calzan con el empalme del Banco Central.")
    cc = ipc.CAMBIO_CANASTA
    enlace = nuevo.loc[cc] / hist.loc[cc]
    if (enlace - 1).abs().max() > 0.001:
        raise RuntimeError(f"Factor de enlace dic-2023 fuera de lo esperado: {enlace.abs().max():.4f}")
    niveles = pd.concat([hist.loc[:cc], nuevo.loc[nuevo.index > cc] / enlace])

    # ruta analítica por mayoría de ponderación de los productos de la subclase
    mat = pd.read_csv(insumos / "matriz_productos.csv").set_index("codigo")
    prod = ine[ine["prod"] > 0].drop_duplicates(["div", "grp", "clase", "subclase", "prod"])
    prod = prod.assign(codigo=prod[["div", "grp", "clase", "subclase", "prod"]].astype(str).agg("".join, axis=1)
                       .astype(int), clave=list(zip(prod["div"], prod["grp"], prod["clase"], prod["subclase"])))
    prod = prod.join(mat[["ponderacion", "categoria", "subcategoria", "subsubcategoria"]], on="codigo")
    rutas = {}
    for clave, grupo in prod.groupby("clave"):
        w = grupo.groupby(["categoria", "subcategoria", "subsubcategoria"])["ponderacion"].sum()
        ruta = " › ".join(_nombre_analitico(x) for x in w.idxmax())
        rutas[clave] = ruta + (" (mayoría de ponderación)" if len(w) > 1 else "")
    meta = orden.set_index("codigo")[["glosa", "pond", "clave"]]
    meta["pond"] = meta["pond"].astype(float)
    meta["ruta"] = meta["clave"].map(rutas)
    return {"niveles": niveles, "meta": meta}


def _nombre_analitico(x):
    """Nombre legible de las categorías de la matriz (usa los de la minuta cuando existen)."""
    return ipc.AGREGADOS.get(x, x)


def _titulo(glosa):
    return glosa.strip().capitalize()


def _num(x, d=1):
    return g._num(d)(x)


def figura_subclase(ruta, codigo, sub: dict, ipc_total: pd.Series, p: pd.Period):
    I = sub["niveles"][codigo].dropna()
    m = sub["meta"].loc[codigo]
    mm = (I / I.shift(1, freq="M") - 1).dropna() * 100
    aa = (I / I.shift(12, freq="M") - 1).dropna() * 100
    inc = m.pond * (I[p] - I[p - 1]) / ipc_total[p - 1] if p - 1 in I.index else float("nan")

    fig, ejes = plt.subplots(2, 2, figsize=(10, 6.4))
    fig.subplots_adjust(top=0.82, bottom=0.06, left=0.06, right=0.985, hspace=0.48, wspace=0.16)
    fig.text(0.01, 0.985, _titulo(m.glosa), fontsize=13, fontweight="bold", va="top")
    datos = (f"Peso: {_num(m.pond, 2)}   |   m/m: {_num(mm.get(p, float('nan')), 1)}%   |   "
             f"Incidencia m/m: {_num(inc, 3)} pp.   |   a/a: {_num(aa.get(p, float('nan')), 1)}%")
    fig.text(0.01, 0.935, datos, fontsize=10, va="top")
    fig.text(0.01, 0.895, f"Clasificación analítica: {m.ruta}", fontsize=9, color="#555555", va="top",
             style="italic")

    # 1) perfil m/m por mes calendario
    ax = ejes[0, 0]
    x = range(1, 13)
    base = mm[(mm.index.year >= PERIODO_PERFIL[0]) & (mm.index.year <= PERIODO_PERFIL[1])]
    if base.index.year.nunique() >= 4:
        por_mes = base.groupby(base.index.month)
        prom, de = por_mes.mean().reindex(x), por_mes.std().reindex(x)
        ax.fill_between(x, prom - de, prom + de, color=g.GRIS, alpha=0.6, lw=0,
                        label=f"±1 d.e. {PERIODO_PERFIL[0]}-{str(PERIODO_PERFIL[1])[2:]}")
        ax.plot(x, prom, color=g.AZUL, lw=2, label=f"Prom. {PERIODO_PERFIL[0]}-{str(PERIODO_PERFIL[1])[2:]}")
    anios = [p.year - k for k in range(ANIOS_PERFIL, 0, -1)]
    for anio, color in zip(anios, COLORES_ANIOS):
        s = mm[mm.index.year == anio]
        if len(s):
            ax.plot(s.index.month, s.values, color=color, lw=1.6, label=str(anio))
    s = mm[(mm.index.year == p.year) & (mm.index <= p)]
    ax.plot(s.index.month, s.values, color=g.ROJO, lw=2.4, marker="o", ms=3.5, label=str(p.year))
    ax.axhline(0, color="#7F7F7F", lw=0.8)
    ax.set_xticks(list(x), [mes[:1].upper() + mes[1:] for mes in g.MESES], fontsize=8)
    ax.set_title("Variación mensual por mes calendario (%)", fontsize=10, loc="left")
    ax.legend(fontsize=7.5, ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.1), handlelength=1.2,
              columnspacing=0.8)

    desde = pd.Period(f"{p.year - ANIOS_RECIENTES}-{p.month:02d}", "M")
    # 2) m/m últimos años
    ax = ejes[0, 1]
    s = mm.loc[desde:p]
    ax.bar(range(len(s)), s.values, color=[g.ROJO if q == p else g.AZUL for q in s.index], width=0.75)
    _eje_fechas(ax, s.index)
    ax.axhline(0, color="#7F7F7F", lw=0.8)
    ax.set_title(f"Variación mensual, últimos {ANIOS_RECIENTES} años (%)", fontsize=10, loc="left")

    # 3) niveles
    ax = ejes[1, 0]
    s = I.loc[pd.Period(INICIO_NIVELES, "M"):p]
    ax.plot(range(len(s)), s.values, color=g.AZUL, lw=1.8)
    _eje_fechas(ax, s.index, cada=2)
    if ipc.CAMBIO_CANASTA in s.index:
        ax.axvline(list(s.index).index(ipc.CAMBIO_CANASTA) + 0.5, color=g.NEGRO, ls=":", lw=1)
    ax.set_title("Índice (promedio 2023 = 100)", fontsize=10, loc="left")

    # 4) a/a últimos años
    ax = ejes[1, 1]
    s = aa.loc[desde:p]
    ax.plot(range(len(s)), s.values, color=g.AZUL, lw=1.8)
    if len(s):
        ax.plot(len(s) - 1, s.iloc[-1], "o", color=g.ROJO, ms=5)
    _eje_fechas(ax, s.index)
    ax.axhline(0, color="#7F7F7F", lw=0.8)
    ax.set_title(f"Variación anual, últimos {ANIOS_RECIENTES} años (%)", fontsize=10, loc="left")

    for ax in ejes.flat:
        ax.tick_params(labelsize=8)
        ax.yaxis.set_major_formatter(g._num(0 if ax is ejes[1, 0] else 1))
    fig.savefig(ruta, dpi=130)
    plt.close(fig)
    # paleta de 64 colores: el anexo pasa de ~23 MB a ~7 MB sin pérdida visible
    Image.open(ruta).convert("RGB").quantize(colors=64, method=Image.Quantize.MEDIANCUT).save(ruta, optimize=True)
    return ruta


def _eje_fechas(ax, periodos, cada=1):
    marcas = [i for i, q in enumerate(periodos) if q.month == 1 and q.year % cada == 0]
    ax.set_xticks(marcas, [str(periodos[i].year) for i in marcas], fontsize=8)
    ax.set_xlim(-1, len(periodos))


def generar(carpeta: Path, datos: Path, insumos: Path, ipc_total: pd.Series, p: pd.Period) -> list:
    """Genera las 172 figuras en carpeta/ y devuelve sus rutas, en el orden de la estructura del IPC."""
    carpeta.mkdir(parents=True, exist_ok=True)
    for f in carpeta.glob("sub*.png"):
        f.unlink()
    sub = cargar_subclases(datos, insumos)
    return [figura_subclase(carpeta / f"sub{c:03d}.png", c, sub, ipc_total, p) for c in sub["meta"].index]
