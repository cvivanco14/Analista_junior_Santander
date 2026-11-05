# -*- coding: utf-8 -*-
r"""Minuta Remuneraciones - Gerencia de Estudios (Santander Chile).

Minuta mensual en PDF con los tabulados IR-ICL del INE (base anual 2023=100), con textos generados
automáticamente a partir del último dato, en colores Santander. Lee los Excel que deja
datos_remuneraciones.py en Inputs.

Carpetas (relativas a este script; sirven para cualquier usuario):
    <Minuta Remuneraciones>\Code\     datos_remuneraciones.py y minuta_remuneraciones.py
    <Minuta Remuneraciones>\Inputs\   tabulados del INE (los deja datos_remuneraciones.py)
    <Minuta Remuneraciones>\Outputs\  AAAAMMDD_Minuta_Remuneraciones.pdf y figuras\

Uso (no hay que editar nada):
    python datos_remuneraciones.py               # descarga los tabulados del INE a Inputs
    python minuta_remuneraciones.py              # minuta del último mes publicado
    python minuta_remuneraciones.py --mes 2026-06  # minuta de un mes anterior
"""

import argparse
import datetime as dt
import os
import re
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import matplotlib.ticker as mticker  # noqa: E402
import pandas as pd  # noqa: E402
from reportlab.lib import colors  # noqa: E402
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY  # noqa: E402
from reportlab.lib.pagesizes import letter  # noqa: E402
from reportlab.lib.styles import ParagraphStyle  # noqa: E402
from reportlab.lib.units import cm  # noqa: E402
from reportlab.platypus import (Image, KeepTogether, ListFlowable, ListItem, PageBreak,  # noqa: E402
                                Paragraph, SimpleDocTemplate, Table, TableStyle)
from reportlab.platypus.flowables import HRFlowable  # noqa: E402

import datos_remuneraciones as dr  # noqa: E402

AUTOR = "Andrés Sansone"
ANIO_INICIO_BANDA = 2010  # gráficos estacionales: banda y promedio desde este año hasta (año actual - 3)

CODE = dr.CODE
INPUTS, OUT = dr.INPUTS, dr.OUT
FIGURAS = OUT / "figuras"
MANUAL_ESTILO = Path("Administrativos") / "Manual de Estilo"
BIBLIOTECA = Path("Estudios - Documentos") / "General"
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
         "noviembre", "diciembre"]

# colores Santander (los mismos de formato_graficos.py)
ROJO_SANTANDER = colors.Color(236 / 255, 0, 0)
GRIS_TEXTO = colors.Color(89 / 255, 89 / 255, 89 / 255)
GRIS_LINEA = colors.Color(191 / 255, 191 / 255, 191 / 255)
GRIS_FONDO = colors.Color(242 / 255, 242 / 255, 242 / 255)


def _logo():
    """Logo Santander (assets/logo.png del Manual de Estilo), buscado como en build_chartpack: variable
    FORMATO_GRAFICOS_DIR, subiendo desde Code y en las carpetas sincronizadas de SharePoint. Opcional."""
    candidatos = [Path(os.environ["FORMATO_GRAFICOS_DIR"])] if os.environ.get("FORMATO_GRAFICOS_DIR") else []
    for p in [CODE, *CODE.parents]:
        candidatos += [p / MANUAL_ESTILO, p / "formato"]
    casa = Path.home()
    for base in (casa, casa / "Library" / "CloudStorage"):
        if base.is_dir():
            for raiz in base.iterdir():
                if raiz.is_dir() and re.search(r"santander|onedrive|sharepoint", raiz.name, re.I):
                    candidatos.append(raiz / BIBLIOTECA / MANUAL_ESTILO)
    return next((c / "assets" / "logo.png" for c in candidatos if (c / "assets" / "logo.png").is_file()), None)


# ------------------------------------------------------------------ formato y texto
def num(x, d=1):
    s = f"{x:,.{d}f}"
    if float(s.replace(",", "")) == 0:
        s = s.lstrip("-")
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def pct(x, d=1):
    return f"{num(x, d)}%"


def signo(x, sube, baja, igual):
    x = round(x, 1)
    return sube if x > 0 else baja if x < 0 else igual


def lista(partes):
    return partes[0] if len(partes) == 1 else ", ".join(partes[:-1]) + " y " + partes[-1]


def var(s: pd.Series, rezago: int) -> pd.Series:
    return (s / s.shift(rezago, freq="M") - 1) * 100


def ranking(tabla: pd.DataFrame, sujeto: str) -> str:
    """tabla con columnas nombre, var_12 -> frase de alzas y caídas a/a."""
    t = tabla.dropna(subset=["var_12"])
    pos = t[t.var_12.round(1) > 0].sort_values("var_12", ascending=False).head(3)
    neg = t[t.var_12.round(1) < 0].sort_values("var_12").head(3)
    fr = lambda df: lista([f"{r.nombre.lower()} ({pct(r.var_12)})" for r in df.itertuples()])  # noqa: E731
    partes = []
    if len(pos):
        partes.append(f"Entre los {sujeto}, las mayores alzas se registraron en {fr(pos)}")
    if len(neg):
        partes.append(("mientras que las mayores caídas se observaron en " if partes else
                       f"Entre los {sujeto}, las mayores caídas se observaron en ") + fr(neg))
    return (", ".join(partes) + ".") if partes else f"Ningún {sujeto[:-1]} registró variación anual."


class Mes:
    def __init__(self, d, p):
        self.d, self.p = d, p
        self.mm_sa = {k: var(d["sa"][k], 1) for k in d["sa"]}
        self.aa = {k: var(d["nsa"][k], 12) for k in d["nsa"]}
        self.mm_nsa = {k: var(d["nsa"][k], 1) for k in d["nsa"]}

    def horas(self, tipo):
        h = self.d["horas"]
        fila = h[(h.index == self.p) & (h["estimación"] == tipo)].iloc[0]
        act = self.d["horas_act"]
        grp = self.d["horas_grp"]
        sel = lambda x: x[(x.index == self.p) & (x["estimación"] == tipo)][["nombre", "var_12"]]  # noqa: E731
        return fila, sel(act), sel(grp)


def titulo(m: Mes):
    p = m.p
    r_mm, r_aa = m.mm_sa["ir_real"][p], m.aa["ir_real"][p]
    return (f"Remuneraciones {MESES[p.month - 1]} {p.year} | En {MESES[p.month - 1]}, los salarios reales "
            + signo(r_mm, f"aumentaron {pct(abs(r_mm))}", f"cayeron {pct(abs(r_mm))}", "no variaron")
            + " m/m (sa) y " + signo(r_aa, f"crecen {pct(abs(r_aa))}", f"caen {pct(abs(r_aa))}", "no varían")
            + " en doce meses.")


def destacados(m: Mes):
    p, pa = m.p, m.p - 1
    mes, mes_a = MESES[p.month - 1], MESES[pa.month - 1]
    ir_mm, ir_mm_a = m.mm_sa["ir"][p], m.mm_sa["ir"][pa]
    hist = m.mm_nsa["ir"]
    prom = hist[(hist.index.month == p.month) & (hist.index.year >= 2010) & (hist.index.year <= 2019)].mean()
    b = [f"<b>El Índice de Remuneraciones nominal (IR) registró "
         + signo(ir_mm, "un aumento", "una caída", "una variación nula") + f" de {pct(ir_mm)} m/m (sa), "
         + signo(ir_mm - ir_mm_a, "por sobre", "por debajo de", "en línea con")
         + f" lo observado en {mes_a} ({pct(ir_mm_a)}).</b> En términos anuales, el índice muestra "
         + signo(m.aa["ir"][p], "una expansión", "una contracción", "una variación nula")
         + f" de {pct(m.aa['ir'][p])}. La variación mensual de {mes} en la serie original ({pct(m.mm_nsa['ir'][p])}) se "
         + "ubica " + signo(m.mm_nsa["ir"][p] - prom, "por sobre el", "por debajo del", "en línea con el")
         + f" promedio 2010-2019 para el mismo mes ({pct(prom)}). Por su parte, el Índice de Costos Laborales "
         f"(ICL) varió {pct(m.mm_sa['icl'][p])} m/m (sa) y {pct(m.aa['icl'][p])} a/a."]

    r_aa = m.aa["ir_real"]
    inflacion = ((1 + m.aa["ir"][p] / 100) / (1 + r_aa[p] / 100) - 1) * 100
    racha, signo_actual = 0, r_aa[p] > 0
    for q in sorted(r_aa.dropna().index[r_aa.dropna().index <= p], reverse=True):
        if (r_aa[q] > 0) != signo_actual:
            break
        racha += 1
    b.append(f"<b>Los salarios reales registraron "
             + signo(m.mm_sa["ir_real"][p], "un aumento", "una caída", "una variación nula")
             + f" de {pct(m.mm_sa['ir_real'][p])} m/m (sa) en {mes}.</b> En términos anuales, considerando una "
             f"inflación de {pct(inflacion)} (IPC implícito en el IR real del INE), los salarios reales "
             + signo(r_aa[p], "se expandieron", "se contrajeron", "no variaron") + f" {pct(abs(r_aa[p]))}, con lo que "
             f"acumulan {racha} {'mes' if racha == 1 else 'meses'} consecutivos de "
             + ("alzas" if signo_actual else "caídas") + " anuales.")

    for tipo, nombre in (("HONT", "ordinarias"), ("HENT", "extraordinarias")):
        fila, act, grp = m.horas(tipo)
        b.append(f"<b>Las horas {nombre} medias por persona trabajadora "
                 + signo(fila.var_12, "aumentaron", "cayeron", "no variaron") + f" {pct(abs(fila.var_12))} a/a</b> "
                 f"({num(fila.valor, 1)} horas al mes). {ranking(act, 'sectores')} {ranking(grp, 'grupos ocupacionales')}")
    return b


# ------------------------------------------------------------------ figuras
plt.rcParams.update({
    "font.family": "Helvetica", "font.size": 11, "axes.spines.top": False, "axes.spines.right": False,
    "axes.spines.left": False, "axes.edgecolor": "#AFB3B3", "axes.grid": True, "axes.grid.axis": "y",
    "grid.color": "#AFB3B3", "grid.linewidth": 0.6, "ytick.left": False, "legend.frameon": False,
    "savefig.dpi": 200,
})
MES_CORTO = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


def _fig(titulo, subtitulo):
    fig, ax = plt.subplots(figsize=(6.6, 4.6))
    fig.subplots_adjust(top=0.78, bottom=0.14, left=0.11, right=0.95)
    fig.text(0.02, 0.965, titulo, fontsize=13, fontweight="bold", va="top")
    fig.text(0.02, 0.905, subtitulo, fontsize=11, va="top")
    fig.text(0.02, 0.015, "Fuente: INE.", fontsize=9, va="bottom")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _p: num(x, 1)))
    return fig, ax


def _guardar(fig, n):
    FIGURAS.mkdir(parents=True, exist_ok=True)
    ruta = FIGURAS / f"fig{n}.png"
    fig.savefig(ruta)
    plt.close(fig)
    return ruta


def _eje_tiempo(ax, periodos):
    marcas = [i for i, q in enumerate(periodos) if q.month == 1]
    ax.set_xticks(marcas, [str(periodos[i].year) for i in marcas], fontsize=10)
    ax.set_xlim(-1, len(periodos))


def figuras(m: Mes):
    p, d = m.p, m.d
    out = []
    s = d["sa"]["ir_real"].loc[p - 71:p]
    fig, ax = _fig("Figura 1: Índice de Remuneraciones real", "(serie desestacionalizada, 2023=100)")
    ax.plot(range(len(s)), s.values, color="darkgreen", lw=2)
    _eje_tiempo(ax, s.index)
    out.append(_guardar(fig, 1))

    ini = pd.Period(f"{p.year - 7}-01", "M")
    fig, ax = _fig("Figura 2: Variación anual de los índices", "(%, serie original)")
    idx = m.aa["ir"].loc[ini:p].index
    for k, etq, c in (("ir", "IR", "#1F3FBF"), ("icl", "ICL", "#D62728"), ("ir_real", "IR real", "darkgreen")):
        ax.plot(range(len(idx)), m.aa[k].reindex(idx).values, color=c, lw=2, label=etq)
    ax.axhline(0, color="#7F7F7F", lw=0.8)
    _eje_tiempo(ax, idx)
    fig.legend(loc="upper center", bbox_to_anchor=(0.5, 0.865), ncol=3, fontsize=10)
    out.append(_guardar(fig, 2))

    base = list(range(ANIO_INICIO_BANDA, p.year - 2))
    for n, k, nombre in ((3, "ir", "nominal"), (4, "ir_real", "real")):
        v = m.mm_sa[k].dropna()
        v = v[v.index <= p]
        t = pd.DataFrame({"v": v.values, "a": v.index.year, "m": v.index.month}).pivot(index="m", columns="a",
                                                                                         values="v")
        b = t[[a for a in base if a in t]]
        fig, ax = _fig(f"Figura {n}: Índice de Remuneraciones {nombre}",
                       f"(var. % m/m, serie desestacionalizada)")
        x = list(range(1, 13))
        h = [ax.fill_between(x, b.min(axis=1), b.max(axis=1), color="#CBCCCB", lw=0,
                             label=f"Mín-Máx {base[0]}-{base[-1]}")]
        h += ax.plot(x, b.mean(axis=1), color="#1F3FBF", lw=2, label=f"Prom. {base[0]}-{base[-1]}")
        for a, c in ((p.year - 2, "black"), (p.year - 1, "#D62728")):
            if a in t:
                h += ax.plot(x, t[a].reindex(x).values, color=c, lw=1.6, label=str(a))
        h += ax.plot(x, t[p.year].reindex(x).values, ls="none", marker="D", ms=6, color="darkgreen",
                     label=str(p.year))
        ax.axhline(0, color="#7F7F7F", lw=0.8)
        ax.set_xticks(x, MES_CORTO, fontsize=10)
        fig.legend(handles=h, loc="upper center", bbox_to_anchor=(0.5, 0.865), ncol=5, fontsize=9.5,
                   handlelength=1.2, columnspacing=1.0)
        out.append(_guardar(fig, n))
    return out


# ------------------------------------------------------------------ PDF
E = {
    "titulo": ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=14, leading=18,
                             textColor=ROJO_SANTANDER),
    "fecha": ParagraphStyle("f", fontName="Helvetica-Bold", fontSize=10.5, spaceBefore=6),
    "autor": ParagraphStyle("a", fontName="Helvetica-Oblique", fontSize=9, spaceBefore=2, textColor=GRIS_TEXTO),
    "h2": ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=13, textColor=ROJO_SANTANDER,
                         spaceBefore=8, spaceAfter=6),
    "cuerpo": ParagraphStyle("c", fontName="Helvetica", fontSize=10, leading=13.5, alignment=TA_JUSTIFY),
    "caption": ParagraphStyle("cap", fontName="Helvetica", fontSize=9, alignment=TA_CENTER, spaceBefore=12,
                              spaceAfter=4),
    "nota": ParagraphStyle("n", fontName="Helvetica", fontSize=7.5, leading=9, spaceBefore=3,
                           textColor=GRIS_TEXTO),
}


def cuadro(m: Mes, ancho):
    p, pa = m.p, m.p - 1
    corto = lambda q: f"{MES_CORTO[q.month - 1]}-{str(q.year)[2:]}"  # noqa: E731
    filas = [["Índice", f"m/m (sa) {corto(pa)}", f"m/m (sa) {corto(p)}", f"a/a {corto(pa)}", f"a/a {corto(p)}"]]
    for k, etq in (("ir", "IR nominal"), ("icl", "ICL nominal"), ("ir_real", "IR real")):
        filas.append([etq, pct(m.mm_sa[k][pa], 2), pct(m.mm_sa[k][p], 2), pct(m.aa[k][pa], 2), pct(m.aa[k][p], 2)])
    t = Table(filas, colWidths=[ancho * 0.24] + [ancho * 0.16] * 4)
    t.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "Helvetica", 8.5), ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8.5),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, 0), ROJO_SANTANDER), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("LINEBELOW", (0, 1), (-1, -2), 0.5, GRIS_LINEA), ("LINEBELOW", (0, -1), (-1, -1), 1, ROJO_SANTANDER),
        ("BACKGROUND", (0, 2), (-1, 2), GRIS_FONDO),
    ]))
    return KeepTogether([Paragraph("Cuadro 1: Variaciones de los índices de remuneraciones y costos laborales",
                                   E["caption"]), t,
                         Paragraph("Nota: sa = serie desestacionalizada por el INE. El último dato es provisorio. "
                                   "Fuente: INE.", E["nota"])])


def armar_pdf(m: Mes, figs, ruta):
    izq, der = 2 * cm, 1.84 * cm
    ancho = letter[0] - izq - der

    logo = _logo()

    def pagina(canvas, doc):
        canvas.saveState()
        if logo:   # logo Santander arriba a la derecha (proporción del logo de las láminas)
            canvas.drawImage(str(logo), letter[0] - der - 0.9 * cm, letter[1] - 1.75 * cm, 0.9 * cm, 1.16 * cm,
                             mask="auto", preserveAspectRatio=True)
        canvas.setStrokeColor(ROJO_SANTANDER)
        canvas.setLineWidth(0.75)
        canvas.line(izq, 1.95 * cm, letter[0] - der, 1.95 * cm)
        canvas.setFillColor(GRIS_TEXTO)
        canvas.setFont("Helvetica-Bold", 9)
        canvas.drawRightString(letter[0] - der, 1.5 * cm, str(doc.page))
        canvas.setFont("Helvetica", 8.5)
        canvas.drawString(izq, 1.5 * cm, AUTOR)
        canvas.restoreState()

    doc = SimpleDocTemplate(str(ruta), pagesize=letter, leftMargin=izq, rightMargin=der, topMargin=2.2 * cm,
                            bottomMargin=2.3 * cm, title=f"Minuta Remuneraciones {MESES[m.p.month - 1]} {m.p.year}",
                            author=AUTOR)
    hoy = dt.date.today()
    w = ancho / 2 - 0.1 * cm
    imgs = Table([[Image(str(figs[i]), width=w, height=w * 4.6 / 6.6) for i in (j, j + 1)] for j in (0, 2)],
                 colWidths=[ancho / 2] * 2)
    imgs.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    h = [
        Paragraph(titulo(m), E["titulo"]),
        Paragraph(f"{hoy.day} de {MESES[hoy.month - 1]} de {hoy.year}", E["fecha"]),
        Paragraph(f"Autor: {AUTOR}", E["autor"]),
        HRFlowable(width="100%", thickness=1, color=ROJO_SANTANDER, spaceBefore=4, spaceAfter=2),
        Paragraph("Principales resultados", E["h2"]),
        ListFlowable([ListItem(Paragraph(x, E["cuerpo"]), leftIndent=12) for x in destacados(m)],
                     bulletType="bullet", start="•", leftIndent=12, bulletFontSize=9,
                     bulletColor=ROJO_SANTANDER),
        cuadro(m, ancho),
        PageBreak(),
        Paragraph("Figuras", E["h2"]),
        imgs,
    ]
    doc.build(h, onFirstPage=pagina, onLaterPages=pagina)


def main(hoy=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--mes", help="mes del reporte AAAA-MM (por defecto, el último publicado)")
    args = ap.parse_args()
    hoy = pd.Timestamp(hoy or dt.date.today())
    d = dr.cargar(INPUTS)
    p = pd.Period(args.mes, "M") if args.mes else min(d["sa"].dropna().index[-1], d["horas"].index[-1])
    print(f"Mes del reporte: {MESES[p.month - 1]} {p.year}")
    m = Mes(d, p)
    OUT.mkdir(exist_ok=True)
    ruta = OUT / f"{hoy:%Y%m%d}_Minuta_Remuneraciones.pdf"
    armar_pdf(m, figuras(m), ruta)
    print(ruta)


if __name__ == "__main__":
    main()
