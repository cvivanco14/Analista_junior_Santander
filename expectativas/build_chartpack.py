"""Arma el Chartpack de Expectativas (PPTX + PDF) con el formato de Estudios Económicos.

Uso:
    python fetch_data.py         # descarga series a data/
    python build_chartpack.py    # genera output/<fecha>_Chartpack_Expectativas.pptx y .pdf
"""
import datetime as dt
import subprocess
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib import font_manager
from matplotlib.ticker import FixedLocator, FuncFormatter
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Pt

ROOT = Path(__file__).parent
DATA = ROOT / "data"
ASSETS = ROOT / "assets"
OUT = ROOT / "output"
CHARTS = OUT / "charts"

MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
MESES_LARGO = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
               "septiembre", "octubre", "noviembre", "diciembre"]

ROJO_SAN = "#EC0000"
GRIS_TXT = "#595959"
GRIS_GRID = "#D9D9D9"
GRIS_EJE = "#BFBFBF"

# paletas tomadas del chartpack original
AZUL, VERDE, ROJO = "#0070C0", "#00B050", "#FF0000"
NAVY, ROJO_OSC = "#0E2841", "#C00000"
AZUL_ACERO, NARANJO, DORADO = "#1F5F8B", "#ED7D31", "#E0B020"
VERDE_CLARO, CELESTE, VERDE_OSC, MORADO = "#70BF5A", "#00A0E0", "#1E7B34", "#A0209F"

for f in font_manager.findSystemFonts():
    if "Carlito" in f:
        font_manager.fontManager.addfont(f)
plt.rcParams.update({
    "font.family": "Carlito",
    "font.size": 10.5,
    "axes.edgecolor": GRIS_EJE,
    "axes.linewidth": 0.75,
    "xtick.color": GRIS_EJE,
    "ytick.color": GRIS_EJE,
    "xtick.labelcolor": GRIS_TXT,
    "ytick.labelcolor": GRIS_TXT,
    "legend.frameon": False,
    "legend.fontsize": 10.5,
    "legend.handlelength": 2.2,
    "legend.labelcolor": GRIS_TXT,
})


# ---------------------------------------------------------------- datos

def serie(code):
    s = pd.read_csv(DATA / f"{code}.csv", index_col=0, parse_dates=True)["valor"]
    return s.dropna().sort_index()


def desde(s, inicio):
    return s[s.index >= pd.Timestamp(inicio)]


# ---------------------------------------------------------------- formatos

def coma(dec):
    return FuncFormatter(lambda v, _: f"{v:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", "."))


def fecha_lbl(estilo):
    def f(d):
        if estilo == "anio":
            return f"{d.year}"
        sep = "-" if estilo == "guion" else "."
        return f"{MESES[d.month - 1]}{sep}{d.year % 100:02d}"
    return f


def ticks_fecha(inicio, fin, meses):
    t, out = pd.Timestamp(inicio), []
    while t <= pd.Timestamp(fin):
        out.append(t)
        t += pd.DateOffset(months=meses)
    return out


# ---------------------------------------------------------------- graficos

def nuevo_grafico(w_pt, h_pt):
    fig, ax = plt.subplots(figsize=(w_pt / 72, h_pt / 72))
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", color=GRIS_GRID, lw=0.75)
    ax.set_axisbelow(True)
    ax.tick_params(axis="y", length=0, pad=6)
    ax.tick_params(axis="x", length=4, direction="out", pad=4)
    return fig, ax


def ejes(ax, ylim, ystep, xini, xfin, xticks, estilo, dec=0, rot=0):
    lo, hi = ylim
    n = round((hi - lo) / ystep)
    ax.set_ylim(lo, hi)
    ax.yaxis.set_major_locator(FixedLocator([lo + i * ystep for i in range(n + 1)]))
    ax.yaxis.set_major_formatter(coma(dec))
    ax.set_xlim(pd.Timestamp(xini), pd.Timestamp(xfin))
    lbl = fecha_lbl(estilo)
    ax.set_xticks(xticks)
    ax.set_xticklabels([lbl(t) for t in xticks], rotation=rot,
                       va="top", ha="center" if rot == 0 else "center")


def unidad(ax, texto, x=0.05, y=0.93, ha="left"):
    ax.text(x, y, texto, transform=ax.transAxes, color=GRIS_TXT, fontsize=10.5, ha=ha, va="top")


def linea50(ax, xini, xfin):
    ax.plot([pd.Timestamp(xini), pd.Timestamp(xfin)], [50, 50], ls=(0, (3, 2)), color="#7F7F7F", lw=1.1, zorder=1)


def guardar(fig, nombre):
    CHARTS.mkdir(parents=True, exist_ok=True)
    p = CHARTS / f"{nombre}.png"
    fig.savefig(p, dpi=250)
    plt.close(fig)
    return p


def rango(*series, paso, piso=None, techo=None, margen=0.0):
    """Limites redondeados a 'paso' que contienen todas las series."""
    lo = min(s.min() for s in series) - margen
    hi = max(s.max() for s in series) + margen
    lo = (lo // paso) * paso
    hi = -((-hi) // paso) * paso
    if piso is not None:
        lo = min(lo, piso)
    if techo is not None:
        hi = max(hi, techo)
    return round(lo, 6), round(hi, 6)


# medidas de las cajas de grafico en la lamina (pt)
DOS = [(36, 148, 434, 344), (497, 148, 430, 344)]
TRES = [(26, 148, 300, 344), (336, 148, 300, 344), (646, 148, 300, 344)]
UNO = (40, 148, 892, 344)


def g_imce(codigos, nombre, fin):
    w, h = DOS[0][2:]
    fig, ax = nuevo_grafico(w, h)
    for code, lbl, c in zip(codigos, ["Comercio", "Construcción", "Industria"], [AZUL, VERDE, ROJO]):
        s = desde(serie(code), "2018-01-01")
        ax.plot(s.index, s.values, color=c, lw=1.6, label=lbl)
    xini, xfin = "2017-12-01", fin
    linea50(ax, "2018-01-01", fin)
    ejes(ax, (0, 100), 10, xini, xfin, ticks_fecha("2018-01-01", fin, 12), "guion")
    unidad(ax, "(Índice)", 0.05, 0.95)
    ax.legend(loc="lower right", bbox_to_anchor=(1.0, 0.03), ncol=3, columnspacing=1.2, handlelength=2.6)
    fig.subplots_adjust(left=0.085, right=0.975, top=0.97, bottom=0.085)
    return guardar(fig, nombre)


def g_dos_lineas(s1, s2, l1, l2, c1, c2, nombre, inicio, fin, ylim, ystep, xticks, estilo, unidad_txt,
                 leg_loc, dec=0, caja=DOS[0], lw=1.6, leg_cols=2, rot=0, bottom=0.085, linea_50=False):
    fig, ax = nuevo_grafico(*caja[2:])
    ax.plot(s1.index, s1.values, color=c1, lw=lw, label=l1)
    ax.plot(s2.index, s2.values, color=c2, lw=lw, label=l2)
    if linea_50:
        linea50(ax, inicio, fin)
    ejes(ax, ylim, ystep, inicio, fin, xticks, estilo, dec, rot)
    unidad(ax, unidad_txt)
    ax.legend(loc="center", bbox_to_anchor=leg_loc, ncol=leg_cols, columnspacing=1.2)
    fig.subplots_adjust(left=0.085, right=0.975, top=0.97, bottom=bottom)
    return guardar(fig, nombre)


def g_multi(series, nombre, inicio, fin, ylim, ystep, unidad_txt, unidad_pos, leg_kw, caja):
    fig, ax = nuevo_grafico(*caja[2:])
    for s, lbl, c in series:
        ax.plot(s.index, s.values, color=c, lw=1.6, label=lbl)
    linea50(ax, inicio, fin)
    xt = ticks_fecha(inicio, fin, 6)
    ejes(ax, ylim, ystep, pd.Timestamp(inicio) - pd.DateOffset(days=20), fin, xt, "punto", rot=90)
    unidad(ax, unidad_txt, *unidad_pos)
    ax.legend(**leg_kw)
    fig.subplots_adjust(left=0.115, right=0.98, top=0.97, bottom=0.15)
    return guardar(fig, nombre)


def g_eee(s_act, s_sig, l_act, l_sig, nombre, inicio, fin, caja, paso, leg_loc):
    fig, ax = nuevo_grafico(*caja[2:])
    ax.plot(s_act.index, s_act.values, color=DORADO, lw=1.8, label=l_act)
    mk = "D" if len(s_sig) < 4 else None
    ax.plot(s_sig.index, s_sig.values, color=AZUL_ACERO, lw=2.2, label=l_sig, marker=mk, ms=5)
    ylim = rango(s_act, s_sig, paso=paso, margen=paso * 0.5)
    xt = ticks_fecha(inicio, fin, 3)
    ejes(ax, ylim, paso, pd.Timestamp(inicio) - pd.DateOffset(days=10), pd.Timestamp(fin) + pd.DateOffset(days=10),
         xt, "punto", dec=1)
    unidad(ax, "(Var. anual, %)", 0.05, 0.95)
    h, lb = ax.get_legend_handles_labels()
    ax.legend(h[::-1], lb[::-1], loc="center", bbox_to_anchor=leg_loc, ncol=2, columnspacing=1.2)
    fig.subplots_adjust(left=0.12, right=0.97, top=0.97, bottom=0.085)
    return guardar(fig, nombre)


def doble_eje(ax, ylim, ystep, dec=0):
    ax2 = ax.twinx()
    ax2.spines[["top", "left"]].set_visible(False)
    ax2.spines["right"].set_visible(True)
    ax2.tick_params(axis="y", length=4, direction="out", pad=6)
    lo, hi = ylim
    ax2.set_ylim(lo, hi)
    ax2.yaxis.set_major_locator(FixedLocator([lo + i * ystep for i in range(round((hi - lo) / ystep) + 1)]))
    ax2.yaxis.set_major_formatter(coma(dec))
    return ax2


# ---------------------------------------------------------------- pptx

def rgb(hexa):
    return RGBColor.from_string(hexa.lstrip("#"))


def texto(slide, x, y, w, h, txt, size, color="#000000", bold=False, align=PP_ALIGN.LEFT,
          anchor=MSO_ANCHOR.TOP, font="Calibri"):
    tb = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, linea in enumerate(txt.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = linea
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.name = font
        r.font.color.rgb = rgb(color)
    return tb


def linea(slide, x1, y1, x2, y2, color, ancho=1.0):
    # rectangulo delgado: los conectores heredan la sombra del tema
    ln = slide.shapes.add_shape(1, Pt(x1), Pt(y1 - ancho / 2), Pt(x2 - x1), Pt(max(y2 - y1, ancho)))
    ln.fill.solid()
    ln.fill.fore_color.rgb = rgb(color)
    ln.line.fill.background()
    ln.shadow.inherit = False
    estilo = ln._element.find(qn("p:style"))
    if estilo is not None:
        ln._element.remove(estilo)
    return ln


def lamina(prs, num, titulo, fuente="Fuente: Banco Central de Chile"):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    texto(s, 36, 11, 820, 34, titulo, 24, ROJO_SAN, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    linea(s, 37, 57.5, 891, 57.5, GRIS_GRID, 1.0)
    s.shapes.add_picture(str(ASSETS / "logo.png"), Pt(910), Pt(0), Pt(34), Pt(44))
    linea(s, 41, 496, 919, 496, GRIS_GRID, 1.0)
    if fuente:
        texto(s, 44, 505, 400, 14, fuente, 9, "#404040")
    texto(s, 860, 512, 42, 20, str(num), 15, "#7F7F7F", bold=True, align=PP_ALIGN.RIGHT)
    return s


def grafico(slide, caja, png, titulo, tx=None):
    x, y, w, h = caja
    texto(slide, tx if tx is not None else x + 8, 97, w, 44, titulo, 14)
    slide.shapes.add_picture(str(png), Pt(x), Pt(y), Pt(w), Pt(h))


def portada(prs, fecha_txt):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(str(ASSETS / "portada.png"), Pt(11), Pt(10), Pt(938), Pt(520))
    s.shapes.add_picture(str(ASSETS / "logo_portada.png"), Pt(813), Pt(9), Pt(116), Pt(145))
    texto(s, 675, 232, 260, 100, "Estudios\nEconómicos", 36, ROJO_SAN)
    linea(s, 673, 334, 884, 334, "#1F4E79", 1.0)
    texto(s, 679, 339, 200, 20, fecha_txt, 14, ROJO_SAN)


# ---------------------------------------------------------------- main

def main(hoy=None):
    hoy = pd.Timestamp(hoy or dt.date.today())
    OUT.mkdir(exist_ok=True)

    # fin de eje = ultimo dato disponible, redondeado al mes siguiente
    def fin_de(*ss):
        return max(s.index.max() for s in ss) + pd.DateOffset(days=15)

    # --- lamina 2: IMCE
    precios = ["G089.IME.IND.A14.M", "G089.IME.IND.A34.M", "G089.IME.IND.A27.M"]
    costos = ["G089.IME.IND.A15.M", "G089.IME.IND.A35.M", "G089.IME.IND.A29.M"]
    fin = fin_de(*[serie(c) for c in precios + costos])
    p2a = g_imce(precios, "02a_imce_precios", fin)
    p2b = g_imce(costos, "02b_imce_costos", fin)

    # --- lamina 3: expectativas de inflacion EEE / EOF
    eee1, eee2 = desde(serie("F089.IPC.V12.14.M"), "2015-01-01"), desde(serie("F089.IPC.V12.15.M"), "2015-01-01")
    eof1, eof2 = desde(serie("F089.EOF.VII.12MS.D"), "2015-01-01"), desde(serie("F089.EOF.VII.S12M.D"), "2015-01-01")
    fin = fin_de(eee1, eof1)
    yt = ticks_fecha("2015-01-01", fin, 12)
    ylim = rango(eee1, eee2, eof1, eof2, paso=1, piso=1, techo=9)
    p3a = g_dos_lineas(eee1, eee2, "EEE 1 AÑO", "EEE 2 AÑOS", NAVY, ROJO_OSC, "03a_eee_inflacion",
                       "2015-01-01", fin, ylim, 1, yt, "anio", "(Var. anual, %)", (0.33, 0.77))
    p3b = g_dos_lineas(eof1, eof2, "EOF 1 AÑO", "EOF 2 AÑOS", NAVY, ROJO_OSC, "03b_eof_inflacion",
                       "2015-01-01", fin, ylim, 1, yt, "anio", "(Var. anual, %)", (0.33, 0.74), caja=DOS[1])

    # --- lamina 4: compensacion inflacionaria
    c1 = desde(serie("F022.SWSP.TAS.AN01.Z.Z.D"), "2022-01-01")
    c2 = desde(serie("F022.SWSP.TAS.AN02.Z.Z.D"), "2022-01-01")
    c11 = desde(serie("F022.SWFW.TAS.E1X1.Z.Z.D"), "2022-01-01")
    fin = fin_de(c1, c11)
    xt = ticks_fecha("2022-01-01", fin, 6)
    p4a = g_dos_lineas(c1, c2, "1 año", "2 años", ROJO, AZUL_ACERO, "04a_compensaciones", "2022-01-01", fin,
                       rango(c1, c2, paso=2, piso=0, techo=12), 2, xt, "punto", "(Var. anual, %)", (0.72, 0.64),
                       lw=1.3)
    fig, ax = nuevo_grafico(*DOS[1][2:])
    ax.plot(c11.index, c11.values, color=NAVY, lw=1.3)
    ejes(ax, rango(c11, paso=1, piso=0, techo=7), 1, "2022-01-01", fin, xt, "punto")
    unidad(ax, "(%)", 0.08, 0.93)
    fig.subplots_adjust(left=0.085, right=0.975, top=0.97, bottom=0.085)
    p4b = guardar(fig, "04b_swap_1en1")

    # --- lamina 5: EDEP
    ini = "2021-11-01"
    cl, pi, cmt, cf = (desde(serie(f"F089.EDEP.IND.88{k}.M"), ini) for k in ["CL", "PI", "CMT", "CF"])
    fin = fin_de(cl)
    p5a = g_multi([(cl, "Costos laborales", ROJO_OSC), (pi, "Precios de insumos", DORADO),
                   (cmt, "Costo de maquinarias o\ntécnología", VERDE_CLARO), (cf, "Costos fijos", CELESTE)],
                  "05a_edep_costos", ini, fin, rango(cl, pi, cmt, cf, paso=5, piso=40, techo=90), 5,
                  "(Índice de difusión)", (0.08, 0.13),
                  dict(loc="upper right", bbox_to_anchor=(1.02, 1.0), labelspacing=0.9), TRES[0])
    nv, pd_, mg, di, co = (desde(serie(f"F089.EDEP.IND.84IF{k}.M"), ini) for k in ["NV", "PD", "MG", "DI", "C"])
    p5b = g_multi([(nv, "Niveles de venta", NAVY), (pd_, "Precio del dólar", NARANJO),
                   (mg, "Márgenes de ganancias", VERDE_OSC), (di, "Disponibilidad de insumos", CELESTE),
                   (co, "Costos", MORADO)],
                  "05b_edep_influencia", ini, fin, rango(nv, pd_, mg, di, co, paso=10, piso=0, techo=90), 10,
                  "(Índice de difusión)", (0.1, 0.97),
                  dict(loc="lower right", bbox_to_anchor=(1.02, 0.0), labelspacing=1.0), TRES[1])
    cl3, pi3, cmt3, cf3 = (desde(serie(f"F089.EDEP.IND.90{k}.M"), ini) for k in ["CL", "PI", "CMT", "CF"])
    p5c = g_multi([(cl3, "Costos laborales", ROJO_OSC), (pi3, "Precios de insumos", DORADO),
                   (cmt3, "Costo de maquinarias o\ntécnología", VERDE_CLARO), (cf3, "Costos fijos", AZUL)],
                  "05c_epn_costos", ini, fin, rango(cl3, pi3, cmt3, cf3, paso=5, piso=40, techo=80), 5,
                  "(Índice de difusión)", (0.08, 0.13),
                  dict(loc="upper right", bbox_to_anchor=(1.02, 1.03), labelspacing=0.9), TRES[2])

    # --- lamina 6: EDEP expectativas de inflacion y GSCPI
    e12, e24 = desde(serie("F089.EDEP.POR.91EI12.M"), ini), desde(serie("F089.EDEP.POR.91EI24.M"), ini)
    fin = fin_de(e12)
    p6a = g_dos_lineas(e12, e24, "EDEP a 1 año", "EDEP a 2 años", AZUL_ACERO, NARANJO, "06a_edep_inflacion",
                       pd.Timestamp(ini) - pd.DateOffset(days=20), fin,
                       rango(e12, e24, paso=2, piso=0, techo=12), 2, ticks_fecha(ini, fin, 6), "punto",
                       "(Var. anual, %)", (0.42, 0.17))
    g = desde(serie("GSCPI"), "1998-01-01")
    fin_g = fin_de(g)
    fig, ax = nuevo_grafico(*DOS[1][2:])
    ax.plot(g.index, g.values, color=NAVY, lw=1.3)
    ax.axhline(0, color=GRIS_EJE, lw=0.75)
    ejes(ax, rango(g, paso=1, piso=-2, techo=5), 1, "1997-10-01", fin_g,
         ticks_fecha("1998-01-01", fin_g, 12), "anio", rot=90)
    unidad(ax, "(Desviación estandar del valor\npromedio)", 0.04, 0.96)
    fig.subplots_adjust(left=0.07, right=0.985, top=0.97, bottom=0.135)
    p6b = guardar(fig, "06b_gscpi")

    # --- lamina 7: TCN y TCM
    tcn, tcm = desde(serie("F073.TCO.PRE.Z.D"), "2018-01-01"), desde(serie("F073.TCM.IND.199502.D"), "2018-01-01")
    fin = fin_de(tcn)
    fig, ax = nuevo_grafico(*UNO[2:])
    ax.plot(tcn.index, tcn.values, color=ROJO_OSC, lw=1.3, label="TCN")
    ejes(ax, rango(tcn, paso=100, piso=500, techo=1100), 100, "2018-01-01", fin,
         ticks_fecha("2018-01-01", fin, 12), "anio")
    ax.spines["left"].set_visible(True)
    ax2 = doble_eje(ax, rango(tcm, paso=10, piso=90, techo=150), 10)
    ax2.plot(tcm.index, tcm.values, color=NAVY, lw=1.3, label="TCM (DERECHA)")
    unidad(ax, "(Pesos chilenos)", 0.012, 0.95)
    unidad(ax, "(Índice)", 0.93, 0.94, ha="center")
    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, loc="center", bbox_to_anchor=(0.68, 0.17), ncol=2, columnspacing=8)
    fig.subplots_adjust(left=0.06, right=0.94, top=0.97, bottom=0.085)
    p7 = guardar(fig, "07_tcn_tcm")

    # --- lamina 8: tasas SPC
    ini8 = f"{hoy.year - 1}-01-01"
    sp1, sp2 = desde(serie("F022.SPC.TPR.D360.NO.Z.D"), ini8), desde(serie("F022.SPC.TIN.AN02.NO.Z.D"), ini8)
    su1, su2 = desde(serie("F022.SPC.TIN.AN01.UF.Z.D"), ini8), desde(serie("F022.SPC.TIN.AN02.UF.Z.D"), ini8)
    fin = fin_de(sp1)
    xt = ticks_fecha(ini8, fin, 3)
    p8a = g_dos_lineas(sp2, sp1, "SPC pesos 2 años", "SPC pesos 1 año", NAVY, ROJO, "08a_spc_pesos", ini8, fin,
                       rango(sp1, sp2, paso=0.5, margen=0.1), 0.5, xt, "punto", "(%)", (0.67, 0.74), dec=1,
                       lw=1.3)
    p8b = g_dos_lineas(su2, su1, "SPC UF 2 años", "SPC UF 1 año", NAVY, ROJO, "08b_spc_uf", ini8, fin,
                       rango(su1, su2, paso=0.5, margen=0.1), 0.5, xt, "punto", "(%)", (0.5, 0.12), dec=1,
                       caja=DOS[1], lw=1.3)

    # --- lamina 9: EEE actividad (anio en curso y siguiente)
    a0, a1 = hoy.year, hoy.year + 1
    ini9 = f"{a0 - 1}-01-01"
    eee = {k: (desde(serie(f"F089.{k}.V12.{a0}.M"), ini9), desde(serie(f"F089.{k}.V12.{a1}.M"), ini9))
           for k in ["PIB", "CHOG", "FBKF"]}
    fin = max(s.index.max() for par in eee.values() for s in par)
    p9 = [g_eee(*eee[k], str(a0), str(a1), f"09{l}_eee_{k.lower()}", ini9, fin, caja, paso, leg)
          for k, l, caja, paso, leg in [("PIB", "a", TRES[0], 0.5, (0.33, 0.25)),
                                        ("CHOG", "b", TRES[1], 0.5, (0.33, 0.2)),
                                        ("FBKF", "c", TRES[2], 1.0, (0.33, 0.2))]]

    # --- lamina 10: IMCE expectativas de inversion vs Imacec no minero
    imce_inv = (serie("G089.IME.IND.A17.M") + serie("G089.IME.IND.A2B.M")) / 2
    imc = serie("F032.IMC.IND.Z.Z.EP18.N03.Z.0.M")
    imc_tri = (imc.rolling(3).sum().pct_change(12) * 100).dropna()
    imce_inv, imc_tri = desde(imce_inv, "2010-01-01"), desde(imc_tri, "2010-01-01")
    fin = fin_de(imce_inv, imc_tri)
    fig, ax = nuevo_grafico(*UNO[2:])
    ax.plot(imc_tri.index, imc_tri.values, color=DORADO, lw=2.0, label="Imacec No Minero")
    ejes(ax, rango(imc_tri, paso=5, piso=-20, techo=25), 5, "2009-11-01", fin,
         ticks_fecha("2010-01-01", fin, 12), "anio", rot=90)
    ax.spines["left"].set_visible(True)
    ax.axhline(0, color=GRIS_EJE, lw=0.75)
    ax2 = doble_eje(ax, rango(imce_inv, paso=10, piso=10, techo=80), 10)
    ax2.plot(imce_inv.index, imce_inv.values, color=NAVY, lw=2.0, label="IMCE expectativas de inversión (derecha)")
    unidad(ax, "(Var. Trimestral a/a)", 0.09, 0.97)
    unidad(ax, "(Índice)", 0.85, 0.97)
    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, loc="center", bbox_to_anchor=(0.32, 0.22), ncol=2, columnspacing=3)
    fig.subplots_adjust(left=0.045, right=0.955, top=0.97, bottom=0.12)
    p10 = guardar(fig, "10_imce_imacec")

    # ------------------------------------------------------------ presentacion
    prs = Presentation()
    prs.slide_width, prs.slide_height = Pt(960), Pt(540)
    portada(prs, f"{MESES_LARGO[hoy.month - 1]} de {hoy.year}")

    s = lamina(prs, 2, "IMCE")
    grafico(s, DOS[0], p2a, "IMCE: Indicador de Precios")
    grafico(s, DOS[1], p2b, "IMCE: Indicador de Costos")

    s = lamina(prs, 3, "Expectativas de Inflación")
    grafico(s, DOS[0], p3a, "Mediana EEE 1 y 2 años\n(Var. anual, %)")
    grafico(s, DOS[1], p3b, "Mediana EOF 1 y 2 años\n(Var. anual, %)")

    s = lamina(prs, 4, "Compensación Inflacionaria")
    grafico(s, DOS[0], p4a, "Compensaciones inflacionarias\n(Var. Anual, %)")
    grafico(s, DOS[1], p4b, "Compensación inflacionaria 1 en 1: Swap\n(%)")

    s = lamina(prs, 5, "EDEP")
    grafico(s, TRES[0], p5a, "EDEP: Evolución costos de la empresa\n(índice de difusión)", tx=36)
    grafico(s, TRES[1], p5b, "EDEP: Influencia sobre precio de venta\n(índice de difusión)", tx=348)
    grafico(s, TRES[2], p5c, "EPN: Evolución de los costos\n(índice de difusión, percepción evolución\n"
                             "últimos 3 meses)", tx=660)

    s = lamina(prs, 6, "IMCE, EDEP y Costos de transporte")
    grafico(s, DOS[0], p6a, "EDEP: Expectativas de Inflación\n(Var. Anual, %)")
    grafico(s, DOS[1], p6b, "Costos de transporte\n\n(desviación estándar del valor promedio)")
    texto(s, 505, 507, 400, 14, "Fuente: Reserva Federal de Nueva York", 9, "#404040")

    s = lamina(prs, 7, "TCN TCM")
    grafico(s, UNO, p7, "Tipo de cambio nominal spread Tipo de cambio multilateral\n(Pesos chilenos) (Índice)",
            tx=44)

    s = lamina(prs, 8, "Tasas SPC")
    grafico(s, DOS[0], p8a, "SPC pesos: 1 y 2 años\n(%)")
    grafico(s, DOS[1], p8b, "SPC UF 1 y 2 años\n(%)")

    s = lamina(prs, 9, "EEE")
    grafico(s, TRES[0], p9[0], "Expectativas de PIB\n(Var. Anual, %)", tx=36)
    grafico(s, TRES[1], p9[1], "Expectativas de Consumo de los hogares\n(Var. Anual, %)", tx=348)
    grafico(s, TRES[2], p9[2], "Expectativas FBCF\n(Var. Anual, %)", tx=660)

    s = lamina(prs, 10, "Imce expectativas v/s Imacec no minero")
    grafico(s, UNO, p10, "Imce expectativas v/s Imacec no minero\n(Índice) (Var. Trimestral a/a)", tx=44)

    nombre = f"{hoy:%Y%m%d}_Chartpack_Expectativas"
    pptx = OUT / f"{nombre}.pptx"
    prs.save(pptx)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(OUT), str(pptx)],
                   check=True, capture_output=True)
    print(pptx)
    print(OUT / f"{nombre}.pdf")


if __name__ == "__main__":
    main()
