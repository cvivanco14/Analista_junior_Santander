"""Minuta IPC: descarga los índices del INE, calcula los agregados con la matriz de volátiles
y genera el reporte en PDF.

Uso (desde minuta_ipc/Code, o con F5 en Spyder / VS Code):
    python minuta_ipc.py                    # descarga del INE y genera la minuta del último mes
    python minuta_ipc.py --sin-descarga     # usa los Excel ya guardados en Inputs/INE/
    python minuta_ipc.py --mes 2026-07      # genera la minuta de un mes anterior
    python minuta_ipc.py --sin-anexo        # omite el anexo de 172 subclases (más rápido)

Carpetas (relativas a este archivo, funcionan en cualquier computador):
    Code/          minuta_ipc.py, ipc.py, graficos_ipc.py, anexo_subclases.py
    Inputs/        insumos fijos (matriz_productos.csv, historico_bcch_indices.csv, ...)
    Inputs/INE/    Excel descargados del INE (se reemplazan en cada corrida)
    Outputs/       "Minuta IPC <mes> <año>.pdf"
    Outputs/figuras/  gráficos PNG de la minuta y del anexo (se regeneran)

Antes de cada minuta, completar EXPECTATIVAS (y COMENTARIO si corresponde).
"""

import argparse
import datetime as dt
import re
import sys
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (Image, KeepTogether, ListFlowable, ListItem, PageBreak, Paragraph,
                                SimpleDocTemplate, Table, TableStyle)
from reportlab.platypus.flowables import HRFlowable

# Carpetas del proyecto: Code/ (este archivo), Inputs/ y Outputs/
CODE = Path(__file__).resolve().parent
PROYECTO = CODE.parent
INPUTS = PROYECTO / "Inputs"
OUTPUTS = PROYECTO / "Outputs"
INSUMOS = INPUTS                 # insumos fijos (CSV)
DATOS = INPUTS / "INE"           # Excel descargados del INE
FIGURAS = OUTPUTS / "figuras"    # PNG intermedios
SALIDA = OUTPUTS                 # PDF de la minuta
sys.path.insert(0, str(CODE))    # para importar los módulos hermanos aunque se corra desde otra carpeta

import anexo_subclases  # noqa: E402
import graficos_ipc as g  # noqa: E402
import ipc  # noqa: E402

# ------------------------------------------------------------------ configuración (cambiar cada mes)
AUTOR = "Andrés Sansone"
# Expectativas de variación mensual del IPC para el mes del reporte, en % (ej. 0.3). None = sin dato.
EXPECTATIVAS = {
    "seguros": None,           # seguros de inflación (tarde del día anterior)
    "bloomberg": None,         # mediana encuesta Bloomberg
    "bloomberg_rango": None,   # (mín, máx) de la encuesta Bloomberg, ej. (0.1, 0.4)
    "eee": None,               # Encuesta de Expectativas Económicas (BCCh)
    "eof": None,               # Encuesta de Operadores Financieros (BCCh)
}
EXPECTATIVA_SV = None          # EEE para el IPC sin volátiles, en % m/m
COMENTARIO = ""                # párrafo libre al final de los destacados (opcional)
INICIO_GRAFICOS = "2019-01"

MESES_LARGOS = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
                "septiembre", "octubre", "noviembre", "diciembre"]
AZUL_OXFORD = colors.Color(0.0, 0.13, 0.28)


# ------------------------------------------------------------------ formato
def num(x, d=1):
    s = f"{x:,.{d}f}"
    if float(s.replace(",", "")) == 0:
        s = s.lstrip("-")  # evita "-0,0"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def pct(x, d=1):
    return f"{num(x, d)}%"


def mes_largo(p):
    return MESES_LARGOS[p.month - 1]


def mes_corto(p):
    return f"{MESES_LARGOS[p.month - 1][:3]}-{str(p.year)[2:]}"


def verbo(x, sube, baja, igual, d=1):
    x = round(x, d)
    return sube if x > 0 else baja if x < 0 else igual


def comparar(real, esperado):
    dif = round(real, 1) - esperado
    if dif >= 0.1 - 1e-9:
        return "por sobre"
    if dif >= 0.05 - 1e-9:
        return "levemente por sobre"
    if dif <= -0.1 + 1e-9:
        return "por debajo de"
    if dif <= -0.05 + 1e-9:
        return "levemente por debajo de"
    return "en línea con"


def lista(partes):
    return partes[0] if len(partes) == 1 else ", ".join(partes[:-1]) + " y " + partes[-1]


def comillas(glosa):
    return f"“{glosa.strip().lower()}”"


# ------------------------------------------------------------------ cálculos del mes
class Mes:
    def __init__(self, d, p):
        self.d, self.p = d, p
        self.mm = {a: ipc.variacion(d, a, 1) for a in ipc.AGREGADOS}
        self.aa = {a: ipc.variacion(d, a, 12) for a in ipc.AGREGADOS}
        gen = d["general_ine"]
        # IPC total: variaciones oficiales del INE. En 2024 la a/a se calcula contra la serie
        # referencial 2023 (misma canasta), igual que los agregados; desde 2025 coincide con la oficial.
        hist = lambda s: s[s.index <= ipc.CAMBIO_CANASTA]  # noqa: E731
        mm_of = gen["var_mm"].dropna()
        self.mm["IPC"] = pd.concat([hist(self.mm["IPC"]), mm_of[mm_of.index > ipc.CAMBIO_CANASTA]])
        aa_ref = (gen["indice"] / gen["indice"].shift(12, freq="M") - 1) * 100
        aa_ref = aa_ref[(aa_ref.index > ipc.CAMBIO_CANASTA) & (aa_ref.index.year == 2024)].dropna()
        aa_of = gen["var_aa"].dropna()
        self.aa["IPC"] = pd.concat([hist(self.aa["IPC"]), aa_ref, aa_of[aa_of.index.year >= 2025]])
        self.inc = {a: ipc.incidencia(d, a, 1) for a in ipc.AGREGADOS}
        self.prod = ipc.productos(d, p)
        self.cba = ipc.cba(d)

    def acumulado(self, a):
        if a == "IPC":
            return self.d["general_ine"]["var_acum"][self.p]
        I = self.d["agregados"][a]
        return (I[self.p] / I[pd.Period(f"{self.p.year - 1}-12", "M")] - 1) * 100


# ------------------------------------------------------------------ texto
def expectativas_texto(real):
    e = {k: v for k, v in EXPECTATIVAS.items() if v is not None and k != "bloomberg_rango"}
    if not e:
        return None, None
    valores = list(e.values())
    lo, hi = min(valores), max(valores)
    rango = pct(lo) if round(lo, 1) == round(hi, 1) else f"{pct(lo)}-{pct(hi)}"
    comp = comparar(real, lo) if round(real, 1) <= lo else comparar(real, hi) if round(real, 1) >= hi else "en línea con"
    partes = []
    if "seguros" in e:
        partes.append(f"los seguros de inflación la situaban en {pct(e['seguros'], 2)}")
    if "bloomberg" in e:
        r = EXPECTATIVAS.get("bloomberg_rango")
        partes.append(f"la encuesta Bloomberg en {pct(e['bloomberg'])}"
                      + (f" (rango {pct(r[0])}-{pct(r[1])})" if r else ""))
    if "eee" in e:
        partes.append(f"la EEE en {pct(e['eee'])}")
    if "eof" in e:
        partes.append(f"la EOF en {pct(e['eof'])}")
    return f"{comp} las expectativas de mercado ({rango})", (
        f"La cifra se ubicó {comp} las expectativas: {lista(partes)} m/m.")


def titulo(m: Mes):
    p = m.p
    v, v_sv = m.mm["IPC"][p], m.mm["IPC sin volátiles"][p]
    comp, _ = expectativas_texto(v)
    t = (f"IPC {mes_largo(p)} {p.year} | El IPC "
         + verbo(v, f"aumentó {pct(abs(v))}", f"cayó {pct(abs(v))}", "no varió") + " mensual"
         + (f", {comp}" if comp else "") + ", mientras que, al descontarle los componentes volátiles, "
         + verbo(v_sv, f"creció {pct(abs(v_sv))}", f"cayó {pct(abs(v_sv))}", "no varió") + " m/m")
    if EXPECTATIVA_SV is not None:
        t += f", {comparar(v_sv, EXPECTATIVA_SV)} lo esperado por la EEE ({pct(EXPECTATIVA_SV)})"
    return t + "."


def destacados(m: Mes):
    p, pa = m.p, m.p - 1
    mm = lambda a, q=p: m.mm[a][q]  # noqa: E731
    aa = lambda a, q=p: m.aa[a][q]  # noqa: E731

    def mov(a, plural=False, q=p):  # "creció 0,5% m/m (2,3% a/a)"
        sube, baja, igual = ("crecieron", "cayeron", "no variaron") if plural else ("creció", "cayó", "no varió")
        return (verbo(mm(a, q), f"{sube} {pct(abs(mm(a, q)))}", f"{baja} {pct(abs(mm(a, q)))}", igual)
                + f" m/m ({pct(aa(a, q))} a/a)")

    b = []
    b.append(f"<b>La inflación anual se ubicó en {pct(aa('IPC'))} en {mes_largo(p)}</b>, "
             + verbo(aa("IPC") - aa("IPC", pa), "aumentando", "disminuyendo", "manteniéndose")
             + f" respecto de {mes_largo(pa)} ({pct(aa('IPC', pa))}). El IPC sin volátiles pasó de un "
             f"{pct(aa('IPC sin volátiles', pa))} anual en {mes_largo(pa)} a un {pct(aa('IPC sin volátiles'))} "
             f"en {mes_largo(p)}.")

    _, frase_exp = expectativas_texto(mm("IPC"))
    b.append(f"<b>El IPC total registró "
             + verbo(mm("IPC"), "un crecimiento", "una disminución", "una variación nula")
             + f" de {pct(mm('IPC'))} m/m en {mes_largo(p)}, alcanzando {pct(aa('IPC'))} a/a.</b>"
             + (f" {frase_exp}" if frase_exp else ""))

    t = (f"<b>El IPC sin volátiles registró una variación de {pct(mm('IPC sin volátiles'))} m/m y "
         f"{pct(aa('IPC sin volátiles'))} a/a</b> ({pct(aa('IPC sin volátiles', pa))} en {mes_largo(pa)})")
    if EXPECTATIVA_SV is not None:
        t += f", {comparar(mm('IPC sin volátiles'), EXPECTATIVA_SV)} lo esperado por la EEE ({pct(EXPECTATIVA_SV)} m/m)"
    b.append(t + f". Entre sus componentes, los bienes sin volátiles {mov('Bienes sin volátiles', True)} y los "
             f"servicios sin volátiles {mov('Servicios sin volátiles', True)}. Los servicios sin volátiles excluyendo "
             f"administrados e indexados {mov('Servicios_ Resto_sv', True)}.")

    v_vol = mm("IPC volátiles")
    b.append(f"<b>El IPC volátiles "
             + verbo(v_vol, f"creció {pct(abs(v_vol))}", f"cayó {pct(abs(v_vol))}", "no varió")
             + f" m/m y registró {pct(aa('IPC volátiles'))} a/a</b> ({pct(aa('IPC volátiles', pa))} en "
             f"{mes_largo(pa)}). Entre sus componentes, los alimentos volátiles {mov('Alim. Volát.', True)}, "
             f"la energía volátil {mov('Energía Volátiles')} y el resto de volátiles {mov('Resto volátiles')}.")

    b.append(f"Los alimentos no volátiles registraron una variación de {pct(aa('Alimentos no volátiles'))} a/a "
             f"y {pct(mm('Alimentos no volátiles'))} m/m, mientras que los alimentos volátiles variaron "
             f"{pct(aa('Alim. Volát.'))} a/a y {pct(mm('Alim. Volát.'))} m/m.")

    pr = m.prod
    alim = pr[pr.subsubcategoria.isin(["Alimentos no volátiles", "Alim. Resto_v", "FyV (proxy)_v"])]
    suben = alim[alim.var_mm.round(1) > 0].sort_values("var_mm", ascending=False).head(5)
    bajan = alim[alim.var_mm.round(1) < 0].sort_values("var_mm").head(4)
    fr = lambda df: lista([f"{comillas(r.glosa)} ({pct(r.var_mm)} m/m)" for r in df.itertuples()])  # noqa: E731
    t = ""
    if len(suben):
        t += f"Los alimentos que más aumentaron de precio en {mes_largo(p)} fueron {fr(suben)}. "
    if len(bajan):
        t += f"Por su parte, los que más disminuyeron fueron {fr(bajan)}."
    if t:
        b.append(t.strip())

    en = pr[pr.subcategoria == "Energía volátiles"].copy()
    en = en.reindex(en.var_aa.abs().sort_values(ascending=False).index).head(4)
    b.append(f"La energía volátil registró una variación de {pct(aa('Energía Volátiles'))} a/a y "
             f"{pct(mm('Energía Volátiles'))} m/m. Entre sus productos, "
             + lista([f"{comillas(r.glosa)} varió {pct(r.var_aa)} a/a ({pct(r.var_mm)} m/m)" for r in en.itertuples()])
             + ".")

    alza = pr[pr.inc_mm > 0].sort_values("inc_mm", ascending=False).head(3)
    baja = pr[pr.inc_mm < 0].sort_values("inc_mm").head(3)
    fi = lambda df: lista([f"{comillas(r.glosa)} ({pct(r.var_mm)} m/m; {num(r.inc_mm, 3)} pp.)"  # noqa: E731
                           for r in df.itertuples()])
    b.append(f"Los productos que más contribuyeron al alza fueron {fi(alza)}. Por su parte, los que más "
             f"contribuyeron a la baja fueron {fi(baja)}.")

    dif = ipc.difusion(m.d, "IPC")
    b.append(f"En términos de difusión, el {pct(dif[p])} de los productos registró una variación mensual "
             f"positiva en {mes_largo(p)}, versus {pct(dif[pa])} en {mes_largo(pa)}.")

    c = m.cba
    c_mm, c_mm_a = (c[p] / c[pa] - 1) * 100, (c[pa] / c[pa - 1] - 1) * 100
    c_aa, c_aa_a = (c[p] / c[p - 12] - 1) * 100, (c[pa] / c[pa - 12] - 1) * 100
    b.append(f"La Canasta Básica de Alimentos (CBA, metodología 2024) se ubicó en ${num(c[p], 0)} por persona, con una variación "
             f"de {pct(c_mm)} m/m ({pct(c_mm_a)} en {mes_largo(pa)}) y de {pct(c_aa)} a/a ({pct(c_aa_a)} en "
             f"{mes_largo(pa)}), equivalente a "
             + verbo(c[p] - c[p - 12], "un aumento", "una disminución", "una variación", 0)
             + f" de ${num(abs(c[p] - c[p - 12]), 0)} en doce meses.")
    if COMENTARIO:
        b.append(COMENTARIO)
    return b


# ------------------------------------------------------------------ cuadros
FILAS_CUADRO = [  # (etiqueta, agregado, sangría)
    ("<b>IPC</b>", "IPC", 0), ("<b>IPC sin volátiles</b>", "IPC sin volátiles", 0),
    ("Bienes sin volátiles", "Bienes sin volátiles", 1), ("Alimentos no volátiles", "Alimentos no volátiles", 2),
    ("Bienes SV exc. alimentos", "Bienes sv. ex. alim.", 2), ("Servicios sin volátiles", "Servicios sin volátiles", 1),
    ("Administrados e indexados", "Servicios_A&I_sv", 2), ("Resto (exc. A&I)", "Servicios_ Resto_sv", 2),
    ("<b>IPC volátiles</b>", "IPC volátiles", 0), ("Energía volátiles", "Energía Volátiles", 1),
    ("Electricidad", "Electricidad_v", 2), ("Combustibles", "Combustible_v", 2),
    ("Alimentos volátiles", "Alim. Volát.", 1), ("Alimentos exc. FyV", "Alim. Resto_v", 2),
    ("Frutas y verduras", "FyV (proxy)_v", 2), ("Resto volátiles", "Resto volátiles", 1),
    ("Bienes volátiles", "Bienes volát.", 2), ("Servicios volátiles", "Servicios volát.", 2),
    ("Administrados e indexados", "Servicios_A&I_v", 3), ("Resto (exc. A&I)", "Servicios_Resto_v", 3),
]


def cuadro_agregados(m: Mes):
    p, pa = m.p, m.p - 1
    enc = ["", "Peso", f"m/m\n{mes_corto(pa)}", f"m/m\n{mes_corto(p)}", f"Incid. m/m\n{mes_corto(p)}",
           f"a/a\n{mes_corto(pa)}", f"a/a\n{mes_corto(p)}", f"Acum.\n{p.year}"]
    filas = [enc]
    for etq, a, s in FILAS_CUADRO:
        inc = sum(m.inc[x][p] for x in ("IPC sin volátiles", "IPC volátiles")) if a == "IPC" else m.inc[a][p]
        filas.append(["&nbsp;" * 4 * s + etq, num(m.d["pesos"][a]), pct(m.mm[a][pa]), pct(m.mm[a][p]),
                       num(inc, 2), pct(m.aa[a][pa]), pct(m.aa[a][p]), pct(m.acumulado(a))])
    c = m.cba
    dic = pd.Period(f"{p.year - 1}-12", "M")
    filas.append(["CBA ($ por persona)", f"${num(c[p], 0)}", pct((c[pa] / c[pa - 1] - 1) * 100), pct((c[p] / c[pa] - 1) * 100),
                  "-", pct((c[pa] / c[pa - 12] - 1) * 100), pct((c[p] / c[p - 12] - 1) * 100),
                  pct((c[p] / c[dic] - 1) * 100)])
    titulo = "Cuadro 1: Inflación mensual, anual e incidencias (%, pp.) y Canasta Básica de Alimentos."
    nota = ("Nota: agregados calculados con los índices por producto del INE (canasta 2023=100) y la clasificación "
            "de productos volátiles del Banco Central; A&I = administrados e indexados. Las cifras del IPC total "
            "corresponden a las publicadas por el INE. La CBA replica la metodología 2024 del MDSF ($66.896 a marzo "
            "de 2022, actualizada con el IPC de alimentos dentro y fuera del hogar); puede diferir en pocos pesos de "
            "la cifra oficial por redondeo de los índices. Fuente: INE, MDSF y Banco Central de Chile.")
    return titulo, filas, nota


def cuadro_productos(m: Mes):
    pr = m.prod
    top = pr.reindex(pr.inc_mm.abs().sort_values(ascending=False).index).head(15).sort_values("inc_mm", ascending=False)
    filas = [["Categoría", "Producto", "Pond.", "var. m/m", "inc. m/m"]]
    for r in top.itertuples():
        filas.append([r.subcategoria, r.glosa.capitalize(), num(r.ponderacion, 2), pct(r.var_mm), num(r.inc_mm, 3)])
    titulo = (f"Cuadro 2: 15 productos con mayor incidencia absoluta en la variación mensual del IPC, "
              f"{mes_largo(m.p)} {m.p.year} (%, pp.)")
    return titulo, filas, "Fuente: INE."


# ------------------------------------------------------------------ figuras
def generar_figuras(m: Mes):
    FIGURAS.mkdir(parents=True, exist_ok=True)
    for f in FIGURAS.glob("fig*.png"):
        f.unlink()
    p, ini = m.p, pd.Period(INICIO_GRAFICOS, "M")
    rng = lambda s: s.loc[ini:p]  # noqa: E731
    figs = []
    nueva = lambda: FIGURAS / f"fig{len(figs) + 1:02d}.png"  # noqa: E731
    t = lambda x: f"Gráfico {len(figs) + 1}: {x}"  # noqa: E731
    nota_bc = "Fuente: INE y Banco Central de Chile."
    cc = ipc.CAMBIO_CANASTA

    figs.append(g.lineas_barras(nueva(), t("Inflación total"), "(var. % a/a y m/m)",
                                {"Var. % a/a": rng(m.aa["IPC"])}, rng(m.mm["IPC"]), nota_bc, cc))

    cats = {"IPC total": "IPC", "Alim. volátiles": "Alim. Volát.", "Bienes SV": "Bienes sin volátiles",
            "Energía volátil": "Energía Volátiles", "Resto volátiles": "Resto volátiles",
            "Servicios SV": "Servicios sin volátiles"}
    var = pd.Series({k: m.mm[a][p] for k, a in cats.items()})
    inc = pd.Series({k: (var["IPC total"] if a == "IPC" else m.inc[a][p]) for k, a in cats.items()})
    figs.append(g.barras_mes(nueva(), t("IPC y agregados con y sin volátiles"),
                             f"(var. % m/m e incidencias en pp., {mes_largo(p)} {p.year})", var, inc))

    comp = {"Bienes SV": "Bienes sin volátiles", "Servicios SV": "Servicios sin volátiles",
            "Energía volátil": "Energía Volátiles", "Alim. volátiles": "Alim. Volát.",
            "Resto volátiles": "Resto volátiles"}
    figs.append(g.incidencias_apiladas(nueva(), t("Incidencias en la inflación anual"), "(var. % a/a e incidencias en pp.)",
                                       rng(m.aa["IPC"]), {k: rng(ipc.incidencia(m.d, a, 12)) for k, a in comp.items()},
                                       linea_canasta=cc))
    figs.append(g.lineas_barras(nueva(), t("IPC total e IPC sin volátiles"), "(var. % a/a)",
                                {"IPC total": rng(m.aa["IPC"]), "IPC sin volátiles": rng(m.aa["IPC sin volátiles"])},
                                nota=nota_bc, linea_canasta=cc))
    figs.append(g.lineas_barras(nueva(), t("IPC servicios sin volátiles"), "(var. % a/a)",
                                {"Servicios SV": rng(m.aa["Servicios sin volátiles"]),
                                 "Administrados e indexados": rng(m.aa["Servicios_A&I_sv"]),
                                 "Resto (exc. A&I)": rng(m.aa["Servicios_ Resto_sv"])}, nota=nota_bc, linea_canasta=cc))
    c = m.cba
    figs.append(g.lineas_barras(nueva(), t("Canasta Básica de Alimentos"), "(costo en $)",
                                {"CBA": rng(c)}, nota="Fuente: elaboración propia con datos del INE y MDSF.", dec=0))
    figs.append(g.lineas_barras(nueva(), t("Canasta Básica de Alimentos"), "(var. % a/a y m/m)",
                                {"Var. % a/a": rng((c / c.shift(12) - 1) * 100)}, rng((c / c.shift(1) - 1) * 100),
                                nota="Fuente: elaboración propia con datos del INE y MDSF."))

    hist = m.d["dif_hist"]
    for a, nombre in [("IPC", "IPC total"), ("IPC sin volátiles", "IPC sin volátiles"),
                      ("Bienes sin volátiles", "IPC bienes sin volátiles"),
                      ("Servicios sin volátiles", "IPC servicios sin volátiles"),
                      ("Alimentos no volátiles", "IPC alimentos no volátiles"),
                      ("Bienes sv. ex. alim.", "IPC bienes sin volátiles exc. alimentos"),
                      ("Servicios_A&I_sv", "IPC servicios sin volátiles A&I"),
                      ("Servicios_ Resto_sv", "IPC servicios sin volátiles exc. A&I")]:
        h = hist[hist.agregado == a].pivot(index="mes", columns="estadistico", values="valor")
        dif = ipc.difusion(m.d, a).loc[:p]
        por_anio = lambda y: dif[dif.index.year == y].set_axis(dif[dif.index.year == y].index.month)  # noqa: E731
        figs.append(g.difusion(nueva(), t(nombre), h, por_anio(p.year - 1), por_anio(p.year), p.year))
    return figs


# ------------------------------------------------------------------ PDF
E = {
    "titulo": ParagraphStyle("titulo", fontName="Helvetica-Bold", fontSize=14, leading=18, textColor=AZUL_OXFORD),
    "fecha": ParagraphStyle("fecha", fontName="Helvetica-Bold", fontSize=10.5, spaceBefore=6),
    "autor": ParagraphStyle("autor", fontName="Helvetica-Oblique", fontSize=9, spaceBefore=2),
    "h2": ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=13, textColor=AZUL_OXFORD, spaceBefore=8,
                         spaceAfter=6),
    "cuerpo": ParagraphStyle("cuerpo", fontName="Helvetica", fontSize=10, leading=13.5, alignment=TA_JUSTIFY),
    "celda": ParagraphStyle("celda", fontName="Helvetica", fontSize=7.8, leading=9.2),
    "celda_c": ParagraphStyle("celda_c", fontName="Helvetica-Bold", fontSize=7.8, leading=9.2, alignment=TA_CENTER),
    "caption": ParagraphStyle("caption", fontName="Helvetica", fontSize=9, leading=11, alignment=TA_CENTER,
                              spaceBefore=10, spaceAfter=4),
    "nota": ParagraphStyle("nota", fontName="Helvetica", fontSize=7.3, leading=9, spaceBefore=3,
                           alignment=TA_JUSTIFY),
}


def esc(texto):
    """Escapa '&' para reportlab (A&I), respetando entidades ya escritas."""
    return re.sub(r"&(?!(amp|nbsp|lt|gt);)", "&amp;", str(texto))


def tabla_pdf(titulo, filas, nota, ancho, anchos):
    datos = [[Paragraph(esc(c).replace("\n", "<br/>"), E["celda_c"]) for c in filas[0]]]
    datos += [[Paragraph(esc(c), E["celda"]) if j < 2 and not str(c)[:1].isdigit() and not str(c).startswith("$")
               else c for j, c in enumerate(f)] for f in filas[1:]]
    t = Table(datos, colWidths=[ancho * a for a in anchos], repeatRows=1)
    estilo = [("FONT", (0, 0), (-1, -1), "Helvetica", 7.8), ("ALIGN", (1, 1), (-1, -1), "CENTER"),
              ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LINEABOVE", (0, 0), (-1, 0), 1, colors.black),
              ("LINEBELOW", (0, 0), (-1, 0), 0.6, colors.black), ("LINEBELOW", (0, -1), (-1, -1), 1, colors.black),
              ("TOPPADDING", (0, 0), (-1, -1), 1.8), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.8)]
    estilo += [("BACKGROUND", (0, i), (-1, i), colors.HexColor("#EEEEEE")) for i in range(2, len(datos), 2)]
    t.setStyle(TableStyle(estilo))
    return KeepTogether([Paragraph(esc(titulo), E["caption"]), t, Paragraph(esc(nota), E["nota"])])


def pares(rutas, ancho):
    w = ancho / 2 - 0.1 * cm
    filas = [[Image(str(r), width=w, height=w * 4.6 / 6.6) for r in rutas[i:i + 2]] for i in range(0, len(rutas), 2)]
    if len(filas[-1]) == 1:
        filas[-1].append("")
    t = Table(filas, colWidths=[ancho / 2] * 2)
    t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    return t


def anexo(figs_sub, ancho):
    nota = ("Cada página muestra dos subclases de la canasta 2023, en el orden de la estructura del IPC. "
            "Perfil: variación mensual por mes calendario, con el promedio 2013-2019 y una banda de ±1 desviación "
            "estándar (solo subclases con al menos cuatro años de datos en ese período). Incidencia en pp. de la "
            "variación mensual del IPC. Clasificación analítica según la matriz de volátiles del Banco Central, por "
            "mayoría de ponderación de los productos de la subclase. Series: empalme de subclases del Banco "
            "Central (1989-2023) y, desde enero de 2024, índices por subclase del INE base 2023. "
            "Fuente: INE y Banco Central de Chile.")
    w = ancho * 0.86  # dos subclases por página, también en la primera (con título y nota)
    h = [PageBreak(), Paragraph("Anexo: gráficos por subclase", E["h2"]), Paragraph(esc(nota), E["nota"])]
    h += [Image(str(r), width=w, height=w * 6.4 / 10, hAlign="CENTER") for r in figs_sub]
    return h


def armar_pdf(m: Mes, figs, ruta, figs_sub=()):
    izq, der = 2 * cm, 1.84 * cm
    ancho = letter[0] - izq - der

    def pagina(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 9)
        canvas.drawCentredString(letter[0] / 2, 1.5 * cm, str(doc.page))
        canvas.setFillColor(colors.gray)
        canvas.setFont("Helvetica", 8.5)
        canvas.drawString(izq, 1.5 * cm, AUTOR)
        canvas.restoreState()

    doc = SimpleDocTemplate(str(ruta), pagesize=letter, leftMargin=izq, rightMargin=der, topMargin=2.2 * cm,
                            bottomMargin=2.3 * cm, title=f"Minuta IPC {mes_largo(m.p)} {m.p.year}", author=AUTOR)
    hoy = dt.date.today()
    nota_graf = ("Las series hasta diciembre de 2023 corresponden a los agregados empalmados del Banco Central "
                 "(canasta 2018); desde enero de 2024, a la canasta 2023=100 del INE con variaciones anuales "
                 "respecto de la serie referencial 2023. La línea punteada marca el cambio de canasta.")
    nota_dif = ("La banda y el promedio 2019-2022 corresponden a la difusión de la canasta 2018 (subclases); "
                "las líneas y puntos, a los productos de la canasta 2023.")
    h = [
        Paragraph(titulo(m), E["titulo"]),
        Paragraph(f"{hoy.day} de {MESES_LARGOS[hoy.month - 1]} de {hoy.year}", E["fecha"]),
        Paragraph(f"Autor: {AUTOR}", E["autor"]),
        HRFlowable(width="100%", thickness=0.5, color=colors.black, spaceBefore=4, spaceAfter=2),
        Paragraph("Destacados", E["h2"]),
        ListFlowable([ListItem(Paragraph(esc(x), E["cuerpo"]), leftIndent=12) for x in destacados(m)],
                     bulletType="bullet", start="•", leftIndent=12, bulletFontSize=9),
        PageBreak(),
        Paragraph("Cuadros", E["h2"]),
        tabla_pdf(*cuadro_agregados(m), ancho, [0.27, 0.09, 0.09, 0.09, 0.1, 0.09, 0.09, 0.09]),
        tabla_pdf(*cuadro_productos(m), ancho, [0.22, 0.48, 0.1, 0.1, 0.1]),
        PageBreak(),
        Paragraph("Gráficos", E["h2"]),
        pares(figs[0:4], ancho),
        Paragraph(nota_graf, E["nota"]),
        PageBreak(),
        pares(figs[4:7], ancho),
        PageBreak(),
        Paragraph("Difusión inflacionaria", E["h2"]),
        pares(figs[7:15], ancho),
        Paragraph(nota_dif, E["nota"]),
    ]
    if figs_sub:
        h += anexo(figs_sub, ancho)
    doc.build(h, onFirstPage=pagina, onLaterPages=pagina)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--sin-descarga", action="store_true", help="usar los Excel ya guardados en Inputs/INE/")
    ap.add_argument("--mes", help="mes del reporte AAAA-MM (por defecto, el último publicado)")
    ap.add_argument("--sin-anexo", action="store_true", help="omitir el anexo de gráficos por subclase")
    args = ap.parse_args()
    faltan = [f for f in ("matriz_productos.csv", "historico_bcch_indices.csv", "pesos_canasta_2018.csv",
                          "difusion_historica.csv", "ipc_base2018_alimentos.csv", "subclases_bcch_1989_2023.csv")
              if not (INSUMOS / f).is_file()]
    if faltan:
        sys.exit(f"Faltan insumos en {INSUMOS}:\n    " + "\n    ".join(faltan))
    if not args.sin_descarga:
        print("Descargando índices del INE...")
        ipc.descargar(DATOS)
    d = ipc.cargar(DATOS, INSUMOS)
    p = pd.Period(args.mes, "M") if args.mes else d["agregados"].index[-1]
    print(f"Mes del reporte: {mes_largo(p)} {p.year}")
    if all(v is None for v in EXPECTATIVAS.values()):
        print("Aviso: EXPECTATIVAS está vacío; la minuta no compara con expectativas de mercado.")
    m = Mes(d, p)
    figs = generar_figuras(m)
    figs_sub = []
    if not args.sin_anexo:
        print("Generando anexo de subclases...")
        figs_sub = anexo_subclases.generar(FIGURAS / "subclases", DATOS, INSUMOS, d["general_ine"]["indice"], p)
    SALIDA.mkdir(parents=True, exist_ok=True)
    ruta = SALIDA / f"Minuta IPC {mes_largo(p)} {p.year}.pdf"
    armar_pdf(m, figs, ruta, figs_sub)
    print(f"Listo: {ruta}")


if __name__ == "__main__":
    main()
