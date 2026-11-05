# -*- coding: utf-8 -*-
r"""Minuta de Remuneraciones (INE): PPTX con gráficos nativos (editables), PDF y Excel de salida.

Láminas: portada, resumen (textos automáticos del último dato), remuneraciones y costo laboral, y
masa salarial real (variación anual con contribuciones de ocupados, horas y salario real, y nivel).
Formato de formato_graficos.py (Administrativos\Manual de Estilo), como la minuta IMCE.

Documento: Minuta_Remuneraciones_<mes>_<año>.pdf (reportlab), con colores Santander: título con el
resultado del mes, principales resultados (IR, ICL, salarios reales y masa salarial), cuadro de
variaciones m/m (sa) y a/a, y cuatro figuras (IR real, variación anual de los índices y variación m/m
del IR nominal y real por mes calendario frente a su rango histórico).

Carpetas (relativas a este script; sirven para cualquier usuario):
    <Remuneraciones>\Code\      datos_remuneraciones.py y minuta_remuneraciones.py
    <Remuneraciones>\Inputs\    archivos del INE (los deja datos_remuneraciones.py)
    <Remuneraciones>\Outputs\   Masa_salarial_Base.xlsx y Cifras_empleo.xlsx (bases) y
                                AAAAMMDD_Minuta_Remuneraciones.pptx / .pdf / _datos.xlsx,
                                Minuta_Remuneraciones_<mes>_<año>.pdf (documento) y Masa_salarial.png

Uso (no hay que editar nada):
    python datos_remuneraciones.py     # descarga del INE y actualiza los Excel de bases
    python minuta_remuneraciones.py    # arma la minuta en Outputs (con los archivos de Inputs)
"""
import datetime as dt
import math
import os
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

import datos_remuneraciones as dr

CODE = Path(__file__).resolve().parent
OUT = dr.OUT
PNG = OUT / "Masa_salarial.png"
MANUAL_ESTILO = Path("Administrativos") / "Manual de Estilo"
BIBLIOTECA = Path("Estudios - Documentos") / "General"
ICLOUD = Path.home() / "Library" / "Mobile Documents" / "com~apple~CloudDocs"

TITULO = "Remuneraciones"
AUTOR = "Andrés Sansone"   # aparece en la portada, en el pie de las láminas y en el documento
FUENTE = "INE (Índice de Remuneraciones y Encuesta Nacional de Empleo)"
NOTA_MASA = ("Masa salarial real = ocupados (desestacionalizados por el INE) × horas habituales semanales × "
             "índice de remuneraciones real (ambos desestacionalizados con STL). Contribuciones: descomposición "
             "logarítmica de la variación anual. Periodo: último mes del trimestre móvil.")
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre",
         "octubre", "noviembre", "diciembre"]
DESDE = "2015-01-01"

# series de la hoja 'remuneraciones' (originales; la variación a/a se calcula sobre ellas)
IR, ICL, IR_REAL = "IR", "ICL", "IR Real"


# ─────────────────────────────────────────────────────────────────────────────
# formato_graficos.py (se busca como en la minuta IMCE)
# ─────────────────────────────────────────────────────────────────────────────

def _raices_sincronizadas():
    casa = Path.home()
    raices = []
    for base in (casa, casa / "Library" / "CloudStorage", ICLOUD):
        if base.is_dir():
            for raiz in sorted(base.iterdir()):
                if raiz.is_dir() and re.search(r"santander|onedrive|sharepoint", raiz.name, re.I):
                    raices.append(raiz)
    return raices


def _formato_en(raiz, profundidad=6):
    rutas = [raiz / BIBLIOTECA / MANUAL_ESTILO, raiz / "General" / MANUAL_ESTILO, raiz / MANUAL_ESTILO]
    try:
        subs = sorted(p for p in raiz.iterdir() if p.is_dir() and not p.name.startswith("."))
    except OSError:
        subs = []
    for sub in subs:
        rutas += [sub / BIBLIOTECA / MANUAL_ESTILO, sub / "General" / MANUAL_ESTILO, sub / MANUAL_ESTILO]
    for r in rutas:
        if (r / "formato_graficos.py").is_file():
            return r
    if ICLOUD not in raiz.parents:
        return None
    base = len(raiz.parts)
    for carpeta, dirs, archivos in os.walk(raiz):
        dirs[:] = [d for d in dirs if not d.startswith(".")] if len(Path(carpeta).parts) - base < profundidad else []
        if "formato_graficos.py" in archivos:
            return Path(carpeta)
        if ".formato_graficos.py.icloud" in archivos:
            sys.exit(f"formato_graficos.py está en iCloud pero no descargado en este Mac:\n    {carpeta}\n"
                     "Ábrelo en Finder y elige 'Descargar ahora' (o 'Mantener descargado'), y vuelve a correr.")
    return None


def _carpeta_formato():
    candidatos = [Path(os.environ["FORMATO_GRAFICOS_DIR"])] if os.environ.get("FORMATO_GRAFICOS_DIR") else []
    for p in [CODE, *CODE.parents]:
        candidatos += [p / MANUAL_ESTILO, p / "formato"]
    for carpeta in candidatos:
        if (carpeta / "formato_graficos.py").is_file():
            return carpeta
    for raiz in _raices_sincronizadas():
        carpeta = _formato_en(raiz)
        if carpeta:
            return carpeta
    sys.exit("No encuentro formato_graficos.py. Debe estar en\n"
             f"    ...\\{BIBLIOTECA}\\{MANUAL_ESTILO}\n"
             "(también se busca en OneDrive/SharePoint y en la carpeta Santander de iCloud Drive).\n"
             "Para indicar otra carpeta: variable FORMATO_GRAFICOS_DIR.")


def _importar_formato():
    carpeta = _carpeta_formato()
    sys.path.insert(0, str(carpeta))
    import formato_graficos
    print(f"  Formato: {carpeta / 'formato_graficos.py'}")
    return formato_graficos


# ─────────────────────────────────────────────────────────────────────────────
# utilidades
# ─────────────────────────────────────────────────────────────────────────────

def num(x, dec=1):
    """1.234,5 (coma decimal, punto de miles)."""
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def signo(x, dec=1):
    x = round(x, dec) + 0.0   # evita "-0,0"
    return f"{'+' if x > 0 else ''}{num(x, dec)}"


def _aa(s):
    return 100 * (s / s.shift(12) - 1)


def _limites(valores, paso, incluir=None):
    v = np.asarray(valores, dtype=float).ravel()
    v = v[np.isfinite(v)]
    lo, hi = v.min(), v.max()
    if incluir is not None:
        lo, hi = min(lo, incluir), max(hi, incluir)
    return paso * math.floor(lo / paso), paso * math.ceil(hi / paso)


def _paso(lo, hi):
    rango = hi - lo
    return 1 if rango <= 8 else 2 if rango <= 16 else 5 if rango <= 40 else 10


def _limites_apilado(total, barras):
    lo = min(barras.clip(upper=0).sum(axis=1).min(), total.min(), 0)
    hi = max(barras.clip(lower=0).sum(axis=1).max(), total.max(), 0)
    paso = _paso(lo, hi)
    return paso * math.floor(lo / paso), paso * math.ceil(hi / paso), paso


def _apilar(chart):
    """Columnas de formato_graficos (agrupadas) -> apiladas."""
    from pptx.oxml.ns import qn
    barra = chart._chartSpace.find(qn("c:chart")).find(qn("c:plotArea")).find(qn("c:barChart"))
    barra.find(qn("c:grouping")).set("val", "stacked")
    gap = barra.find(qn("c:gapWidth"))
    if barra.find(qn("c:overlap")) is None:
        gap.addnext(barra.makeelement(qn("c:overlap"), {"val": "100"}))
    gap.set("val", "30")


# ─────────────────────────────────────────────────────────────────────────────
# Gráfico PNG (para correos o reportes fuera de la minuta)
# ─────────────────────────────────────────────────────────────────────────────

def grafico_png(ms, ruta=PNG, desde=DESDE):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.dates as mdates
    import matplotlib.pyplot as plt

    d = ms[ms.index >= desde].dropna(subset=[dr.MASA_AA])
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5.2), dpi=150)
    colores = ["#" + c for c in (dr.ROJO, dr.GRIS, dr.GRIS_CLARO)]

    pos = np.zeros(len(d))
    neg = np.zeros(len(d))
    for nombre, color in zip(dr.CONTRIB.values(), colores):
        v = d[nombre].to_numpy()
        ax1.bar(d.index, v, bottom=np.where(v >= 0, pos, neg), width=25, color=color,
                label=nombre.replace(" (pp)", ""))
        pos, neg = pos + np.clip(v, 0, None), neg + np.clip(v, None, 0)
    ax1.plot(d.index, d[dr.MASA_AA], color="#" + dr.NEGRO, lw=1.8, label="Masa salarial real")
    ax1.axhline(0, color="#" + dr.NEGRO, lw=0.6)
    ax1.annotate(num(d[dr.MASA_AA].iloc[-1]), (d.index[-1], d[dr.MASA_AA].iloc[-1]), xytext=(6, 0),
                 textcoords="offset points", va="center", fontsize=9, fontweight="bold")
    ax1.set_title("Masa salarial real: variación anual y contribuciones", loc="left", fontsize=12,
                  fontweight="bold")
    ax1.set_ylabel("(Var. a/a %, contribuciones en pp)")

    ax2.plot(d.index, d[dr.MASA], color="#" + dr.ROJO, lw=2, label="Ocupados × horas × salario real")
    ax2.plot(d.index, d[dr.MASA_SH], color="#" + dr.GRIS, lw=1.6, label="Ocupados × salario real")
    ax2.axhline(100, color="#" + dr.GRIS_CLARO, lw=0.8, ls="--")
    ax2.set_title("Masa salarial real: nivel", loc="left", fontsize=12, fontweight="bold")
    ax2.set_ylabel(f"(Índice {dr.ANIO_BASE} = 100, desestacionalizado)")

    for ax in (ax1, ax2):
        ax.xaxis.set_major_locator(mdates.YearLocator(2 if len(d) > 100 else 1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
        ax.spines[["top", "right"]].set_visible(False)
        ax.grid(axis="y", color="#E5E5E5", lw=0.6)
        ax.set_axisbelow(True)
        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=2, frameon=False, fontsize=8.5)
    fig.text(0.01, 0.005, f"Fuente: {FUENTE}. Último dato: trimestre móvil terminado en "
             f"{d.index[-1]:%m-%Y}. Ocupados desestacionalizados por el INE; horas e IR real "
             "desestacionalizados con STL.", fontsize=7.5, color="#" + dr.GRIS)
    fig.tight_layout(rect=(0, 0.03, 1, 1))
    fig.savefig(ruta)
    plt.close(fig)
    return ruta


# ─────────────────────────────────────────────────────────────────────────────
# Documento PDF (minuta escrita): textos automáticos, cuadro y figuras
# ─────────────────────────────────────────────────────────────────────────────

# colores Santander (formato_graficos.PALETA): rojo, negro, turquesa, amarillo, gris
C_ROJO, C_NEGRO, C_TURQUESA, C_AMARILLO, C_GRIS = "#EC0000", "#000000", "#4DC4CA", "#FFAE42", "#B2B2B2"
C_BANDA, C_GRILLA, C_TEXTO = "#DADADA", "#E5E5E5", "#595959"
MESES_CORTOS = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


def _pct(x, dec=1):
    return f"{num(round(x, dec) + 0.0, dec)}%"


def _comparar(a, b, dec=1, articulo=""):
    """'por encima de' / 'por debajo de' / 'en línea con' comparando a y b redondeados.
    articulo="el": 'por encima del' / 'por debajo del' / 'en línea con el'."""
    a, b = round(a, dec), round(b, dec)
    if a == b:
        return "en línea con" + (f" {articulo}" if articulo else "")
    base = "por encima" if a > b else "por debajo"
    return f"{base} del" if articulo == "el" else f"{base} de" + (f" {articulo}" if articulo else "")


def _racha(s):
    """Meses consecutivos al final de la serie con el mismo signo que el último dato."""
    s = s.dropna()
    signo_ultimo = np.sign(s.iloc[-1])
    n = 0
    for v in s.iloc[::-1]:
        if np.sign(v) != signo_ultimo or v == 0:
            break
        n += 1
    return n


def _estilo_ejes(ax, tam=7):
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.spines["bottom"].set_color(C_GRIS)
    ax.grid(axis="y", color=C_GRILLA, lw=0.6)
    ax.set_axisbelow(True)
    ax.tick_params(labelsize=tam, length=0, colors="#333333")
    from matplotlib.ticker import MaxNLocator
    ax.yaxis.set_major_locator(MaxNLocator(nbins=6, steps=[1, 2, 5, 10]))
    ax.yaxis.set_major_formatter(_formato_coma())


def _formato_coma(dec=1):
    from matplotlib.ticker import FuncFormatter
    return FuncFormatter(lambda v, _: num(v, dec))


def _png(fig):
    import io
    import matplotlib.pyplot as plt
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200)
    plt.close(fig)
    buf.seek(0)
    return buf


def _figura_lineas(df, colores, leyenda=True, tam=(3.45, 2.25)):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.dates as mdates
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(figsize=tam)
    for col, color in zip(df.columns, colores):
        ax.plot(df.index, df[col], color=color, lw=1.4, label=col)
    _estilo_ejes(ax)
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    if df.min().min() < 0 < df.max().max():
        ax.axhline(0, color=C_GRIS, lw=0.8)
    if leyenda:
        ax.legend(loc="lower center", bbox_to_anchor=(0.5, 1.0), ncol=len(df.columns), frameon=False,
                  fontsize=7, handlelength=1.6)
    fig.tight_layout(pad=0.4)
    return _png(fig)


def _figura_estacional(mm, t):
    """Variación m/m por mes calendario: banda mín-máx y promedio de los años de referencia, los dos años
    anteriores en líneas y el año en curso en rombos. Devuelve (png, DataFrame con los datos)."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    tabla = mm.dropna().to_frame("v").assign(anio=lambda d: d.index.year, mes=lambda d: d.index.month)
    tabla = tabla.pivot(index="mes", columns="anio", values="v").reindex(range(1, 13))
    completos = [a for a in tabla.columns if tabla[a].notna().sum() == 12]
    ref = [a for a in completos if 2010 <= a <= t.year - 3] or [a for a in completos if a <= t.year - 3]
    a0, a1 = min(ref), max(ref)
    banda = tabla[ref]
    datos = pd.DataFrame({f"Mín {a0}-{a1}": banda.min(axis=1), f"Máx {a0}-{a1}": banda.max(axis=1),
                          f"Prom. {a0}-{a1}": banda.mean(axis=1)})
    for a in (t.year - 2, t.year - 1, t.year):
        if a in tabla:
            datos[str(a)] = tabla[a]
    datos.index = pd.Index(MESES_CORTOS, name="Mes")

    x = np.arange(12)
    fig, ax = plt.subplots(figsize=(3.45, 2.25))
    ax.fill_between(x, datos.iloc[:, 0], datos.iloc[:, 1], color=C_BANDA, lw=0, label=f"Mín-Máx {a0}-{a1}")
    ax.plot(x, datos.iloc[:, 2], color=C_TURQUESA, lw=1.4, label=f"Prom. {a0}-{a1}")
    for a, color in ((t.year - 2, C_NEGRO), (t.year - 1, C_AMARILLO)):
        if str(a) in datos:
            ax.plot(x, datos[str(a)], color=color, lw=1.2, label=str(a))
    if str(t.year) in datos:
        ax.scatter(x, datos[str(t.year)], marker="D", s=14, color=C_ROJO, zorder=5, label=str(t.year))
    ax.axhline(0, color=C_GRIS, lw=0.8)
    _estilo_ejes(ax)
    ax.set_xticks(x, MESES_CORTOS)
    ax.set_xlim(-0.5, 11.5)
    ax.legend(loc="lower center", bbox_to_anchor=(0.5, 1.0), ncol=5, frameon=False, fontsize=6,
              handlelength=1.4, columnspacing=0.9)
    fig.tight_layout(pad=0.4)
    return _png(fig), datos


# ─────────────────────────────────────────────────────────────────────────────
# Minuta
# ─────────────────────────────────────────────────────────────────────────────

class Minuta:
    def __init__(self, hoy=None):
        self.fg = _importar_formato()
        datos = dr.cargar(descarga=False, hojas=["ajuste_estacional", "horas", "remuneraciones"])
        faltan = [h for h in ("ajuste_estacional", "horas", "remuneraciones") if h not in datos]
        if faltan:
            sys.exit(f"Faltan datos en Inputs ({', '.join(faltan)}). Corre antes datos_remuneraciones.py.")
        self.rem = datos["remuneraciones"].set_index("Periodo").sort_index()
        self.aa = self.rem[[IR, ICL, IR_REAL]].apply(_aa)
        self.ms = dr.masa_salarial(datos)
        self.t = self.rem[IR].dropna().index.max()
        self.mes = f"{MESES[self.t.month - 1]} {self.t.year}"
        self.tm = self.ms.dropna(subset=[dr.MASA_AA]).index.max()   # último mes de la masa salarial
        self.hoy = pd.Timestamp(hoy or dt.date.today())
        self.prs = self.fg.nueva_presentacion()
        self.n = 1
        self.pie = f"{TITULO} {self.mes}  |  {AUTOR}"
        self.datos = {}   # hoja del Excel de salida -> DataFrame

    # ---- maqueta (compatible con versiones nuevas y anteriores de formato_graficos)
    def lamina(self, titulo):
        self.n += 1
        try:
            return self.fg.lamina(self.prs, titulo, self.n, pie_texto=self.pie)
        except TypeError:
            return self.fg.lamina(self.prs, titulo, self.n)

    def texto(self, lam, texto, x, y, w, h, tam=13):
        from pptx.util import Cm
        self.fg._caja_texto(lam, texto, Cm(x), Cm(y), Cm(w), Cm(h), self.fg.tipografia("fuente"), tam,
                            color=getattr(self.fg, "GRIS_TEXTO", (89, 89, 89)))

    def vinetas(self, lam, items, x, y, w, h, tam=14):
        """Viñetas con cuadrado rojo Santander."""
        from pptx.oxml.ns import qn
        from pptx.util import Cm, Pt
        fg = self.fg
        caja = lam.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
        tf = caja.text_frame
        tf.word_wrap = True
        for i, texto in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(tam * 0.9)
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", str(Cm(0.6)))
            pPr.set("indent", str(-Cm(0.6)))
            clr = pPr.makeelement(qn("a:buClr"), {})
            clr.append(clr.makeelement(qn("a:srgbClr"), {"val": "EC0000"}))
            pPr.append(clr)
            pPr.append(pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"}))
            pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "■"}))
            run = p.add_run()
            run.text = texto
            run.font.name = fg.tipografia("fuente")
            run.font.size = Pt(tam)
            run.font.color.rgb = fg._rgb(*getattr(fg, "GRIS_TEXTO", (89, 89, 89)))
        return caja

    def grafico(self, lam, df, x, w, hoja, y=5.0, alto=10.0, nota=None, apilado=False, **kw):
        """Gráfico de formato_graficos + pie; guarda los datos para el Excel de salida."""
        from pptx.util import Cm
        df = df.dropna(how="all")
        chart = self.fg.agregar_grafico(lam, df, Cm(x), Cm(y), Cm(w), Cm(alto), **kw)
        if apilado:
            _apilar(chart)
        self.fg.pie(lam, Cm(x), Cm(y + alto + 0.15), Cm(w), FUENTE, nota, h=Cm(1.6))
        self.datos[f"{self.n:02d} {hoja}"] = df
        return chart

    def lineas_aa(self, lam, cols, x, w, titulo, hoja, nota=None):
        df = self.aa.loc[self.aa.index >= DESDE, list(cols)].rename(columns=cols)
        y_min, y_max = _limites(df.to_numpy(), 1, incluir=0)
        paso = _paso(y_min, y_max)
        y_min, y_max = _limites(df.to_numpy(), paso, incluir=0)
        return self.grafico(lam, df, x, w, hoja, titulo=titulo, unidad="(Var. a/a, %)",
                            posicion_unidad="arriba", y_min=y_min, y_max=y_max, y_paso=paso, linea_ref=0,
                            meses_entre_etiquetas=12, formato_fecha="yyyy", nota=nota, tam=11)

    # ================================================================ láminas
    def portada(self):
        fecha = f"{self.hoy.day:02d} de {MESES[self.hoy.month - 1]} de {self.hoy.year}"
        try:
            self.fg.portada(self.prs, "Remuneraciones y\nmasa salarial", self.mes.capitalize(),
                            autor=AUTOR, fecha=fecha)
        except TypeError:   # versión anterior: portada(prs, fecha_texto, titulo)
            self.fg.portada(self.prs, f"{self.mes.capitalize()}\n{fecha}", titulo="Remuneraciones")

    def v(self, col, rez=0):
        return self.aa[col].get(self.t - pd.DateOffset(months=rez), np.nan)

    def resumen(self):
        ms, tm = self.ms, self.tm
        lam = self.lamina("Resumen")
        prom = self.aa.loc[self.aa.index.year == self.t.year]
        c = {k: ms.loc[tm, v] for k, v in dr.CONTRIB.items()}
        nombres = {dr.OCUPADOS: "los ocupados", dr.HORAS: "las horas trabajadas", dr.IR_REAL: "el salario real"}
        motor = max(c, key=lambda k: abs(c[k]))
        items = [
            f"En {self.mes}, el índice de remuneraciones (IR) nominal creció {num(self.v(IR))}% a/a "
            f"({signo(self.v(IR) - self.v(IR, 1))} pp respecto del mes anterior) y el índice de costo laboral "
            f"(ICL) {num(self.v(ICL))}% a/a.",
            f"En términos reales, las remuneraciones crecieron {num(self.v(IR_REAL))}% a/a (mes anterior: "
            f"{num(self.v(IR_REAL, 1))}%). En lo que va de {self.t.year} promedian {num(prom[IR_REAL].mean())}% "
            f"a/a (nominal: {num(prom[IR].mean())}%).",
            f"La masa salarial real (ocupados × horas × salario real) varió {signo(ms.loc[tm, dr.MASA_AA])}% a/a "
            f"en el trimestre móvil terminado en {MESES[tm.month - 1]} {tm.year}: ocupados "
            f"{signo(c[dr.OCUPADOS])} pp, horas {signo(c[dr.HORAS])} pp y salario real {signo(c[dr.IR_REAL])} pp; "
            f"el principal factor fue {nombres[motor]}. Sin ajustar por horas, "
            f"{signo(ms.loc[tm, dr.MASA_SH_AA])}% a/a.",
        ]
        self.vinetas(lam, items, 1.3, 3.0, 15.2, 14)
        self.lineas_aa(lam, {IR: "IR nominal", IR_REAL: "IR real"}, 17.5, 15, "Índice de remuneraciones",
                       "IR nominal y real")

    def remuneraciones(self):
        lam = self.lamina("Remuneraciones y costo laboral")
        self.lineas_aa(lam, {IR: "Remuneraciones (IR)", ICL: "Costo laboral (ICL)"}, 1.3, 15,
                       "Remuneraciones y costo laboral nominales", "IR e ICL nominal")
        df = self.rem.loc[self.rem.index >= DESDE, [IR_REAL, "IR Real (Desest.)"]]
        df = df.rename(columns={IR_REAL: "IR real", "IR Real (Desest.)": "IR real desestacionalizado"})
        y_min, y_max = _limites(df.to_numpy(), 5)
        self.grafico(lam, df, 17.5, 15, "IR real nivel", titulo="Índice de remuneraciones real",
                     unidad=f"(Índice, base anual {dr.ANIO_BASE} = 100)", posicion_unidad="arriba",
                     y_min=y_min, y_max=y_max, y_paso=5, meses_entre_etiquetas=12, formato_fecha="yyyy",
                     etiqueta_ultimo=["IR real desestacionalizado"], tam=11,
                     nota="Desestacionalizado con STL.")

    def masa_salarial(self):
        ms, tm = self.ms, self.tm
        lam = self.lamina("Masa salarial real")
        d = ms[(ms.index >= DESDE)].dropna(subset=[dr.MASA_AA])
        u = d.iloc[-1]
        self.texto(lam, f"En el trimestre móvil terminado en {MESES[tm.month - 1]} {tm.year}, la masa salarial "
                        f"real varió {signo(u[dr.MASA_AA])}% a/a ({signo(u[dr.MASA_AA] - d[dr.MASA_AA].iloc[-2])} pp "
                        f"respecto del mes anterior). Sin ajustar por horas (ocupados × salario real), "
                        f"{signo(u[dr.MASA_SH_AA])}% a/a.", 1.3, 2.4, 31, 1.6)
        barras = {v: v.replace("Contribución ", "").replace(" (pp)", "").capitalize() for v in dr.CONTRIB.values()}
        g1 = pd.concat([d[dr.MASA_AA].rename("Masa salarial real"), d[list(barras)].rename(columns=barras)], axis=1)
        y_min, y_max, paso = _limites_apilado(g1.iloc[:, 0], g1.iloc[:, 1:])
        self.grafico(lam, g1, 1.3, 15, "Masa salarial a-a", y=5.4, alto=9.6, apilado=True,
                     titulo="Variación anual y contribuciones", unidad="(Var. a/a, %; contribuciones en pp)",
                     posicion_unidad="arriba", barras=list(barras.values()),
                     etiqueta_ultimo=["Masa salarial real"], y_min=y_min, y_max=y_max, y_paso=paso,
                     meses_entre_etiquetas=12, formato_fecha="yyyy", tam=11, nota=NOTA_MASA)
        g2 = d[[dr.MASA, dr.MASA_SH]].rename(columns={dr.MASA: "Ocupados × horas × salario real",
                                                       dr.MASA_SH: "Ocupados × salario real"})
        y_min, y_max = _limites(g2.to_numpy(), 5)
        self.grafico(lam, g2, 17.5, 15, "Masa salarial nivel", y=5.4, alto=9.6, titulo="Nivel",
                     unidad=f"(Índice {dr.ANIO_BASE} = 100, desestacionalizado)", posicion_unidad="arriba",
                     y_min=y_min, y_max=y_max, y_paso=5, meses_entre_etiquetas=12, formato_fecha="yyyy",
                     tam=11)

    # ================================================================ documento PDF
    def sa(self, base):
        """Serie desestacionalizada: la del INE si el cuadro la trae; si no, la de STL."""
        col = f"{base} (SA INE)"
        return self.rem[col] if col in self.rem and self.rem[col].notna().any() else self.rem[f"{base} (Desest.)"]

    def documento(self):
        """Minuta escrita en PDF (reportlab), con colores Santander. Devuelve la ruta del PDF."""
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_JUSTIFY
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import (HRFlowable, Image, KeepTogether, ListFlowable, ListItem, PageBreak,
                                        Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle)

        t = self.t
        ant = t - pd.DateOffset(months=1)
        mes, mes_ant = MESES[t.month - 1], MESES[ant.month - 1]
        ine_sa = all(f"{b} (SA INE)" in self.rem for b in (IR, ICL, IR_REAL))
        mm = {b: 100 * (self.sa(b) / self.sa(b).shift(1) - 1) for b in (IR, ICL, IR_REAL)}
        aa = self.aa
        mm_orig = 100 * (self.rem[IR] / self.rem[IR].shift(1) - 1)
        hist = mm_orig[(mm_orig.index.month == t.month) & (mm_orig.index.year.isin(range(2010, 2020)))]
        v = lambda s, f=t: s.get(f, np.nan)  # noqa: E731

        # ---- textos
        r_mm, r_aa = v(mm[IR_REAL]), v(aa[IR_REAL])
        verbo = "subieron" if round(r_mm, 1) > 0 else "cayeron" if round(r_mm, 1) < 0 else "no variaron"
        titulo = (f"Remuneraciones {mes} {t.year} | En {mes}, los salarios reales {verbo}"
                  f"{' ' + _pct(abs(r_mm)) if verbo != 'no variaron' else ''} m/m (sa) y "
                  f"{'crecen' if r_aa >= 0 else 'caen'} {_pct(abs(r_aa))} en doce meses.")
        ir_mm, ir_mm_ant = v(mm[IR]), v(mm[IR], ant)
        mov = (f"un aumento de {_pct(ir_mm)}" if round(ir_mm, 1) > 0 else
               f"una caída de {_pct(abs(ir_mm))}" if round(ir_mm, 1) < 0 else "una variación nula")
        b1 = (f"<b>El Índice de Remuneraciones nominal (IR) registró {mov} m/m (sa), "
              f"{_comparar(ir_mm, ir_mm_ant)} lo observado en {mes_ant} ({_pct(ir_mm_ant)}).</b> "
              f"En términos anuales, el índice muestra una {'expansión' if v(aa[IR]) >= 0 else 'contracción'} de "
              f"{_pct(abs(v(aa[IR])))}.")
        if len(hist):
            b1 += (f" La variación mensual de {mes} en la serie original ({_pct(v(mm_orig))}) se ubica "
                   f"{_comparar(v(mm_orig), hist.mean(), articulo='el')} "
                   f"promedio {hist.index.year.min()}-{hist.index.year.max()} para el mismo mes ({_pct(hist.mean())}).")
        b1 += (f" Por su parte, el Índice de Costos Laborales (ICL) varió {_pct(v(mm[ICL]))} m/m (sa) y "
               f"{_pct(v(aa[ICL]))} a/a.")
        inflacion = 100 * ((1 + v(aa[IR]) / 100) / (1 + r_aa / 100) - 1)
        racha = _racha(aa[IR_REAL][aa.index <= t])
        mov_r = (f"un alza de {_pct(r_mm)}" if round(r_mm, 1) > 0 else
                 f"una caída de {_pct(abs(r_mm))}" if round(r_mm, 1) < 0 else "una variación nula")
        b2 = (f"<b>Los salarios reales registraron {mov_r} m/m (sa) en {mes}.</b> En términos anuales, "
              f"considerando una inflación de {_pct(inflacion)} (IPC implícito en el IR real del INE), los salarios "
              f"reales se {'expandieron' if r_aa >= 0 else 'contrajeron'} {_pct(abs(r_aa))}, con lo que acumulan "
              f"{racha} {'mes consecutivo' if racha == 1 else 'meses consecutivos'} de "
              f"{'alzas' if r_aa >= 0 else 'caídas'} anuales.")
        ms, tm = self.ms, self.tm
        b3 = (f"<b>La masa salarial real varió {signo(ms.loc[tm, dr.MASA_AA])}% a/a</b> en el trimestre móvil "
              f"terminado en {MESES[tm.month - 1]} (ocupados × horas habituales × salario real): ocupados "
              f"{signo(ms.loc[tm, dr.CONTRIB[dr.OCUPADOS]])} pp, horas {signo(ms.loc[tm, dr.CONTRIB[dr.HORAS]])} pp "
              f"y salario real {signo(ms.loc[tm, dr.CONTRIB[dr.IR_REAL]])} pp.")

        # ---- estilos (colores Santander)
        rojo, gris = colors.HexColor(C_ROJO), colors.HexColor(C_TEXTO)
        st = {
            "titulo": ParagraphStyle("titulo", fontName="Helvetica-Bold", fontSize=15, leading=19, textColor=rojo),
            "fecha": ParagraphStyle("fecha", fontName="Helvetica-Bold", fontSize=10.5, leading=14, spaceBefore=8),
            "autor": ParagraphStyle("autor", fontName="Helvetica-Oblique", fontSize=9, leading=12, textColor=gris),
            "h1": ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=14, leading=18, textColor=rojo,
                                 spaceBefore=6, spaceAfter=4),
            "texto": ParagraphStyle("texto", fontName="Helvetica", fontSize=10, leading=14, alignment=TA_JUSTIFY),
            "cuadro": ParagraphStyle("cuadro", fontName="Helvetica", fontSize=9, leading=12, alignment=1),
            "nota": ParagraphStyle("nota", fontName="Helvetica", fontSize=7.5, leading=10, textColor=gris),
            "fig": ParagraphStyle("fig", fontName="Helvetica", fontSize=8.5, leading=10.5),
            "unidad": ParagraphStyle("unidad", fontName="Helvetica", fontSize=7, leading=9, textColor=gris),
            "fuente": ParagraphStyle("fuente", fontName="Helvetica", fontSize=6.5, leading=8, textColor=gris),
        }
        fecha = f"{self.hoy.day} de {MESES[self.hoy.month - 1]} de {self.hoy.year}"

        # ---- cuadro 1
        etq_ant, etq = f"{MESES_CORTOS[ant.month - 1]}-{ant:%y}", f"{MESES_CORTOS[t.month - 1]}-{t:%y}"
        filas = [["Índice", f"m/m (sa) {etq_ant}", f"m/m (sa) {etq}", f"a/a {etq_ant}", f"a/a {etq}"]]
        cuadro_df = {}
        for nombre, b in (("IR nominal", IR), ("ICL nominal", ICL), ("IR real", IR_REAL)):
            vals = [v(mm[b], ant), v(mm[b]), v(aa[b], ant), v(aa[b])]
            filas.append([nombre] + [_pct(x, 2) for x in vals])
            cuadro_df[nombre] = vals
        self.datos["Doc Cuadro 1"] = pd.DataFrame(cuadro_df, index=filas[0][1:]).T.rename_axis("Índice")
        cuadro = Table(filas, colWidths=[4.0 * cm] + [3.2 * cm] * 4)
        cuadro.setStyle(TableStyle([
            ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8.5), ("FONT", (0, 1), (-1, -1), "Helvetica", 8.5),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("BACKGROUND", (0, 0), (-1, 0), rojo),
            ("ALIGN", (1, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F2")]),
            ("LINEBELOW", (0, -1), (-1, -1), 0.8, rojo),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        nota_sa = ("sa = serie desestacionalizada por el INE" if ine_sa else
                   "sa = serie desestacionalizada con STL a partir de la serie original del INE")

        # ---- figuras
        w, h = 8.6 * cm, 5.6 * cm
        desde_nivel = pd.Timestamp(year=t.year - 5, month=1, day=1)
        f1 = self.sa(IR_REAL)[lambda s: s.index >= desde_nivel].dropna().rename("IR real (sa)").to_frame()
        f2 = aa.loc[aa.index >= pd.Timestamp(year=t.year - 7, month=1, day=1)].rename(
            columns={IR: "IR", ICL: "ICL", IR_REAL: "IR real"})
        png3, d3 = _figura_estacional(mm[IR], t)
        png4, d4 = _figura_estacional(mm[IR_REAL], t)
        figuras = [
            ("Figura 1: Índice de Remuneraciones real", f"(serie desestacionalizada, {dr.ANIO_BASE}=100)",
             _figura_lineas(f1, [C_ROJO], leyenda=False)),
            ("Figura 2: Variación anual de los índices", "(%, serie original)",
             _figura_lineas(f2, [C_ROJO, C_NEGRO, C_TURQUESA])),
            ("Figura 3: Índice de Remuneraciones nominal", "(var. % m/m, serie desestacionalizada)", png3),
            ("Figura 4: Índice de Remuneraciones real", "(var. % m/m, serie desestacionalizada)", png4),
        ]
        for k, df in (("Doc Fig1 IR real", f1), ("Doc Fig2 var anual", f2), ("Doc Fig3 IR nominal mm", d3),
                      ("Doc Fig4 IR real mm", d4)):
            self.datos[k] = df

        def bloque(tit, uni, png):
            return [Paragraph(tit, st["fig"]), Paragraph(uni, st["unidad"]), Image(png, width=w, height=h),
                    Paragraph("Fuente: INE.", st["fuente"])]

        celdas = [bloque(*f) for f in figuras]
        grilla = Table([[celdas[0], celdas[1]], [celdas[2], celdas[3]]], colWidths=[9.0 * cm, 9.0 * cm])
        grilla.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                    ("BOTTOMPADDING", (0, 0), (-1, -1), 14)]))

        # ---- armado
        bullets = ListFlowable([ListItem(Paragraph(b, st["texto"]), value="■", bulletColor=rojo,
                                         bulletFontSize=7) for b in (b1, b2, b3)],
                               bulletType="bullet", start="■", leftIndent=14, bulletOffsetY=-1)
        historia = [
            Paragraph(titulo, st["titulo"]), Paragraph(fecha, st["fecha"]), Paragraph(f"Autor: {AUTOR}", st["autor"]),
            Spacer(1, 4), HRFlowable(width="100%", thickness=1.2, color=rojo, spaceAfter=8),
            Paragraph("Principales resultados", st["h1"]), bullets, Spacer(1, 14),
            KeepTogether([Paragraph("Cuadro 1: Variaciones de los índices de remuneraciones y costos laborales",
                                    st["cuadro"]), Spacer(1, 4), cuadro, Spacer(1, 4),
                          Paragraph(f"Nota: {nota_sa}. El último dato es provisorio. Fuente: INE.", st["nota"])]),
            PageBreak(), Paragraph("Figuras", st["h1"]), Spacer(1, 4), grilla,
        ]

        def pie(canvas, doc):
            canvas.saveState()
            canvas.setFont("Helvetica", 8)
            canvas.setFillColor(gris)
            canvas.drawString(2 * cm, 1.3 * cm, AUTOR)
            canvas.setFillColor(colors.black)
            canvas.drawCentredString(letter[0] / 2, 1.3 * cm, str(doc.page))
            canvas.setStrokeColor(rojo)
            canvas.setLineWidth(2)
            canvas.line(2 * cm, letter[1] - 1.4 * cm, letter[0] - 2 * cm, letter[1] - 1.4 * cm)
            canvas.restoreState()

        ruta = OUT / f"Minuta_Remuneraciones_{mes}_{t.year}.pdf"
        doc = SimpleDocTemplate(str(ruta), pagesize=letter, leftMargin=2 * cm, rightMargin=2 * cm,
                                topMargin=2.2 * cm, bottomMargin=2.2 * cm, title=f"Remuneraciones {mes} {t.year}",
                                author=AUTOR)
        try:
            doc.build(historia, onFirstPage=pie, onLaterPages=pie)
        except PermissionError:
            sys.exit(f"No pude guardar {ruta}: está abierto. Ciérralo y vuelve a correr.")
        return ruta

    # ---- armar y guardar
    def construir(self):
        OUT.mkdir(parents=True, exist_ok=True)
        self.portada()
        self.resumen()
        self.remuneraciones()
        self.masa_salarial()
        pdf_doc = self.documento()
        print(f"  ✅ Documento: {pdf_doc}")
        nombre = f"{self.hoy:%Y%m%d}_Minuta_Remuneraciones"
        pptx = OUT / f"{nombre}.pptx"
        try:
            self.prs.save(pptx)
            with pd.ExcelWriter(OUT / f"{nombre}_datos.xlsx") as xl:
                for hoja, df in self.datos.items():
                    df.rename_axis("Fecha").to_excel(xl, sheet_name=hoja[:31])
        except PermissionError as e:
            sys.exit(f"No pude guardar {e.filename}: está abierto. Ciérralo y vuelve a correr.")
        print(f"  ✅ PPTX: {pptx} ({self.n} láminas; IR a {self.mes})")
        print(f"  ✅ Datos de los gráficos: {OUT / f'{nombre}_datos.xlsx'}")
        print(f"  ✅ Gráfico masa salarial: {grafico_png(self.ms)}")
        try:
            pdf = self.fg.exportar_pdf(pptx)
            if pdf:
                print(f"  ✅ PDF:  {pdf}")
        except Exception as e:  # noqa: BLE001
            print(f"  ⚠️ No pude exportar el PDF ({e}); el PPTX quedó listo.")
        return pptx


if __name__ == "__main__":
    print("=" * 60 + "\n  Minuta de Remuneraciones\n" + "=" * 60)
    Minuta().construir()
