# -*- coding: utf-8 -*-
r"""Minuta de Remuneraciones (INE): PPTX con gráficos nativos (editables), PDF y Excel de salida.

Láminas: portada, resumen (textos automáticos del último dato), remuneraciones y costo laboral, y
masa salarial real (variación anual con contribuciones de ocupados, horas y salario real, y nivel).
Formato de formato_graficos.py (Administrativos\Manual de Estilo), como la minuta IMCE.

Carpetas (relativas a este script; sirven para cualquier usuario):
    <Remuneraciones>\Code\      datos_remuneraciones.py y minuta_remuneraciones.py
    <Remuneraciones>\Inputs\    archivos del INE (los deja datos_remuneraciones.py)
    <Remuneraciones>\Outputs\   Masa_salarial_Base.xlsx y Cifras_empleo.xlsx (bases) y
                                AAAAMMDD_Minuta_Remuneraciones.pptx / .pdf / _datos.xlsx y Masa_salarial.png

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
AUTOR = "Estudios Santander"   # aparece en la portada y en el pie de cada lámina
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

    # ---- armar y guardar
    def construir(self):
        OUT.mkdir(parents=True, exist_ok=True)
        self.portada()
        self.resumen()
        self.remuneraciones()
        self.masa_salarial()
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
