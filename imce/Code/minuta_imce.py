# -*- coding: utf-8 -*-
r"""Minuta mensual del IMCE (Índice Mensual de Confianza Empresarial, ICARE/UAI).

Réplica de la presentación mensual de Coordinación Macroeconómica, con textos generados
automáticamente a partir del último dato y gráficos nativos (editables) de PowerPoint con el formato
de formato_graficos.py (00_Administrativos\Manual de Estilo).

Carpetas (relativas a este script; sirven para cualquier usuario):
    <IMCE>\Code\      datos_imce.py y minuta_imce.py
    <IMCE>\Inputs\    imce.csv (lo deja datos_imce.py)
    <IMCE>\Outputs\   Base_IMCE.xlsx (bases) y Presentación_IMCE_<mes>_<año>.pptx / .pdf / _datos.xlsx;
                      las presentaciones anteriores pasan a Outputs\<año>\ con la fecha al inicio
    <IMCE>\run_minuta_imce.py   lanzador (F5 en Spyder)

Uso (no hay que editar nada): abrir run_minuta_imce.py (carpeta IMCE) en Spyder y F5; o correr este
script, que primero ejecuta datos_imce.py
    python minuta_imce.py                 # descarga, actualiza Base_IMCE.xlsx y arma la minuta
    python minuta_imce.py --sin-descarga  # lo mismo con el Inputs\imce.csv que ya existe
(datos_imce.py se puede seguir corriendo solo, para actualizar únicamente los datos y el Excel.)
"""
import datetime as dt
import math
import sys

import numpy as np
import pandas as pd
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

import datos_imce as di

fg = di.fg
OUT = di.OUT

TITULO = "IMCE"
AUTOR = "Andrés Sansone"   # aparece en la portada y en el pie de cada lámina
FUENTE = "ICARE/UAI, publicado por el Banco Central de Chile"
NOTA_50 = "Sobre (bajo) 50 indica optimismo (pesimismo)."
MESES = fg.MESES_LARGOS
SECTORES = {"Comercio": "comercio", "Industria": "industria", "Construcción": "construccion", "Minería": "mineria"}

# índices promedio (simple) de los sectores, definidos en datos_imce.PROMEDIOS: nombre -> (series, hoja)
PROMEDIOS = {nombre: (series, origen) for nombre, series, origen in di.PROMEDIOS}
SECTOR = {"com": "comercio", "con": "construcción", "ind": "industria"}

# posiciones de gráficos (cm), como en el chartpack de Expectativas
Y0, ALTO = Cm(fg.Y_CONTENIDO_CM), Cm(fg.ALTO_CM)
DOS, ANCHO_DOS = [Cm(1.3), Cm(17.5)], Cm(fg.ANCHO_CM)
TRES, ANCHO_TRES, TAM_TRES = [Cm(1.0), Cm(11.8), Cm(22.6)], Cm(10.2), 10
X_UNO, ANCHO_UNO = Cm(1.3), Cm(31.2)


# ─────────────────────────────────────────────────────────────────────────────
# utilidades de texto
# ─────────────────────────────────────────────────────────────────────────────

def num(x, dec=1):
    """1.234,5 (coma decimal, punto de miles)."""
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def signo(x):
    return f"{'+' if x > 0 else ''}{num(x)}"


def variacion(d, umbral=0.05):
    """'aumentó 1,2 p.' / 'cayó 0,4 p.' / 'se mantuvo sin variación'."""
    if pd.isna(d):
        return "sin dato"
    if abs(d) < umbral:
        return "se mantuvo sin variación"
    return f"{'aumentó' if d > 0 else 'cayó'} {num(abs(d))} p."


def zona(v):
    return "optimista" if v > 50 else "pesimista" if v < 50 else "neutral"


def lista(xs):
    xs = [x.lower() for x in xs]
    return ", ".join(xs[:-1]) + (" y " if len(xs) > 1 else "") + xs[-1] if xs else ""


# ─────────────────────────────────────────────────────────────────────────────
# Minuta
# ─────────────────────────────────────────────────────────────────────────────

class Minuta:
    def __init__(self, hoy=None):
        self.m = di.cargar()
        self.sa = di.desestacionalizar(self.m)
        self.comp = di.compuestos(self.sa)
        self.prom = {}   # nombre -> (serie promedio, nota de formato_graficos)
        for nombre, (cols, origen) in PROMEDIOS.items():
            base = (self.sa if origen == "SA" else self.m)[cols]
            base.columns = [SECTOR[c[:3]] for c in cols]
            df, nota = fg.promedio(base, f"Índice promedio de {nombre.lower().replace(' (desest.)', '')}")
            self.prom[nombre] = (df.iloc[:, 0], nota)
        self.t = self.m["total"].dropna().index.max()
        self.mes = MESES[self.t.month - 1]
        self.mes_ant = MESES[(self.t - pd.DateOffset(months=1)).month - 1]
        self.hoy = pd.Timestamp(hoy or dt.date.today())
        self.prs = fg.nueva_presentacion()
        self.n = 1
        self.pie = f"{TITULO} {self.mes} {self.t.year}  |  {AUTOR}"   # pie de las láminas
        self.datos = {}   # hoja del Excel de respaldo -> DataFrame

    # ---- datos
    def v(self, k, rez=0, sa=False):
        s = self.sa[k] if sa else self.m[k]
        return s.get(self.t - pd.DateOffset(months=rez), np.nan)

    def dm(self, k, sa=False):
        return self.v(k, sa=sa) - self.v(k, 1, sa=sa)

    def da(self, k, sa=False):
        return self.v(k, sa=sa) - self.v(k, 12, sa=sa)

    def mm(self, k, sa=False):
        return f"{signo(self.dm(k, sa))} p. m/m"

    def aa(self, k, sa=False):
        return f"{signo(self.da(k, sa))} p. a/a"

    def _dmc(self, s):
        return s[self.t] - s[self.t - pd.DateOffset(months=1)]

    # ---- maqueta
    def lamina(self, titulo):
        self.n += 1
        return fg.lamina(self.prs, titulo, self.n, pie_texto=self.pie)

    def vinetas(self, lam, items, x, y, w, h, tam=14):
        """Viñetas con cuadrado rojo Santander."""
        caja = lam.shapes.add_textbox(x, y, w, h)
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
            run.font.color.rgb = fg._rgb(*fg.GRIS_TEXTO)
        return caja

    def grafico(self, lam, df, x, w, hoja, y=Y0, alto=ALTO, nota=None, apilado=False, **kw):
        """Gráfico de formato_graficos (sin etiqueta del último dato) + pie; guarda los datos para el Excel."""
        kw.setdefault("etiqueta_ultimo", False)
        chart = fg.agregar_grafico(lam, df, x, y, w, alto, **kw)
        if apilado:
            _apilar(chart)
        fg.pie(lam, x, y + alto + Cm(0.15), w, FUENTE, nota, h=Cm(1.6))
        self.datos[f"{self.n:02d} {hoja}"] = df
        return chart

    def serie(self, cols, sa=False, desde="2012-01-01"):
        """DataFrame de las series [(clave, etiqueta)] desde 'desde'."""
        df = pd.DataFrame({lab: (self.sa[k] if sa else self.m[k]) for k, lab in cols})
        return df[df.index >= desde].dropna(how="all")

    def promedio(self, nombre, desde="2012-01-01"):
        s, nota = self.prom[nombre]
        etiqueta = "Promedio " + " y ".join(", ".join(SECTOR[c[:3]] for c in PROMEDIOS[nombre][0]).rsplit(", ", 1))
        return s[s.index >= desde].rename(etiqueta).to_frame(), nota

    def vineta(self, nombre, df):
        """'Nombre: comercio 60,4 (+2,0 p.), ...; todos por sobre agosto.' con el dato del mes de cada serie."""
        datos = [(c, df[c].get(self.t, np.nan), df[c].get(self.t, np.nan) - df[c].get(self.t - pd.DateOffset(months=1), np.nan))
                 for c in df.columns]
        datos = [(c, v, d) for c, v, d in datos if not pd.isna(v)]
        if not datos:
            return f"{nombre}: sin dato para {self.mes}."
        pos = lambda d: ("por sobre" if d > 0.05 else "por debajo de" if d < -0.05 else "igual a")  # noqa: E731
        if len(datos) == 1:
            _, v, d = datos[0]
            return f"{nombre}: {num(v)} puntos en {self.mes}, {pos(d)} {self.mes_ant} ({signo(d)} p.)."
        minus = lambda c: c if c[:2].isupper() else c[0].lower() + c[1:]  # noqa: E731 - "IMCE" no cambia
        partes = [f"{minus(c)} {num(v)} ({signo(d)} p.)" for c, v, d in datos]
        texto = f"{nombre} en {self.mes}: " + ", ".join(partes[:-1]) + " y " + partes[-1]
        grupos = {}
        for c, v, d in datos:
            grupos.setdefault(pos(d), []).append(minus(c))
        if len(grupos) == 1:
            p = next(iter(grupos))
            todos = "ambos" if len(datos) == 2 else "todos"
            return f"{texto}; {todos} {p} {self.mes_ant}."
        return f"{texto}; " + "; ".join(f"{', '.join(xs[:-1]) + ' y ' + xs[-1] if len(xs) > 1 else xs[0]} "
                                         f"{p} {self.mes_ant}" for p, xs in grupos.items()) + "."

    def lamina_graficos(self, titulo, graficos, tam_vineta=12):
        """Lámina con una viñeta por gráfico (el dato del mes) arriba y 1, 2 o 3 gráficos debajo.
        graficos: lista de dict(df, hoja, vineta, ...argumentos de grafico/agregar_grafico)."""
        lam = self.lamina(titulo)
        vinetas = [g.pop("vineta") for g in graficos]
        alto_v = sum(0.62 * math.ceil(len(t) / 135) + 0.3 for t in vinetas)
        self.vinetas(lam, vinetas, Cm(1.3), Cm(2.4), Cm(31), Cm(alto_v + 0.3), tam=tam_vineta)
        y = 2.4 + alto_v + 1.45
        alto = 15.2 - y
        n = len(graficos)
        pos = {1: [(X_UNO, ANCHO_UNO)], 2: list(zip(DOS, [ANCHO_DOS] * 2)), 3: list(zip(TRES, [ANCHO_TRES] * 3))}[n]
        for (x, w), g in zip(pos, graficos):
            df = g.pop("df")
            if not g.pop("apilado", False):
                y_min, y_max = _limites(df)
                g = {"unidad": "(Índice)", "posicion_unidad": "arriba", "y_min": y_min, "y_max": y_max,
                     "y_paso": 10, "linea_ref": 50, "meses_entre_etiquetas": 24, "formato_fecha": "yyyy",
                     **({"tam": TAM_TRES} if n == 3 else {}), **g}
            else:
                g["apilado"] = True
            self.grafico(lam, df, x, w, g.pop("hoja"), y=Cm(y), alto=Cm(alto), **g)
        return lam

    # ================================================================ láminas
    def portada(self):
        fecha = f"{self.hoy.day:02d} de {MESES[self.hoy.month - 1]} de {self.hoy.year}"
        fg.portada(self.prs, "Índice Mensual de Confianza\nEmpresarial (IMCE)",
                   f"{self.mes.capitalize()} {self.t.year}", autor=AUTOR, fecha=fecha)

    def calendario(self, lam, clave, nombre):
        """Año en curso frente a años de referencia, por mes calendario."""
        m, t = self.m[clave], self.t
        anios = [t.year, t.year - 1, t.year - 2, 2018, 2010]
        idx = pd.date_range("2000-01-01", periods=12, freq="MS")
        df = pd.DataFrame({str(a): pd.Series(m[m.index.year == a].values, index=idx[:(m.index.year == a).sum()])
                           for a in anios}, index=idx)
        y_min, y_max = _limites(df, paso=5, piso=None, techo=None)
        self.grafico(lam, df, DOS[1], ANCHO_DOS, f"{nombre} calendario",
                     titulo=f"{nombre}: comparación con años previos", unidad="(Índice por mes calendario)",
                     posicion_unidad="arriba", y_min=y_min, y_max=y_max, y_paso=5, linea_ref=50,
                     meses_entre_etiquetas=1, formato_fecha="mmm", etiqueta_ultimo=False,
                     nota=NOTA_50)
        self.datos[f"{self.n:02d} {nombre} calendario"] = df.set_index(pd.Index([x[:3] for x in MESES], name="Mes"))

    def resumen(self):
        t, m = self.t, self.m
        lam = self.lamina("Resumen")
        tot, d1, d12 = self.v("total"), self.dm("total"), self.da("total")
        sm1, sm12 = self.dm("sin_mineria"), self.da("sin_mineria")
        anio = m["total"][m.index.year == t.year]
        sube = [s for s, k in SECTORES.items() if self.dm(k) > 0.05]
        baja = [s for s, k in SECTORES.items() if self.dm(k) < -0.05]
        opt = [s for s, k in SECTORES.items() if self.v(k) > 50]
        pes = [s for s, k in SECTORES.items() if self.v(k) < 50]
        minimo = [s for s, k in SECTORES.items() if self.v(k) <= m[k][m.index.year == t.year].min()]
        items = [
            f"El IMCE total se ubicó en {num(tot)} puntos en {self.mes}, en terreno {zona(tot)}; "
            f"{variacion(d1)} respecto del mes anterior y {'subió' if d12 > 0 else 'bajó'} {num(abs(d12))} p. "
            f"en doce meses. En lo que va de {t.year} promedia {num(anio.mean())} puntos "
            f"(2019: {num(m['total']['2019'].mean())}).",
            f"El IMCE sin minería {variacion(sm1)} en el mes, hasta {num(self.v('sin_mineria'))} puntos, y "
            f"{'aumentó' if sm12 > 0 else 'cayó'} {num(abs(sm12))} p. a/a.",
        ]
        partes = []
        if sube:
            partes.append(f"{'subió' if len(sube) == 1 else 'subieron'} {lista(sube)}")
        if baja:
            partes.append(f"{'cayó' if len(baja) == 1 else 'cayeron'} {lista(baja)}")
        txt = ("Por sectores, en el mes " + ", mientras que ".join(partes) + ". ") if partes else ""
        if opt:
            txt += f"{lista(opt).capitalize()} {'se ubica' if len(opt) == 1 else 'se ubican'} en terreno optimista"
            txt += f" y {lista(pes)} en terreno pesimista." if pes else "."
        elif pes:
            txt += "Todos los sectores se ubican en terreno pesimista."
        if minimo:
            txt += f" {lista(minimo).capitalize()} {'marca' if len(minimo) == 1 else 'marcan'} su menor valor del año."
        items.append(txt.strip())
        self.vinetas(lam, items, Cm(1.3), Cm(3.0), Cm(15.2), Cm(14))
        self.calendario(lam, "total", "IMCE total")

    def resumen_sin_mineria(self):
        m, t, k = self.m, self.t, "sin_mineria"
        lam = self.lamina("Resumen: IMCE sin minería")
        v, d1, d12 = self.v(k), self.dm(k), self.da(k)
        anio = m[k][m.index.year == t.year].dropna()
        prev = m[k][m.index.year == t.year - 1].dropna()
        mismo = prev[prev.index.month <= t.month]
        maxi, mini = anio.idxmax(), anio.idxmin()
        brecha = self.v("total") - v
        sect = {s: c for s, c in SECTORES.items() if s != "Minería"}
        items = [
            f"El IMCE sin minería se ubicó en {num(v)} puntos en {self.mes}, en terreno {zona(v)}; "
            f"{variacion(d1)} respecto del mes anterior y {'subió' if d12 > 0 else 'bajó'} {num(abs(d12))} p. "
            f"en doce meses.",
            f"En lo que va de {t.year} promedia {num(anio.mean())} puntos, frente a {num(mismo.mean())} en igual "
            f"período de {t.year - 1} y {num(m[k]['2019'].mean())} en 2019. Su máximo del año fue "
            f"{num(anio.max())} ({MESES[maxi.month - 1]}) y su mínimo {num(anio.min())} ({MESES[mini.month - 1]}).",
            "Por sectores: " + ", ".join(f"{n.lower()} {num(self.v(c))} ({self.mm(c)})" for n, c in sect.items())
            + f". La minería ({num(self.v('mineria'))} pts) deja al IMCE total {num(abs(brecha))} p. "
              f"{'por encima' if brecha > 0 else 'por debajo'} del indicador sin minería.",
        ]
        self.vinetas(lam, items, Cm(1.3), Cm(3.0), Cm(15.2), Cm(14))
        self.calendario(lam, k, "IMCE sin minería")

    def resultados(self):
        mm, aa, v = self.mm, self.aa, self.v
        lam = self.lamina("Principales resultados")
        items = [
            f"IMCE total: {num(v('total'))} puntos en {self.mes} ({mm('total')}; {aa('total')}), en terreno "
            f"{zona(v('total'))}; el mes anterior fue {num(v('total', 1))}.",
            "Sectores: " + ", ".join(f"{n.lower()} {num(v(c))} ({mm(c)}; {aa(c)})"
                                     for n, c in list(SECTORES.items())[:-1])
            + f" y minería {num(v('mineria'))} ({mm('mineria')}; {aa('mineria')}).",
            f"Expectativas de costos: comercio {num(v('com_costos'))} ({mm('com_costos')}), construcción "
            f"{num(v('con_costos'))} ({mm('con_costos')}) e industria {num(v('ind_costos'))} ({mm('ind_costos')}).",
            f"Expectativas de salarios: construcción {num(v('con_salarios'))} ({mm('con_salarios')}; "
            f"{aa('con_salarios')}) e industria {num(v('ind_salarios'))} ({mm('ind_salarios')}; "
            f"{aa('ind_salarios')}). Expectativas de precios: comercio {num(v('com_precios'))} "
            f"({mm('com_precios')}), construcción {num(v('con_precios'))} ({mm('con_precios')}) e industria "
            f"{num(v('ind_precios'))} ({mm('ind_precios')}).",
            f"Expectativas de empleo: comercio {num(v('com_empleo'))} ({mm('com_empleo')}; {aa('com_empleo')}), "
            f"construcción {num(v('con_empleo'))} ({mm('con_empleo')}; {aa('con_empleo')}) e industria "
            f"{num(v('ind_empleo'))} ({mm('ind_empleo')}; {aa('ind_empleo')}).",
        ]
        self.vinetas(lam, items, Cm(1.3), Cm(3.0), Cm(31), Cm(13.5))
        fg.pie(lam, Cm(1.3), Cm(16.1), Cm(31), FUENTE, NOTA_50)

        lam = self.lamina("Principales resultados (continuación)")
        c = {k: tot for k, (_, tot) in self.comp.items()}
        cm = lambda k: f"{num(c[k][self.t])} ({signo(self._dmc(c[k]))} p. m/m)"  # noqa: E731
        items = [
            f"Inventarios (situación actual): comercio {num(v('com_inventarios'))} ({mm('com_inventarios')}) e "
            f"industria {num(v('ind_inventarios'))} ({mm('ind_inventarios')}).",
            f"Demanda (situación actual, desestacionalizada): comercio {num(v('com_demanda', sa=True))} "
            f"({mm('com_demanda', True)}), industria {num(v('ind_demanda', sa=True))} ({mm('ind_demanda', True)}) "
            f"y construcción {num(v('con_demanda', sa=True))} ({mm('con_demanda', True)}).",
            f"Expectativas de inversión: comercio {num(v('com_inversion'))} ({mm('com_inversion')}) e industria "
            f"{num(v('ind_inversion'))} ({mm('ind_inversion')}). Expectativas de ventas en comercio (desest.) "
            f"{num(v('com_ventas', sa=True))} ({mm('com_ventas', True)}) y de producción en industria (desest.) "
            f"{num(v('ind_prod_esp', sa=True))} ({mm('ind_prod_esp', True)}).",
            f"Índices compuestos (50 = 0, desestacionalizados): situación actual general (OCSI) {cm('OCSI')}, "
            f"situación actual personal (PCSI) {cm('PCSI')}, expectativas para el país (CFEI) {cm('CFEI')} y "
            f"expectativas generales (OFEI) {cm('OFEI')}.",
        ]
        self.vinetas(lam, items, Cm(1.3), Cm(3.0), Cm(31), Cm(13.5))
        fg.pie(lam, Cm(1.3), Cm(16.1), Cm(31), FUENTE,
               "Demanda, ventas, producción e índices compuestos desestacionalizados con STL. " + NOTA_50)

    def graficos(self):
        def g(nombre, df, **kw):
            return dict(df=df, vineta=self.vineta(nombre, df), **kw)

        def gp(nombre, titulo, hoja, nota_extra="", **kw):
            df, nota = self.promedio(nombre)
            return dict(df=df, vineta=self.vineta(titulo, df), titulo=titulo, hoja=hoja,
                        nota=f"{nota} {nota_extra}".strip(), **kw)

        df = self.serie([("total", "IMCE total"), ("sin_mineria", "IMCE sin minería")], desde="2004-01-01")
        self.lamina_graficos("IMCE total y sin minería", [
            g("Índices", df, titulo="IMCE total y sin minería", hoja="IMCE total",
              nota=NOTA_50 + " Oct-19: estallido social; mar-20: Covid-19.")])

        df = self.serie([(k, s) for s, k in SECTORES.items() if s != "Minería"], desde="2011-01-01")
        self.lamina_graficos("IMCE por sectores", [
            g("IMCE por sectores", df, titulo="IMCE por sectores", hoja="IMCE sectores", meses_entre_etiquetas=12,
              nota=NOTA_50)])

        self.lamina_graficos("Presiones de salarios, costos y precios", [
            g("Expectativas de salarios", self.serie([("con_salarios", "Construcción"), ("ind_salarios", "Industria")]),
              titulo="Expectativas de salarios", hoja="Salarios", nota="Comercio no informa salarios."),
            g("Expectativas de costos", self.serie([("com_costos", "Comercio"), ("con_costos", "Construcción"),
                                                    ("ind_costos", "Industria")]),
              titulo="Expectativas de costos", hoja="Costos"),
            g("Expectativas de precios", self.serie([("com_precios", "Comercio"), ("con_precios", "Construcción"),
                                                     ("ind_precios", "Industria")]),
              titulo="Expectativas de precios", hoja="Precios")])

        self.lamina_graficos("Presiones de salarios, costos y precios: promedio de los sectores", [
            gp("Salarios", "Expectativas de salarios (promedio)", "Promedio salarios"),
            gp("Costos", "Expectativas de costos (promedio)", "Promedio costos"),
            gp("Precios", "Expectativas de precios (promedio)", "Promedio precios")])

        self.lamina_graficos("Expectativas de empleo", [
            g("Expectativas de empleo", self.serie([("com_empleo", "Comercio"), ("con_empleo", "Construcción"),
                                                    ("ind_empleo", "Industria")], desde="2011-01-01"),
              titulo="Expectativas de empleo por sector", hoja="Empleo", meses_entre_etiquetas=12, nota=NOTA_50)])

        self.lamina_graficos("Expectativas de empleo: promedio de los sectores", [
            gp("Empleo", "Expectativas de empleo (promedio)", "Promedio empleo", NOTA_50, meses_entre_etiquetas=12)])

        self.lamina_graficos("Situación actual de inventarios y demanda", [
            g("Inventarios", self.serie([("com_inventarios", "Comercio"), ("ind_inventarios", "Industria")],
                                        desde="2011-01-01"),
              titulo="Situación actual de inventarios", hoja="Inventarios"),
            g("Demanda (desest.)", self.serie([("com_demanda", "Comercio"), ("ind_demanda", "Industria"),
                                               ("con_demanda", "Construcción")], sa=True, desde="2011-01-01"),
              titulo="Situación actual de la demanda", hoja="Demanda SA", unidad="(Índice, desestacionalizado)",
              nota=di.NOTA_SA)])

        self.lamina_graficos("Situación actual de inventarios y demanda: promedio de los sectores", [
            gp("Inventarios", "Inventarios (promedio)", "Promedio inventarios"),
            gp("Demanda (desest.)", "Demanda (promedio, desest.)", "Promedio demanda SA", di.NOTA_SA,
               unidad="(Índice, desestacionalizado)")])

        self.lamina_graficos("Expectativas de inversión, ventas y producción", [
            g("Expectativas de inversión", self.serie([("com_inversion", "Comercio"), ("ind_inversion", "Industria")]),
              titulo="Expectativas de inversión", hoja="Inversion", nota=NOTA_50),
            g("Ventas y producción (desest.)", self.serie([("com_ventas", "Comercio: ventas"),
                                                           ("ind_prod_esp", "Industria: producción")], sa=True),
              titulo="Expectativas de ventas y producción", hoja="Ventas y produccion SA",
              unidad="(Índice, desestacionalizado)", nota=di.NOTA_SA)])

    def compuestos(self):
        nombres = di.NOMBRES_COMPUESTOS
        pesos = ", ".join(f"{s.lower()} {num(100 * p, 0)}%" for s, p in di.PESOS.items())
        nota = ("Columnas: contribución de cada sector. Basado en Chanut, Marcel y Medel (2018): promedio de "
                "índices de difusión desestacionalizados (STL) menos 50 de cada sector, ponderado por su "
                f"participación en el PIB 2018 ({pesos}).")
        for par, titulo in ((("OCSI", "PCSI"), "Índices compuestos: situación actual"),
                            (("CFEI", "OFEI"), "Índices compuestos: expectativas")):
            graficos = []
            for k in par:
                c, tot = self.comp[k]
                d = self._dmc(tot)
                dc = c.loc[self.t] - c.loc[self.t - pd.DateOffset(months=1)]
                motor = dc.idxmax() if d > 0 else dc.idxmin()
                pos = "por sobre" if d > 0.05 else "por debajo de" if d < -0.05 else "igual a"
                vineta = (f"{nombres[k]} ({k}): {num(tot[self.t])} en {self.mes}, {pos} {self.mes_ant} "
                          f"({signo(d)} p.), principalmente por {motor.lower()}.")
                df = pd.concat([tot.rename(f"IMCE {k}"), c[di.SECTORES]], axis=1)
                df = df[df.index >= "2018-09-01"].dropna(how="all")
                y_min, y_max, paso = _limites_apilado(df)
                graficos.append(dict(df=df, vineta=vineta, hoja=k, apilado=True, nota=nota,
                                     titulo=f"IMCE: {nombres[k]} ({k})",
                                     unidad="(Índice 50 = 0, desestacionalizado)", posicion_unidad="arriba",
                                     barras=di.SECTORES, y_min=y_min, y_max=y_max, y_paso=paso,
                                     meses_entre_etiquetas=12, formato_fecha="yyyy", tam=11))
            self.lamina_graficos(titulo, graficos)

    # ---- armar y guardar
    def construir(self):
        OUT.mkdir(exist_ok=True)
        _archivar_anteriores()
        self.portada()
        self.resumen()
        self.resumen_sin_mineria()
        self.resultados()
        self.graficos()
        self.compuestos()
        nombre = f"Presentación_{TITULO}_{self.mes}_{self.t.year}"
        pptx = OUT / f"{nombre}.pptx"
        datos_xlsx = OUT / f"{nombre}_datos.xlsx"
        try:
            self.prs.save(pptx)
            with pd.ExcelWriter(datos_xlsx) as xl:
                for hoja, df in self.datos.items():
                    df.rename_axis(df.index.name or "Fecha").to_excel(xl, sheet_name=hoja[:31])
        except PermissionError as e:
            sys.exit(f"No pude guardar {e.filename}: está abierto. Ciérralo y vuelve a correr.")
        print(f"  ✅ PPTX: {pptx} ({self.n} láminas; datos a {self.mes} de {self.t.year})")
        print(f"  ✅ Datos de los gráficos: {datos_xlsx}")
        pdf = fg.exportar_pdf(pptx)
        if pdf:
            print(f"  ✅ PDF:  {pdf}")
        return pptx


# ─────────────────────────────────────────────────────────────────────────────
# archivo de presentaciones anteriores
# ─────────────────────────────────────────────────────────────────────────────

def _archivar_anteriores():
    """Mueve las presentaciones anteriores (PPTX, PDF y su _datos.xlsx sueltos en Outputs) a Outputs/<año>/, con
    la fecha al inicio del nombre (AAAAMMDD_<nombre>): la que el nombre ya traiga o, si no tiene, la de su
    última modificación. Base_IMCE.xlsx se queda en Outputs."""
    import re
    for archivo in sorted(OUT.iterdir()):
        if (not archivo.is_file() or archivo.name.startswith("~$") or archivo == di.EXCEL
                or not (archivo.suffix.lower() in (".pdf", ".pptx") or archivo.name.endswith("_datos.xlsx"))):
            continue
        fecha_nombre = re.match(r"(\d{8})_", archivo.name)
        if fecha_nombre:
            fecha, nombre = pd.to_datetime(fecha_nombre.group(1), format="%Y%m%d", errors="coerce"), archivo.name
        else:
            fecha = pd.Timestamp(dt.datetime.fromtimestamp(archivo.stat().st_mtime))
            nombre = f"{fecha:%Y%m%d}_{archivo.name}"
        if pd.isna(fecha):
            continue
        destino = OUT / str(fecha.year) / nombre
        destino.parent.mkdir(parents=True, exist_ok=True)
        try:
            archivo.replace(destino)
            print(f"  📁 {archivo.name} → {destino.parent.name}/{destino.name}")
        except PermissionError:
            print(f"  ⚠️ No pude mover {archivo.name}: está abierto. Ciérralo y vuelve a correr.")


# ─────────────────────────────────────────────────────────────────────────────
# auxiliares de gráficos
# ─────────────────────────────────────────────────────────────────────────────

def _limites(df, paso=10, piso=0, techo=80):
    """Eje Y en múltiplos de 'paso' que contiene las series, el 50 y (si se indican) piso y techo."""
    y = df.to_numpy(dtype=float)
    y = y[np.isfinite(y)]
    lo, hi = min(y.min(), 50), max(y.max(), 50)
    if piso is not None:
        lo = min(lo, piso)
    if techo is not None:
        hi = max(hi, techo)
    return paso * math.floor(lo / paso), paso * math.ceil(hi / paso)


def _limites_apilado(df):
    """Límites para columnas apiladas (positivas y negativas por separado) más la línea del total."""
    barras = df.iloc[:, 1:]
    lo = min(barras.clip(upper=0).sum(axis=1).min(), df.iloc[:, 0].min(), 0)
    hi = max(barras.clip(lower=0).sum(axis=1).max(), df.iloc[:, 0].max(), 0)
    paso = 5 if hi - lo <= 40 else 10
    return paso * math.floor(lo / paso), paso * math.ceil(hi / paso), paso


def _apilar(chart):
    """Columnas de formato_graficos (agrupadas) -> apiladas."""
    barra = chart._chartSpace.find(qn("c:chart")).find(qn("c:plotArea")).find(qn("c:barChart"))
    barra.find(qn("c:grouping")).set("val", "stacked")
    barra.find(qn("c:gapWidth")).addnext(barra.makeelement(qn("c:overlap"), {"val": "100"}))


def main(argv=None):
    """Proceso completo: datos (descarga, Inputs/imce.csv y Outputs/Base_IMCE.xlsx) y minuta."""
    import argparse
    ap = argparse.ArgumentParser(description="Minuta IMCE: descarga los datos, actualiza Base_IMCE.xlsx y "
                                             "arma la minuta (PPTX y PDF) en Outputs.")
    ap.add_argument("--sin-descarga", action="store_true",
                    help="no descarga: usa el Inputs/imce.csv existente (igual actualiza Base_IMCE.xlsx)")
    a = ap.parse_args(argv)
    di.main(["--sin-descarga"] if a.sin_descarga else [])   # 1) datos
    print("\n" + "=" * 60 + "\n  Minuta IMCE\n" + "=" * 60)       # 2) minuta
    return Minuta().construir()


if __name__ == "__main__":
    main()
