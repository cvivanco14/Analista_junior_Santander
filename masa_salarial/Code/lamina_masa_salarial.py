# -*- coding: utf-8 -*-
r"""Lámina de masa salarial real en PPTX (gráficos nativos editables) y PDF, con formato_graficos.py.

Arma Outputs\AAAAMMDD_Masa_salarial.pptx (portada + lámina) y su PDF. masa_salarial.py la llama al
final; también se puede correr sola (usa los archivos que ya están en Inputs):
    python lamina_masa_salarial.py

formato_graficos.py (y su carpeta assets) vive en Estudios - Documentos\General\Administrativos\
Manual de Estilo; se busca igual que en la minuta IMCE: subiendo desde Code, en OneDrive/SharePoint y
en la carpeta Santander de iCloud Drive. FORMATO_GRAFICOS_DIR permite indicar otra carpeta.
"""
import datetime as dt
import math
import os
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

CODE = Path(__file__).resolve().parent
MANUAL_ESTILO = Path("Administrativos") / "Manual de Estilo"
BIBLIOTECA = Path("Estudios - Documentos") / "General"
ICLOUD = Path.home() / "Library" / "Mobile Documents" / "com~apple~CloudDocs"

TITULO = "Masa salarial real"
AUTOR = "Estudios Santander"   # aparece en la portada y en el pie de la lámina
FUENTE = "INE (Encuesta Nacional de Empleo e Índice de Remuneraciones)"
NOTA = ("Masa salarial real = ocupados (desestacionalizados por el INE) × horas habituales semanales × "
        "índice de remuneraciones real (ambos desestacionalizados con STL). Contribuciones: descomposición "
        "logarítmica de la variación anual. Periodo: último mes del trimestre móvil.")
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre",
         "octubre", "noviembre", "diciembre"]


# ─────────────────────────────────────────────────────────────────────────────
# formato_graficos.py
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
            print(f"  ⚠️ formato_graficos.py está en iCloud pero no descargado en este Mac:\n    {carpeta}\n"
                  "     Ábrelo en Finder y elige 'Descargar ahora'.")
            return None
    return None


def carpeta_formato():
    """Carpeta de formato_graficos.py, o None si no se encuentra."""
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
    return None


def importar_formato():
    carpeta = carpeta_formato()
    if carpeta is None:
        return None
    sys.path.insert(0, str(carpeta))
    import formato_graficos
    print(f"  Formato: {carpeta / 'formato_graficos.py'}")
    return formato_graficos


# ─────────────────────────────────────────────────────────────────────────────
# Lámina
# ─────────────────────────────────────────────────────────────────────────────

def _num(x, dec=1):
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _signo(x, dec=1):
    return f"{'+' if x > 0 else ''}{_num(x, dec)}"


def _limites(valores, paso):
    v = np.asarray(valores, dtype=float)
    v = v[np.isfinite(v)]
    return paso * math.floor(v.min() / paso), paso * math.ceil(v.max() / paso)


def _limites_apilado(total, barras):
    lo = min(barras.clip(upper=0).sum(axis=1).min(), total.min(), 0)
    hi = max(barras.clip(lower=0).sum(axis=1).max(), total.max(), 0)
    paso = 1 if hi - lo <= 8 else 2 if hi - lo <= 16 else 5
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


def _portada(fg, prs, subtitulo, fecha):
    """Portada con la firma de la versión del Manual de Estilo que esté instalada."""
    try:
        return fg.portada(prs, "Masa salarial real", subtitulo, autor=AUTOR, fecha=fecha)
    except TypeError:   # versión anterior: portada(prs, fecha_texto, titulo)
        return fg.portada(prs, f"{subtitulo}\n{fecha}", titulo="Masa salarial\nreal")


def _lamina(fg, prs, titulo, n, pie):
    try:
        return fg.lamina(prs, titulo, n, pie_texto=pie)
    except TypeError:   # versión anterior, sin pie de lámina
        return fg.lamina(prs, titulo, n)


def construir(ms, out_dir, hoy=None, desde="2015-01-01"):
    """ms: DataFrame de masa_salarial.masa_salarial(). Devuelve la ruta del PPTX o None."""
    import masa_salarial as m
    fg = importar_formato()
    if fg is None:
        print("  ⚠️ No encontré formato_graficos.py (Manual de Estilo): no se genera la lámina PPTX.\n"
              "     Para indicar la carpeta: variable FORMATO_GRAFICOS_DIR.")
        return None
    from pptx.util import Cm

    hoy = pd.Timestamp(hoy or dt.date.today())
    d = ms[ms.index >= desde].dropna(subset=[m.MASA_AA])
    t = d.index[-1]
    mes = f"{MESES[t.month - 1]} {t.year}"
    trimestre = f"trimestre móvil terminado en {mes}"
    prs = fg.nueva_presentacion()
    _portada(fg, prs, mes.capitalize(), f"{hoy.day:02d} de {MESES[hoy.month - 1]} de {hoy.year}")
    lam = _lamina(fg, prs, "Masa salarial real", 2, f"{TITULO} {mes}  |  {AUTOR}")

    u = d.iloc[-1]
    c = {k: u[v] for k, v in m.CONTRIB.items()}
    motor = max(c, key=lambda k: abs(c[k]))
    nombres = {m.OCUPADOS: "los ocupados", m.HORAS: "las horas trabajadas", m.IR_REAL: "el salario real"}
    texto = (f"En el {trimestre}, la masa salarial real varió {_signo(u[m.MASA_AA])}% a/a "
             f"({_signo(u[m.MASA_AA] - d[m.MASA_AA].iloc[-2])} pp respecto del mes anterior). Aportaron: "
             f"ocupados {_signo(c[m.OCUPADOS])} pp, horas {_signo(c[m.HORAS])} pp y salario real "
             f"{_signo(c[m.IR_REAL])} pp; el principal factor fue {nombres[motor]}. Sin ajustar por horas "
             f"(ocupados × salario real), la variación fue {_signo(u[m.MASA_SH_AA])}% a/a.")
    fg._caja_texto(lam, texto, Cm(1.3), Cm(2.4), Cm(31), Cm(1.8), fg.tipografia("fuente"), 13,
                   color=getattr(fg, "GRIS_TEXTO", (89, 89, 89)))

    y, alto = Cm(5.6), Cm(9.6)
    xs, ancho = [Cm(1.3), Cm(17.5)], Cm(getattr(fg, "ANCHO_CM", 15))
    datos = {}

    # 1) variación anual y contribuciones (columnas apiladas + línea)
    barras = {v: v.replace("Contribución ", "").replace(" (pp)", "").capitalize() for v in m.CONTRIB.values()}
    g1 = pd.concat([d[m.MASA_AA].rename("Masa salarial real"), d[list(barras)].rename(columns=barras)], axis=1)
    y_min, y_max, paso = _limites_apilado(g1.iloc[:, 0], g1.iloc[:, 1:])
    chart = fg.agregar_grafico(lam, g1, xs[0], y, ancho, alto, titulo="Variación anual y contribuciones",
                               unidad="(Var. a/a, %; contribuciones en pp)", posicion_unidad="arriba",
                               barras=list(barras.values()), etiqueta_ultimo=["Masa salarial real"],
                               y_min=y_min, y_max=y_max, y_paso=paso, meses_entre_etiquetas=12,
                               formato_fecha="yyyy", tam=11)
    _apilar(chart)
    datos["Var. anual y contribuciones"] = g1

    # 2) nivel
    g2 = d[[m.MASA, m.MASA_SH]].rename(columns={m.MASA: "Ocupados × horas × salario real",
                                                  m.MASA_SH: "Ocupados × salario real"})
    y_min, y_max = _limites(g2.to_numpy(), 5)
    fg.agregar_grafico(lam, g2, xs[1], y, ancho, alto, titulo="Nivel",
                       unidad=f"(Índice {m.ANIO_BASE} = 100, desestacionalizado)", posicion_unidad="arriba",
                       y_min=y_min, y_max=y_max, y_paso=5, meses_entre_etiquetas=12, formato_fecha="yyyy",
                       tam=11)
    datos["Nivel"] = g2
    fg.pie(lam, Cm(1.3), y + alto + Cm(0.15), Cm(31), FUENTE, NOTA, h=Cm(1.5))

    nombre = f"{hoy:%Y%m%d}_Masa_salarial"
    pptx = out_dir / f"{nombre}.pptx"
    try:
        prs.save(pptx)
        with pd.ExcelWriter(out_dir / f"{nombre}_datos.xlsx") as xl:
            for hoja, df in datos.items():
                df.rename_axis("Fecha").to_excel(xl, sheet_name=hoja[:31])
    except PermissionError as e:
        sys.exit(f"No pude guardar {e.filename}: está abierto. Ciérralo y vuelve a correr.")
    print(f"  ✅ PPTX: {pptx}")
    try:
        pdf = fg.exportar_pdf(pptx)
        if pdf:
            print(f"  ✅ PDF:  {pdf}")
    except Exception as e:  # noqa: BLE001
        print(f"  ⚠️ No pude exportar el PDF ({e}); el PPTX quedó listo.")
    return pptx


if __name__ == "__main__":
    import datos_empleo as de
    import masa_salarial as m
    datos = de.cargar(descarga=False, hojas=["ajuste_estacional", "horas", "remuneraciones"])
    construir(m.masa_salarial(datos), de.OUT)
