"""
macro.py - Datos y graficos de la Gerencia de Estudios (Santander Chile)

Uso basico:
    import macro
    df = macro.bde(["F073.TCO.PRE.Z.D"], nombres=["Dolar observado"], desde="2025-01-01")
    macro.grafico_lineas(df, titulo="El dolar sube", unidad="CLP por USD",
                         fuente="Banco Central de Chile", archivo="graficos.pptx")

Para armar presentaciones de varias laminas (ej. chartpacks):
    prs = macro.nueva_presentacion()
    macro.portada(prs, "octubre de 2026")
    lam = macro.lamina(prs, "IMCE", numero=2)
    macro.agregar_grafico(lam, df, x, y, ancho, alto, titulo=..., unidad=...)
    macro.pie(lam, x, y, ancho, fuente="Banco Central de Chile", nota=...)

Requisitos (una vez, en la terminal de VS Code):
    python -m pip install bcchapi pandas openpyxl requests python-pptx matplotlib

Credenciales (variables de entorno de Windows, comillas SIMPLES en PowerShell):
    setx BDE_TOKEN 'tu_token'        # API Key Token del BCCh (tambien acepta BCCH_API_token)
    setx FRED_API_KEY 'tu_key'       # FRED (St. Louis Fed) (tambien acepta FRED_API)
    setx IMF_API_KEY 'tu_key'        # opcional, FMI

Reglas que respeta este modulo (manual de estilo de Estudios):
    - Ninguna transformacion se aplica en silencio: cada funcion de transformacion
      devuelve tambien el texto de nota metodologica que corresponde.
    - No desestacionaliza, no deflacta, no empalma, no interpola, no imputa.
    - Graficos nativos de PowerPoint (editables), formato Santander.
"""
import copy
import datetime as dt
import io
import os
import sys
from pathlib import Path

import pandas as pd
import requests

ASSETS = Path(__file__).parent / "assets"

# =============================================================================
# 1. CREDENCIALES
# =============================================================================

# nombres alternativos con que la variable puede estar definida
_ALIAS = {
    "BDE_TOKEN": ["BCCH_API_token", "BCCH_API_TOKEN"],
    "FRED_API_KEY": ["FRED_API"],
}


def _credencial(nombre, obligatoria=True):
    valor = os.environ.get(nombre)
    for alias in _ALIAS.get(nombre, []):
        valor = valor or os.environ.get(alias)
    if obligatoria and not valor:
        sys.exit(f"No encuentro la variable {nombre}. En PowerShell: setx {nombre} 'valor' "
                 f"y luego cierra y abre VS Code.")
    return valor


# =============================================================================
# 2. DATOS
# =============================================================================

_SIETE = None


def _siete():
    """Conexion unica a la API BDE (se reutiliza entre llamadas)."""
    global _SIETE
    if _SIETE is None:
        import bcchapi
        token = _credencial("BDE_TOKEN")
        # La documentacion oficial del BCCh usa Siete(token=...).
        try:
            _SIETE = bcchapi.Siete(token=token)
        except TypeError:
            _SIETE = bcchapi.Siete(api_key=token)
    return _SIETE


def bde(series, nombres=None, desde=None, hasta=None, quitar_vacios=True):
    """
    Descarga series del BCCh (API BDE) en su frecuencia original.
    series  : lista de codigos BDE, ej. ["F073.TCO.PRE.Z.D"]
    nombres : lista de nombres para las columnas (mismo orden)
    desde / hasta : "AAAA-MM-DD"
    quitar_vacios : elimina filas donde TODAS las series estan vacias
                    (fines de semana y feriados en series diarias). No imputa nada.
    """
    if isinstance(series, str):
        series = [series]
    df = _siete().cuadro(series=series, nombres=nombres, desde=desde, hasta=hasta)
    df.index = pd.to_datetime(df.index)
    if quitar_vacios:
        df = df.dropna(how="all")
    df.attrs["fuente"] = "Banco Central de Chile"
    return df


def buscar_bde(texto):
    """Busca codigos en el catalogo BDE. Ej: buscar_bde('dolar observado')."""
    r = _siete().buscar(texto)
    return r[["seriesId", "frequencyCode", "spanishTitle"]]


def fred(series, nombres=None, desde=None, hasta=None):
    """Descarga series de FRED. series: lista de IDs, ej. ["FEDFUNDS", "DGS10"]."""
    if isinstance(series, str):
        series = [series]
    key = _credencial("FRED_API_KEY")
    columnas = []
    for i, sid in enumerate(series):
        params = {"series_id": sid, "api_key": key, "file_type": "json"}
        if desde:
            params["observation_start"] = desde
        if hasta:
            params["observation_end"] = hasta
        r = requests.get("https://api.stlouisfed.org/fred/series/observations",
                         params=params, timeout=30)
        r.raise_for_status()
        obs = pd.DataFrame(r.json()["observations"])
        s = pd.to_numeric(obs["value"], errors="coerce")  # "." = sin dato
        s.index = pd.to_datetime(obs["date"])
        s.name = nombres[i] if nombres else sid
        columnas.append(s)
    df = pd.concat(columnas, axis=1).dropna(how="all")
    df.attrs["fuente"] = "Federal Reserve Bank of St. Louis (FRED)"
    return df


def fmi_weo(indicador, paises, desde=None, hasta=None):
    """
    Datos anuales del WEO del FMI via DataMapper (sin key).
    indicador: ej. "NGDP_RPCH" (crecimiento PIB real), "PCPIPCH" (inflacion promedio)
    paises   : lista de codigos ISO3, ej. ["CHL", "PER", "COL"]
    Ojo: incluye proyecciones del FMI. Marcar como proyeccion en el grafico.
    """
    if isinstance(paises, str):
        paises = [paises]
    url = f"https://www.imf.org/external/datamapper/api/v1/{indicador}/{'/'.join(paises)}"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    valores = r.json()["values"][indicador]
    df = pd.DataFrame(valores)
    df.index = pd.to_datetime(df.index.astype(str) + "-01-01")
    df = df.sort_index()
    if desde:
        df = df[df.index >= desde]
    if hasta:
        df = df[df.index <= hasta]
    df.attrs["fuente"] = "Fondo Monetario Internacional (WEO)"
    return df


def nyfed_gscpi(desde=None, hasta=None, nombre="GSCPI"):
    """
    Global Supply Chain Pressure Index (mensual) desde la Fed de Nueva York (sin key).
    Unidad: desviaciones estandar respecto del promedio historico.
    """
    url = ("https://www.newyorkfed.org/medialibrary/research/interactives/gscpi/"
           "downloads/gscpi_data.xlsx")
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    crudo = pd.read_excel(io.BytesIO(r.content), sheet_name="GSCPI Monthly Data", header=None)
    crudo = crudo.iloc[:, :2].dropna()
    fechas = pd.to_datetime(crudo.iloc[:, 0], format="%d-%b-%Y", errors="coerce")
    df = pd.DataFrame({nombre: pd.to_numeric(crudo.iloc[:, 1], errors="coerce").values},
                      index=fechas.dt.to_period("M").dt.to_timestamp().values)
    df = df[df.index.notna()].dropna(how="all").sort_index()
    if desde:
        df = df[df.index >= desde]
    if hasta:
        df = df[df.index <= hasta]
    df.attrs["fuente"] = "Reserva Federal de Nueva York"
    return df


# =============================================================================
# 3. TRANSFORMACIONES (siempre explicitas, devuelven la nota)
# =============================================================================

def var_anual(df, periodos_por_anio=12, ya_es_anual=False):
    """
    Variacion % a/a. periodos_por_anio: 12 mensual, 4 trimestral.
    Si la serie ya viene expresada a/a, pasa ya_es_anual=True y no la transforma.
    Devuelve (df, nota).
    """
    if ya_es_anual:
        return df, ""
    return df.pct_change(periodos_por_anio, fill_method=None) * 100, \
        "Variación porcentual respecto del mismo período del año anterior."


def media_movil(df, n, unidad="observaciones"):
    """Media movil simple de n observaciones (no dias calendario). Devuelve (df, nota)."""
    return df.rolling(n).mean(), f"Promedio móvil de {n} {unidad}."


def base100(df, fecha):
    """Indice base 100 en la fecha dada. Devuelve (df, nota)."""
    base = df.loc[pd.to_datetime(fecha)]
    etiqueta = pd.to_datetime(fecha).strftime("%Y")
    return df / base * 100, f"Índice {etiqueta}=100."


def promedio(df, nombre):
    """
    Promedio simple de las columnas de df (solo fechas con todas las series).
    Devuelve (df de una columna, nota).
    """
    s = df.dropna().mean(axis=1).rename(nombre)
    cols = " y ".join(str(c) for c in df.columns) if df.shape[1] == 2 else ", ".join(map(str, df.columns))
    return s.to_frame(), f"{nombre}: promedio simple de {cols}."


# =============================================================================
# 4. ESTILO SANTANDER
# =============================================================================

def _rgb(r, g, b):
    from pptx.dml.color import RGBColor
    return RGBColor(r, g, b)


PALETA = [(236, 0, 0),      # rojo Santander (serie principal)
          (0, 0, 0),        # negro
          (77, 196, 202),   # turquesa
          (255, 174, 66),   # amarillo
          (178, 178, 178)]  # gris
COLOR_EJES = (191, 191, 191)  # "Blanco, fondo 1, oscuro 25%"
GRIS_TEXTO = (89, 89, 89)
FUENTES_SANTANDER = {
    "titulo": "Santander Headline",
    "ejes": "Santander Micro Text",
    "fuente": "Santander Text",
}
FUENTE_RESPALDO = "Arial"   # siempre instalada en Windows, sans serif neutra
FORMATO_FECHA = "mmm.yy"    # ene.26 (meses en minusculas con Office en espanol)
# fija el idioma de los formatos en espanol (Chile) aunque el Office este en otro idioma
IDIOMA_FORMATO = "[$-340A]"
ANCHO_CM, ALTO_CM = 15, 11
SLIDE_ANCHO_CM, SLIDE_ALTO_CM = 33.867, 19.05  # 16:9


def _fuente_instalada(nombre):
    """Revisa si la tipografia esta instalada en este computador."""
    try:
        from matplotlib import font_manager
        nombres = {f.name.lower() for f in font_manager.fontManager.ttflist}
        if nombre.lower() in nombres:
            return True
        # fuentes instaladas solo para el usuario en Windows
        carpeta = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")
        if os.path.isdir(carpeta):
            clave = nombre.lower().replace(" ", "")
            return any(clave in f.lower().replace(" ", "").replace("-", "")
                       for f in os.listdir(carpeta))
    except Exception:
        pass
    return False


_TIPOGRAFIAS = {}


def tipografia(rol):
    """Devuelve la tipografia Santander del rol, o Arial si no esta instalada."""
    if rol not in _TIPOGRAFIAS:
        nombre = FUENTES_SANTANDER[rol]
        if _fuente_instalada(nombre):
            _TIPOGRAFIAS[rol] = nombre
        else:
            print(f"Aviso: no encuentro '{nombre}'; uso {FUENTE_RESPALDO}.")
            _TIPOGRAFIAS[rol] = FUENTE_RESPALDO
    return _TIPOGRAFIAS[rol]


def _formato_numero(decimales):
    return "#,##0" if decimales == 0 else "#,##0." + "0" * decimales


# =============================================================================
# 5. GRAFICOS NATIVOS DE POWERPOINT (editables)
# =============================================================================

def nueva_presentacion():
    from pptx import Presentation
    from pptx.util import Cm
    prs = Presentation()
    prs.slide_width, prs.slide_height = Cm(SLIDE_ANCHO_CM), Cm(SLIDE_ALTO_CM)
    return prs


def _abrir_ppt(archivo):
    from pptx import Presentation
    if os.path.exists(archivo):
        return Presentation(archivo)
    return nueva_presentacion()


def _caja_texto(slide, texto, x, y, w, h, nombre_fuente, tam, negrita=False, color=(0, 0, 0),
                alineacion=None):
    """Caja de texto; cada salto de linea en 'texto' es un parrafo."""
    from pptx.util import Pt
    caja = slide.shapes.add_textbox(x, y, w, h)
    tf = caja.text_frame
    tf.word_wrap = True
    for i, linea in enumerate(texto.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if alineacion is not None:
            p.alignment = alineacion
        run = p.add_run()
        run.text = linea
        run.font.name = nombre_fuente
        run.font.size = Pt(tam)
        run.font.bold = negrita
        run.font.color.rgb = _rgb(*color)
    return caja


def _linea_horizontal(slide, x, y, w, color=COLOR_EJES, grosor_pt=0.75):
    """Linea separadora (rectangulo delgado sin sombra ni estilo de tema)."""
    from pptx.oxml.ns import qn
    from pptx.util import Pt
    forma = slide.shapes.add_shape(1, x, y, w, Pt(grosor_pt))
    forma.fill.solid()
    forma.fill.fore_color.rgb = _rgb(*color)
    forma.line.fill.background()
    forma.shadow.inherit = False
    estilo = forma._element.find(qn("p:style"))
    if estilo is not None:
        forma._element.remove(estilo)
    return forma


def portada(prs, fecha_texto, titulo="Estudios\nEconómicos"):
    """Portada institucional: foto, logo Santander, titulo y fecha."""
    from pptx.util import Cm
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(str(ASSETS / "portada.png"), Cm(0.39), Cm(0.35), Cm(33.09), Cm(18.34))
    s.shapes.add_picture(str(ASSETS / "logo_portada.png"), Cm(28.68), Cm(0.32), Cm(4.09), Cm(5.12))
    _caja_texto(s, titulo, Cm(23.6), Cm(7.9), Cm(9.5), Cm(3.6), tipografia("titulo"), 36,
                color=PALETA[0])
    _linea_horizontal(s, Cm(23.74), Cm(11.78), Cm(7.44), color=(31, 78, 121))
    _caja_texto(s, fecha_texto, Cm(23.8), Cm(11.9), Cm(7), Cm(0.9), tipografia("titulo"), 14,
                color=PALETA[0])
    return s


def lamina(prs, titulo_lamina, numero=None):
    """Lamina en blanco con titulo rojo (24, negrita), logo, separadores y numero."""
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Cm
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _caja_texto(s, titulo_lamina, Cm(1.1), Cm(0.25), Cm(29), Cm(1.5), tipografia("titulo"), 24,
                negrita=True, color=PALETA[0])
    _linea_horizontal(s, Cm(1.3), Cm(2.03), Cm(30.1))
    s.shapes.add_picture(str(ASSETS / "logo.png"), Cm(32.1), Cm(0), Cm(1.2), Cm(1.55))
    _linea_horizontal(s, Cm(1.45), Cm(17.5), Cm(31))
    if numero is not None:
        _caja_texto(s, str(numero), Cm(30.2), Cm(17.75), Cm(2), Cm(0.9), tipografia("titulo"), 14,
                    negrita=True, color=(127, 127, 127), alineacion=PP_ALIGN.RIGHT)
    return s


def pie(slide, x, y, w, fuente, nota=None, h=None):
    """Nota metodologica (si hay) y fuente(s) bajo el grafico: 'Nota: ...' antes de 'Fuente: ...'."""
    from pptx.util import Cm
    fuentes = fuente if isinstance(fuente, (list, tuple)) else [fuente]
    etiqueta = "Fuentes: " if len(fuentes) > 1 else "Fuente: "
    texto = ""
    if nota:
        texto += f"Nota: {nota}\n"
    texto += etiqueta + "; ".join(fuentes) + "."
    return _caja_texto(slide, texto, x, y, w, h or Cm(1.2), tipografia("fuente"), 10)


def _sub(padre, tag, **attrs):
    from pptx.oxml.ns import qn
    from lxml import etree
    el = etree.SubElement(padre, qn(tag))
    for k, v in attrs.items():
        el.set(k, str(v))
    return el


def _num_fmt(eje_xml, codigo):
    from pptx.oxml.ns import qn
    nf = eje_xml.find(qn("c:numFmt"))
    nf.set("formatCode", codigo)
    nf.set("sourceLinked", "0")


def _escala(eje, minimo, maximo, paso):
    if minimo is not None:
        eje.minimum_scale = minimo
    if maximo is not None:
        eje.maximum_scale = maximo
    if paso is not None:
        eje.major_unit = paso


def _formato_eje_valor(eje, decimales, paso, tam, nombre_fuente):
    from pptx.enum.chart import XL_TICK_MARK
    from pptx.util import Pt
    eje.major_tick_mark = XL_TICK_MARK.OUTSIDE
    eje.minor_tick_mark = XL_TICK_MARK.NONE
    eje.has_major_gridlines = False
    # eje sin decimales salvo que el paso sea fraccional (ej. 0,5)
    dec_eje = 0 if decimales <= 1 else decimales
    if paso is not None and paso != int(paso):
        dec_eje = max(dec_eje, len(f"{paso:g}".split(".")[1]))
    eje.tick_labels.number_format = IDIOMA_FORMATO + _formato_numero(dec_eje)
    eje.tick_labels.number_format_is_linked = False
    eje.tick_labels.font.size = Pt(tam)
    eje.tick_labels.font.name = nombre_fuente
    eje.tick_labels.font.color.rgb = _rgb(0, 0, 0)
    eje.format.line.width = Pt(0.75)
    eje.format.line.color.rgb = _rgb(*COLOR_EJES)


def _eje_derecho(chart, columnas_der):
    """Mueve las series indicadas a un grafico de lineas con eje Y secundario (derecha)."""
    from pptx.oxml.ns import qn
    plot_area = chart._chartSpace.find(qn("c:chart")).find(qn("c:plotArea"))
    linea1 = plot_area.find(qn("c:lineChart"))
    fecha1 = plot_area.find(qn("c:dateAx"))
    valor1 = plot_area.find(qn("c:valAx"))
    id_x2, id_y2 = "50050", "50060"

    linea2 = copy.deepcopy(linea1)
    for el in list(linea2):
        if el.tag in (qn("c:ser"), qn("c:axId"), qn("c:dLbls")):
            linea2.remove(el)
    for ser in list(linea1.findall(qn("c:ser"))):
        nombre = ser.find(qn("c:tx")).find(".//" + qn("c:v")).text
        if nombre in columnas_der:
            linea1.remove(ser)
            linea2.insert(len(linea2) - 1 if linea2.find(qn("c:marker")) is not None else len(linea2), ser)
    _sub(linea2, "c:axId", val=id_x2)
    _sub(linea2, "c:axId", val=id_y2)
    linea1.addnext(linea2)

    fecha2 = copy.deepcopy(fecha1)
    fecha2.find(qn("c:axId")).set("val", id_x2)
    fecha2.find(qn("c:delete")).set("val", "1")
    fecha2.find(qn("c:crossAx")).set("val", id_y2)
    cruce_x = fecha2.find(qn("c:crossesAt"))
    if cruce_x is not None:
        cruce_x.addprevious(_sub(fecha2, "c:crosses", val="autoZero"))
        fecha2.remove(cruce_x)
    valor2 = copy.deepcopy(valor1)
    # escala propia: no hereda minimo, maximo ni paso del eje izquierdo
    for tag in ("c:max", "c:min"):
        el = valor2.find(qn("c:scaling")).find(qn(tag))
        if el is not None:
            el.getparent().remove(el)
    paso = valor2.find(qn("c:majorUnit"))
    if paso is not None:
        valor2.remove(paso)
    valor2.find(qn("c:axId")).set("val", id_y2)
    valor2.find(qn("c:axPos")).set("val", "r")
    valor2.find(qn("c:crossAx")).set("val", id_x2)
    cruce = valor2.find(qn("c:crosses"))
    if cruce is None:
        cruce = valor2.find(qn("c:crossesAt"))
    nuevo = _sub(valor2, "c:crosses", val="max")
    cruce.addprevious(nuevo)
    valor2.remove(cruce)
    valor1.addnext(fecha2)
    fecha2.addnext(valor2)
    return valor2


def agregar_grafico(slide, df, x, y, w, h, titulo=None, unidad=None, decimales=1,
                    proyecciones=None, etiqueta_ultimo=True,
                    y_min=None, y_max=None, y_paso=None,
                    eje_derecho=None, unidad_derecha=None, y2_min=None, y2_max=None, y2_paso=None,
                    meses_entre_etiquetas=None, formato_fecha=FORMATO_FECHA,
                    linea_ref=None, posicion_unidad="abajo", tam=12):
    """
    Grafico de lineas nativo (editable) en formato Santander, en la posicion dada de una lamina.
    df           : DataFrame indexado por fecha; cada columna es una serie (max. 5)
    titulo       : titulo del grafico (14, negrita), sobre el grafico
    unidad       : descriptor dentro del grafico, ej. "(%)"
    proyecciones : lista de columnas que se dibujan con linea discontinua
    eje_derecho  : lista de columnas que van al eje Y secundario (derecha)
    meses_entre_etiquetas : distancia entre etiquetas del eje X (ej. 12 = una por anio)
    linea_ref    : valor del eje Y donde cruza el eje X (ej. 50 en indices de difusion);
                   las fechas siguen abajo
    posicion_unidad : "abajo" (izquierda abajo, por defecto) o "arriba" (izquierda arriba)
    """
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import (XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION,
                                 XL_MARKER_STYLE, XL_TICK_LABEL_POSITION, XL_TICK_MARK)
    from pptx.enum.dml import MSO_LINE_DASH_STYLE
    from pptx.oxml.ns import qn
    from pptx.util import Cm, Pt

    if df.shape[1] > 5:
        raise ValueError("Máximo 5 series por gráfico (legibilidad y paleta).")
    proyecciones = proyecciones or []
    eje_derecho = eje_derecho or []
    f_tit, f_ejes = tipografia("titulo"), tipografia("ejes")

    if titulo:
        _caja_texto(slide, titulo, x, y - Cm(1.1), w, Cm(1), f_tit, 14, negrita=True)

    # --- datos ---
    datos = CategoryChartData(number_format=_formato_numero(decimales))
    datos.categories = [d.date() for d in df.index]
    for col in df.columns:
        datos.add_series(str(col), [None if pd.isna(v) else float(v) for v in df[col]])
    marco = slide.shapes.add_chart(XL_CHART_TYPE.LINE, x, y, w, h, datos)
    chart = marco.chart
    chart.has_title = False   # el titulo va en una caja de texto sobre el grafico
    chart.font.name = f_ejes
    chart.font.size = Pt(tam)
    # texto negro aplicado a ejes y leyenda (no al grafico completo: algunos visores,
    # como LibreOffice, pintan con ese color los simbolos de la leyenda)

    # --- leyenda dentro del grafico (se puede mover a mano en PowerPoint) ---
    chart.has_legend = df.shape[1] > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.TOP
        # la leyenda reserva su espacio arriba, sin tapar las series
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(tam)
        chart.legend.font.name = f_ejes
        chart.legend.font.color.rgb = _rgb(0, 0, 0)

    # --- eje X: fechas siempre abajo, formato ene.26 ---
    ex = chart.category_axis
    ex.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    ex.major_tick_mark = XL_TICK_MARK.OUTSIDE
    ex.minor_tick_mark = XL_TICK_MARK.NONE
    ex.tick_labels.number_format = IDIOMA_FORMATO + formato_fecha
    ex.tick_labels.number_format_is_linked = False
    ex.tick_labels.font.size = Pt(tam)
    ex.tick_labels.font.name = f_ejes
    ex.tick_labels.font.color.rgb = _rgb(0, 0, 0)
    ex.format.line.width = Pt(0.75)
    ex.format.line.color.rgb = _rgb(*COLOR_EJES)
    ex.has_major_gridlines = False
    if meses_entre_etiquetas:
        _sub(ex._element, "c:majorUnit", val=meses_entre_etiquetas)
        _sub(ex._element, "c:majorTimeUnit", val="months")

    # --- eje Y ---
    ey = chart.value_axis
    _formato_eje_valor(ey, decimales, y_paso, tam, f_ejes)
    _escala(ey, y_min, y_max, y_paso)
    if linea_ref is not None:
        ey.crosses_at = linea_ref
        ex.major_tick_mark = XL_TICK_MARK.NONE
        ex.format.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        ex.format.line.color.rgb = _rgb(127, 127, 127)

    # --- eje Y secundario ---
    if eje_derecho:
        valor2 = _eje_derecho(chart, [str(c) for c in eje_derecho])
        from pptx.chart.axis import ValueAxis
        ey2 = ValueAxis(valor2)
        _formato_eje_valor(ey2, decimales, y2_paso, tam, f_ejes)
        _escala(ey2, y2_min, y2_max, y2_paso)

    # --- series (en el orden de df, aunque esten en dos ejes) ---
    series = {s.name: s for plot in chart.plots for s in plot.series}
    for i, col in enumerate(df.columns):
        serie = series[str(col)]
        color = _rgb(*PALETA[i])
        serie.smooth = False
        serie.format.line.width = Pt(1.5)
        serie.format.line.color.rgb = color
        serie.marker.style = XL_MARKER_STYLE.NONE
        # el marcador (invisible) lleva el color de la linea: algunos visores (LibreOffice)
        # dibujan el simbolo de la leyenda con el color de fuente si no
        serie.marker.format.line.color.rgb = color
        if col in proyecciones:
            serie.format.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        if etiqueta_ultimo:
            idx = df[col].last_valid_index()
            if idx is not None:
                pos = df.index.get_loc(idx)
                etq = serie.points[pos].data_label
                etq.has_text_frame = False
                etq.show_value = True
                etq.number_format = IDIOMA_FORMATO + _formato_numero(decimales)
                etq.number_format_is_linked = False
                etq.font.size = Pt(10)
                etq.font.bold = True
                etq.font.color.rgb = color
                etq.position = XL_LABEL_POSITION.RIGHT

    # --- descriptor de unidad dentro del grafico ---
    y_unidad = y + h - Cm(2.2) if posicion_unidad == "abajo" else y + Cm(1.3)
    if unidad:
        _caja_texto(slide, unidad, x + Cm(2.0), y_unidad, Cm(8), Cm(0.8), f_ejes, tam)
    if unidad_derecha:
        from pptx.enum.text import PP_ALIGN
        _caja_texto(slide, unidad_derecha, x + w - Cm(9.5), y_unidad, Cm(8), Cm(0.8), f_ejes, tam,
                    alineacion=PP_ALIGN.RIGHT)
    return chart


def grafico_lineas(df, titulo, unidad, fuente, nota=None, archivo="graficos.pptx",
                   decimales=1, proyecciones=None, etiqueta_ultimo=True,
                   titulo_lamina=None):
    """
    Agrega una lamina con un grafico de lineas nativo (editable) en formato Santander.
    df           : DataFrame indexado por fecha; cada columna es una serie (max. 5)
    titulo       : titulo del grafico (preferir mensaje economico)
    unidad       : descriptor dentro del grafico, ej. "Tasas a 10 años (%)"
    fuente       : texto despues de "Fuente:" (o lista para "Fuentes:")
    nota         : nota metodologica (va antes de la fuente)
    proyecciones : lista de columnas que se dibujan con linea discontinua
    titulo_lamina: titulo grande de la lamina (rojo, 24); opcional
    """
    from pptx.util import Cm
    prs = _abrir_ppt(archivo)
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # lamina en blanco
    x0, y0 = Cm(2), Cm(3.2)
    if titulo_lamina:
        _caja_texto(slide, titulo_lamina, Cm(1.5), Cm(0.8), Cm(30), Cm(1.8),
                    tipografia("titulo"), 24, negrita=True, color=PALETA[0])
    agregar_grafico(slide, df, x0, y0, Cm(ANCHO_CM), Cm(ALTO_CM), titulo=titulo, unidad=unidad,
                    decimales=decimales, proyecciones=proyecciones, etiqueta_ultimo=etiqueta_ultimo)
    pie(slide, x0, y0 + Cm(ALTO_CM) + Cm(0.2), Cm(ANCHO_CM), fuente, nota)
    prs.save(archivo)
    print(f"Gráfico agregado a {archivo} (lámina {len(prs.slides)}).")
    return archivo


def vista_previa(df, titulo="", decimales=1):
    """Vista rapida con matplotlib (solo para revisar, no para presentar)."""
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(figsize=(15 / 2.54, 11 / 2.54))
    for i, col in enumerate(df.columns):
        ax.plot(df.index, df[col], color=[c / 255 for c in PALETA[i]], lw=1.5, label=col)
    ax.set_title(titulo, fontweight="bold")
    if df.shape[1] > 1:
        ax.legend(frameon=False)
    for lado in ("top", "right"):
        ax.spines[lado].set_visible(False)
    plt.tight_layout()
    plt.show()


# =============================================================================
# 6. EJEMPLO
# =============================================================================

if __name__ == "__main__":
    df = bde(["F073.TCO.PRE.Z.D"], nombres=["Dólar observado"], desde="2025-01-01")
    df.to_excel(f"dolar_observado_{dt.date.today():%Y%m%d}.xlsx")
    grafico_lineas(df,
                   titulo="Dólar observado",
                   unidad="(CLP por USD)",
                   fuente="Banco Central de Chile",
                   decimales=0,
                   archivo="prueba_macro.pptx")
