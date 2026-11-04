"""
formato_graficos.py - Datos y graficos de la Gerencia de Estudios (Santander Chile)

Uso basico:
    import formato_graficos as fg
    df = fg.bde(["F073.TCO.PRE.Z.D"], nombres=["Dolar observado"], desde="2025-01-01")
    fg.grafico_lineas(df, titulo="El dolar sube", unidad="CLP por USD",
                      fuente="Banco Central de Chile", archivo="graficos.pptx")

Para armar presentaciones de varias laminas (ej. chartpacks):
    prs = fg.nueva_presentacion()
    fg.portada(prs, "Chartpack de\nExpectativas", "Noviembre 2026", autor="Estudios Económicos")
    lam = fg.lamina(prs, "IMCE", numero=2, pie_texto="Chartpack de Expectativas noviembre 2026 | Estudios Económicos")
    fg.agregar_grafico(lam, df, x, y, ancho, alto, titulo=..., unidad=...)   # lineas
    fg.agregar_barras(lam, df_anual, x, y, ancho, alto, titulo=..., unidad=...)   # barras
    fg.pie(lam, x, y, ancho, fuente="Banco Central de Chile", nota=...)

Requisitos (una vez, en la terminal de VS Code):
    python -m pip install bcchapi pandas openpyxl requests python-pptx matplotlib

Credenciales (una vez por computador; nunca en el codigo):
  macOS, en la consola de Spyder (quedan en el Llavero):
    import keyring, getpass; keyring.set_password('estudios_santander', 'BDE_TOKEN', getpass.getpass('Token BCCh: '))
    (igual para 'FRED_API_KEY')
  Windows, variables de entorno (Anaconda Prompt: comillas DOBLES; PowerShell: SIMPLES):
    setx BDE_TOKEN "tu_token"        # API Key Token del BCCh (tambien acepta BCCH_API_token)
    setx BDE_USER "tu_correo"        # usuario y clave de la cuenta del BDE: alternativa al token;
    setx BDE_PASS "tu_clave"         #   se usan si no hay token o si el BCCh lo rechaza
  Si no hay token valido ni usuario y clave guardados, el modulo los pide en pantalla.
    setx FRED_API_KEY "tu_key"       # FRED (St. Louis Fed) (tambien acepta FRED_API)
    setx IMF_API_KEY "tu_key"        # opcional, FMI

Reglas que respeta este modulo (manual de estilo de Estudios):
    - Ninguna transformacion se aplica en silencio: cada funcion de transformacion
      devuelve tambien el texto de nota metodologica que corresponde.
    - No desestacionaliza, no deflacta, no empalma, no interpola, no imputa.
    - Graficos nativos de PowerPoint (editables), formato Santander:
        lineas de 2,25 pt (series mensuales o menos frecuentes) y 1,75 pt (diarias);
        fechas del eje X horizontales; numeros sin ceros sobrantes (2 y no 2,0);
        etiqueta del ultimo dato (en series diarias incluye el dia);
        descriptor de la serie (entre parentesis) en tamano 11;
        nota y fuente en Santander Text Light tamano 9.
"""
import copy
import datetime as dt
import io
import os
import sys
from pathlib import Path

import pandas as pd
import requests

MESES_LARGOS = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
                "septiembre", "octubre", "noviembre", "diciembre"]

# =============================================================================
# 1. CREDENCIALES
# =============================================================================

# nombres alternativos con que la variable puede estar definida
_ALIAS = {
    # API Key Token del BDE: metodo de acceso propio (no es la clave de la cuenta)
    "BDE_TOKEN": ["BDE_API_TOKEN", "BCCH_API_token", "BCCH_API_TOKEN"],
    # usuario y clave de la cuenta del BDE: van siempre juntos
    "BDE_USER": ["BCCH_USER"],
    "BDE_PASS": ["BCCH_PASS"],
    "FRED_API_KEY": ["FRED_API"],
}


LLAVERO = "estudios_santander"   # servicio en el Llavero de macOS / Administrador de credenciales


def _limpiar(valor):
    """Sin espacios ni comillas alrededor (setx 'x' en Anaconda Prompt/cmd las guarda)."""
    return valor.strip().strip("'\"").strip() if valor else valor


def _llavero(nombre):
    try:
        import keyring
        return keyring.get_password(LLAVERO, nombre)
    except Exception:  # noqa: BLE001 - sin keyring o sin llavero disponible
        return None


def _candidatos(nombre):
    """[(origen, valor)] en orden de prioridad: variable, sus alias y el llavero; sin repetidos."""
    vistos, out = set(), []
    for origen in [nombre, *_ALIAS.get(nombre, [])]:
        v = _limpiar(os.environ.get(origen))
        if v and v not in vistos:
            vistos.add(v)
            out.append((f"variable {origen}", v))
    v = _limpiar(_llavero(nombre))
    if v and v not in vistos:
        out.append(("Llavero", v))
    return out


def _como_guardar(nombre):
    if sys.platform == "darwin":
        return (f"En la consola de Spyder (una vez; pide el valor sin mostrarlo):\n"
                f"    import keyring, getpass; keyring.set_password('{LLAVERO}', '{nombre}', "
                f"getpass.getpass('{nombre}: '))")
    return (f"En Anaconda Prompt: setx {nombre} \"valor\" (en PowerShell: setx {nombre} 'valor') "
            f"y luego cierra y abre Spyder / VS Code.")


def _credencial(nombre, obligatoria=True):
    """Variable de entorno (o sus alias) y, si no esta, el llavero del sistema (keyring)."""
    candidatos = _candidatos(nombre)
    if obligatoria and not candidatos:
        sys.exit(f"No encuentro {nombre}. " + _como_guardar(nombre))
    return candidatos[0][1] if candidatos else None


# =============================================================================
# 2. DATOS
# =============================================================================

_SIETE = None


def _nueva_siete(token=None, usuario=None, clave=None):
    import bcchapi
    if token is None:
        return bcchapi.Siete(usuario, clave)        # usuario y clave de la cuenta del BDE
    # La documentacion oficial del BCCh usa Siete(token=...).
    try:
        return bcchapi.Siete(token=token)
    except TypeError:
        return bcchapi.Siete(api_key=token)


def _es_rechazo(e):
    texto = f"{type(e).__name__} {e}".lower()
    return "invalidcredentials" in texto or "invalid username or password" in texto


def _probar_siete(siete):
    """None si la conexion funciona; la excepcion si el BCCh rechaza la credencial."""
    try:
        siete.cuadro(series=["F073.TCO.PRE.Z.D"],
                     desde=(dt.date.today() - dt.timedelta(days=10)).isoformat())
        return None
    except Exception as e:  # noqa: BLE001
        if _es_rechazo(e):
            return e
        raise


def _pedir_usuario_clave():
    """Pide usuario y clave del BDE en pantalla (la clave no se muestra). None si no hay consola."""
    import getpass
    try:
        print("Ingresa tu usuario y clave de la API del BDE (si3.bcentral.cl):")
        usuario = _limpiar(input("  Usuario (correo): "))
        clave = getpass.getpass("  Clave: ")
    except (EOFError, OSError, KeyboardInterrupt):
        return None
    return (usuario, clave) if usuario and clave else None


def _ofrecer_guardar(usuario, clave):
    try:
        import keyring
        r = input("¿Guardar usuario y clave en el llavero para la próxima vez? [s/N]: ").strip().lower()
        if r in ("s", "si", "sí", "y", "yes"):
            keyring.set_password(LLAVERO, "BDE_USER", usuario)
            keyring.set_password(LLAVERO, "BDE_PASS", clave)
            print("Guardados en el llavero.")
    except Exception:  # noqa: BLE001 - sin keyring o sin consola: no se guarda
        pass


def _siete():
    """
    Conexion unica a la API BDE (se reutiliza entre llamadas).
      1) API Key Token: BDE_TOKEN (o sus alias) en variables de entorno y en el llavero.
      2) Si no hay token o el BCCh lo rechaza: usuario y clave de la cuenta (BDE_USER y BDE_PASS,
         juntos), desde variables de entorno o el llavero; si no estan, se piden en pantalla.
    El token nunca se usa como clave ni al reves.
    """
    global _SIETE
    if _SIETE is not None:
        return _SIETE
    rechazados = []

    # 1) API Key Token
    for origen, token in _candidatos("BDE_TOKEN"):
        siete = _nueva_siete(token=token)
        if _probar_siete(siete) is None:
            for o in rechazados:
                print(f"Aviso: el token de {o} no es válido; uso el de {origen}.")
            print(f"BDE: conectado con API token ({origen}).")
            _SIETE = siete
            return _SIETE
        rechazados.append(f"token ({origen}, {len(token)} caracteres)")

    # 2) usuario y clave, siempre juntos
    usuarios, claves = _candidatos("BDE_USER"), _candidatos("BDE_PASS")
    pares = []
    for o_u, u in usuarios:
        for o_c, c in claves:
            # misma procedencia primero: variable con variable, llavero con llavero
            if (o_u == "Llavero") == (o_c == "Llavero"):
                pares.append((f"{o_u} / {o_c}", u, c))
    for origen, u, c in pares:
        siete = _nueva_siete(usuario=u, clave=c)
        if _probar_siete(siete) is None:
            print(f"BDE: conectado con usuario y clave ({origen}).")
            _SIETE = siete
            return _SIETE
        rechazados.append(f"usuario y clave ({origen})")

    if rechazados:
        print("El Banco Central rechazó: " + "; ".join(rechazados) + ".")
    else:
        print("No encuentro API token ni usuario y clave del BDE.")
    par = _pedir_usuario_clave()
    if par:
        siete = _nueva_siete(usuario=par[0], clave=par[1])
        if _probar_siete(siete) is None:
            print("BDE: conectado con usuario y clave (ingresados).")
            _ofrecer_guardar(*par)
            _SIETE = siete
            return _SIETE
        sys.exit("El Banco Central rechazó el usuario y la clave ingresados (Invalid username or password).")
    sys.exit("No pude conectarme al BDE. Guarda el API token o el usuario y la clave.\n"
             "API token: " + _como_guardar("BDE_TOKEN") + "\n"
             "Usuario y clave (juntos): " + _como_guardar("BDE_USER") + " Y lo mismo para BDE_PASS.")


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


def _motor_excel(fuente):
    """Lector de Excel segun el contenido (no la extension): xlrd para .xls antiguos, openpyxl
    para .xlsx. Explicito para no depender del lector por defecto de cada version de pandas."""
    if isinstance(fuente, (bytes, bytearray)):
        firma = bytes(fuente[:4])
    elif hasattr(fuente, "getvalue"):
        firma = fuente.getvalue()[:4]
    else:
        with open(fuente, "rb") as f:
            firma = f.read(4)
    return "xlrd" if firma == b"\xd0\xcf\x11\xe0" else "openpyxl"


def nyfed_gscpi(desde=None, hasta=None, nombre="GSCPI"):
    """
    Global Supply Chain Pressure Index (mensual) desde la Fed de Nueva York (sin key).
    Unidad: desviaciones estandar respecto del promedio historico.
    """
    url = ("https://www.newyorkfed.org/medialibrary/research/interactives/gscpi/"
           "downloads/gscpi_data.xlsx")
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    crudo = pd.read_excel(io.BytesIO(r.content), sheet_name="GSCPI Monthly Data", header=None,
                          engine=_motor_excel(r.content))
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
          (70, 196, 202),   # turquesa
          (255, 174, 66),   # amarillo
          (178, 178, 178)]  # gris
# barras (plantilla Santander): gris oscuro, rojo y arena; luego turquesa y amarillo
PALETA_BARRAS = [(68, 68, 68),     # gris oscuro
                 (236, 0, 0),      # rojo Santander
                 (210, 173, 141),  # arena / beige
                 (70, 196, 202),   # turquesa
                 (255, 174, 66)]   # amarillo
COLOR_GUIAS = (217, 217, 217)      # lineas guia horizontales de las barras
COLOR_EJES = (191, 191, 191)  # "Blanco, fondo 1, oscuro 25%"
GRIS_TEXTO = (89, 89, 89)
FUENTES_SANTANDER = {
    "titulo": "Santander Headline",
    "ejes": "Santander Micro Text",
    "fuente": "Santander Text Light",
}
FUENTE_RESPALDO = "Arial"   # siempre instalada en Windows, sans serif neutra
FORMATO_FECHA = "mmm.yy"    # ene.26; en Windows, donde el mes ya trae punto, se usa "mmmyy"
# fija el idioma de los formatos en espanol (Chile) aunque el Office este en otro idioma
IDIOMA_FORMATO = "[$-340A]"
_MES_CON_PUNTO = None


def _mes_abreviado_con_punto():
    """
    True si Windows abrevia los meses en espanol de Chile con punto ("ene."). Office usa esa
    abreviatura para "mmm", y con FORMATO_FECHA "mmm.yy" las fechas saldrian "ene..26".
    En macOS (y LibreOffice) la abreviatura no lleva punto.
    """
    global _MES_CON_PUNTO
    if _MES_CON_PUNTO is None:
        _MES_CON_PUNTO = False
        if sys.platform.startswith("win"):
            try:
                import ctypes
                from ctypes import wintypes

                class SYSTEMTIME(ctypes.Structure):
                    _fields_ = [(c, wintypes.WORD) for c in ("wYear", "wMonth", "wDayOfWeek", "wDay",
                                                             "wHour", "wMinute", "wSecond",
                                                             "wMilliseconds")]
                fecha = SYSTEMTIME(2026, 1, 0, 15, 0, 0, 0, 0)
                buf = ctypes.create_unicode_buffer(32)
                n = ctypes.windll.kernel32.GetDateFormatEx("es-CL", 0, ctypes.byref(fecha), "MMM",
                                                           buf, len(buf), None)
                _MES_CON_PUNTO = bool(n) and buf.value.strip().endswith(".")
            except Exception:  # noqa: BLE001 - si no se puede consultar, se deja el formato tal cual
                _MES_CON_PUNTO = False
    return _MES_CON_PUNTO


def _formato_fecha_local(formato):
    """'mmm.yy' -> 'mmmyy' cuando el mes abreviado ya trae punto, para que salga 'ene.26'."""
    if "mmm." in formato and "mmmm" not in formato and _mes_abreviado_con_punto():
        return formato.replace("mmm.", "mmm")
    return formato
ANCHO_CM, ALTO_CM = 15, 11
GROSOR_MENSUAL = 2.25       # pt, series mensuales (o menos frecuentes)
GROSOR_DIARIO = 1.75        # pt, series diarias
TAM_UNIDAD = 11             # descriptor de la serie, ej. "(Var. anual, %)"
TAM_FUENTE = 9              # nota y fuente
TAM_ETIQUETA = 10           # etiqueta del ultimo dato
MESES_ABREV = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
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


def _decimales_sin_ceros(valor, decimales):
    """Decimales que necesita 'valor' redondeado a 'decimales', sin ceros sobrantes (2,50 -> 1)."""
    texto = f"{round(float(valor), decimales):.{decimales}f}"
    return len(texto.split(".")[1].rstrip("0")) if "." in texto else 0


def _numero_texto(valor, decimales):
    """Numero con coma decimal, punto de miles y sin ceros sobrantes: 1.234,5 / 3 / 0,25."""
    d = _decimales_sin_ceros(valor, decimales)
    return f"{round(float(valor), d):,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def es_diaria(serie):
    """True si la serie tiene datos con frecuencia diaria (separacion mediana < 20 dias)."""
    fechas = serie.dropna().index
    if len(fechas) < 3:
        return False
    return pd.Series(fechas).diff().dt.days.median() < 20


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


# Plantilla de presentacion (minuta mensual de Estudios): portada roja y laminas con franja roja
# superior, titulo rojo, linea gris inferior, pie "documento | autor" y numero de pagina en rojo.
GRIS_PIE = (111, 119, 121)      # #6F7779
GRIS_LINEA = (200, 205, 208)    # #C8CDD0
BLANCO = (255, 255, 255)
MARGEN_PT = 43.2                # 4,5% del ancho de la lamina
Y_CONTENIDO_CM = 4.0            # borde superior sugerido para graficos (bajo el titulo)


def _texto_plantilla(slide, texto, x_pt, y_pt, w_pt, h_pt, nombre_fuente, tam, color, negrita=False,
                     alineacion=None, ancla="top"):
    """Caja de texto sin margenes internos, para ubicar exactamente los elementos de la plantilla."""
    from pptx.enum.text import MSO_ANCHOR
    from pptx.util import Pt
    caja = _caja_texto(slide, texto, Pt(x_pt), Pt(y_pt), Pt(w_pt), Pt(h_pt), nombre_fuente, tam,
                       negrita=negrita, color=color, alineacion=alineacion)
    tf = caja.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE}[ancla]
    return caja


def _rectangulo(slide, x_pt, y_pt, w_pt, h_pt, color):
    from pptx.util import Pt
    forma = _linea_horizontal(slide, Pt(x_pt), Pt(y_pt), Pt(w_pt), color=color, grosor_pt=h_pt)
    return forma


def portada(prs, titulo, subtitulo="", autor="", fecha=None):
    """
    Portada roja: titulo en blanco (puede tener saltos de linea), barra blanca, subtitulo
    (ej. "Noviembre 2026") y, abajo, autor y fecha (por defecto, la de hoy).
    """
    s = prs.slides.add_slide(prs.slide_layouts[6])
    ancho_pt, alto_pt = prs.slide_width.pt, prs.slide_height.pt
    _rectangulo(s, 0, 0, ancho_pt, alto_pt, PALETA[0])
    x = 0.08 * ancho_pt
    lineas = titulo.split("\n")
    for k, linea in enumerate(lineas):           # 38 pt, lineas cada 48,6 pt
        _texto_plantilla(s, linea, x, 175 + 48.6 * k - 48.6 * (len(lineas) - 2), 0.84 * ancho_pt, 46,
                         tipografia("titulo"), 38, BLANCO)
    _rectangulo(s, x, 283, 0.12 * ancho_pt, 3.24, BLANCO)
    if subtitulo:
        _texto_plantilla(s, subtitulo, x, 305, 0.84 * ancho_pt, 28, tipografia("titulo"), 22, BLANCO)
    if fecha is None:
        hoy = dt.date.today()
        fecha = f"{hoy.day:02d} de {MESES_LARGOS[hoy.month - 1]} de {hoy.year}"
    if autor:
        _texto_plantilla(s, autor, x, 447, 0.6 * ancho_pt, 18, tipografia("titulo"), 14, BLANCO, negrita=True)
    if fecha:
        _texto_plantilla(s, fecha, x, 470, 0.6 * ancho_pt, 16, tipografia("titulo"), 12, BLANCO)
    return s


def lamina(prs, titulo_lamina, numero=None, pie_texto=""):
    """
    Lamina de contenido: franja roja superior, titulo rojo (20), linea gris inferior, pie a la
    izquierda (ej. "Chartpack de Expectativas noviembre 2026 | Estudios Economicos") y numero.
    """
    from pptx.enum.text import PP_ALIGN
    s = prs.slides.add_slide(prs.slide_layouts[6])
    ancho_pt, alto_pt = prs.slide_width.pt, prs.slide_height.pt
    util = ancho_pt - 2 * MARGEN_PT
    _rectangulo(s, 0, 0, ancho_pt, 0.035 * alto_pt, PALETA[0])
    _texto_plantilla(s, titulo_lamina, MARGEN_PT, 0.095 * alto_pt - 14, util, 28,
                     tipografia("titulo"), 20, PALETA[0], ancla="middle")
    _rectangulo(s, MARGEN_PT, 0.955 * alto_pt, util, 0.65, GRIS_LINEA)
    if pie_texto:
        _texto_plantilla(s, pie_texto, MARGEN_PT, 0.975 * alto_pt - 6, 0.7 * util, 12,
                         tipografia("fuente"), 8, GRIS_PIE, ancla="middle")
    if numero is not None:
        _texto_plantilla(s, str(numero), ancho_pt - MARGEN_PT - 60, 0.975 * alto_pt - 6, 60, 12,
                         tipografia("titulo"), 8.5, PALETA[0], negrita=True,
                         alineacion=PP_ALIGN.RIGHT, ancla="middle")
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
    return _caja_texto(slide, texto, x, y, w, h or Cm(1.2), tipografia("fuente"), TAM_FUENTE)


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
    # sin ceros sobrantes: los decimales del paso (1 -> "2"; 0,5 -> "2,5"); con escala
    # automatica, formato General (Office no agrega ceros)
    if paso is not None:
        formato = _formato_numero(_decimales_sin_ceros(paso, 4))
    else:
        formato = "General"
    eje.tick_labels.number_format = IDIOMA_FORMATO + formato
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


def _a_barras(chart, columnas):
    """Mueve las series indicadas a un grafico de columnas que usa los mismos ejes."""
    from pptx.oxml.ns import qn
    plot_area = chart._chartSpace.find(qn("c:chart")).find(qn("c:plotArea"))
    linea = plot_area.find(qn("c:lineChart"))
    barra = linea.makeelement(qn("c:barChart"), {})
    _sub(barra, "c:barDir", val="col")
    _sub(barra, "c:grouping", val="clustered")
    _sub(barra, "c:varyColors", val="0")
    for ser in list(linea.findall(qn("c:ser"))):
        if ser.find(qn("c:tx")).find(".//" + qn("c:v")).text in columnas:
            linea.remove(ser)
            # una serie de columnas no lleva marcador ni suavizado
            for tag in ("c:marker", "c:smooth"):
                el = ser.find(qn(tag))
                if el is not None:
                    ser.remove(el)
            previo = ser.find(qn("c:spPr"))
            if previo is None:
                previo = ser.find(qn("c:tx"))
            previo.addnext(ser.makeelement(qn("c:invertIfNegative"), {"val": "0"}))
            barra.append(ser)
    _sub(barra, "c:gapWidth", val="40")
    for eje in linea.findall(qn("c:axId")):
        _sub(barra, "c:axId", val=eje.get("val"))
    linea.addprevious(barra)  # las columnas se dibujan detras de las lineas


def _leyenda_libre(chart, df, w, h, tam, y_min, y_max, eje_derecho, y2_min, y2_max, barras,
                   etiqueta_ultimo, unidad, unidad_derecha, posicion_unidad):
    """
    Ubica la leyenda dentro del grafico, en el rincon donde no tapa lineas, columnas, etiquetas del
    ultimo dato ni el descriptor de unidad. Si no hay un rincon libre, la deja arriba, fuera del
    area de las series (como antes). Estima la posicion de las series a partir de los datos; la
    leyenda queda editable y se puede mover a mano en PowerPoint.
    """
    import numpy as np
    from pptx.oxml.ns import qn
    w_pt, h_pt = w / 12700, h / 12700
    der = [str(c) for c in (eje_derecho or [])]
    # area de las series dentro del marco del grafico (fracciones), aproximada como la deja Office
    izq = (tam * 2.6 + 6) / w_pt
    dcha = (tam * 3.4 if der else tam * 0.8) / w_pt + (0.06 if etiqueta_ultimo else 0)
    arriba, abajo = 0.03, (tam * 2.2) / h_pt
    pw, ph = 1 - izq - dcha, 1 - arriba - abajo

    def rango(cols, lo, hi):
        v = df[cols].to_numpy(dtype=float)
        v = v[np.isfinite(v)]
        if v.size == 0:
            return 0.0, 1.0
        lo = v.min() if lo is None else lo
        hi = v.max() if hi is None else hi
        if barras and any(c in barras for c in cols):
            lo, hi = min(lo, 0), max(hi, 0)
        margen = 0.05 * (hi - lo or 1)
        return (lo - margen if y_min is None else lo), (hi + margen if y_max is None else hi)

    izq_cols = [str(c) for c in df.columns if str(c) not in der]
    r_izq = rango(izq_cols, y_min, y_max) if izq_cols else (0, 1)
    r_der = rango(der, y2_min, y2_max) if der else (0, 1)
    # posicion horizontal como fraccion del periodo (independiente de la unidad de las fechas)
    inicio, total = df.index.min(), (df.index.max() - df.index.min()) or pd.Timedelta(days=1)
    malla = np.linspace(0, 1, 400)
    puntos = []                                     # (x, y) en fracciones del marco; y desde arriba
    for col in df.columns:
        s = df[col].dropna()
        if s.empty:
            continue
        lo, hi = r_der if str(col) in der else r_izq
        xs = np.asarray((s.index - inicio) / total, dtype=float)
        ys = (s.to_numpy(dtype=float) - lo) / ((hi - lo) or 1)
        if str(col) in barras:                      # columnas: toda la altura desde 0
            base = (0 - lo) / ((hi - lo) or 1)
            for xx, yy in zip(xs, ys):
                for f in np.linspace(min(base, yy), max(base, yy), 6):
                    puntos.append((izq + xx * pw, arriba + (1 - f) * ph))
        else:                                       # lineas: muestreo fino entre observaciones
            yi = np.interp(malla, xs, ys, left=np.nan, right=np.nan)
            puntos += [(izq + xx * pw, arriba + (1 - yy) * ph) for xx, yy in zip(malla, yi) if np.isfinite(yy)]
        if str(col) in (etiqueta_ultimo or []):     # etiqueta del ultimo dato, a la derecha
            px, py = izq + xs[-1] * pw, arriba + (1 - ys[-1]) * ph
            puntos += [(px + dx, py + dy) for dx in np.linspace(0, 0.09, 6) for dy in (-0.025, 0, 0.025)]
    puntos = np.array(puntos) if puntos else np.zeros((0, 2))

    # descriptores de unidad (cajas de 8 x 0,8 cm a 2 cm del borde, arriba o abajo)
    obstaculos = []
    y_u = (1.3 / (h_pt / 28.35)) if posicion_unidad != "abajo" else 1 - 2.2 / (h_pt / 28.35)
    alto_u = 0.8 / (h_pt / 28.35)
    if unidad:
        obstaculos.append((2.0 / (w_pt / 28.35), y_u, 2.0 / (w_pt / 28.35) + len(unidad) * 0.55 * 11 / w_pt,
                           y_u + alto_u))
    if unidad_derecha:
        fin = 1 - 1.5 / (w_pt / 28.35)
        obstaculos.append((fin - len(unidad_derecha) * 0.55 * 11 / w_pt, y_u, fin, y_u + alto_u))

    # tamano de la leyenda: entradas en filas que caben en 60% del ancho de las series
    anchos = [(len(str(c)) * 0.6 * tam + 3.4 * tam) / w_pt for c in df.columns]
    filas, fila = [], []
    for a in anchos:
        if fila and sum(fila) + a > 0.6 * pw:
            filas.append(fila)
            fila = []
        fila.append(a)
    filas.append(fila)
    lw, lh = max(sum(f) for f in filas) + 0.02, (len(filas) * 1.7 * tam + 6) / h_pt

    pad = 0.015
    candidatos = []
    for vert, y0 in (("arriba", arriba + pad), ("abajo", arriba + ph - lh - pad)):
        for hor, x0 in (("izquierda", izq + pad), ("centro", izq + (pw - lw) / 2),
                        ("derecha", izq + pw - lw - pad)):
            candidatos.append((x0, y0))

    def choca(x0, y0):
        x1, y1 = x0 + lw, y0 + lh
        n = int(((puntos[:, 0] > x0 - pad) & (puntos[:, 0] < x1 + pad) &
                 (puntos[:, 1] > y0 - pad) & (puntos[:, 1] < y1 + pad)).sum()) if len(puntos) else 0
        n += sum(1000 for (a0, b0, a1, b1) in obstaculos
                 if a0 - pad < x1 and x0 < a1 + pad and b0 - pad < y1 and y0 < b1 + 2 * pad)
        return n

    libres = [(x0, y0) for x0, y0 in candidatos if 0 <= x0 and x0 + lw <= 1 and choca(x0, y0) == 0]
    if not libres:
        return False                                # queda arriba, fuera del area de las series
    x0, y0 = libres[0]
    chart.legend.include_in_layout = True           # superpuesta: no achica el area de las series
    leg = chart.legend._element
    capa = leg.find(qn("c:layout"))
    if capa is None:
        capa = _sub(leg, "c:layout")
        leg.find(qn("c:overlay")).addprevious(capa)
    for hijo in list(capa):
        capa.remove(hijo)
    manual = _sub(capa, "c:manualLayout")
    for tag, val in (("c:xMode", "edge"), ("c:yMode", "edge"), ("c:x", x0), ("c:y", y0),
                     ("c:w", lw), ("c:h", lh)):
        _sub(manual, tag, val=round(val, 4) if isinstance(val, float) else val)
    return True


def agregar_grafico(slide, df, x, y, w, h, titulo=None, unidad=None, decimales=1,
                    proyecciones=None, etiqueta_ultimo=True,
                    y_min=None, y_max=None, y_paso=None,
                    eje_derecho=None, unidad_derecha=None, y2_min=None, y2_max=None, y2_paso=None,
                    meses_entre_etiquetas=None, formato_fecha=FORMATO_FECHA,
                    linea_ref=None, posicion_unidad="abajo", tam=12, barras=None, leyenda="auto"):
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
    etiqueta_ultimo : True (todas las series), False o lista de columnas con etiqueta
    barras       : lista de columnas que se dibujan como columnas (mismo eje Y que las lineas)
    leyenda      : "auto" (por defecto: dentro del grafico, en el rincon donde no tapa series ni
                   rotulos; si no hay espacio libre, arriba fuera del area de las series),
                   "arriba" o "abajo"
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
    barras = [str(c) for c in (barras or [])]
    if etiqueta_ultimo is True:
        etiqueta_ultimo = list(df.columns)
    etiqueta_ultimo = [str(c) for c in (etiqueta_ultimo or [])]
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
        if leyenda == "abajo":
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        elif leyenda == "auto":
            _leyenda_libre(chart, df, w, h, tam, y_min, y_max, eje_derecho, y2_min, y2_max, barras,
                           etiqueta_ultimo, unidad, unidad_derecha, posicion_unidad)

    # --- eje X: fechas siempre abajo, formato ene.26 ---
    ex = chart.category_axis
    ex.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    ex.major_tick_mark = XL_TICK_MARK.OUTSIDE
    ex.minor_tick_mark = XL_TICK_MARK.NONE
    ex.tick_labels.number_format = IDIOMA_FORMATO + _formato_fecha_local(formato_fecha)
    ex.tick_labels.number_format_is_linked = False
    ex.tick_labels.font.size = Pt(tam)
    ex.tick_labels.font.name = f_ejes
    ex.tick_labels.font.color.rgb = _rgb(0, 0, 0)
    # fechas horizontales
    cuerpo = ex._element.find(qn("c:txPr")).find(qn("a:bodyPr"))
    cuerpo.set("rot", "0")
    cuerpo.set("vert", "horz")
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

    # --- columnas (comparten ejes con las lineas y quedan detras de ellas) ---
    if barras:
        _a_barras(chart, barras)

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
        diaria = es_diaria(df[col])
        if str(col) in barras:
            serie.format.fill.solid()
            serie.format.fill.fore_color.rgb = color
            serie.format.line.fill.background()
        else:
            serie.smooth = False
            serie.format.line.width = Pt(GROSOR_DIARIO if diaria else GROSOR_MENSUAL)
            serie.format.line.color.rgb = color
            serie.marker.style = XL_MARKER_STYLE.NONE
            # el marcador (invisible) lleva el color de la linea: algunos visores (LibreOffice)
            # dibujan el simbolo de la leyenda con el color de fuente si no
            serie.marker.format.line.color.rgb = color
            if col in proyecciones:
                serie.format.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        idx = df[col].last_valid_index()
        if str(col) in etiqueta_ultimo and idx is not None:
            valor = df[col].loc[idx]
            etq = serie.points[df.index.get_loc(idx)].data_label
            if diaria:
                # series diarias: valor y dia del ultimo dato, ej. "977 (22-oct)"
                etq.has_text_frame = True
                etq.text_frame.text = f"{_numero_texto(valor, decimales)} ({idx.day}-{MESES_ABREV[idx.month - 1]})"
                fuente_etq = etq.text_frame.paragraphs[0].runs[0].font
            else:
                etq.has_text_frame = False
                etq.show_value = True
                etq.number_format = IDIOMA_FORMATO + _formato_numero(_decimales_sin_ceros(valor, decimales))
                etq.number_format_is_linked = False
                fuente_etq = etq.font
            fuente_etq.size = Pt(TAM_ETIQUETA)
            fuente_etq.bold = True
            fuente_etq.name = f_ejes
            fuente_etq.color.rgb = color
            etq.position = XL_LABEL_POSITION.OUTSIDE_END if str(col) in barras else XL_LABEL_POSITION.RIGHT

    # --- descriptor de unidad dentro del grafico ---
    y_unidad = y + h - Cm(2.2) if posicion_unidad == "abajo" else y + Cm(1.3)
    if unidad:
        _caja_texto(slide, unidad, x + Cm(2.0), y_unidad, Cm(8), Cm(0.8), f_ejes, TAM_UNIDAD)
    if unidad_derecha:
        from pptx.enum.text import PP_ALIGN
        _caja_texto(slide, unidad_derecha, x + w - Cm(9.5), y_unidad, Cm(8), Cm(0.8), f_ejes, TAM_UNIDAD,
                    alineacion=PP_ALIGN.RIGHT)
    return chart


def _texto_categoria(c, formato):
    """Etiqueta de categoria: texto tal cual; fechas como año (enero) o 'ene.26'."""
    if not isinstance(c, (pd.Timestamp, dt.date)):
        return str(c)
    c = pd.Timestamp(c)
    if formato:
        return c.strftime(formato)
    return str(c.year) if c.month == 1 and c.day == 1 else f"{MESES_ABREV[c.month - 1]}.{c.year % 100:02d}"


def agregar_barras(slide, df, x, y, w, h, titulo=None, unidad=None, decimales=1,
                   apiladas=False, horizontal=False, etiquetas=True, leyenda="abajo",
                   formato_categoria=None, colores=None, tam=12, posicion_unidad="arriba"):
    """
    Grafico de barras nativo (editable) con el formato de la plantilla Santander.
    df           : una fila por categoria (indice: texto, años o fechas) y una columna por serie (max. 5)
    apiladas     : barras apiladas en vez de agrupadas
    horizontal   : barras horizontales en vez de columnas
    etiquetas    : valor sobre cada barra, en el color de su serie (sin ceros sobrantes); con
                   etiquetas no se muestra el eje de valores, solo lineas guia suaves
    leyenda      : "abajo" (por defecto, como la plantilla), "arriba" o None
    formato_categoria : formato strftime para fechas, ej. "%Y"; por defecto año o 'ene.26'
    """
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_MARK
    from pptx.enum.chart import XL_TICK_LABEL_POSITION
    from pptx.util import Cm, Pt

    if df.shape[1] > 5:
        raise ValueError("Máximo 5 series por gráfico (legibilidad y paleta).")
    colores = colores or PALETA_BARRAS
    f_tit, f_ejes = tipografia("titulo"), tipografia("ejes")
    if titulo:
        _caja_texto(slide, titulo, x, y - Cm(1.1), w, Cm(1), f_tit, 14, negrita=True)

    datos = CategoryChartData(number_format="General")
    datos.categories = [_texto_categoria(c, formato_categoria) for c in df.index]
    for col in df.columns:
        # valores redondeados a 'decimales': con formato General no quedan ceros sobrantes
        datos.add_series(str(col), [None if pd.isna(v) else round(float(v), decimales) for v in df[col]])
    tipo = {(False, False): XL_CHART_TYPE.COLUMN_CLUSTERED, (True, False): XL_CHART_TYPE.COLUMN_STACKED,
            (False, True): XL_CHART_TYPE.BAR_CLUSTERED, (True, True): XL_CHART_TYPE.BAR_STACKED}
    chart = slide.shapes.add_chart(tipo[(apiladas, horizontal)], x, y, w, h, datos).chart
    chart.has_title = False
    chart.font.name = f_ejes
    chart.font.size = Pt(tam)

    plot = chart.plots[0]
    plot.gap_width = 60
    plot.overlap = 100 if apiladas else 0
    if horizontal:   # primera categoria arriba, como se lee
        chart.category_axis.reverse_order = True

    # eje de categorias: linea gris, sin marcas
    ex = chart.category_axis
    ex.major_tick_mark = XL_TICK_MARK.NONE
    ex.minor_tick_mark = XL_TICK_MARK.NONE
    ex.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    ex.format.line.width = Pt(0.75)
    ex.format.line.color.rgb = _rgb(*COLOR_EJES)
    ex.tick_labels.font.size = Pt(tam)
    ex.tick_labels.font.name = f_ejes
    ex.tick_labels.font.color.rgb = _rgb(0, 0, 0)
    ex.has_major_gridlines = False

    # eje de valores: con etiquetas, oculto y con lineas guia suaves; sin etiquetas, eje normal
    ey = chart.value_axis
    if etiquetas:
        ey.visible = False
        ey.has_major_gridlines = True
        ey.major_gridlines.format.line.color.rgb = _rgb(*COLOR_GUIAS)
        ey.major_gridlines.format.line.width = Pt(0.5)
    else:
        _formato_eje_valor(ey, decimales, None, tam, f_ejes)

    for i, serie in enumerate(plot.series):
        color = _rgb(*colores[i])
        serie.format.fill.solid()
        serie.format.fill.fore_color.rgb = color
        serie.format.line.fill.background()
        serie.invert_if_negative = False
        if etiquetas:
            dl = serie.data_labels
            dl.show_value = True
            dl.number_format = IDIOMA_FORMATO + "General"
            dl.number_format_is_linked = False
            dl.font.size = Pt(TAM_ETIQUETA + 1)
            dl.font.name = f_ejes
            # apiladas: valor al centro en blanco (gris claro y arena: en negro); agrupadas: sobre la barra
            if apiladas:
                dl.position = XL_LABEL_POSITION.CENTER
                claro = sum(colores[i]) > 450
                dl.font.color.rgb = _rgb(0, 0, 0) if claro else _rgb(255, 255, 255)
            else:
                dl.position = XL_LABEL_POSITION.OUTSIDE_END
                dl.font.color.rgb = color

    chart.has_legend = df.shape[1] > 1 and leyenda is not None
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.TOP if leyenda == "arriba" else XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(tam - 1)
        chart.legend.font.name = f_ejes
        chart.legend.font.color.rgb = _rgb(0, 0, 0)

    if unidad:
        y_unidad = y + h - Cm(2.2) if posicion_unidad == "abajo" else y + Cm(0.1)
        _caja_texto(slide, unidad, x + Cm(0.4), y_unidad, Cm(8), Cm(0.8), f_ejes, TAM_UNIDAD)
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


def exportar_pdf(archivo_pptx):
    """
    Guarda un PDF junto al PPTX. En Windows usa PowerPoint (requiere pywin32 o comtypes); en
    macOS, PowerPoint via AppleScript; si no esta disponible, LibreOffice si esta instalado.
    Devuelve la ruta del PDF, o None si no se pudo generar.
    """
    import shutil
    import subprocess
    pptx = Path(archivo_pptx).resolve()
    pdf = pptx.with_suffix(".pdf")
    if sys.platform.startswith("win"):
        try:
            try:
                from win32com.client import Dispatch
            except ImportError:
                try:
                    from comtypes.client import CreateObject as Dispatch
                except ImportError:
                    raise ImportError("falta pywin32; instálalo una vez con: python -m pip install pywin32")
            app = Dispatch("PowerPoint.Application")
            pres = app.Presentations.Open(str(pptx), True, False, False)  # solo lectura, sin ventana
            pres.SaveAs(str(pdf), 32)  # 32 = ppSaveAsPDF
            pres.Close()
            if app.Presentations.Count == 0:
                app.Quit()
            return pdf
        except Exception as e:  # noqa: BLE001 - se intenta con LibreOffice
            print(f"Aviso: PowerPoint no pudo exportar el PDF ({e}).")
    if sys.platform == "darwin":
        # PowerPoint para Mac via AppleScript (puede pedir permiso para la carpeta la primera vez)
        script = (f'tell application "Microsoft PowerPoint"\n'
                  f'  open POSIX file "{pptx}"\n'
                  f'  set p to active presentation\n'
                  f'  save p in (POSIX file "{pdf}") as save as PDF\n'
                  f'  close p saving no\n'
                  f'end tell')
        r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
        if r.returncode == 0 and pdf.exists():
            return pdf
        print(f"Aviso: PowerPoint para Mac no pudo exportar el PDF ({r.stderr.strip()[:200]}).")
    soffice = (shutil.which("soffice") or shutil.which("libreoffice")
               or next((p for p in ["/Applications/LibreOffice.app/Contents/MacOS/soffice"]
                        if os.path.exists(p)), None))
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(pptx.parent),
                        str(pptx)], check=True, capture_output=True)
        return pdf
    print(f"Aviso: no pude generar el PDF; ábrelo en PowerPoint y usa Guardar como PDF: {pptx}")
    return None


def vista_previa(df, titulo="", decimales=1):
    """Vista rapida con matplotlib (solo para revisar, no para presentar)."""
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(figsize=(15 / 2.54, 11 / 2.54))
    for i, col in enumerate(df.columns):
        ax.plot(df.index, df[col], color=[c / 255 for c in PALETA[i]], label=col,
                lw=GROSOR_DIARIO if es_diaria(df[col]) else GROSOR_MENSUAL)
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
                   archivo="prueba_formato_graficos.pptx")
