# -*- coding: utf-8 -*-
r"""Chartpack de Expectativas - Gerencia de Estudios (Santander Chile).

Define el contenido (series, transformaciones, laminas). Los datos del BDE, el formato y los
graficos nativos de PowerPoint vienen de formato_graficos.py (00_Administrativos\Manual de Estilo).
Las laminas de dispersion EEE y umbrales EOF leen los Excel que genera encuestas_bcch.py.

Carpetas (relativas a este script, sirven para cualquier usuario):
    <Chartpack de Expectativas>\Code\     build_chartpack.py y encuestas_bcch.py
    <Chartpack de Expectativas>\Inputs\   eee\ y eof\ (planillas raw y Excel resultante)
    <Chartpack de Expectativas>\Outputs\  AAAAMMDD_Chartpack_Expectativas.pptx / .pdf

Base de datos Chartpack_Expectativas.xlsx (una hoja por grafico, siempre con ese nombre):
    - Windows con OneDrive/SharePoint: ...\General\10_Bases y Analisis\Local\Expectativas y Encuestas
    - En otro caso (por ejemplo macOS): Outputs\
    Si el archivo ya existe se abre y solo se reescriben las columnas de cada hoja del chartpack (desde
    A1); las demas hojas, columnas, formulas y graficos se conservan. Si lo guardo Excel y tiene
    graficos, se actualiza con Excel (xlwings) cuando esta disponible; si no, con openpyxl dejando un
    respaldo en Inputs\respaldo.

Uso (no hay que editar nada):
    python encuestas_bcch.py     # actualiza Inputs (EEE y EOF)
    python build_chartpack.py    # genera el chartpack en Outputs
"""
import datetime as dt
import os
import re
import shutil
import sys
import unicodedata
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd
from pptx.util import Cm

CODE = Path(__file__).resolve().parent
BASE = CODE.parent
INPUTS = BASE / "Inputs"
OUT = BASE / "Outputs"
EEE_DECILES = INPUTS / "eee" / "eee_inflacion_deciles.xlsx"
EOF_UMBRALES = INPUTS / "eof" / "eof_inflacion_2a_umbrales.xlsx"
MANUAL_ESTILO = Path("00_Administrativos") / "Manual de Estilo"
MANUAL_ESTILO_ANTERIOR = Path("Administrativos") / "Manual de Estilo"
BIBLIOTECA = Path("Estudios - Documentos") / "General"
BASE_EXCEL = "Chartpack_Expectativas.xlsx"
# carpeta de bases en SharePoint (Windows): ...\General\10_Bases y Analisis\Local\Expectativas y Encuestas
BASES = r"10_bases y an[aá]lisis"
BASES_SUB = Path("Local") / "Expectativas y Encuestas"


ICLOUD = Path.home() / "Library" / "Mobile Documents" / "com~apple~CloudDocs"   # iCloud Drive (macOS)


def _raices_sincronizadas():
    """Carpetas de OneDrive/SharePoint/iCloud con 'Santander' (u OneDrive) en el nombre, de cualquier
    usuario: C:\\Users\\<usuario>\\... (Windows), ~/Library/CloudStorage/... e iCloud Drive (macOS)."""
    casa = Path.home()
    raices = []
    for base in (casa, casa / "Library" / "CloudStorage", ICLOUD):
        if base.is_dir():
            for raiz in sorted(base.iterdir()):
                if raiz.is_dir() and re.search(r"santander|onedrive|sharepoint", raiz.name, re.I):
                    raices.append(raiz)
    return raices


def _formato_en(raiz, profundidad=6):
    """Busca formato_graficos.py bajo 'raiz': primero en las rutas habituales y, en iCloud, en toda la
    carpeta (hasta 'profundidad' niveles). Avisa si iCloud lo tiene solo en la nube."""
    rutas = []
    for me in (MANUAL_ESTILO, MANUAL_ESTILO_ANTERIOR):
        rutas += [raiz / BIBLIOTECA / me, raiz / "General" / me, raiz / me]
        for sub in sorted(p for p in raiz.iterdir() if p.is_dir() and not p.name.startswith(".")):
            rutas += [sub / BIBLIOTECA / me, sub / "General" / me, sub / me]
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
        if ".formato_graficos.py.icloud" in archivos:   # iCloud lo dejo solo en la nube
            sys.exit(f"formato_graficos.py está en iCloud pero no descargado en este Mac:\n    {carpeta}\n"
                     "Ábrelo en Finder y elige 'Descargar ahora' (o 'Mantener descargado'), y vuelve a correr.")
    return None


def _carpeta_formato() -> Path:
    """Carpeta de formato_graficos.py: 00_Administrativos/Manual de Estilo de la biblioteca de Estudios."""
    candidatos = []
    if os.environ.get("FORMATO_GRAFICOS_DIR"):
        candidatos.append(Path(os.environ["FORMATO_GRAFICOS_DIR"]))
    # 1) subiendo desde la carpeta de este script (...\General\...\Code)
    for p in [CODE, *CODE.parents]:
        candidatos += [p / MANUAL_ESTILO, p / MANUAL_ESTILO_ANTERIOR, p / "formato"]   # "formato": repositorio
    for carpeta in candidatos:
        if (carpeta / "formato_graficos.py").is_file():
            return carpeta
    # 2) carpetas sincronizadas: OneDrive/SharePoint (Windows y macOS) e iCloud Drive (macOS)
    for raiz in _raices_sincronizadas():
        carpeta = _formato_en(raiz)
        if carpeta:
            return carpeta
    sys.exit("No encuentro formato_graficos.py. Debe estar en\n"
             f"    ...\\{BIBLIOTECA}\\{MANUAL_ESTILO}\n"
             "(también se busca en OneDrive/SharePoint y en la carpeta Santander de iCloud Drive).\n"
             "Para indicar otra carpeta: variable FORMATO_GRAFICOS_DIR.")


sys.path.insert(0, str(_carpeta_formato()))
import formato_graficos as fg  # noqa: E402

BCCH = "Banco Central de Chile"
IPEC_FUENTE = "GfK Adimark, publicado por el Banco Central de Chile"
TITULO = "Chartpack de Expectativas"
AUTOR = "Estudios Económicos"   # aparece en la portada y en el pie de cada lámina
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]

# posiciones de graficos (cm): dos por lamina, tres por lamina y uno ancho
Y0, ALTO = Cm(fg.Y_CONTENIDO_CM), Cm(fg.ALTO_CM)
DOS = [Cm(1.3), Cm(17.5)]
ANCHO_DOS = Cm(fg.ANCHO_CM)
TRES = [Cm(1.0), Cm(11.8), Cm(22.6)]
ANCHO_TRES = Cm(10.2)
TAM_TRES = 10   # graficos de 10 cm de ancho: fuente 10 para que leyenda y ejes no se superpongan
ANCHO_UNO = Cm(31.2)

DATOS = {}   # hoja de Chartpack_Expectativas.xlsx -> DataFrame (nombres fijos: el Excel puede tener formulas)
_NUM = [1]   # numero de la lamina actual (1 = portada)
_PIE = [""]  # pie de las laminas: "Chartpack de Expectativas <mes año> | <autor>"


def lamina(prs, titulo):
    """Lamina nueva con numeracion automatica."""
    _NUM[0] += 1
    return fg.lamina(prs, titulo, _NUM[0], pie_texto=_PIE[0])


def graficar(lam, df, x, w, hoja, fuente=BCCH, nota=None, **kw):
    """Grafico + pie (nota y fuente) bajo el grafico; guarda los datos para el Excel."""
    fg.agregar_grafico(lam, df, x, Y0, w, ALTO, **kw)
    fg.pie(lam, x, Y0 + ALTO + Cm(0.15), w, fuente, nota, h=Cm(1.6))
    DATOS[hoja[:31]] = df


def _input(archivo):
    if not archivo.is_file():
        sys.exit(f"No encuentro {archivo}. Corre primero: python encuestas_bcch.py")
    return archivo


def eee_deciles(horizonte):
    """Mediana, deciles 1 y 9 y rango 9-1 de la EEE (Excel de encuestas_bcch.py)."""
    d = pd.read_excel(_input(EEE_DECILES), sheet_name="deciles", engine="openpyxl")
    d = d[d["horizonte"] == horizonte]
    d.index = pd.to_datetime(d["fecha"])
    return d[["mediana", "decil_1", "decil_9", "rango_9_1"]].rename(columns={
        "mediana": "Mediana", "decil_1": "Decil 1", "decil_9": "Decil 9", "rango_9_1": "Decil 9 - 1"})


def eof_umbrales():
    """% de operadores EOF con inflacion a 2 anios sobre/bajo umbrales (Excel de encuestas_bcch.py)."""
    d = pd.read_excel(_input(EOF_UMBRALES), sheet_name="mensual", engine="openpyxl")
    d.index = pd.to_datetime(d["mes"])
    return d


# ─────────────────────────────────────────────────────────────────────────────
# Chartpack_Expectativas.xlsx: escribe solo en sus columnas (conserva el resto del libro)
# ─────────────────────────────────────────────────────────────────────────────

def _normal(texto):
    return unicodedata.normalize("NFC", texto).lower()


def _hija(padre, patron):
    """Subcarpeta de 'padre' cuyo nombre calza con 'patron' (sin distinguir mayusculas)."""
    try:
        return next((p for p in sorted(padre.iterdir()) if p.is_dir() and re.fullmatch(patron, _normal(p.name))),
                    None)
    except OSError:
        return None


def carpeta_base():
    """Windows con OneDrive/SharePoint: ...\\General\\10_Bases y Analisis\\Local\\Expectativas y Encuestas
    (si existe 10_Bases y Analisis); en otro caso (por ejemplo macOS), Outputs."""
    if sys.platform.startswith("win"):
        padres = list(CODE.parents)
        for raiz in _raices_sincronizadas():
            padres += [raiz / BIBLIOTECA, raiz / "General", raiz]
        for padre in padres:
            bases = _hija(padre, BASES) if padre.is_dir() else None
            if bases:
                return bases / BASES_SUB
        print("  ⚠️ No encontré la carpeta '10_Bases y Analisis' en OneDrive; la base queda en Outputs.")
    return OUT


def _valor_celda(v):
    if v is None or v is pd.NaT or (isinstance(v, float) and np.isnan(v)):
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if isinstance(v, (np.floating, np.integer)):
        return None if np.isnan(v) else float(v)
    return v


def _filas(df):
    """Encabezado + filas (indice como primera columna)."""
    d = df.reset_index()
    return [list(d.columns)] + [[_valor_celda(v) for v in fila] for fila in d.itertuples(index=False)]


def _formato_columna(serie):
    if pd.api.types.is_datetime64_any_dtype(serie):
        return "mmm-yyyy"
    return "0.00" if pd.api.types.is_numeric_dtype(serie) else None


class _LibroOpenpyxl:
    def __init__(self, ruta):
        from openpyxl import Workbook, load_workbook
        self.ruta = ruta
        if ruta.exists():
            self.wb = load_workbook(ruta)
        else:
            self.wb = Workbook()
            self.wb.remove(self.wb.active)

    def escribir(self, hoja, filas, formatos):
        from openpyxl.styles import Font
        ws = self.wb[hoja] if hoja in self.wb.sheetnames else self.wb.create_sheet(hoja)
        n_cols = len(filas[0])
        previo = ws.max_row
        for i, fila in enumerate(filas, start=1):
            for j, v in enumerate(fila, start=1):
                c = ws.cell(i, j)
                c.value = v
                if i == 1:
                    c.font = Font(bold=True)
                elif formatos[j - 1] and v is not None:
                    c.number_format = formatos[j - 1]
        for i in range(len(filas) + 1, previo + 1):   # filas sobrantes de la corrida anterior (solo sus columnas)
            for j in range(1, n_cols + 1):
                ws.cell(i, j).value = None

    def guardar(self):
        self.wb.save(self.ruta)

    def cerrar(self):
        pass


class _LibroExcel:
    """Con Excel (xlwings): conserva todo lo que Excel guarda (graficos, formatos, tablas)."""

    def __init__(self, ruta):
        import xlwings as xw
        self.app = xw.App(visible=False, add_book=False)
        self.app.display_alerts = False
        self.wb = self.app.books.open(str(ruta))

    def escribir(self, hoja, filas, formatos):
        nombres = [s.name for s in self.wb.sheets]
        sh = self.wb.sheets[hoja] if hoja in nombres else self.wb.sheets.add(hoja, after=self.wb.sheets[-1])
        n_cols = len(filas[0])
        previo = sh.range((sh.cells.last_cell.row, 1)).end("up").row
        if previo > len(filas):
            sh.range((len(filas) + 1, 1), (previo, n_cols)).clear_contents()
        sh.range((1, 1)).value = filas
        sh.range((1, 1), (1, n_cols)).font.bold = True
        for j, fmt in enumerate(formatos, start=1):
            if fmt and len(filas) > 1:
                sh.range((2, j), (len(filas), j)).number_format = fmt

    def guardar(self):
        self.wb.save()

    def cerrar(self):
        self.wb.close()
        self.app.quit()


def _abrir_libro(ruta):
    """Si el archivo lo guardo Excel y tiene graficos o tablas, se edita con Excel (xlwings); si no hay
    Excel, con openpyxl y respaldo en Inputs\\respaldo."""
    if ruta.exists():
        with zipfile.ZipFile(ruta) as z:
            nombres = z.namelist()
            app = z.read("docProps/app.xml").decode("utf-8", "ignore") if "docProps/app.xml" in nombres else ""
        con_objetos = any(n.startswith(("xl/charts/", "xl/drawings/", "xl/pivotTables/", "xl/tables/"))
                          for n in nombres)
        if con_objetos and "openpyxl" not in app.lower():
            try:
                return _LibroExcel(ruta)
            except Exception as e:  # noqa: BLE001
                respaldo = INPUTS / "respaldo" / f"{ruta.stem}_{dt.date.today():%Y%m%d}.xlsx"
                respaldo.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(ruta, respaldo)
                print(f"  ⚠️ No pude abrir Excel con xlwings ({str(e)[:80]}); actualizo con openpyxl, que puede "
                      f"simplificar graficos o formatos. Respaldo: {respaldo}")
    return _LibroOpenpyxl(ruta)


def escribir_base(hojas, ruta=None):
    """hojas: {nombre: DataFrame}. Cada DataFrame va desde A1 (fecha en la columna A); solo se tocan esas
    columnas y filas (las sobrantes de la corrida anterior se vacian en esas mismas columnas)."""
    ruta = ruta or carpeta_base() / BASE_EXCEL
    ruta.parent.mkdir(parents=True, exist_ok=True)
    libro = _abrir_libro(ruta)
    try:
        for hoja, df in hojas.items():
            d = df.rename_axis("Fecha")
            libro.escribir(hoja[:31], _filas(d), [_formato_columna(c) for _, c in d.reset_index().items()])
        libro.guardar()
    except PermissionError:
        sys.exit(f"No pude guardar {ruta}: está abierto. Ciérralo y vuelve a correr.")
    finally:
        libro.cerrar()
    return ruta


def main(hoy=None):
    hoy = pd.Timestamp(hoy or dt.date.today())
    a0, a1 = hoy.year, hoy.year + 1
    OUT.mkdir(exist_ok=True)

    prs = fg.nueva_presentacion()
    mes = f"{MESES[hoy.month - 1]} {hoy.year}"
    fg.portada(prs, "Chartpack de\nExpectativas", mes.capitalize(), autor=AUTOR)
    _PIE[0] = f"{TITULO} {mes}  |  {AUTOR}"
    _NUM[0] = 1
    DATOS.clear()

    # --- IMCE: precios y costos esperados por sector
    sectores = ["Comercio", "Construcción", "Industria"]
    precios = fg.bde(["G089.IME.IND.A14.M", "G089.IME.IND.A34.M", "G089.IME.IND.A27.M"],
                        nombres=sectores, desde="2018-01-01")
    costos = fg.bde(["G089.IME.IND.A15.M", "G089.IME.IND.A35.M", "G089.IME.IND.A29.M"],
                       nombres=sectores, desde="2018-01-01")
    lam = lamina(prs, "IMCE")
    for x, df, tit, hoja in [(DOS[0], precios, "IMCE: Indicador de Precios", "IMCE precios"),
                             (DOS[1], costos, "IMCE: Indicador de Costos", "IMCE costos")]:
        graficar(lam, df, x, ANCHO_DOS, hoja, titulo=tit, unidad="(Índice)", y_min=0, y_max=100,
                 y_paso=10, linea_ref=50, meses_entre_etiquetas=12)

    # --- Expectativas de inflacion EEE y EOF
    eee = fg.bde(["F089.IPC.V12.14.M", "F089.IPC.V12.15.M"], nombres=["EEE 1 año", "EEE 2 años"],
                    desde="2015-01-01")
    eof = fg.bde(["F089.EOF.VII.12MS.D", "F089.EOF.VII.S12M.D"], nombres=["EOF 1 año", "EOF 2 años"],
                    desde="2015-01-01")
    lam = lamina(prs, "Expectativas de Inflación")
    graficar(lam, eee, DOS[0], ANCHO_DOS, "EEE inflacion", titulo="Mediana EEE 1 y 2 años",
             unidad="(Var. anual, %)", posicion_unidad="arriba", y_min=1, y_max=9, y_paso=1,
             meses_entre_etiquetas=12, formato_fecha="yyyy",
             nota="Inflación esperada a 11 meses (1 año) y a 23 meses (2 años).")
    graficar(lam, eof, DOS[1], ANCHO_DOS, "EOF inflacion", titulo="Mediana EOF 1 y 2 años",
             unidad="(Var. anual, %)", posicion_unidad="arriba", y_min=1, y_max=9, y_paso=1,
             meses_entre_etiquetas=12, formato_fecha="yyyy",
             nota="Inflación esperada para los 12 meses siguientes (1 año) y de 13 a 24 meses (2 años).")

    # --- EEE: dispersion de expectativas de inflacion (deciles)
    lam = lamina(prs, "EEE: Dispersión de expectativas de inflación")
    for x, h, tit in [(DOS[0], "11m", "EEE: Inflación en 11 meses"),
                      (DOS[1], "23m", "EEE: Inflación en 23 meses")]:
        graficar(lam, eee_deciles(h), x, ANCHO_DOS, f"EEE deciles {h}", titulo=tit,
                 fuente="Banco Central de Chile, Encuesta de Expectativas Económicas",
                 unidad="(Var. anual, %)", posicion_unidad="arriba", barras=["Decil 9 - 1"],
                 etiqueta_ultimo=["Mediana", "Decil 9"], y_min=0, y_max=12 if h == "11m" else 6,
                 y_paso=2 if h == "11m" else 1, meses_entre_etiquetas=24, formato_fecha="yyyy",
                 nota="Columnas: diferencia entre el decil 9 y el decil 1 de las respuestas.")

    # --- EOF: operadores con inflacion a 2 anios sobre/bajo umbrales
    umbrales = eof_umbrales()
    lam = lamina(prs, "EOF: Inflación a 2 años")
    for x, col, tit in [(TRES[0], "pct_sobre_3_0", "Sobre 3,0%"), (TRES[1], "pct_sobre_3_2", "Sobre 3,2%"),
                        (TRES[2], "pct_bajo_2_8", "Bajo 2,8%")]:
        df = umbrales[[col]].rename(columns={col: f"Operadores {tit.lower()}"})
        graficar(lam, df, x, ANCHO_TRES, f"EOF {tit}", titulo=f"Inflación a 2 años {tit.lower()}",
                 fuente="Banco Central de Chile, Encuesta de Operadores Financieros",
                 unidad="(% de operadores)", posicion_unidad="arriba", tam=TAM_TRES,
                 y_min=0, y_max=100, y_paso=20, linea_ref=50, meses_entre_etiquetas=24,
                 formato_fecha="yyyy",
                 nota="Inflación esperada 13 a 24 meses adelante; promedio simple de las encuestas "
                      "que cerraron en cada mes (nov-21 sin encuesta).")

    # --- Compensacion inflacionaria (swaps promedio camara)
    comp = fg.bde(["F022.SWSP.TAS.AN01.Z.Z.D", "F022.SWSP.TAS.AN02.Z.Z.D"],
                     nombres=["1 año", "2 años"], desde="2022-01-01")
    c11 = fg.bde(["F022.SWFW.TAS.E1X1.Z.Z.D"], nombres=["Swap 1 en 1"], desde="2022-01-01")
    lam = lamina(prs, "Compensación Inflacionaria")
    graficar(lam, comp, DOS[0], ANCHO_DOS, "Compensaciones", titulo="Compensaciones inflacionarias",
             unidad="(Var. anual, %)", posicion_unidad="arriba", y_min=0, y_paso=2,
             meses_entre_etiquetas=12, nota="Compensación spot a partir de swaps promedio cámara.")
    graficar(lam, c11, DOS[1], ANCHO_DOS, "Swap 1 en 1", titulo="Compensación inflacionaria 1 en 1: Swap",
             unidad="(%)", posicion_unidad="arriba", y_min=0, y_paso=1, meses_entre_etiquetas=12,
             nota="Compensación forward 1 año en 1 año a partir de swaps promedio cámara.")

    # --- EDEP: costos y determinantes del precio de venta
    ini = "2021-11-01"
    costos_edep = ["Costos laborales", "Precios de insumos", "Costo de maquinarias o tecnología",
                   "Costos fijos"]
    edep_c = fg.bde([f"F089.EDEP.IND.88{k}.M" for k in ["CL", "PI", "CMT", "CF"]],
                       nombres=costos_edep, desde=ini)
    edep_i = fg.bde([f"F089.EDEP.IND.84IF{k}.M" for k in ["NV", "PD", "MG", "DI", "C"]],
                       nombres=["Niveles de venta", "Precio del dólar", "Márgenes de ganancias",
                                "Disponibilidad de insumos", "Costos"], desde=ini)
    edep_p = fg.bde([f"F089.EDEP.IND.90{k}.M" for k in ["CL", "PI", "CMT", "CF"]],
                       nombres=costos_edep, desde=ini)
    lam = lamina(prs, "EDEP")
    comun = dict(unidad="(Índice de difusión)", linea_ref=50, meses_entre_etiquetas=12, tam=TAM_TRES)
    graficar(lam, edep_c, TRES[0], ANCHO_TRES, "EDEP costos", titulo="EDEP: Evolución de costos",
             y_min=40, y_paso=5, nota="Últimos 3 meses; datos de trimestre móvil.", **comun)
    graficar(lam, edep_i, TRES[1], ANCHO_TRES, "EDEP precio venta",
             titulo="EDEP: Influencia en precio de venta", y_min=0, y_paso=10,
             nota="Últimos 3 meses; datos de trimestre móvil.", **comun)
    graficar(lam, edep_p, TRES[2], ANCHO_TRES, "EDEP costos esperados",
             titulo="EDEP: Costos esperados", y_min=40, y_paso=5,
             nota="Próximos 3 meses; datos de trimestre móvil.", **comun)

    # --- EDEP expectativas de inflacion y costos de transporte
    edep_inf = fg.bde(["F089.EDEP.POR.91EI12.M", "F089.EDEP.POR.91EI24.M"],
                         nombres=["EDEP a 1 año", "EDEP a 2 años"], desde=ini)
    gscpi = fg.nyfed_gscpi(desde="1998-01-01", nombre="GSCPI")
    lam = lamina(prs, "IMCE, EDEP y Costos de transporte")
    graficar(lam, edep_inf, DOS[0], ANCHO_DOS, "EDEP inflacion", titulo="EDEP: Expectativas de Inflación",
             unidad="(Var. anual, %)", y_min=0, y_paso=2, meses_entre_etiquetas=12,
             nota="Mediana de inflación esperada por las empresas, 1 a 12 y 13 a 24 meses.")
    graficar(lam, gscpi, DOS[1], ANCHO_DOS, "GSCPI", fuente=gscpi.attrs["fuente"],
             titulo="Costos de transporte", unidad="(Desv. estándar del promedio)",
             posicion_unidad="arriba", y_paso=1, meses_entre_etiquetas=36, formato_fecha="yyyy",
             nota="Global Supply Chain Pressure Index (GSCPI).")

    # --- Tipo de cambio nominal y multilateral
    tc = fg.bde(["F073.TCO.PRE.Z.D", "F073.TCM.IND.199502.D"], nombres=["TCN", "TCM (der.)"],
                   desde="2018-01-01")
    lam = lamina(prs, "TCN TCM")
    graficar(lam, tc, DOS[0], ANCHO_UNO, "TCN TCM",
             titulo="Tipo de cambio nominal y tipo de cambio multilateral", decimales=0,
             unidad="(Pesos chilenos)", unidad_derecha="(Índice)", posicion_unidad="arriba",
             eje_derecho=["TCM (der.)"], y_min=500, y_paso=100, y2_min=90, y2_paso=10,
             meses_entre_etiquetas=12, formato_fecha="yyyy",
             nota="TCM: tipo de cambio multilateral, índice 2 de enero de 1998=100.")

    # --- Tasas swap promedio camara
    ini8 = f"{a0 - 1}-01-01"
    spc_clp = fg.bde(["F022.SPC.TPR.D360.NO.Z.D", "F022.SPC.TIN.AN02.NO.Z.D"],
                        nombres=["SPC pesos 1 año", "SPC pesos 2 años"], desde=ini8)
    spc_uf = fg.bde(["F022.SPC.TIN.AN01.UF.Z.D", "F022.SPC.TIN.AN02.UF.Z.D"],
                       nombres=["SPC UF 1 año", "SPC UF 2 años"], desde=ini8)
    lam = lamina(prs, "Tasas SPC")
    graficar(lam, spc_clp, DOS[0], ANCHO_DOS, "SPC pesos", titulo="SPC pesos: 1 y 2 años",
             unidad="(%)", y_paso=0.5, meses_entre_etiquetas=3, decimales=1,
             nota="SPC pesos 1 año: tasa swap promedio cámara nominal a 360 días.")
    graficar(lam, spc_uf, DOS[1], ANCHO_DOS, "SPC UF", titulo="SPC UF: 1 y 2 años",
             unidad="(%)", y_paso=0.5, meses_entre_etiquetas=3, decimales=1)

    # --- EEE: actividad, anio en curso y siguiente
    ini9 = f"{a0 - 1}-01-01"
    lam = lamina(prs, "EEE")
    for x, (cod, tit, paso) in zip(TRES, [("PIB", "Expectativas de PIB", 0.5),
                                          ("CHOG", "Expectativas de Consumo hogares", 0.5),
                                          ("FBKF", "Expectativas de FBCF", 1)]):
        df = fg.bde([f"F089.{cod}.V12.{a0}.M", f"F089.{cod}.V12.{a1}.M"],
                       nombres=[str(a0), str(a1)], desde=ini9)
        graficar(lam, df, x, ANCHO_TRES, f"EEE {cod}", titulo=tit, unidad="(Var. anual, %)", y_paso=paso,
                 meses_entre_etiquetas=6, tam=TAM_TRES, nota="Mediana de la EEE para cada año calendario.")

    # --- IPEC: percepcion de la economia y expectativas de consumo (GfK Adimark, en la BDE)
    ipec = fg.bde(["F089.IPE.IND.75M2.M", "F089.IPE.IND.74M2.M", "F089.IPE.IND.72M2.M", "F089.IPE.IND.70M2.M"],
                  nombres=["Índice de percepción de la economía",
                           "Expectativas de consumo de artículos para el hogar",
                           "Situación económica del país a 12 meses", "Situación económica personal actual"],
                  desde="2010-01-01")
    comun = dict(fuente=IPEC_FUENTE, unidad="(Índices de difusión, 0-100)", zonas=("optimista", "pesimista"),
                 y_min=10, y_paso=10, linea_ref=50, meses_entre_etiquetas=12, formato_fecha="yyyy",
                 nota="Sobre 50 indica optimismo; bajo 50, pesimismo.")
    lam = lamina(prs, "IPEC: Índice de percepción de la economía")
    graficar(lam, ipec[ipec.index >= "2018-01-01"], DOS[0], ANCHO_UNO, "IPEC",
             titulo="Índice de percepción de la economía (IPEC) y sus componentes", y_max=70, **comun)
    lam = lamina(prs, "Expectativas de consumo")
    graficar(lam, ipec[["Expectativas de consumo de artículos para el hogar"]], DOS[0], ANCHO_UNO,
             "IPEC expectativas consumo", titulo="IPEC: Expectativas de consumo de artículos para el hogar",
             y_max=80, **comun)

    # --- IMCE expectativas de inversion vs Imacec no minero
    imce = fg.bde(["G089.IME.IND.A17.M", "G089.IME.IND.A2B.M"],
                     nombres=["IMCE Comercio", "IMCE Industria"], desde="2010-01-01")
    imce_inv, nota_imce = fg.promedio(imce, "IMCE expectativas de inversión (der.)")
    imacec = fg.bde(["F032.IMC.IND.Z.Z.EP18.N03.Z.0.M"], nombres=["Imacec No Minero"],
                       desde="2008-01-01")
    imacec, nota_mm = fg.media_movil(imacec, 3, "meses")
    imacec, nota_va = fg.var_anual(imacec, 12)
    df10 = pd.concat([imacec, imce_inv], axis=1)
    df10 = df10[df10.index >= "2010-01-01"].dropna(how="all")
    lam = lamina(prs, "Imce expectativas v/s Imacec no minero")
    graficar(lam, df10, DOS[0], ANCHO_UNO, "IMCE vs Imacec",
             titulo="IMCE expectativas de inversión v/s Imacec no minero",
             unidad="(Var. trimestral a/a, %)", unidad_derecha="(Índice)", posicion_unidad="arriba",
             eje_derecho=["IMCE expectativas de inversión (der.)"], y_min=-20, y_max=25, y_paso=5,
             y2_min=10, y2_max=80, y2_paso=10, meses_entre_etiquetas=12, formato_fecha="yyyy",
             nota=f"Imacec no minero: {nota_mm.lower()} {nota_va} {nota_imce}")

    # --- guardar
    nombre = f"{hoy:%Y%m%d}_Chartpack_Expectativas"
    pptx = OUT / f"{nombre}.pptx"
    try:
        prs.save(pptx)
    except PermissionError:
        sys.exit(f"No pude guardar {pptx}: está abierto en PowerPoint. Ciérralo y vuelve a correr.")
    print(pptx)
    print(f"Base: {escribir_base(DATOS)}")
    pdf = fg.exportar_pdf(pptx)
    if pdf:
        print(pdf)


if __name__ == "__main__":
    main()
