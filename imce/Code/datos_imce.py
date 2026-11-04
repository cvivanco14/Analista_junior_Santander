# -*- coding: utf-8 -*-
r"""Datos de la Minuta IMCE (ICARE/UAI, publicado en la BDE del Banco Central de Chile).

1. Descarga las series del IMCE desde la API del BCCh (con formato_graficos.py) y deja el CSV en
   Inputs\imce.csv.
2. Actualiza el Excel de bases Outputs\Base_IMCE.xlsx (siempre con ese nombre) manteniendo el
   formato de la planilla IMCE (hojas IMCE y Grafs):
     - si no existe, lo crea copiando Inputs\Base_IMCE_plantilla.xlsx (o la ubicación anterior
       IMCE\Base\IMCE.xlsx);
     - en la hoja IMCE solo escribe las columnas de las series; los meses nuevos copian el formato de
       la fila de arriba y los gráficos (hoja Grafs) se extienden hasta el último mes;
     - las hojas Léeme, SA, Pesos y Compuestos son del script y se reescriben en cada corrida
       (Compuestos queda con fórmulas de Excel). Los cálculos y gráficos propios van en otras hojas.
   Si el archivo lo guardó Excel y tiene gráficos, lo actualiza con Excel (xlwings) cuando está
   disponible; si no, con openpyxl dejando un respaldo en Outputs\respaldo.

Carpetas (relativas a este script; sirven para cualquier usuario):
    <IMCE>\Code\      datos_imce.py y minuta_imce.py
    <IMCE>\Inputs\    imce.csv (descarga), Base_IMCE_plantilla.xlsx y hacienda_imce_sa.csv (referencia)
    <IMCE>\Outputs\   Base_IMCE.xlsx (bases y cálculos) y la minuta: PPTX, PDF y _datos.xlsx

Uso (no hay que editar nada):
    python datos_imce.py                 # descarga y actualiza Inputs y Outputs\Base_IMCE.xlsx
    python datos_imce.py --sin-descarga  # solo rehace Base_IMCE.xlsx con el Inputs\imce.csv existente

Credenciales: el token del BCCh lo lee formato_graficos.py desde la variable de entorno BDE_TOKEN
(o BCCH_API_token) o desde el Llavero; nunca se escribe en el código (ver README principal).
"""
import argparse
import datetime as dt
import math
import os
import re
import shutil
import sys
import unicodedata
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pandas as pd

CODE = Path(__file__).resolve().parent
PROYECTO = CODE.parent
INPUTS = PROYECTO / "Inputs"
OUT = PROYECTO / "Outputs"
CSV = INPUTS / "imce.csv"
HACIENDA = INPUTS / "hacienda_imce_sa.csv"
EXCEL = OUT / "Base_IMCE.xlsx"                 # salida: siempre Base_<proyecto>.xlsx
PLANTILLA = INPUTS / "Base_IMCE_plantilla.xlsx"  # formato de la planilla (hojas IMCE y Grafs)
FECHA_INICIO = "2003-01-01"

MANUAL_ESTILO = Path("Administrativos") / "Manual de Estilo"
BIBLIOTECA = Path("Estudios - Documentos") / "General"
# ubicación anterior de IMCE.xlsx en la carpeta compartida (la usaba imce_principal.py)
RUTA_ANTERIOR = Path("Bases - Chartpack") / "Bases" / "Local" / "IMCE" / "Base" / "IMCE.xlsx"


# ─────────────────────────────────────────────────────────────────────────────
# Carpetas
# ─────────────────────────────────────────────────────────────────────────────

def _raices_onedrive():
    """Carpetas sincronizadas de OneDrive/SharePoint (Windows: C:\\Users\\<usuario>; macOS: CloudStorage)."""
    raices = [Path(os.environ[v]) for v in ("OneDriveCommercial", "OneDrive") if os.environ.get(v)]
    casa = Path.home()
    for base in (casa, casa / "Library" / "CloudStorage"):
        if base.is_dir():
            raices += [p for p in base.iterdir()
                       if p.is_dir() and re.search(r"santander|onedrive|sharepoint", p.name, re.I)]
    return list({str(r).lower(): r for r in raices if r.is_dir()}.values())


def _carpeta_formato() -> Path:
    """Carpeta de formato_graficos.py: Administrativos/Manual de Estilo de la biblioteca de Estudios."""
    candidatos = []
    if os.environ.get("FORMATO_GRAFICOS_DIR"):
        candidatos.append(Path(os.environ["FORMATO_GRAFICOS_DIR"]))
    # 1) subiendo desde la carpeta de este script (...\General\Expectativas y Encuestas\IMCE\Code)
    for p in [CODE, *CODE.parents]:
        candidatos += [p / MANUAL_ESTILO, p / "formato"]   # "formato": estructura del repositorio
    # 2) en la biblioteca sincronizada del usuario que corre el script
    for raiz in _raices_onedrive():
        candidatos.append(raiz / BIBLIOTECA / MANUAL_ESTILO)
    for carpeta in candidatos:
        if (carpeta / "formato_graficos.py").is_file():
            return carpeta
    sys.exit("No encuentro formato_graficos.py. Debe estar en\n"
             f"    ...\\{BIBLIOTECA}\\{MANUAL_ESTILO}\n"
             "(con su carpeta assets). Para indicar otra carpeta: variable FORMATO_GRAFICOS_DIR.")


sys.path.insert(0, str(_carpeta_formato()))
import formato_graficos as fg  # noqa: E402


# ─────────────────────────────────────────────────────────────────────────────
# Series: (clave, sufijo del código G089.IME.IND.<x>.M, nombre de la columna en IMCE.xlsx)
# Los nombres son los de la planilla IMCE.xlsx de siempre, para poder actualizarla.
# ─────────────────────────────────────────────────────────────────────────────

SERIES = [
    ("total", "A0", "IMCE Total"), ("sin_mineria", "A04", "IMCE: Sin Minería"),
    ("comercio", "A1", "Comercio"), ("industria", "A2", "Industria"),
    ("construccion", "A3", "Construcción"), ("mineria", "A4", "Minería"),
    # comercio
    ("com_inventarios", "A11", "COM_Situacion actual - Inventarios"),
    ("com_demanda", "A12", "COM_Situacion actual - Demanda"),
    ("com_situacion", "A13", "COM_Situacion actual - Situación general de la empresa"),
    ("com_precios", "A14", "COM_Expectativas - Precios"),
    ("com_costos", "A15", "COM_Expectativas - Costos"),
    ("com_empleo", "A16", "COM_Expectativas - Empleo"),
    ("com_inversion", "A17", "COM_Expectativas - Inversiones"),
    ("com_pais", "A18", "COM_Expectativas - Situación económica global del país"),
    ("com_negocio", "A19", "COM_Expectativas - Situación general del negocio"),
    ("com_ventas", "A1A", "COM_Expectativas - Ventas"),
    ("com_financiera", "A1B", "COM_Expectativas - Situación financiera empresa"),
    ("com_inflacion", "A1C", "COM_Expectativas - Inflación esperada 12 meses"),
    # construcción
    ("con_demanda", "A31", "CONS_Situacion actual - Demanda"),
    ("con_actividad", "A32", "CONS_Situacion actual - Nivel de actividad (3 meses)"),
    ("con_situacion", "A33", "CONS_Situacion actual - Situación general de la empresa"),
    ("con_precios", "A34", "CONS_Expectativas - Precios"),
    ("con_costos", "A35", "CONS_Expectativas - Costos"),
    ("con_salarios", "A36", "CONS_Expectativas - Salarios"),
    ("con_empleo", "A37", "CONS_Expectativas - Empleo"),
    ("con_pais", "A38", "CONS_Expectativas - Situación económica global del país"),
    ("con_financiera", "A39", "CONS_Expectativas - Situación financiera empresa"),
    # industria
    ("ind_produccion", "A21", "IND_Situacion actual - Producción"),
    ("ind_inventarios", "A22", "IND_Situacion actual - Inventarios"),
    ("ind_demanda", "A23", "IND_Situacion actual - Demanda"),
    ("ind_capacidad", "A24", "IND_Situacion actual - Uso de capacidad"),
    ("ind_situacion", "A25", "IND_Situacion actual - Situación general de la empresa"),
    ("ind_prod_esp", "A26", "IND_Expectativas - Producción"),
    ("ind_precios", "A27", "IND_Expectativas - Precios"),
    ("ind_salarios", "A28", "IND_Expectativas - Salarios"),
    ("ind_costos", "A29", "IND_Expectativas - Costos"),
    ("ind_empleo", "A2A", "IND_Expectativas - Empleo"),
    ("ind_inversion", "A2B", "IND_Expectativas - Inversiones"),
    ("ind_pais", "A2C", "IND_Expectativas - Situación económica global del país"),
    ("ind_negocio", "A2D", "IND_Expectativas - Situación general del negocio"),
    ("ind_inflacion", "A2E", "IND_Expectativas - Inflación esperada 12 meses"),
]
CODIGO = {k: f"G089.IME.IND.{s}.M" for k, s, _ in SERIES}
NOMBRE = {k: n for k, _, n in SERIES}
# otras columnas de IMCE.xlsx con la misma serie
ALIAS_EXCEL = {"IND_Indice de difusión sectorial": "industria", "COM_Indice de difusión sectorial": "comercio",
               "CONS_Indice de difusión sectorial": "construccion"}

# ─────────────────────────────────────────────────────────────────────────────
# Índices compuestos: fórmulas de la planilla de Coordinación Macroeconómica (Hacienda, IMCE.xlsx,
# hoja «Indicadores (SA)»), sobre índices de difusión desestacionalizados y centrados en 50.
# Cada sector: [(serie, signo)] -> promedio de signo * (serie - 50). Pesos: PIB 2018 normalizado.
# ─────────────────────────────────────────────────────────────────────────────

PESOS = {"Comercio": 0.4142243648672061, "Construcción": 0.2357331330488559, "Industria": 0.3500425020839379}
SECTORES = list(PESOS)
DEFINICION = {
    # Expectativas futuras del país: situación económica global del país
    "CFEI": {"Comercio": [("com_pais", 1)], "Construcción": [("con_pais", 1)], "Industria": [("ind_pais", 1)]},
    # Sentimiento actual general: situación de la empresa − inventarios + demanda/actividad/producción
    "OCSI": {"Comercio": [("com_situacion", 1), ("com_inventarios", -1), ("com_ventas", 1)],
             "Construcción": [("con_situacion", 1), ("con_actividad", 1), ("con_demanda", 1)],
             "Industria": [("ind_situacion", 1), ("ind_inventarios", -1), ("ind_demanda", 1), ("ind_produccion", 1)]},
    # Expectativas futuras generales
    "OFEI": {"Comercio": [("com_pais", 1), ("com_negocio", 1), ("com_ventas", 1), ("com_financiera", 1)],
             "Construcción": [("con_pais", 1), ("con_financiera", 1), ("con_empleo", 1)],
             "Industria": [("ind_situacion", 1), ("ind_negocio", 1), ("ind_prod_esp", 1)]},
    # Situación actual «personal» (de la empresa)
    "PCSI": {"Comercio": [("com_situacion", 1), ("com_inventarios", -1), ("com_demanda", 1)],
             "Construcción": [("con_situacion", 1), ("con_actividad", 1), ("con_demanda", 1)],
             "Industria": [("ind_situacion", 1), ("ind_inventarios", -1), ("ind_demanda", 1), ("ind_prod_esp", 1)]},
}
NOMBRES_COMPUESTOS = {"OCSI": "Sentimiento actual general", "PCSI": "Situación actual personal",
                      "CFEI": "Expectativas futuras del país", "OFEI": "Expectativas futuras generales"}
# series que se desestacionalizan (las de los compuestos; incluye demanda, ventas y producción esperada)
ITEMS_SA = list(dict.fromkeys(k for ind in DEFINICION.values() for comp in ind.values() for k, _ in comp))
NOTA_SA = "Desestacionalización propia: STL robusto, período 12 (statsmodels)."


# ─────────────────────────────────────────────────────────────────────────────
# Cálculos
# ─────────────────────────────────────────────────────────────────────────────

def desest(s: pd.Series) -> pd.Series:
    """Serie menos su componente estacional STL (robusto, período 12)."""
    from statsmodels.tsa.seasonal import STL
    x = s.dropna()
    r = STL(x, period=12, robust=True).fit()
    return (x - r.seasonal).reindex(s.index)


def desestacionalizar(m: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame({k: desest(m[k]) for k in ITEMS_SA})


def compuestos(sa: pd.DataFrame):
    """{índice: (DataFrame de contribuciones por sector, total)}, en desviaciones de 50."""
    out = {}
    for nombre, sectores in DEFINICION.items():
        c = pd.DataFrame({s: PESOS[s] * sum(signo * (sa[k] - 50) for k, signo in items) / len(items)
                          for s, items in sectores.items()})
        out[nombre] = (c, c.sum(axis=1, min_count=len(sectores)))
    return out


def cargar() -> pd.DataFrame:
    """Series del IMCE desde Inputs/imce.csv (columnas = claves de SERIES)."""
    if not CSV.is_file():
        sys.exit(f"No encuentro {CSV}. Corre primero: python datos_imce.py")
    m = pd.read_csv(CSV, index_col=0, parse_dates=True).asfreq("MS")
    return m.reindex(columns=[k for k, _, _ in SERIES])   # series que falten (ej. CSV antiguo) quedan vacías


# ─────────────────────────────────────────────────────────────────────────────
# Descarga (API BDE del BCCh vía formato_graficos)
# ─────────────────────────────────────────────────────────────────────────────

def _serie(clave):
    df = fg.bde([CODIGO[clave]], nombres=[clave], desde=FECHA_INICIO)
    s = df.iloc[:, 0]
    s.index = s.index.to_period("M").to_timestamp()
    return s


def descargar() -> pd.DataFrame:
    INPUTS.mkdir(exist_ok=True)
    try:   # la primera serie prueba la credencial antes de lanzar las demás en paralelo
        primera = _serie("total")
    except Exception as e:  # noqa: BLE001
        texto = f"{type(e).__name__} {e}".lower()
        if "invalid" in texto and ("password" in texto or "credential" in texto):
            sys.exit("El Banco Central rechazó el token (variable BDE_TOKEN / BCCH_API_token o Llavero).\n"
                     "Revisa la sección «Si aparece Invalid username or password» del README de imce.")
        sys.exit(f"No pude conectarme a la API del BCCh: {e}")
    anterior = pd.read_csv(CSV, index_col=0, parse_dates=True) if CSV.is_file() else pd.DataFrame()

    def bajar(clave):
        try:
            return _serie(clave)
        except Exception as e:  # noqa: BLE001 - se sigue con las demás y se conserva el dato anterior
            print(f"  ⚠️ {clave} ({CODIGO[clave]}): {e}")
            return anterior[clave] if clave in anterior else pd.Series(dtype=float, name=clave)

    claves = [k for k, _, _ in SERIES if k != "total"]
    with ThreadPoolExecutor(8) as ex:
        res = dict(zip(claves, ex.map(bajar, claves)))
    m = pd.DataFrame({"total": primera, **res})[[k for k, _, _ in SERIES]].sort_index()
    m.index.name = "fecha"
    m.to_csv(CSV)
    print(f"  ✅ {CSV.name}: {m.shape[1]} series, último dato {m['total'].dropna().index.max():%Y-%m}")
    return m


# ─────────────────────────────────────────────────────────────────────────────
# Excel de bases (Outputs/Base_IMCE.xlsx)
# ─────────────────────────────────────────────────────────────────────────────

FILA_CODIGO, FILA_TITULOS, FILA_DATOS = 2, 3, 4     # misma estructura que la planilla IMCE.xlsx
HOJAS_DEL_SCRIPT = ["Léeme", "SA", "Pesos", "Compuestos"]


def _norm(s):
    s = unicodedata.normalize("NFD", str(s) if s is not None else "")
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", s.strip().lower()).replace("–", "-").replace("—", "-")


def _mes(v):
    try:
        t = pd.Timestamp(v)
    except (ValueError, TypeError):
        return None
    return None if pd.isna(t) else t.to_period("M").to_timestamp()


def _letra(col):
    """1 -> A, 27 -> AA."""
    s = ""
    while col:
        col, r = divmod(col - 1, 26)
        s = chr(65 + r) + s
    return s


def _valor(v):
    return None if v is None or (isinstance(v, float) and math.isnan(v)) else float(v)


class _LibroOpenpyxl:
    """Edita el .xlsx con openpyxl (no necesita Excel). Conserva estilos y los gráficos que lee."""

    def __init__(self, ruta):
        from openpyxl import load_workbook
        self.ruta = ruta
        self.wb = load_workbook(ruta)

    def hojas(self):
        return self.wb.sheetnames

    def _hoja(self, nombre):
        return self.wb[nombre] if nombre in self.wb.sheetnames else self.wb.create_sheet(nombre)

    def valores(self, nombre):
        return [list(r) for r in self.wb[nombre].iter_rows(values_only=True)]

    def escribir(self, nombre, fila, col, filas, formulas=False):
        from copy import copy
        ws = self._hoja(nombre)
        propia = nombre in HOJAS_DEL_SCRIPT
        for i, f in enumerate(filas):
            for j, v in enumerate(f):
                c = ws.cell(fila + i, col + j)
                if not propia and not c.has_style and c.row > FILA_DATOS:
                    # filas nuevas de la planilla: mismo formato que la fila de arriba (antes del valor,
                    # porque openpyxl le pone formato propio a las fechas)
                    arriba = ws.cell(c.row - 1, c.column)
                    if arriba.has_style:
                        c._style = copy(arriba._style)
                c.value = v
                if propia:   # hojas del script: formato simple
                    if isinstance(v, (dt.date, pd.Timestamp)):
                        c.number_format = "mmm-yy"
                    elif isinstance(v, float) or (formulas and isinstance(v, str) and v.startswith("=")):
                        c.number_format = "0.0"

    def limpiar(self, nombre):
        ws = self._hoja(nombre)
        if ws.max_row:
            ws.delete_rows(1, ws.max_row)

    def ancho(self, nombre, col, ancho):
        if nombre in HOJAS_DEL_SCRIPT:   # en las hojas de la planilla se respetan los anchos
            self._hoja(nombre).column_dimensions[_letra(col)].width = ancho

    def guardar(self):
        self.wb.save(self.ruta)

    def cerrar(self):
        pass


class _LibroExcel:
    """Edita el .xlsx con Excel (xlwings): conserva todo lo que Excel guarda."""

    def __init__(self, ruta):
        import xlwings as xw
        self.app = xw.App(visible=False, add_book=False)
        self.app.display_alerts = False
        self.wb = self.app.books.open(str(ruta))

    def hojas(self):
        return [s.name for s in self.wb.sheets]

    def _hoja(self, nombre):
        if nombre not in self.hojas():
            self.wb.sheets.add(nombre, after=self.wb.sheets[-1])
        return self.wb.sheets[nombre]

    def valores(self, nombre):
        h = self.wb.sheets[nombre]
        ultima = h.used_range.last_cell
        return h.range((1, 1), (ultima.row, ultima.column)).options(ndim=2).value

    def escribir(self, nombre, fila, col, filas, formulas=False):
        if not filas:
            return
        hoja = self._hoja(nombre)
        rango = hoja.range((fila, col)).resize(len(filas), len(filas[0]))
        if nombre not in HOJAS_DEL_SCRIPT and fila > FILA_DATOS:
            # filas nuevas de la planilla: mismo formato que la fila de arriba
            hoja.range((fila - 1, col)).resize(1, len(filas[0])).copy()
            rango.paste(paste="formats")
        if formulas:
            rango.formula = [[("" if v is None else v) for v in f] for f in filas]   # fórmulas en inglés
        else:
            rango.value = filas

    def limpiar(self, nombre):
        self._hoja(nombre).cells.clear_contents()

    def ancho(self, nombre, col, ancho):
        pass

    def guardar(self):
        self.wb.save()

    def cerrar(self):
        self.wb.close()
        self.app.quit()


def _guardado_por_openpyxl(ruta):
    """True si el archivo lo guardó openpyxl por última vez (se reescribe sin cambios de formato)."""
    with zipfile.ZipFile(ruta) as z:
        app = z.read("docProps/app.xml").decode("utf-8", "ignore") if "docProps/app.xml" in z.namelist() else ""
        graficos = any(n.startswith(("xl/charts/", "xl/drawings/", "xl/pivotTables/")) for n in z.namelist())
    return "openpyxl" in app.lower(), graficos


def encontrar_excel():
    """Outputs/Base_IMCE.xlsx; si no existe, lo crea copiando la planilla de Inputs (o la ubicación anterior)."""
    if EXCEL.is_file():
        return EXCEL
    fuentes = [PLANTILLA, OUT / "IMCE.xlsx", PROYECTO / "Base" / "IMCE.xlsx"]
    for raiz in _raices_onedrive():
        fuentes += [raiz / RUTA_ANTERIOR] + [sub / RUTA_ANTERIOR for sub in raiz.iterdir() if sub.is_dir()]
    origen = next((p for p in fuentes if p.is_file()), None)
    if origen is None:
        return None
    shutil.copy2(origen, EXCEL)
    print(f"  📋 Creé {EXCEL.name} a partir de {origen}")
    return EXCEL


def _crear_excel(ruta):
    """Último recurso (sin planilla): libro con la estructura de la hoja IMCE."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "IMCE"
    ws["A1"] = "Índice mensual de confianza empresarial (IMCE), total general e índices de difusión sectoriales"
    ws["A2"] = "Código BDE"
    ws.cell(FILA_TITULOS, 1).value = "Periodo"
    ws.column_dimensions["A"].width = 11
    ws.freeze_panes = ws.cell(FILA_DATOS, 2)
    wb.save(ruta)
    print(f"  🆕 No encontré la planilla {PLANTILLA.name}; creé {ruta.name} sin formato")


def _abrir(ruta):
    de_openpyxl, graficos = _guardado_por_openpyxl(ruta)
    if graficos and not de_openpyxl:
        # guardado por Excel: con Excel (xlwings) no se pierde nada; si no, openpyxl con respaldo
        try:
            return _LibroExcel(ruta)
        except Exception as e:  # noqa: BLE001
            respaldo = OUT / "respaldo" / f"{ruta.stem}_{dt.date.today():%Y%m%d}.xlsx"
            respaldo.parent.mkdir(exist_ok=True)
            shutil.copy2(ruta, respaldo)
            print(f"  ⚠️ No pude abrir Excel con xlwings ({str(e)[:80]}). Actualizo con openpyxl, que "
                  f"conserva los gráficos pero puede simplificar parte de su formato.\n"
                  f"     Respaldo del archivo anterior: {respaldo}")
    return _LibroOpenpyxl(ruta)


def _extender_graficos(ruta, ultima):
    """Rangos de los gráficos sobre la hoja IMCE (filas 4 a N) hasta la última fila con datos."""
    patron = re.compile(r"((?:'IMCE'|IMCE)!\$[A-Z]{1,3}\$%d:\$[A-Z]{1,3}\$)(\d+)" % FILA_DATOS)
    cache = re.compile(r"<(?:\w+:)?(?:numCache|strCache)>.*?</(?:\w+:)?(?:numCache|strCache)>", re.S)
    with zipfile.ZipFile(ruta) as z:
        partes = {i: z.read(i.filename) for i in z.infolist()}
    cambios = 0
    for info in list(partes):
        if not (info.filename.startswith("xl/charts/chart") and info.filename.endswith(".xml")):
            continue
        xml = partes[info].decode("utf-8")
        nuevo = patron.sub(lambda mt: mt.group(1) + str(ultima), xml)
        if nuevo != xml:
            # sin valores en caché: Excel y LibreOffice leen las celdas al abrir
            nuevo = cache.sub("", nuevo)
            partes[info] = nuevo.encode("utf-8")
            cambios += 1
    if cambios:
        tmp = ruta.with_suffix(".tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for info, datos in partes.items():
                z.writestr(info, datos)
        tmp.replace(ruta)
    return cambios


def _actualizar_hoja_imce(libro, m):
    v = libro.valores("IMCE")
    v += [[] for _ in range(max(0, FILA_TITULOS - len(v)))]
    titulos = v[FILA_TITULOS - 1]
    col = {_norm(t): j + 1 for j, t in enumerate(titulos) if t is not None}
    if "periodo" not in col:
        sys.exit(f"No encuentro 'Periodo' en la fila {FILA_TITULOS} de la hoja IMCE.")
    c_per = col["periodo"]
    filas = {}
    for i in range(FILA_DATOS - 1, len(v)):
        f = v[i]
        d = _mes(f[c_per - 1]) if len(f) >= c_per and f[c_per - 1] is not None else None
        if d is None:
            break
        filas[d] = i + 1
    # fechas nuevas al final
    ultima = max(filas.values(), default=FILA_DATOS - 1)
    # meses posteriores al último de la planilla (los anteriores a su primer mes no se agregan)
    desde = max(filas, default=pd.Timestamp.min)
    nuevas = [d for d in m.index if d > desde and m.loc[d].notna().any()]
    for d in nuevas:
        ultima += 1
        filas[d] = ultima
    if nuevas:
        libro.escribir("IMCE", filas[nuevas[0]], c_per, [[d.to_pydatetime()] for d in nuevas])
    # columnas: las que existen (incluidos alias) y las que faltan, al final
    destino = {}
    for nombre, j in col.items():
        clave = next((k for k, _, n in SERIES if _norm(n) == nombre), None) or \
            next((k for a, k in ALIAS_EXCEL.items() if _norm(a) == nombre), None)
        if clave:
            destino.setdefault(clave, []).append(j)
    sig = max(col.values()) + 1
    for k, _, n in SERIES:
        if k not in destino and m[k].notna().any():
            libro.escribir("IMCE", FILA_CODIGO, sig, [[CODIGO[k]], [n]])
            libro.ancho("IMCE", sig, 14)
            destino[k] = [sig]
            sig += 1
    # valores: cada columna completa desde la primera fila de datos (conserva meses sin dato nuevo)
    n_filas = ultima - FILA_DATOS + 1
    por_fila = {r: d for d, r in filas.items()}
    for k, cols in destino.items():
        for j in cols:
            bloque = []
            for r in range(FILA_DATOS, ultima + 1):
                d = por_fila.get(r)
                nuevo = m.at[d, k] if d is not None and d in m.index else float("nan")
                previo = v[r - 1][j - 1] if r - 1 < len(v) and j - 1 < len(v[r - 1]) else None
                bloque.append([_valor(nuevo) if not pd.isna(nuevo) else previo])
            libro.escribir("IMCE", FILA_DATOS, j, bloque)
    return n_filas, len(nuevas), ultima


def _hojas_calculo(libro, m, sa):
    # SA: series desestacionalizadas (valores)
    libro.limpiar("SA")
    libro.escribir("SA", 1, 1, [["IMCE desestacionalizado. " + NOTA_SA],
                                ["Código BDE"] + [CODIGO[k] for k in ITEMS_SA],
                                ["Periodo"] + [NOMBRE[k] for k in ITEMS_SA]])
    sa = sa.dropna(how="all")
    libro.escribir("SA", FILA_DATOS, 1, [[d.to_pydatetime()] + [_valor(x) for x in fila]
                                         for d, fila in zip(sa.index, sa.to_numpy())])
    fila_sa = {d: FILA_DATOS + i for i, d in enumerate(sa.index)}
    col_sa = {k: 2 + i for i, k in enumerate(ITEMS_SA)}
    libro.ancho("SA", 1, 11)

    # Pesos
    libro.limpiar("Pesos")
    libro.escribir("Pesos", 1, 1, [["Sector", "Peso (PIB nominal 2018, normalizado)"]] +
                   [[s, p] for s, p in PESOS.items()])
    celda_peso = {s: f"Pesos!$B${2 + i}" for i, s in enumerate(SECTORES)}

    # Compuestos: fórmulas sobre SA y Pesos (se recalculan al abrir en Excel)
    libro.limpiar("Compuestos")
    encabezado, cols = ["Periodo"], []
    for ind in DEFINICION:
        encabezado += [f"{ind} {s}" for s in SECTORES] + [ind]
        cols.append(ind)
    libro.escribir("Compuestos", 1, 1, [
        ["Índices compuestos del IMCE (50 = 0): contribución de cada sector = peso × promedio de "
         "(serie desestacionalizada − 50); el signo − en inventarios. Fórmulas de Coordinación "
         "Macroeconómica (Hacienda), basadas en Chanut, Marcel y Medel (2018)."],
        [""] + [NOMBRES_COMPUESTOS[i] for i in DEFINICION for _ in range(len(SECTORES) + 1)],
        encabezado])
    completas = sa[ITEMS_SA].dropna().index
    filas_f = []
    for r_out, d in enumerate(completas, start=FILA_DATOS):
        r = fila_sa[d]
        fila, c = [d.to_pydatetime()], 2
        for ind, sectores in DEFINICION.items():
            inicio = c
            for s in SECTORES:
                items = sectores[s]
                suma = "".join(f"{'+' if signo > 0 else '-'}('SA'!{_letra(col_sa[k])}{r}-50)" for k, signo in items)
                fila.append(f"={celda_peso[s]}*({suma.lstrip('+')})/{len(items)}")
                c += 1
            fila.append("=" + "+".join(f"{_letra(j)}{r_out}" for j in range(inicio, c)))
            c += 1
        filas_f.append(fila)
    libro.escribir("Compuestos", FILA_DATOS, 1, [[f[0]] for f in filas_f])
    libro.escribir("Compuestos", FILA_DATOS, 2, [f[1:] for f in filas_f], formulas=True)
    libro.ancho("Compuestos", 1, 11)


def _leeme(libro, m):
    ult = m["total"].dropna().index.max()
    lineas = [
        ["Excel de bases de la Minuta IMCE"],
        [f"Actualizado: {dt.datetime.now():%d-%m-%Y %H:%M}. Último dato: {ult:%m-%Y}."],
        ["Fuente: ICARE/Universidad Adolfo Ibáñez, publicado en la Base de Datos Estadísticos del Banco Central de Chile."],
        [""],
        ["Hoja", "Contenido"],
        ["IMCE", "Índices de difusión originales (fila 2: código BDE; fila 3: nombre). El script solo escribe "
                 "las columnas de las series; columnas y fórmulas propias en esta hoja se conservan."],
        ["SA", "Series desestacionalizadas con STL (las que usan los índices compuestos). Se reescribe en cada corrida."],
        ["Pesos", "Pesos sectoriales (PIB 2018) de los índices compuestos. Se reescribe en cada corrida."],
        ["Compuestos", "CFEI, OCSI, OFEI y PCSI con fórmulas de Excel sobre SA y Pesos. Se reescribe en cada corrida."],
        [""],
        ["Cálculos y gráficos propios: en hojas nuevas (no se tocan) o en columnas nuevas de la hoja IMCE."],
        ["Se actualiza con: python datos_imce.py (carpeta Code)."],
    ]
    libro.limpiar("Léeme")
    libro.escribir("Léeme", 1, 1, lineas)
    libro.ancho("Léeme", 1, 14)
    libro.ancho("Léeme", 2, 110)


def actualizar_excel(m, sa):
    OUT.mkdir(exist_ok=True)
    ruta = encontrar_excel()
    if ruta is None:
        ruta = EXCEL
        _crear_excel(ruta)
    print(f"  📂 {ruta}")
    try:
        libro = _abrir(ruta)
    except PermissionError:
        sys.exit(f"{ruta.name} está abierto en Excel o bloqueado por OneDrive. Ciérralo y vuelve a correr.")
    try:
        n, nuevas, ultima = _actualizar_hoja_imce(libro, m)
        _hojas_calculo(libro, m, sa)
        _leeme(libro, m)
        try:
            libro.guardar()
        except PermissionError:
            sys.exit(f"No pude guardar {ruta.name}: está abierto en Excel. Ciérralo y vuelve a correr.")
    finally:
        libro.cerrar()
    g = _extender_graficos(ruta, ultima)
    print(f"  💾 {ruta.name}: {n} meses en la hoja IMCE ({nuevas} nuevos); hojas SA, Pesos y Compuestos al día"
          + (f"; {g} gráficos extendidos hasta la fila {ultima}" if g else ""))
    return ruta


def validar_hacienda(comp):
    """Compara los compuestos propios con los de la planilla de Hacienda (Inputs/hacienda_imce_sa.csv)."""
    if not HACIENDA.is_file():
        return
    h = pd.read_csv(HACIENDA, index_col=0, parse_dates=True)
    partes = []
    for k, (_, tot) in comp.items():
        if k in h:
            x = pd.concat([tot, h[k]], axis=1).dropna()
            partes.append(f"{k} {x.corr().iloc[0, 1]:.3f}")
    if partes:
        print(f"  Correlación con los compuestos de Hacienda ({HACIENDA.name}): " + ", ".join(partes))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--sin-descarga", action="store_true", help="usa el Inputs/imce.csv existente")
    a = ap.parse_args(argv)
    print("=" * 60 + "\n  Minuta IMCE: datos\n" + "=" * 60)
    m = cargar() if a.sin_descarga else descargar()
    sa = desestacionalizar(m)
    validar_hacienda(compuestos(sa))
    actualizar_excel(m, sa)


if __name__ == "__main__":
    main()
