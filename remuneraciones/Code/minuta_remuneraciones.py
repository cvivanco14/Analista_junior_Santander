# -*- coding: utf-8 -*-
r"""Minuta de Remuneraciones (INE): un solo script que descarga los datos y arma la minuta.

1. Datos (antes datos_remuneraciones.py):
     - Descarga del INE a Inputs\ solo lo que usa la minuta: IR, ICL e IR real en series originales
       (tabulados empalmados) y desestacionalizadas por el INE (tabulados empalmados
       desestacionalizados; se usa la hoja de la publicación más reciente), ajuste_estacional.xlsx
       (ocupados desestacionalizados) y horas.xlsx (horas habituales; el INE no las publica
       desestacionalizadas, así que se desestacionalizan con STL).
     - Calcula la masa salarial real = ocupados (sa INE) × horas habituales (sa STL) × IR real (sa INE),
       con su variación a/a, contribuciones y horas × salario real.
     - Deja en Inputs\ las bases intermedias Masa_salarial_Base.xlsx y Cifras_empleo.xlsx.
2. Minuta, con los mismos cálculos y las mismas figuras:
     - Documento Minuta_Remuneraciones_<mes>_<año>.pdf (reportlab, colores Santander): título con el
       resultado del mes; principales resultados (IR, ICL y salarios reales en sa y nsa frente a su
       promedio histórico del mes, y masa salarial real); Cuadro 1 con variaciones m/m (sa y nsa) y a/a
       y los promedios históricos del mes; y ocho figuras (imágenes en Inputs\Figuras).
     - Presentación AAAAMMDD_Minuta_Remuneraciones.pptx (y .pdf), con formato_graficos.py: los mismos
       textos, el mismo cuadro y las mismas figuras.
     - Outputs\IRL_Base.xlsx: datos completos (toda la historia) y los de cada cuadro y figura. Solo
       escribe en sus columnas: las demás columnas, hojas, fórmulas y gráficos que agregues se conservan.

Carpetas (relativas a este script; sirven para cualquier usuario):
    <Remuneraciones>\Code\      minuta_remuneraciones.py
    <Remuneraciones>\Inputs\    archivos del INE, bases intermedias e Inputs\Figuras\
    <Remuneraciones>\Outputs\   IRL_Base.xlsx, documento y presentación

Uso (no hay que editar nada; en Spyder, F5):
    python minuta_remuneraciones.py                 # descarga del INE y arma todo
    python minuta_remuneraciones.py --sin-descarga  # usa los archivos que ya están en Inputs
"""
import argparse
import datetime as dt
import os
import re
import shutil
import sys
import zipfile
from functools import reduce
from pathlib import Path

import numpy as np
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from statsmodels.tsa.seasonal import STL
from urllib3.util.retry import Retry

# ═════════════════════════════════════════════════════════════════════════════
# 1. DATOS DEL INE
# ═════════════════════════════════════════════════════════════════════════════

CODE = Path(__file__).resolve().parent
PROYECTO = CODE.parent
INPUTS = PROYECTO / "Inputs"
OUT = PROYECTO / "Outputs"
EXCEL_CIFRAS = INPUTS / "Cifras_empleo.xlsx"        # bases intermedias: quedan en Inputs
EXCEL = INPUTS / "Masa_salarial_Base.xlsx"
DESDE_GRAFICO = "2015-01-01"
ANIO_BASE = 2023          # base de los índices (como el IR del INE: base anual 2023 = 100)
FUENTE = "Fuente: INE (Encuesta Nacional de Empleo e Índice de Remuneraciones)."

OCUPADOS = "Ocupados SA (miles)"
HORAS = "Horas habituales semanales SA"
MS_IR_REAL = f"IR real SA ({ANIO_BASE}=100)"
MASA = f"Masa salarial real ({ANIO_BASE}=100)"
MASA_AA = "Masa salarial real a/a (%)"
CONTRIB = {OCUPADOS: "Contribución ocupados (pp)", HORAS: "Contribución horas (pp)",
           MS_IR_REAL: "Contribución salario real (pp)"}
MASA_HS = f"Horas × salario real ({ANIO_BASE}=100)"
MASA_HS_AA = "Horas × salario real a/a (%)"
ROJO, GRIS, GRIS_CLARO, NEGRO = "EC0000", "6F7779", "BFC5C7", "1D1D1B"

_INE = "https://www.ine.gob.cl/docs/default-source/"
_ENE = _INE + "ocupacion-y-desocupacion/cuadros-estadisticos/series-vigentes/"
_IR = _INE + "sueldos-y-salarios/cuadros-estadisticos/ir-icl-base-anual-2023-100/"
URLS = {
    # empleo: ocupados desestacionalizados (INE) y horas habituales
    "ajuste_estacional.xlsx": _ENE + "ajuste_estacional.xlsx?sfvrsn=617886e6_107",
    "horas.xlsx": _ENE + "horas.xlsx?sfvrsn=d294ce0e_171",
    # remuneraciones: series originales
    "tabulado_ir_empalmado.xlsx": _IR + "series-empalmadas/tabulado_ir_empalmado.xlsx?sfvrsn=9bb59cb4_43",
    "tabulado_icl_empalmado.xlsx": _IR + "series-empalmadas/tabulado_icl_empalmado.xlsx?sfvrsn=28e08211_45",
    "tabulado_ir_real_empalmado.xlsx": _IR + "series-empalmadas/tabulado_ir_real_empalmado.xlsx?sfvrsn=6f1d912d_43",
    # remuneraciones: series desestacionalizadas por el INE
    "tabulado_ir_icl_nominal_general_empalmado_desestacionalizado.xlsx":
        _IR + "series-empalmadas-desestacionalizadas/"
              "tabulado_ir_icl_nominal_general_empalmado_desestacionalizado.xlsx?sfvrsn=6f7d86b7_65",
    "tabulado_ir_real_general_empalmado_desestacionalizado.xlsx":
        _IR + "series-empalmadas-desestacionalizadas/"
              "tabulado_ir_real_general_empalmado_desestacionalizado.xlsx?sfvrsn=c6a67c32_68",
}

COL_HORAS_SA = "Promedio de horas habitualmente trabajadas a la semana (desest.)"
REMUNERACIONES = {   # archivo de la serie original -> nombre de la serie; la desestacionalizada es '<serie> (SA)'
    "tabulado_ir_empalmado.xlsx": "IR",
    "tabulado_icl_empalmado.xlsx": "ICL",
    "tabulado_ir_real_empalmado.xlsx": "IR Real",
}
SA_INE = {   # serie -> (archivo desestacionalizado del INE, valor de la columna 'indicador' o None)
    "IR": ("tabulado_ir_icl_nominal_general_empalmado_desestacionalizado.xlsx", "IR"),
    "ICL": ("tabulado_ir_icl_nominal_general_empalmado_desestacionalizado.xlsx", "ICL"),
    "IR Real": ("tabulado_ir_real_general_empalmado_desestacionalizado.xlsx", None),
}


# ─────────────────────────────────────────────────────────────────────────────
# Descarga
# ─────────────────────────────────────────────────────────────────────────────

def descargar(archivos=None):
    """Descarga los archivos del INE a Inputs. Si una descarga falla, conserva el archivo anterior."""
    INPUTS.mkdir(parents=True, exist_ok=True)
    sesion = requests.Session()
    sesion.mount("https://", HTTPAdapter(max_retries=Retry(total=3, backoff_factor=1,
                                                           status_forcelist=[429, 500, 502, 503, 504])))
    sesion.headers["User-Agent"] = "Mozilla/5.0"
    for nombre in archivos or URLS:
        url = URLS[nombre]
        destino = INPUTS / nombre
        # el parámetro sfvrsn es la versión del archivo en la web del INE; si cambió, se pide sin él
        for intento in (url, url.split("?")[0]):
            try:
                r = sesion.get(intento, timeout=60)
                r.raise_for_status()
                if not r.content.startswith(b"PK"):   # un .xlsx es un zip; si no, es una página de error
                    raise ValueError("la respuesta no es un archivo Excel")
                destino.write_bytes(r.content)
                print(f"  ✅ Descargado: {nombre}")
                break
            except Exception as e:  # noqa: BLE001
                error = e
        else:
            aviso = "se usa el archivo anterior" if destino.exists() else "no hay archivo anterior"
            print(f"  ❌ No pude descargar {nombre} ({error}); {aviso}.")
    sesion.close()


# ─────────────────────────────────────────────────────────────────────────────
# Fechas
# ─────────────────────────────────────────────────────────────────────────────

_MESES = {
    "ene": 1, "enero": 1, "jan": 1, "january": 1, "feb": 2, "febrero": 2, "february": 2,
    "mar": 3, "marzo": 3, "march": 3, "abr": 4, "abril": 4, "apr": 4, "april": 4, "may": 5, "mayo": 5,
    "jun": 6, "junio": 6, "june": 6, "jul": 7, "julio": 7, "july": 7, "ago": 8, "agosto": 8, "aug": 8,
    "august": 8, "sep": 9, "sept": 9, "septiembre": 9, "september": 9, "oct": 10, "octubre": 10,
    "october": 10, "nov": 11, "noviembre": 11, "november": 11, "dic": 12, "diciembre": 12, "dec": 12,
    "december": 12,
}
_PATRON_MES = re.compile(r"\b(" + "|".join(sorted(_MESES, key=len, reverse=True)) + r"|\d{1,2})\b", re.I)


def formatear_fecha_inteligente(df, col_anio="Año", col_trim="Trimestre"):
    """Agrega 'Periodo' (primera columna) con el último mes del trimestre móvil de col_trim.
    El año sube en uno solo si el trimestre termina en enero (p. ej. 'Nov - Ene' de 2015 = ene-2016)."""
    periodos = []
    for anio, texto in zip(df[col_anio].to_numpy(), df[col_trim].astype(str).to_numpy()):
        try:
            anio = int(float(anio))
        except (TypeError, ValueError):
            periodos.append(pd.NaT)
            continue
        tokens = _PATRON_MES.findall(texto)
        if not tokens:
            periodos.append(pd.NaT)
            continue
        ultimo = tokens[-1].lower()
        mes = int(ultimo) if ultimo.isdigit() else _MESES.get(ultimo)
        if not mes or not 1 <= mes <= 12:
            periodos.append(pd.NaT)
            continue
        periodos.append(pd.Timestamp(year=anio + (mes == 1), month=mes, day=1))
    out = df.drop(columns=["Periodo"], errors="ignore").copy()
    out.insert(0, "Periodo", periodos)
    return out


# ─────────────────────────────────────────────────────────────────────────────
# Lectura de los cuadros del INE
# ─────────────────────────────────────────────────────────────────────────────

def _fila_anio(df_raw):
    return df_raw[df_raw.apply(lambda x: x.astype(str).str.contains("Año").any(), axis=1)].index[0]


def _limpiar_columnas(df):
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
    return df


def _tabla_ine(archivo, hoja, salto=2):
    """Cuadro del INE con encabezado en la fila que contiene 'Año' y datos 'salto' filas más abajo.
    Deja solo las columnas con datos (más de 5 números) y 'Periodo' como primera columna."""
    df_raw = pd.read_excel(archivo, sheet_name=hoja, header=None)
    idx = _fila_anio(df_raw)
    headers = df_raw.iloc[idx].ffill().tolist()
    df = df_raw.iloc[idx + salto:].copy()
    df.columns = headers[:len(df.columns)]
    return _a_series(df)


def _a_series(df):
    validas = [i for i, c in enumerate(df.columns)
               if pd.to_numeric(df.iloc[:, i], errors="coerce").notnull().sum() > 5
               or any(k in str(c) for k in ("Año", "Trimestre"))]
    df = _limpiar_columnas(df.iloc[:, validas])
    col_anio = next(c for c in df.columns if "Año" in c)
    col_trim = next(c for c in df.columns if "Trimestre" in c)
    df = formatear_fecha_inteligente(df, col_anio, col_trim)
    df = df.drop(columns=[col_anio, col_trim]).dropna(subset=["Periodo"])
    for c in df.columns[1:]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df.sort_values("Periodo").reset_index(drop=True)


def _stl(serie):
    """Serie desestacionalizada con STL (tendencia + residuo), como en el notebook original."""
    serie = serie.astype(float).interpolate(method="linear")
    res = STL(serie, period=12, robust=True).fit()
    return res.trend + res.resid








def ajuste_estacional():
    archivo = INPUTS / "ajuste_estacional.xlsx"
    df_raw = pd.read_excel(archivo, sheet_name="niveles_as", header=None)
    idx = _fila_anio(df_raw)
    df = df_raw.iloc[idx + 1:].copy()
    df.columns = df_raw.iloc[idx].ffill().tolist()[:len(df.columns)]
    df = df.dropna(how="all", subset=df.columns[2:])
    return _a_series(df)




def horas():
    df = _tabla_ine(INPUTS / "horas.xlsx", "AS")
    col = next((c for c in df.columns if "habitualmente" in c), None)
    if col is None:
        print("  ⚠️ No encontré la columna de horas habitualmente trabajadas.")
        return df
    df[COL_HORAS_SA] = _stl(df.set_index("Periodo")[col]).to_numpy()
    return df




def _mes_num(v):
    """Mes como número: 8, '8', 'ago', 'Agosto'."""
    if pd.isna(v):
        return np.nan
    try:
        return int(float(v))
    except (TypeError, ValueError):
        texto = str(v).strip().lower()
        return _MESES.get(texto, _MESES.get(texto[:3], np.nan))


def _serie_mensual(df_raw, idx, col):
    """Serie mensual desde un cuadro del INE: columnas 0 (año) y 1 (mes), datos bajo la fila 'idx'."""
    d = df_raw.iloc[idx + 1:, [0, 1, col]].copy()
    d.columns = ["Anio", "Mes", "Valor"]
    d["Anio"] = d["Anio"].ffill()   # algunos cuadros escriben el año solo en enero
    d["Anio"] = pd.to_numeric(d["Anio"], errors="coerce")
    d["Mes"] = d["Mes"].map(_mes_num)
    d["Valor"] = pd.to_numeric(d["Valor"], errors="coerce")
    d = d.dropna()
    d = d[(d["Mes"] >= 1) & (d["Mes"] <= 12)]
    if d.empty:
        return pd.Series(dtype=float)
    idx_fecha = pd.to_datetime(dict(year=d.Anio.astype(int), month=d.Mes.astype(int), day=1))
    return pd.Series(d["Valor"].to_numpy(), index=idx_fecha).sort_index().rename_axis("Periodo")


def _sa_ine(serie):
    """Serie desestacionalizada del INE (SA_INE): hoja de la publicación más reciente (hojas DDMMAAAA),
    columna 'indicador/Índice desestacionalizado (id)'. Devuelve (serie, origen) o (None, motivo)."""
    archivo, indicador = SA_INE[serie]
    ruta = INPUTS / archivo
    if not ruta.exists():
        return None, f"falta {archivo}"
    libro = pd.read_excel(ruta, sheet_name=None, header=None)
    fechas = {h: pd.to_datetime(h, format="%d%m%Y", errors="coerce") for h in libro if re.fullmatch(r"\d{8}", h)}
    fechas = {h: f for h, f in fechas.items() if pd.notna(f)}
    if not fechas:
        return None, f"{archivo} no tiene hojas de publicación (DDMMAAAA)"
    hoja = max(fechas, key=fechas.get)
    df_raw = libro[hoja]
    idx = df_raw[df_raw.apply(lambda x: x.astype(str).str.strip().str.lower().eq("año").any(), axis=1)].index[0]
    enc = df_raw.iloc[idx].astype(str).str.lower().str.strip()
    col = next((j for j, h in enumerate(enc) if "desestacionalizado (id)" in h), None)
    if col is None:
        return None, f"{archivo}/{hoja} sin columna desestacionalizada (id)"
    datos = df_raw.iloc[idx:].reset_index(drop=True)
    if indicador is not None:
        col_ind = next(j for j, h in enumerate(enc) if h == "indicador")
        datos = pd.concat([datos.iloc[:1], datos.iloc[1:][datos.iloc[1:, col_ind].astype(str).str.strip() == indicador]])
    s = _serie_mensual(datos.reset_index(drop=True), 0, col)
    return s, f"INE ({archivo}, publicación {fechas[hoja]:%d-%m-%Y})"


FUENTE_SA = {}   # serie -> origen de su desestacionalización


def remuneraciones():
    """IR, ICL e IR real (base anual 2023 = 100): series originales (NSA, tabulados empalmados) y
    desestacionalizadas por el INE (SA). Si falta el archivo desestacionalizado, usa STL y avisa."""
    series = []
    for nombre, base in REMUNERACIONES.items():
        archivo = INPUTS / nombre
        if not archivo.exists():
            print(f"  ⚠️ No encontré {nombre}")
            continue
        df_raw = pd.read_excel(archivo, sheet_name="General", header=None)
        idx = df_raw[df_raw.apply(lambda x: x.astype(str).str.contains("Año", case=False).any(), axis=1)].index[0]
        nsa = _serie_mensual(df_raw, idx, 3)
        sa, origen = _sa_ine(base)
        if sa is None:
            print(f"  ⚠️ {base}: no pude leer la serie desestacionalizada del INE ({origen}); uso STL.")
            sa, origen = _stl(nsa), "STL (no se encontró la serie desestacionalizada del INE)"
        FUENTE_SA[base] = origen
        series.append(pd.DataFrame({base: nsa, f"{base} (SA)": sa}))
    if not series:
        return pd.DataFrame(columns=["Periodo"])
    df = reduce(lambda a, b: a.join(b, how="outer"), series)
    return df.sort_index().rename_axis("Periodo").reset_index()


HOJAS = {
    "ajuste_estacional": ajuste_estacional,
    "horas": horas,
    "remuneraciones": remuneraciones,
}


def cargar(descarga=True, hojas=None):
    """Descarga (opcional) y devuelve {hoja: DataFrame}. Una hoja que falla se avisa y se omite."""
    hojas = hojas or list(HOJAS)
    if descarga:
        archivos = [a for a in URLS if any(a.startswith(h) for h in hojas)
                    or (a.startswith("tabulado_") and "remuneraciones" in hojas)]
        descargar(archivos)
    datos = {}
    for hoja in hojas:
        try:
            datos[hoja] = HOJAS[hoja]()
            print(f"  ✅ {hoja}: {len(datos[hoja])} meses")
        except FileNotFoundError as e:
            print(f"  ❌ {hoja}: falta {Path(e.filename).name} en Inputs")
        except Exception as e:  # noqa: BLE001
            print(f"  ❌ {hoja}: {type(e).__name__}: {e}")
    return datos


def guardar_excel(datos, ruta, formato_fecha="mmm-yyyy"):
    """Escribe cada DataFrame en su hoja, con encabezado en negrita y fechas mmm-aaaa."""
    from openpyxl.styles import Alignment, Font, PatternFill

    ruta.parent.mkdir(parents=True, exist_ok=True)
    try:
        with pd.ExcelWriter(ruta, engine="openpyxl") as xl:
            for hoja, df in datos.items():
                df.to_excel(xl, sheet_name=hoja[:31], index=False)
                ws = xl.sheets[hoja[:31]]
                for celda in ws[1]:
                    celda.font = Font(bold=True)
                    celda.fill = PatternFill("solid", start_color="CCEAFF")
                    celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ws.freeze_panes = "B2"
                ws.column_dimensions["A"].width = 12
                for fila in ws.iter_rows(min_row=2, max_col=1):
                    fila[0].number_format = formato_fecha
    except PermissionError:
        sys.exit(f"No pude guardar {ruta}: está abierto. Ciérralo y vuelve a correr.")
    return ruta


# ─────────────────────────────────────────────────────────────────────────────
# Masa salarial real
# ─────────────────────────────────────────────────────────────────────────────

def _columna(df, *palabras, excluir=("tasa", "variación", "variacion"), respaldo=None, hoja=""):
    """Primera columna cuyo nombre contiene todas las 'palabras' (sin mayúsculas) y ninguna de 'excluir'."""
    for c in df.columns:
        n = c.lower()
        if all(p in n for p in palabras) and not any(e in n for e in excluir):
            return c
    if respaldo is not None and respaldo < len(df.columns):
        c = df.columns[respaldo]
        print(f"  ⚠️ En '{hoja}' no encontré {palabras}; uso la columna {respaldo}: '{c}'")
        return c
    raise KeyError(f"En '{hoja}' no encontré una columna con {palabras}. Columnas: {list(df.columns)}")


def _indice(s, anio=ANIO_BASE):
    return 100 * s / s[s.index.year == anio].mean()


def _aa(s):
    return 100 * (s / s.shift(12) - 1)


def masa_salarial(datos):
    """DataFrame mensual (índice = Periodo) con componentes, masa salarial, variación a/a y contribuciones."""
    ae = datos["ajuste_estacional"].set_index("Periodo")
    hr = datos["horas"].set_index("Periodo")
    rem = datos["remuneraciones"].set_index("Periodo")
    # la columna 5 es la que usaba el notebook (usecols=[0, 5] de la hoja ajuste_estacional)
    col_ocu = _columna(ae, "ocupad", "ajustad", excluir=("tasa", "no ajust"), hoja="ajuste_estacional")
    ocupados = pd.to_numeric(ae[col_ocu], errors="coerce")
    if ocupados.median() > 1e5:   # ajuste_estacional.xlsx viene en personas; se pasa a miles
        ocupados = ocupados / 1000
    FUENTE_SA["Ocupados"] = "INE (ajuste_estacional.xlsx, cifras ajustadas)"
    # horas: el INE no publica horas desestacionalizadas; STL sobre horas.xlsx
    horas_sa, FUENTE_SA["Horas"] = hr[COL_HORAS_SA], "STL sobre horas.xlsx (el INE no publica horas desestacionalizadas)"
    col_ir = "IR Real (SA)"
    print(f"  Ocupados: '{col_ocu}'\n  Horas SA: {FUENTE_SA['Horas']}\n"
          f"  IR real SA: {FUENTE_SA.get('IR Real', '?')}")

    df = pd.concat({OCUPADOS: ocupados, HORAS: horas_sa, MS_IR_REAL: rem[col_ir]}, axis=1, join="inner")
    df = df.apply(pd.to_numeric, errors="coerce").dropna().sort_index()
    df.index.name = "Periodo"
    df[MS_IR_REAL] = _indice(df[MS_IR_REAL])

    df[MASA] = _indice(df[OCUPADOS] * df[HORAS] * df[MS_IR_REAL])
    df[MASA_AA] = _aa(df[MASA])
    # contribuciones a la variación a/a (descomposición logarítmica: suman exactamente la variación)
    log_total = np.log(df[MASA] / df[MASA].shift(12))
    escala = (df[MASA_AA] / log_total).where(log_total.abs() > 1e-12, 100.0)
    for comp, nombre in CONTRIB.items():
        df[nombre] = np.log(df[comp] / df[comp].shift(12)) * escala
    df[MASA_HS] = _indice(df[HORAS] * df[MS_IR_REAL])
    df[MASA_HS_AA] = _aa(df[MASA_HS])
    return df


def guardar_base(df, ruta=EXCEL, desde=DESDE_GRAFICO):
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.chart.series import SeriesLabel
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    tabla = df.reset_index()
    leeme = pd.DataFrame({"Masa salarial real": [
        "Masa salarial real = ocupados × horas habituales semanales × índice de remuneraciones real.",
        "Ocupados: ENE, cifras ajustadas estacionalmente por el INE (ajuste_estacional.xlsx), en miles.",
        f"Horas: promedio de horas habitualmente trabajadas a la semana; desestacionalización: {FUENTE_SA.get('Horas')}.",
        f"IR real: índice de remuneraciones real empalmado (base anual 2023 = 100); desestacionalización: "
        f"{FUENTE_SA.get('IR Real')}.",
        "Periodo: último mes del trimestre móvil de la ENE (el IR es mensual y se cruza con ese mes).",
        f"Índices: promedio {ANIO_BASE} = 100. Variación a/a: respecto del mismo mes del año anterior.",
        "Contribuciones: descomposición logarítmica de la variación a/a; las tres suman la variación total.",
        "Horas × salario real: remuneración real por persona ocupada ajustada por jornada (sin el efecto del empleo).",
        FUENTE,
        "Se reescribe en cada corrida con minuta_remuneraciones.py (proyecto remuneraciones del repositorio).",
    ]})
    try:
        with pd.ExcelWriter(ruta, engine="openpyxl") as xl:
            tabla.to_excel(xl, sheet_name="Masa salarial", index=False)
            leeme.to_excel(xl, sheet_name="Léeme", index=False)
            ws = xl.sheets["Masa salarial"]
            for celda in ws[1]:
                celda.font = Font(bold=True)
                celda.fill = PatternFill("solid", start_color="CCEAFF")
                celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.row_dimensions[1].height = 45
            ws.freeze_panes = "B2"
            ws.column_dimensions["A"].width = 11
            for j in range(2, ws.max_column + 1):
                ws.column_dimensions[get_column_letter(j)].width = 15
            formatos = {OCUPADOS: "#,##0.0", HORAS: "0.00"}
            for j, col in enumerate(tabla.columns, start=1):
                fmt = "mmm-yyyy" if j == 1 else formatos.get(col, "0.0")
                for (c,) in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                    c.number_format = fmt
            xl.sheets["Léeme"].column_dimensions["A"].width = 110

            # gráficos de Excel (nativos), desde DESDE_GRAFICO hasta el último mes
            gs = xl.book.create_sheet("Gráficos", 1)
            ini = 2 + int((tabla["Periodo"] < pd.Timestamp(desde)).sum())
            fin = len(tabla) + 1
            col = {c: j for j, c in enumerate(tabla.columns, start=1)}
            fechas = Reference(ws, min_col=1, min_row=ini, max_row=fin)

            barras = BarChart()
            barras.type, barras.grouping, barras.overlap, barras.gapWidth = "col", "stacked", 100, 30
            for nombre, color in zip(CONTRIB.values(), (ROJO, GRIS, GRIS_CLARO)):
                barras.add_data(Reference(ws, min_col=col[nombre], min_row=ini, max_row=fin),
                                titles_from_data=False)
                s = barras.series[-1]
                s.tx = SeriesLabel(v=nombre.replace(" (pp)", ""))
                s.graphicalProperties = GraphicalProperties(solidFill=color)
                s.graphicalProperties.line.noFill = True
            linea = LineChart()
            linea.add_data(Reference(ws, min_col=col[MASA_AA], min_row=ini, max_row=fin), titles_from_data=False)
            linea.series[0].tx = SeriesLabel(v="Masa salarial real")
            linea.series[0].graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=NEGRO, w=22000))
            linea.series[0].smooth = False
            barras.set_categories(fechas)
            barras += linea
            barras.title = "Masa salarial real: variación anual y contribuciones"
            barras.y_axis.title = "Var. a/a %, contribuciones en pp"
            barras.x_axis.number_format = "yyyy"
            barras.x_axis.tickLblSkip = 12
            barras.x_axis.tickMarkSkip = 12
            barras.x_axis.tickLblPos = "low"
            barras.x_axis.delete = barras.y_axis.delete = False
            barras.legend.position = "b"
            barras.height, barras.width = 10, 24
            gs.add_chart(barras, "A1")

            nivel = LineChart()
            for nombre, color, ancho in ((MASA, ROJO, 25000), (MASA_HS, GRIS, 19000)):
                nivel.add_data(Reference(ws, min_col=col[nombre], min_row=ini, max_row=fin), titles_from_data=False)
                s = nivel.series[-1]
                s.tx = SeriesLabel(v=nombre)
                s.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=color, w=ancho))
                s.smooth = False
            nivel.set_categories(fechas)
            tramo = tabla.iloc[ini - 2:][[MASA, MASA_HS]]
            nivel.y_axis.scaling.min = 10 * np.floor(np.nanmin(tramo.to_numpy()) / 10)
            nivel.y_axis.scaling.max = 10 * np.ceil(np.nanmax(tramo.to_numpy()) / 10)
            nivel.title = "Masa salarial real: nivel"
            nivel.y_axis.title = f"Índice {ANIO_BASE} = 100, desestacionalizado"
            nivel.x_axis.number_format = "yyyy"
            nivel.x_axis.tickLblSkip = 12
            nivel.x_axis.tickMarkSkip = 12
            nivel.x_axis.delete = nivel.y_axis.delete = False
            nivel.legend.position = "b"
            nivel.height, nivel.width = 10, 24
            gs.add_chart(nivel, "A22")
            gs["A43"] = FUENTE
    except PermissionError:
        sys.exit(f"No pude guardar {ruta}: está abierto. Ciérralo y vuelve a correr.")
    return ruta


# ═════════════════════════════════════════════════════════════════════════════
# 2. MINUTA
# ═════════════════════════════════════════════════════════════════════════════

FIGURAS = INPUTS / "Figuras"              # imágenes de las figuras (documento y presentación)
IRL_BASE = OUT / "IRL_Base.xlsx"             # datos de la minuta (solo se escriben sus columnas)
MANUAL_ESTILO = Path("Administrativos") / "Manual de Estilo"
BIBLIOTECA = Path("Estudios - Documentos") / "General"
ICLOUD = Path.home() / "Library" / "Mobile Documents" / "com~apple~CloudDocs"

TITULO = "Remuneraciones"
AUTOR = "Gerencia de Estudios"   # aparece en el documento, la portada y el pie de las láminas
ANIO_PROM = 2010                 # inicio de los promedios históricos del mes
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre",
         "octubre", "noviembre", "diciembre"]
MESES_CORTOS = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
IR, ICL, IR_REAL = "IR", "ICL", "IR Real"
NOMBRES = {IR: "IR nominal", ICL: "ICL nominal", IR_REAL: "IR real"}

# colores Santander (formato_graficos.PALETA): rojo, negro, turquesa, amarillo, gris
C_ROJO, C_NEGRO, C_TURQUESA, C_AMARILLO, C_GRIS = "#EC0000", "#000000", "#4DC4CA", "#FFAE42", "#B2B2B2"
C_BANDA, C_GRILLA, C_TEXTO = "#DADADA", "#E5E5E5", "#595959"


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
            print(f"  ⚠️ formato_graficos.py está en iCloud pero no descargado en este Mac:\n    {carpeta}\n"
                  "     Ábrelo en Finder y elige 'Descargar ahora'.")
            return None
    return None


def _importar_formato():
    """formato_graficos (Manual de Estilo) o None si no se encuentra (entonces no se arma la presentación)."""
    candidatos = [Path(os.environ["FORMATO_GRAFICOS_DIR"])] if os.environ.get("FORMATO_GRAFICOS_DIR") else []
    for p in [CODE, *CODE.parents]:
        candidatos += [p / MANUAL_ESTILO, p / "formato"]
    carpeta = next((c for c in candidatos if (c / "formato_graficos.py").is_file()), None)
    if carpeta is None:
        carpeta = next((c for c in map(_formato_en, _raices_sincronizadas()) if c), None)
    if carpeta is None:
        return None
    sys.path.insert(0, str(carpeta))
    import formato_graficos
    print(f"  Formato: {carpeta / 'formato_graficos.py'}")
    return formato_graficos


# ─────────────────────────────────────────────────────────────────────────────
# utilidades de texto
# ─────────────────────────────────────────────────────────────────────────────

def num(x, dec=1):
    """1.234,5 (coma decimal, punto de miles)."""
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def signo(x, dec=1):
    x = round(x, dec) + 0.0   # evita "-0,0"
    return f"{'+' if x > 0 else ''}{num(x, dec)}"


def _pct(x, dec=1):
    return f"{num(round(x, dec) + 0.0, dec)}%"


def _comparar(a, b, dec=1):
    """'por encima del' / 'por debajo del' / 'en línea con el' (a frente a b, redondeados)."""
    a, b = round(a, dec), round(b, dec)
    return "por encima del" if a > b else "por debajo del" if a < b else "en línea con el"


def _movimiento(x, alza="un aumento"):
    if round(x, 1) > 0:
        return f"{alza} de {_pct(x)}"
    return f"una caída de {_pct(abs(x))}" if round(x, 1) < 0 else "una variación nula"


def _racha(s):
    """Meses consecutivos al final de la serie con el mismo signo que el último dato."""
    s = s.dropna()
    ultimo = np.sign(s.iloc[-1])
    n = 0
    for v in s.iloc[::-1]:
        if np.sign(v) != ultimo or v == 0:
            break
        n += 1
    return n


def _del_mes(s, mes, hasta):
    """Valores de la serie en el mes calendario 'mes', de ANIO_PROM al año 'hasta'."""
    s = s.dropna()
    return s[(s.index.month == mes) & (s.index.year >= ANIO_PROM) & (s.index.year <= hasta)]


# ─────────────────────────────────────────────────────────────────────────────
# figuras (matplotlib; las mismas imágenes van al documento y a la presentación)
# ─────────────────────────────────────────────────────────────────────────────

def _plt():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    return plt


def _estilo_ejes(ax, tam=7):
    from matplotlib.ticker import FuncFormatter, MaxNLocator
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.spines["bottom"].set_color(C_GRIS)
    ax.grid(axis="y", color=C_GRILLA, lw=0.6)
    ax.set_axisbelow(True)
    ax.tick_params(labelsize=tam, length=0, colors="#333333")
    ax.yaxis.set_major_locator(MaxNLocator(nbins=6, steps=[1, 2, 5, 10]))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: num(v, 1)))


def _eje_anios(ax, fechas):
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.set_xlim(fechas.min() - pd.Timedelta(days=30), fechas.max() + pd.Timedelta(days=45))


def _leyenda(ax, ncol, tam=6.5):
    ax.legend(loc="lower center", bbox_to_anchor=(0.5, 1.0), ncol=ncol, frameon=False, fontsize=tam,
              handlelength=1.6, columnspacing=1.0)


def _guardar(fig, ruta):
    ruta.parent.mkdir(parents=True, exist_ok=True)
    fig.tight_layout(pad=0.4)
    fig.savefig(ruta, format="png", dpi=200)
    _plt().close(fig)
    return ruta


TAM_FIG = (3.45, 2.1)


def _figura_lineas(df, colores, ruta, leyenda=True, promedios=None):
    """Líneas; 'promedios' {columna: valor} agrega líneas horizontales punteadas del mismo color."""
    plt = _plt()
    fig, ax = plt.subplots(figsize=TAM_FIG)
    for col, color in zip(df.columns, colores):
        ax.plot(df.index, df[col], color=color, lw=1.4, label=col)
    if promedios:
        for (col, valor), color in zip(promedios.items(), colores):
            ax.axhline(valor, color=color, lw=0.9, ls=(0, (4, 2)))
        ax.plot([], [], color=C_TEXTO, lw=0.9, ls=(0, (4, 2)), label="Prom. histórico")
    _estilo_ejes(ax)
    _eje_anios(ax, df.index)
    if df.min().min() < 0 < df.max().max():
        ax.axhline(0, color=C_GRIS, lw=0.8)
    if leyenda:
        _leyenda(ax, len(df.columns) + bool(promedios))
    return _guardar(fig, ruta)


def _figura_contribuciones(total, barras, colores, ruta, promedio=None):
    """Columnas apiladas (contribuciones), línea del total, su promedio histórico y etiqueta del último dato."""
    plt = _plt()
    fig, ax = plt.subplots(figsize=TAM_FIG)
    pos = np.zeros(len(barras))
    neg = np.zeros(len(barras))
    for col, color in zip(barras.columns, colores):
        v = barras[col].to_numpy()
        ax.bar(barras.index, v, bottom=np.where(v >= 0, pos, neg), width=25, color=color, label=col, lw=0)
        pos, neg = pos + np.clip(v, 0, None), neg + np.clip(v, None, 0)
    ax.plot(total.index, total, color=C_NEGRO, lw=1.3, label=total.name)
    if promedio is not None:
        ax.axhline(promedio, color=C_NEGRO, lw=0.9, ls=(0, (4, 2)), label="Prom. histórico")
    ax.annotate(num(total.iloc[-1]), (total.index[-1], total.iloc[-1]), xytext=(4, 0), textcoords="offset points",
                va="center", fontsize=6.5, fontweight="bold")
    ax.axhline(0, color=C_GRIS, lw=0.8)
    _estilo_ejes(ax)
    _eje_anios(ax, total.index)
    _leyenda(ax, 3)
    return _guardar(fig, ruta)


def _tabla_estacional(mm, t):
    """Variación m/m por mes calendario: mín, máx y promedio de los años de referencia (ANIO_PROM a tres
    años antes), los dos años anteriores y el año en curso."""
    tabla = mm.dropna().to_frame("v").assign(anio=lambda d: d.index.year, mes=lambda d: d.index.month)
    tabla = tabla.pivot(index="mes", columns="anio", values="v").reindex(range(1, 13))
    completos = [a for a in tabla.columns if tabla[a].notna().sum() == 12]
    ref = [a for a in completos if ANIO_PROM <= a <= t.year - 3] or [a for a in completos if a <= t.year - 3]
    a0, a1 = min(ref), max(ref)
    datos = pd.DataFrame({f"Mín {a0}-{a1}": tabla[ref].min(axis=1), f"Máx {a0}-{a1}": tabla[ref].max(axis=1),
                          f"Prom. {a0}-{a1}": tabla[ref].mean(axis=1)})
    for a in (t.year - 2, t.year - 1, t.year):
        if a in tabla:
            datos[str(a)] = tabla[a]
    datos.index = pd.Index(MESES_CORTOS, name="Mes")
    return datos, (a0, a1)


def _figura_estacional(datos, t, rango, ruta):
    plt = _plt()
    a0, a1 = rango
    x = np.arange(12)
    fig, ax = plt.subplots(figsize=TAM_FIG)
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
    _leyenda(ax, 5, tam=6)
    return _guardar(fig, ruta)


# ─────────────────────────────────────────────────────────────────────────────
# IRL_Base.xlsx: escribe solo en sus columnas (conserva el resto del libro)
# ─────────────────────────────────────────────────────────────────────────────

def _valor_celda(v):
    if v is None or (isinstance(v, float) and np.isnan(v)) or v is pd.NaT:
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if isinstance(v, (np.floating, np.integer)):
        return None if np.isnan(v) else float(v)
    return v


def _filas(df):
    """Encabezado + filas (índice como primera columna)."""
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
        for i in range(len(filas) + 1, previo + 1):   # filas que sobran de la corrida anterior (solo sus columnas)
            for j in range(1, n_cols + 1):
                ws.cell(i, j).value = None

    def guardar(self):
        self.wb.save(self.ruta)

    def cerrar(self):
        pass


class _LibroExcel:
    """Con Excel (xlwings): conserva todo lo que Excel guarda (gráficos, formatos, tablas)."""

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
    """Si el archivo lo guardó Excel y tiene gráficos o tablas, se edita con Excel (xlwings); si no hay
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
                      f"simplificar gráficos o formatos. Respaldo: {respaldo}")
    return _LibroOpenpyxl(ruta)


def escribir_irl_base(hojas, ruta=IRL_BASE):
    """hojas: {nombre: DataFrame}. Cada DataFrame va desde A1 (índice en la columna A); solo se tocan esas
    columnas y filas (las sobrantes de la corrida anterior se vacían en esas mismas columnas)."""
    ruta.parent.mkdir(parents=True, exist_ok=True)
    libro = _abrir_libro(ruta)
    try:
        for hoja, df in hojas.items():
            d = df.reset_index()
            libro.escribir(hoja[:31], _filas(df), [_formato_columna(d[c]) for c in d.columns])
        libro.guardar()
    except PermissionError:
        sys.exit(f"No pude guardar {ruta}: está abierto. Ciérralo y vuelve a correr.")
    finally:
        libro.cerrar()
    return ruta


# ─────────────────────────────────────────────────────────────────────────────
# Minuta
# ─────────────────────────────────────────────────────────────────────────────

class Minuta:
    def __init__(self, datos=None, hoy=None):
        if datos is None:
            datos = cargar(descarga=False)
        faltan = [h for h in ("ajuste_estacional", "horas", "remuneraciones") if h not in datos]
        if faltan:
            sys.exit(f"Faltan datos del INE ({', '.join(faltan)}). Revisa los mensajes de descarga de arriba.")
        self.rem = datos["remuneraciones"].set_index("Periodo").sort_index()
        self.nsa = {b: self.rem[b] for b in NOMBRES}
        self.sa = {b: self.rem[f"{b} (SA)"] for b in NOMBRES}
        self.mm_sa = {b: 100 * (s / s.shift(1) - 1) for b, s in self.sa.items()}
        self.mm_nsa = {b: 100 * (s / s.shift(1) - 1) for b, s in self.nsa.items()}
        self.aa = {b: 100 * (s / s.shift(12) - 1) for b, s in self.nsa.items()}
        self.ms = masa_salarial(datos)
        self.tm = self.ms.dropna(subset=[MASA_AA]).index.max()
        self.t = self.nsa[IR].dropna().index.max()
        self.ant = self.t - pd.DateOffset(months=1)
        self.mes, self.mes_ant = MESES[self.t.month - 1], MESES[self.ant.month - 1]
        self.hoy = pd.Timestamp(hoy or dt.date.today())
        self.ine_sa = all(FUENTE_SA.get(b, "").startswith("INE") for b in NOMBRES)
        self.nota_sa = ("sa = serie desestacionalizada por el INE; nsa = serie original"
                        if self.ine_sa else
                        "sa = serie desestacionalizada (STL; no se encontró la del INE); nsa = serie original")
        self.textos()
        self.cuadro()
        self.figuras()

    def v(self, s, f=None):
        return s.get(self.t if f is None else f, np.nan)

    def prom_mes(self, s, f=None):
        f = self.t if f is None else f
        return _del_mes(s, f.month, f.year).mean()

    # ---- textos
    def normalidad(self, b):
        """Variación m/m de la serie original frente a su promedio histórico del mes y su rango."""
        s, t = self.mm_nsa[b], self.t
        prom = self.prom_mes(s)
        previos = _del_mes(s, t.month, t.year - 1)
        x = self.v(s)
        texto = (f"en la serie original (nsa) varió {_pct(x)} m/m, {_comparar(x, prom)} promedio "
                 f"{ANIO_PROM}-{t.year} para {self.mes} ({_pct(prom)})")
        if len(previos) >= 3:
            lo, hi = previos.min(), previos.max()
            dentro = round(lo, 1) <= round(x, 1) <= round(hi, 1)
            texto += (f" y {'dentro' if dentro else 'fuera'} de su rango {previos.index.year.min()}-"
                      f"{previos.index.year.max()} ({_pct(lo)} a {_pct(hi)})")
        return texto

    def textos(self):
        t, v = self.t, self.v
        r_mm, r_aa = v(self.mm_sa[IR_REAL]), v(self.aa[IR_REAL])
        verbo = "subieron" if round(r_mm, 1) > 0 else "cayeron" if round(r_mm, 1) < 0 else "no variaron"
        self.titulo = (f"Remuneraciones {self.mes} {t.year} | En {self.mes}, los salarios reales {verbo}"
                       f"{' ' + _pct(abs(r_mm)) if verbo != 'no variaron' else ''} m/m (sa) y "
                       f"{'crecen' if r_aa >= 0 else 'caen'} {_pct(abs(r_aa))} en doce meses.")
        ir, ir_ant = v(self.mm_sa[IR]), v(self.mm_sa[IR], self.ant)
        aa_ir = v(self.aa[IR])
        b1 = (f"<b>El Índice de Remuneraciones nominal (IR) registró {_movimiento(ir)} m/m (sa), "
              f"{_comparar(ir, ir_ant).replace('del', 'de lo').replace('con el', 'con lo')} observado en "
              f"{self.mes_ant} ({_pct(ir_ant)}).</b> En términos anuales, el índice muestra una "
              f"{'expansión' if aa_ir >= 0 else 'contracción'} de {_pct(abs(aa_ir))}. El IR "
              f"{self.normalidad(IR)}. Por su parte, el Índice de Costos Laborales (ICL) varió "
              f"{_pct(v(self.mm_sa[ICL]))} m/m (sa) y {_pct(v(self.aa[ICL]))} a/a; {self.normalidad(ICL)}.")
        inflacion = 100 * ((1 + aa_ir / 100) / (1 + r_aa / 100) - 1)
        racha = _racha(self.aa[IR_REAL][self.aa[IR_REAL].index <= t])
        b2 = (f"<b>Los salarios reales registraron {_movimiento(r_mm, 'un alza')} m/m (sa) en {self.mes}.</b> "
              f"El IR real {self.normalidad(IR_REAL)}. En términos anuales, considerando una inflación de "
              f"{_pct(inflacion)} (IPC implícito en el IR real del INE), los salarios reales se "
              f"{'expandieron' if r_aa >= 0 else 'contrajeron'} {_pct(abs(r_aa))}, con lo que acumulan {racha} "
              f"{'mes consecutivo' if racha == 1 else 'meses consecutivos'} de "
              f"{'alzas' if r_aa >= 0 else 'caídas'} anuales.")
        ms, tm = self.ms, self.tm
        b3 = (f"<b>La masa salarial real varió {signo(ms.loc[tm, MASA_AA])}% a/a</b> en el trimestre móvil "
              f"terminado en {MESES[tm.month - 1]} (ocupados × horas habituales × salario real; promedio "
              f"histórico {num(ms[MASA_AA].mean())}%): ocupados {signo(ms.loc[tm, CONTRIB[OCUPADOS]])} pp, "
              f"horas {signo(ms.loc[tm, CONTRIB[HORAS]])} pp y salario real "
              f"{signo(ms.loc[tm, CONTRIB[MS_IR_REAL]])} pp. Las horas × salario real variaron "
              f"{signo(ms.loc[tm, MASA_HS_AA])}% a/a.")
        self.bullets = [b1, b2, b3]

    # ---- Cuadro 1
    def cuadro(self):
        t, ant = self.t, self.ant
        etq = {f: f"{MESES_CORTOS[f.month - 1]}-{f:%y}" for f in (ant, t)}
        cols = [("m/m (sa)", self.mm_sa), ("m/m (nsa)", self.mm_nsa), ("a/a", self.aa)]
        self.cuadro_cols = [f"{n} {etq[f]}" for n, _ in cols for f in (ant, t)]
        filas = {}
        for b, nombre in NOMBRES.items():
            filas[nombre] = [self.v(d[b], f) for _, d in cols for f in (ant, t)]
            filas[f"Promedio {ANIO_PROM}-{t.year} del mes ({nombre})"] = [self.prom_mes(d[b], f)
                                                                          for _, d in cols for f in (ant, t)]
        self.cuadro_df = pd.DataFrame(filas, index=self.cuadro_cols).T.rename_axis("Índice")

    def cuadro_filas(self):
        """Filas de texto del cuadro (para el documento y la presentación)."""
        filas = [["Índice"] + [c.replace(" ", "\n", 1) for c in self.cuadro_cols]]
        for nombre, valores in self.cuadro_df.iterrows():
            etiqueta = f"Promedio {ANIO_PROM}-{self.t.year}" if nombre.startswith("Promedio") else nombre
            filas.append([etiqueta] + [_pct(x, 2) if pd.notna(x) else "" for x in valores])
        return filas

    # ---- figuras
    def figuras(self):
        t, ms = self.t, self.ms
        desde_nivel = pd.Timestamp(year=t.year - 5, month=1, day=1)
        desde_aa = pd.Timestamp(year=t.year - 7, month=1, day=1)
        f1 = self.sa[IR_REAL][lambda s: s.index >= desde_nivel].dropna().rename("IR real (sa)").to_frame()
        aa = pd.DataFrame({"IR": self.aa[IR], "ICL": self.aa[ICL], "IR real": self.aa[IR_REAL]})
        prom_aa = aa.mean().to_dict()
        f2 = aa[aa.index >= desde_aa].dropna(how="all")
        dm = ms[ms.index >= desde_aa].dropna(subset=[MASA_AA])
        contrib = {v: v.replace("Contribución ", "").replace(" (pp)", "").capitalize() for v in CONTRIB.values()}
        f7 = pd.concat([dm[MASA_AA].rename("Masa salarial real"), dm[list(contrib)].rename(columns=contrib)],
                       axis=1)
        f8 = dm[[MASA, MASA_HS]].rename(columns={MASA: "Ocupados × horas × salario real",
                                                        MASA_HS: "Horas × salario real"})
        fuente_masa = "Gerencia de Estudios en base a INE"
        est = {}
        for b in (IR, IR_REAL):
            for tipo, mm in (("sa", self.mm_sa), ("nsa", self.mm_nsa)):
                est[(b, tipo)] = _tabla_estacional(mm[b], t)
        r = FIGURAS
        # (número, título, unidad, ruta png, fuente, datos)
        self.figs = [
            (1, "Índice de Remuneraciones real", f"(serie desestacionalizada, {ANIO_BASE}=100)",
             _figura_lineas(f1, [C_ROJO], r / "Figura_1_IR_real.png", leyenda=False), "INE", f1),
            (2, "Variación anual de los índices", "(%, serie original; punteadas: promedio histórico)",
             _figura_lineas(f2, [C_ROJO, C_NEGRO, C_TURQUESA], r / "Figura_2_variacion_anual.png",
                            promedios=prom_aa), "INE",
             f2.assign(**{f"Prom. {k}": v for k, v in prom_aa.items()})),
        ]
        self.figs += [
            (4, "Masa salarial real: nivel", f"(índice {ANIO_BASE}=100, series desestacionalizadas)",
             _figura_lineas(f8, [C_ROJO, C_NEGRO], r / "Figura_4_masa_salarial_nivel.png"), fuente_masa + ". Masa salarial = ocupados × horas habituales × IR real", f8),
            (3, "Masa salarial real: variación anual", "(var. % a/a y contribuciones en pp, series desestacionalizadas)",
             _figura_contribuciones(f7.iloc[:, 0], f7.iloc[:, 1:], [C_ROJO, C_GRIS, C_TURQUESA],
                                    r / "Figura_3_masa_salarial_aa.png", promedio=ms[MASA_AA].mean()),
             fuente_masa,
             f7.assign(**{"Prom. histórico": ms[MASA_AA].mean()})),
        ]
        n = 5
        for b, nombre in ((IR, "nominal"), (IR_REAL, "real")):
            for tipo, desc in (("sa", "serie desestacionalizada"), ("nsa", "serie original")):
                datos, rango = est[(b, tipo)]
                png = _figura_estacional(datos, t, rango, r / f"Figura_{n}_IR_{nombre}_mm_{tipo}.png")
                self.figs.append((n, f"Índice de Remuneraciones {nombre} ({tipo})", f"(var. % m/m, {desc})",
                                  png, "INE", datos))
                n += 1

    # ================================================================ documento PDF
    def documento(self):
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_JUSTIFY
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import (HRFlowable, Image, KeepTogether, ListFlowable, ListItem, PageBreak,
                                        Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle)

        rojo, gris = colors.HexColor(C_ROJO), colors.HexColor(C_TEXTO)
        st = {
            "titulo": ParagraphStyle("titulo", fontName="Helvetica-Bold", fontSize=15, leading=19, textColor=rojo),
            "fecha": ParagraphStyle("fecha", fontName="Helvetica-Bold", fontSize=10.5, leading=14, spaceBefore=8),
            "autor": ParagraphStyle("autor", fontName="Helvetica-Oblique", fontSize=9, leading=12, textColor=gris),
            "h1": ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=14, leading=18, textColor=rojo,
                                 spaceBefore=6, spaceAfter=4),
            "texto": ParagraphStyle("texto", fontName="Helvetica", fontSize=9.5, leading=13, alignment=TA_JUSTIFY),
            "cuadro": ParagraphStyle("cuadro", fontName="Helvetica", fontSize=9, leading=12, alignment=1),
            "nota": ParagraphStyle("nota", fontName="Helvetica", fontSize=7.5, leading=10, textColor=gris),
            "fig": ParagraphStyle("fig", fontName="Helvetica", fontSize=8.5, leading=10.5),
            "unidad": ParagraphStyle("unidad", fontName="Helvetica", fontSize=7, leading=9, textColor=gris),
            "fuente": ParagraphStyle("fuente", fontName="Helvetica", fontSize=6.5, leading=8, textColor=gris),
        }
        fecha = f"{self.hoy.day} de {MESES[self.hoy.month - 1]} de {self.hoy.year}"

        filas = self.cuadro_filas()
        cuadro = Table(filas, colWidths=[3.6 * cm] + [2.4 * cm] * 6, repeatRows=1)
        estilo = [
            ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8), ("FONT", (0, 1), (-1, -1), "Helvetica", 8.5),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("BACKGROUND", (0, 0), (-1, 0), rojo),
            ("ALIGN", (1, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LINEBELOW", (0, -1), (-1, -1), 0.8, rojo),
            ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]
        for i in range(1, len(filas)):
            if filas[i][0].startswith("Promedio"):
                estilo += [("FONT", (0, i), (-1, i), "Helvetica-Oblique", 7.5),
                           ("TEXTCOLOR", (0, i), (-1, i), gris), ("LEFTPADDING", (0, i), (0, i), 14),
                           ("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F2F2F2")),
                           ("LINEBELOW", (0, i), (-1, i), 0.4, colors.HexColor(C_GRIS))]
        cuadro.setStyle(TableStyle(estilo))

        w, h = 8.6 * cm, 5.25 * cm

        def bloque(n, tit, uni, png, fuente, _):
            return [Paragraph(f"Figura {n}: {tit}", st["fig"]), Paragraph(uni, st["unidad"]),
                    Image(str(png), width=w, height=h), Paragraph(f"Fuente: {fuente}.", st["fuente"])]

        celdas = [bloque(*f) for f in self.figs]
        grilla = Table([celdas[i:i + 2] for i in range(0, len(celdas), 2)], colWidths=[9.0 * cm, 9.0 * cm])
        grilla.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                    ("BOTTOMPADDING", (0, 0), (-1, -1), 10)]))

        bullets = ListFlowable([ListItem(Paragraph(b, st["texto"]), value="■", bulletColor=rojo, bulletFontSize=7)
                                for b in self.bullets], bulletType="bullet", start="■", leftIndent=14,
                               bulletOffsetY=-1)
        historia = [
            Paragraph(self.titulo, st["titulo"]), Paragraph(fecha, st["fecha"]),
            Paragraph(f"Autor: {AUTOR}", st["autor"]), Spacer(1, 4),
            HRFlowable(width="100%", thickness=1.2, color=rojo, spaceAfter=8),
            Paragraph("Principales resultados", st["h1"]), bullets, Spacer(1, 12),
            KeepTogether([Paragraph("Cuadro 1: Variaciones de los índices de remuneraciones y costos laborales",
                                    st["cuadro"]), Spacer(1, 4), cuadro, Spacer(1, 4),
                          Paragraph(f"Nota: {self.nota_sa}. Promedio: variación promedio del mismo mes calendario, "
                                    f"{ANIO_PROM}-{self.t.year}. El último dato es provisorio. Fuente: INE.",
                                    st["nota"])]),
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

        ruta = OUT / f"Minuta_Remuneraciones_{self.mes}_{self.t.year}.pdf"
        doc = SimpleDocTemplate(str(ruta), pagesize=letter, leftMargin=2 * cm, rightMargin=2 * cm,
                                topMargin=2.2 * cm, bottomMargin=2.2 * cm,
                                title=f"Remuneraciones {self.mes} {self.t.year}", author=AUTOR)
        try:
            doc.build(historia, onFirstPage=pie, onLaterPages=pie)
        except PermissionError:
            sys.exit(f"No pude guardar {ruta}: está abierto. Ciérralo y vuelve a correr.")
        return ruta

    # ================================================================ presentación
    def presentacion(self):
        """Mismos textos, cuadro y figuras del documento, con formato_graficos.py."""
        fg = _importar_formato()
        if fg is None:
            print("  ⚠️ No encontré formato_graficos.py (Manual de Estilo): no se genera la presentación.\n"
                  "     Para indicar la carpeta: variable FORMATO_GRAFICOS_DIR.")
            return None
        from pptx.dml.color import RGBColor
        from pptx.oxml.ns import qn
        from pptx.util import Cm, Pt

        prs = fg.nueva_presentacion()
        n = [1]
        pie_texto = f"{TITULO} {self.mes} {self.t.year}  |  {AUTOR}"
        gris_txt = getattr(fg, "GRIS_TEXTO", (89, 89, 89))

        def lamina(titulo):
            n[0] += 1
            try:
                return fg.lamina(prs, titulo, n[0], pie_texto=pie_texto)
            except TypeError:   # versión anterior de formato_graficos, sin pie de lámina
                return fg.lamina(prs, titulo, n[0])

        fecha = f"{self.hoy.day:02d} de {MESES[self.hoy.month - 1]} de {self.hoy.year}"
        try:
            fg.portada(prs, "Remuneraciones y\nmasa salarial", f"{self.mes.capitalize()} {self.t.year}",
                       autor=AUTOR, fecha=fecha)
        except TypeError:   # versión anterior: portada(prs, fecha_texto, titulo)
            fg.portada(prs, f"{self.mes.capitalize()} {self.t.year}\n{fecha}", titulo="Remuneraciones")

        # principales resultados (mismos textos del documento, sin las marcas <b>)
        lam = lamina("Principales resultados")
        caja = lam.shapes.add_textbox(Cm(1.3), Cm(2.6), Cm(31), Cm(14.5))
        tf = caja.text_frame
        tf.word_wrap = True
        for i, texto in enumerate(self.bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(9)
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", str(Cm(0.6)))
            pPr.set("indent", str(-Cm(0.6)))
            clr = pPr.makeelement(qn("a:buClr"), {})
            clr.append(clr.makeelement(qn("a:srgbClr"), {"val": "EC0000"}))
            pPr.append(clr)
            pPr.append(pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"}))
            pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "■"}))
            for parte, negrita in _partes_negrita(texto):
                run = p.add_run()
                run.text = parte
                run.font.bold = negrita
                run.font.name = fg.tipografia("fuente")
                run.font.size = Pt(14)
                run.font.color.rgb = RGBColor(*gris_txt)

        # cuadro 1 (tabla nativa)
        lam = lamina("Variaciones de los índices de remuneraciones y costos laborales")
        filas = self.cuadro_filas()
        tabla = lam.shapes.add_table(len(filas), len(filas[0]), Cm(1.3), Cm(3.0), Cm(31),
                                     Cm(1.0) * len(filas)).table
        tabla.columns[0].width = Cm(7.0)
        for j in range(1, len(filas[0])):
            tabla.columns[j].width = Cm(4.0)
        for i, fila in enumerate(filas):
            prom = fila[0].startswith("Promedio")
            for j, texto in enumerate(fila):
                celda = tabla.cell(i, j)
                celda.text = texto
                celda.fill.solid()
                celda.fill.fore_color.rgb = RGBColor(0xEC, 0, 0) if i == 0 else (
                    RGBColor(0xF2, 0xF2, 0xF2) if prom else RGBColor(0xFF, 0xFF, 0xFF))
                for p in celda.text_frame.paragraphs:
                    p.alignment = 2 if j else 1   # centro / izquierda
                    for run in p.runs:
                        run.font.size = Pt(13 if i == 0 else 12 if not prom else 11)
                        run.font.bold = i == 0
                        run.font.italic = prom
                        run.font.name = fg.tipografia("fuente")
                        run.font.color.rgb = RGBColor(255, 255, 255) if i == 0 else (
                            RGBColor(*gris_txt) if prom else RGBColor(0, 0, 0))
        fg.pie(lam, Cm(1.3), Cm(3.0) + Cm(1.0) * len(filas) + Cm(0.4), Cm(31), "INE",
               f"{self.nota_sa}. Promedio: variación promedio del mismo mes calendario, {ANIO_PROM}-{self.t.year}.",
               h=Cm(1.4))

        # figuras: dos por lámina, las mismas imágenes del documento
        titulos = {1: "Índice de Remuneraciones", 4: "Masa salarial real",
                   5: "Índice de Remuneraciones nominal: variación mensual",
                   7: "Índice de Remuneraciones real: variación mensual"}
        for k in range(0, len(self.figs), 2):
            par = self.figs[k:k + 2]
            lam = lamina(titulos.get(par[0][0], "Figuras"))
            for (num_fig, tit, uni, png, fuente, _), x in zip(par, (Cm(1.3), Cm(17.4))):
                fg._caja_texto(lam, f"Figura {num_fig}: {tit}", x, Cm(2.5), Cm(15), Cm(0.9),
                               fg.tipografia("titulo"), 14, negrita=True)
                fg._caja_texto(lam, uni, x, Cm(3.35), Cm(15), Cm(0.7), fg.tipografia("fuente"), 10,
                               color=gris_txt)
                lam.shapes.add_picture(str(png), x, Cm(4.1), width=Cm(15))
                fg.pie(lam, x, Cm(13.4), Cm(15), fuente, h=Cm(1.2))

        pptx = OUT / f"{self.hoy:%Y%m%d}_Minuta_Remuneraciones.pptx"
        try:
            prs.save(pptx)
        except PermissionError:
            sys.exit(f"No pude guardar {pptx}: está abierto. Ciérralo y vuelve a correr.")
        print(f"  ✅ Presentación: {pptx} ({n[0]} láminas)")
        try:
            pdf = fg.exportar_pdf(pptx)
            if pdf:
                print(f"  ✅ PDF de la presentación: {pdf}")
        except Exception as e:  # noqa: BLE001
            print(f"  ⚠️ No pude exportar el PDF de la presentación ({e}); el PPTX quedó listo.")
        return pptx

    # ================================================================ IRL_Base
    def hojas_irl(self):
        """Hojas de IRL_Base.xlsx: 'Datos' (toda la historia) y una por cuadro/figura. Columnas fijas."""
        datos = pd.DataFrame(index=self.rem.index)
        for b, nombre in NOMBRES.items():
            datos[f"{nombre} (nsa)"] = self.nsa[b]
            datos[f"{nombre} (sa)"] = self.sa[b]
        for b, nombre in NOMBRES.items():
            datos[f"{nombre} m/m nsa (%)"] = self.mm_nsa[b]
            datos[f"{nombre} m/m sa (%)"] = self.mm_sa[b]
            datos[f"{nombre} a/a (%)"] = self.aa[b]
        masa = self.ms[[OCUPADOS, HORAS, MASA, MASA_AA, *CONTRIB.values(), MASA_HS,
                        MASA_HS_AA]]
        datos = datos.join(masa, how="outer").rename_axis("Periodo")
        hojas = {"Datos": datos, "Cuadro 1": self.cuadro_df}
        for num_fig, tit, _, _, _, df in self.figs:
            hojas[f"Figura {num_fig}"] = df.rename_axis(df.index.name or "Periodo")
        leeme = pd.DataFrame({"Contenido": [
            "Hojas y columnas escritas por minuta_remuneraciones.py en cada corrida (solo esas columnas).",
            "Puedes agregar fórmulas o gráficos en otras columnas u hojas: se conservan.",
            "Datos: series originales (nsa), desestacionalizadas (sa), variaciones y masa salarial, toda la historia.",
            "Desestacionalización: " + "; ".join(f"{k}: {v}" for k, v in FUENTE_SA.items()),
            "Cuadro 1 y Figura N: datos de cada cuadro y figura de la minuta.",
        ]}, index=pd.RangeIndex(1, 6, name="#"))
        hojas["Léeme"] = leeme
        return hojas

    # ---- armar y guardar
    def construir(self):
        OUT.mkdir(parents=True, exist_ok=True)
        print(f"  ✅ Figuras: {FIGURAS}")
        print(f"  ✅ Documento: {self.documento()}")
        self.presentacion()
        print(f"  ✅ IRL_Base: {escribir_irl_base(self.hojas_irl())}")


def _partes_negrita(texto):
    """'<b>a</b> b' -> [('a', True), (' b', False)]."""
    partes = []
    for i, trozo in enumerate(re.split(r"</?b>", texto)):
        if trozo:
            partes.append((trozo, i % 2 == 1))
    return partes


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--sin-descarga", action="store_true", help="usa los archivos del INE que ya están en Inputs")
    args, _ = p.parse_known_args()   # parse_known_args: Spyder agrega sus propios argumentos
    print("=" * 60 + "\n  Minuta de Remuneraciones\n" + "=" * 60)
    print("  1) Datos del INE")
    datos = cargar(descarga=not args.sin_descarga)
    faltan = [h for h in ("ajuste_estacional", "horas", "remuneraciones") if h not in datos]
    if faltan:
        sys.exit(f"Faltan datos del INE ({', '.join(faltan)}). Revisa los mensajes de descarga de arriba.")
    print(f"  ✅ Cifras de empleo: {guardar_excel(datos, EXCEL_CIFRAS)}")
    print(f"  ✅ Bases: {guardar_base(masa_salarial(datos))}")
    print("  2) Minuta")
    Minuta(datos).construir()
