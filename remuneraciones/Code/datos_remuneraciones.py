# -*- coding: utf-8 -*-
r"""Datos de la Minuta de Remuneraciones (INE: Índice de Remuneraciones y Encuesta Nacional de Empleo).

1. Descarga del INE a Inputs\ solo lo que usa la minuta:
     - IR, ICL e IR real: series originales (tabulados empalmados) y desestacionalizadas por el INE
       (tabulados empalmados desestacionalizados; se usa la hoja de la publicación más reciente);
     - ajuste_estacional.xlsx: personas ocupadas desestacionalizadas por el INE;
     - horas.xlsx: horas habitualmente trabajadas (el INE no las publica desestacionalizadas: se
       desestacionalizan con STL).
2. Calcula la masa salarial real:
       ocupados (sa INE) × horas habituales semanales (sa STL) × IR real (sa INE)
   con su variación a/a, las contribuciones de cada componente y horas × salario real.
3. Deja en Inputs\ (bases intermedias que usa minuta_remuneraciones.py):
       Masa_salarial_Base.xlsx   hojas 'Masa salarial' (series), 'Gráficos' (gráficos de Excel) y 'Léeme'
       Cifras_empleo.xlsx        hojas ajuste_estacional, horas y remuneraciones

Carpetas (relativas a este script):
    <Remuneraciones>\Code\      datos_remuneraciones.py y minuta_remuneraciones.py
    <Remuneraciones>\Inputs\    archivos descargados del INE, Masa_salarial_Base.xlsx y Cifras_empleo.xlsx
    <Remuneraciones>\Outputs\   la minuta e IRL_Base.xlsx (los deja minuta_remuneraciones.py)

Uso (no hay que editar nada):
    python datos_remuneraciones.py                 # descarga y actualiza las bases en Inputs
    python datos_remuneraciones.py --sin-descarga  # usa los archivos que ya están en Inputs
"""
import argparse
import re
import sys
from functools import reduce
from pathlib import Path

import numpy as np
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from statsmodels.tsa.seasonal import STL
from urllib3.util.retry import Retry

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
IR_REAL = f"IR real SA ({ANIO_BASE}=100)"
MASA = f"Masa salarial real ({ANIO_BASE}=100)"
MASA_AA = "Masa salarial real a/a (%)"
CONTRIB = {OCUPADOS: "Contribución ocupados (pp)", HORAS: "Contribución horas (pp)",
           IR_REAL: "Contribución salario real (pp)"}
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

    df = pd.concat({OCUPADOS: ocupados, HORAS: horas_sa, IR_REAL: rem[col_ir]}, axis=1, join="inner")
    df = df.apply(pd.to_numeric, errors="coerce").dropna().sort_index()
    df.index.name = "Periodo"
    df[IR_REAL] = _indice(df[IR_REAL])

    df[MASA] = _indice(df[OCUPADOS] * df[HORAS] * df[IR_REAL])
    df[MASA_AA] = _aa(df[MASA])
    # contribuciones a la variación a/a (descomposición logarítmica: suman exactamente la variación)
    log_total = np.log(df[MASA] / df[MASA].shift(12))
    escala = (df[MASA_AA] / log_total).where(log_total.abs() > 1e-12, 100.0)
    for comp, nombre in CONTRIB.items():
        df[nombre] = np.log(df[comp] / df[comp].shift(12)) * escala
    df[MASA_HS] = _indice(df[HORAS] * df[IR_REAL])
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
        "Se reescribe en cada corrida con datos_remuneraciones.py (proyecto remuneraciones del repositorio).",
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


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--sin-descarga", action="store_true", help="usa los archivos que ya están en Inputs")
    args = p.parse_args()
    print("=" * 60 + "\n  Datos Minuta de Remuneraciones (INE)\n" + "=" * 60)
    datos = cargar(descarga=not args.sin_descarga)
    print(f"  ✅ Cifras de empleo: {guardar_excel(datos, EXCEL_CIFRAS)}")
    faltan = [h for h in ("ajuste_estacional", "horas", "remuneraciones") if h not in datos]
    if faltan:
        sys.exit(f"Faltan datos para la masa salarial: {', '.join(faltan)}")
    ms = masa_salarial(datos)
    print(f"  Último dato: {ms.index[-1]:%m-%Y}  masa salarial real {ms[MASA_AA].iloc[-1]:+.1f}% a/a")
    print(f"  ✅ Bases: {guardar_base(ms)}")
