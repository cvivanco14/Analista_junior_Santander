# -*- coding: utf-8 -*-
r"""Cifras de empleo del INE (ENE e Índice de Remuneraciones).

1. Descarga los cuadros estadísticos del INE a Inputs\ (indicadores principales, categoría,
   informalidad, comparación ENE-RRAA, ajuste estacional, horas, rama e índices de remuneraciones).
2. Los ordena como series mensuales (columna 'Periodo' = último mes del trimestre móvil) y
   desestacionaliza con STL las horas y los índices de remuneraciones.
3. Deja todas las hojas en Outputs\Cifras_empleo.xlsx.

Viene del notebook "Series empleo3.ipynb" (Masa_salarial.py): mismas descargas y mismos cálculos, con
rutas relativas al proyecto (sirve para cualquier usuario, en Windows y macOS) y sin xlwings.

Carpetas:
    <Masa salarial>\Code\      datos_empleo.py y masa_salarial.py
    <Masa salarial>\Inputs\    archivos descargados del INE
    <Masa salarial>\Outputs\   Cifras_empleo.xlsx y Masa_salarial_Base.xlsx

Uso:
    python datos_empleo.py                 # descarga y deja Outputs\Cifras_empleo.xlsx
    python datos_empleo.py --sin-descarga  # usa los archivos que ya están en Inputs
"""
import argparse
import re
import sys
from functools import reduce
from pathlib import Path

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from statsmodels.tsa.seasonal import STL
from urllib3.util.retry import Retry

CODE = Path(__file__).resolve().parent
PROYECTO = CODE.parent
INPUTS = PROYECTO / "Inputs"
OUT = PROYECTO / "Outputs"
EXCEL_CIFRAS = OUT / "Cifras_empleo.xlsx"

_INE = "https://www.ine.gob.cl/docs/default-source/"
_ENE = _INE + "ocupacion-y-desocupacion/cuadros-estadisticos/series-vigentes/"
_IR = _INE + "sueldos-y-salarios/cuadros-estadisticos/ir-icl-base-anual-2023-100/series-empalmadas/"
URLS = {
    "indicadores_principales.xlsx": _ENE + "indicadores_principales.xlsx?sfvrsn=afbd6d15_167",
    "categoria.xlsx": _ENE + "categoria.xlsx?sfvrsn=70dc54db_175",
    "informalidad_tasas.xlsx": _INE + "informalidad-y-condiciones-laborales/cuadros-estadisticos/"
                                      "proyecciones_censo_2017/informalidad_tasas.xlsx?sfvrsn=cbf364e1_188",
    "comparacion_ene_rraa.xlsx": _ENE + "comparacion_ene_rraa.xlsx?sfvrsn=3213707b_168",
    "ajuste_estacional.xlsx": _ENE + "ajuste_estacional.xlsx?sfvrsn=617886e6_65",
    "horas.xlsx": _ENE + "horas.xlsx?sfvrsn=d294ce0e_171",
    "rama.xlsx": _ENE + "rama.xlsx?sfvrsn=23a76b12_177",
    "tabulado_ir_empalmado.xlsx": _IR + "tabulado_ir_empalmado.xlsx?sfvrsn=9bb59cb4_43",
    "tabulado_icl_empalmado.xlsx": _IR + "tabulado_icl_empalmado.xlsx?sfvrsn=28e08211_45",
    "tabulado_ir_real_empalmado.xlsx": _IR + "tabulado_ir_real_empalmado.xlsx?sfvrsn=6f1d912d_43",
}

COL_HORAS_SA = "Promedio de horas habitualmente trabajadas a la semana (desest.)"
REMUNERACIONES = {   # archivo -> nombre de la serie desestacionalizada
    "tabulado_ir_empalmado.xlsx": "IR (Desest.)",
    "tabulado_icl_empalmado.xlsx": "ICL (Desest.)",
    "tabulado_ir_real_empalmado.xlsx": "IR Real (Desest.)",
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


def indicadores_principales():
    archivo = INPUTS / "indicadores_principales.xlsx"
    df_raw = pd.read_excel(archivo, sheet_name="AS", header=None)
    idx = _fila_anio(df_raw)
    df = df_raw.iloc[idx + 2:].copy()
    df.columns = df_raw.iloc[idx].ffill().tolist()[:len(df.columns)]
    try:   # tasas de participación de hombres y mujeres (columna 27 de las hojas H y M)
        for hoja, nombre in (("H", "Tasa Part. Hombres"), ("M", "Tasa Part. Mujeres")):
            d = pd.read_excel(archivo, sheet_name=hoja, header=None)
            i = _fila_anio(d)
            df[nombre] = d.iloc[i + 2:i + 2 + len(df), 27].to_numpy()
    except Exception as e:  # noqa: BLE001
        print(f"  ⚠️ No pude agregar las tasas de participación por sexo: {e}")
    return _a_series(df)


def categoria():
    df = _tabla_ine(INPUTS / "categoria.xlsx", "AS")
    ocupados = next(c for c in df.columns if "Población ocupada (total)" in c)
    asalariados = next(c for c in df.columns if "Asalariados/as (Total)" in c)
    df["Resto"] = df[ocupados] - df[asalariados]
    return df


def informalidad_tasas():
    return _tabla_ine(INPUTS / "informalidad_tasas.xlsx", "AS")


def ajuste_estacional():
    archivo = INPUTS / "ajuste_estacional.xlsx"
    df_raw = pd.read_excel(archivo, sheet_name="niveles_as", header=None)
    idx = _fila_anio(df_raw)
    df = df_raw.iloc[idx + 1:].copy()
    df.columns = df_raw.iloc[idx].ffill().tolist()[:len(df.columns)]
    df = df.dropna(how="all", subset=df.columns[2:])
    return _a_series(df)


def comparacion_ene_rraa():
    archivo = INPUTS / "comparacion_ene_rraa.xlsx"
    df_raw = pd.read_excel(archivo, sheet_name="sp_ene", header=None)
    idx = _fila_anio(df_raw)
    df = df_raw.iloc[idx + 2:].copy()
    df.columns = df_raw.iloc[idx].ffill().tolist()
    df = _limpiar_columnas(df.iloc[:, [0, 1, 3, 5]])
    df = formatear_fecha_inteligente(df, df.columns[0], df.columns[1])
    return df.drop(columns=df.columns[1:3]).dropna(subset=["Periodo"]).reset_index(drop=True)


def horas():
    df = _tabla_ine(INPUTS / "horas.xlsx", "AS")
    col = next((c for c in df.columns if "habitualmente" in c), None)
    if col is None:
        print("  ⚠️ No encontré la columna de horas habitualmente trabajadas.")
        return df
    df[COL_HORAS_SA] = _stl(df.set_index("Periodo")[col]).to_numpy()
    return df


def rama():
    return _tabla_ine(INPUTS / "rama.xlsx", "AS")


def remuneraciones():
    """IR, ICL e IR real (base anual 2023 = 100), originales y desestacionalizados con STL."""
    series = []
    for nombre, col_sa in REMUNERACIONES.items():
        archivo = INPUTS / nombre
        if not archivo.exists():
            print(f"  ⚠️ No encontré {nombre}")
            continue
        df_raw = pd.read_excel(archivo, sheet_name="General", header=None)
        idx = df_raw[df_raw.apply(lambda x: x.astype(str).str.contains("Año", case=False).any(), axis=1)].index[0]
        df = df_raw.iloc[idx + 1:, [0, 1, 3]].copy()
        df.columns = ["Anio", "Mes", "Valor"]
        df = df.apply(pd.to_numeric, errors="coerce").dropna()
        df["Periodo"] = pd.to_datetime(dict(year=df.Anio, month=df.Mes, day=1))
        valor = df.sort_values("Periodo").set_index("Periodo")["Valor"]
        series.append(pd.DataFrame({col_sa.replace(" (Desest.)", ""): valor, col_sa: _stl(valor)}))
    if not series:
        return pd.DataFrame(columns=["Periodo"])
    df = reduce(lambda a, b: a.join(b, how="outer"), series)
    return df.sort_index().rename_axis("Periodo").reset_index()


HOJAS = {
    "indicadores_principales": indicadores_principales,
    "categoria": categoria,
    "informalidad_tasas": informalidad_tasas,
    "ajuste_estacional": ajuste_estacional,
    "comparacion_ene_rraa": comparacion_ene_rraa,
    "horas": horas,
    "rama": rama,
    "remuneraciones": remuneraciones,
}


def cargar(descarga=True, hojas=None):
    """Descarga (opcional) y devuelve {hoja: DataFrame}. Una hoja que falla se avisa y se omite."""
    hojas = hojas or list(HOJAS)
    if descarga:
        archivos = [a for a in URLS if any(a.startswith(h) for h in hojas)
                    or (a in REMUNERACIONES and "remuneraciones" in hojas)]
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


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--sin-descarga", action="store_true", help="usa los archivos que ya están en Inputs")
    args = p.parse_args()
    print("=" * 60 + "\n  Cifras de empleo INE\n" + "=" * 60)
    print(f"  ✅ Excel: {guardar_excel(cargar(descarga=not args.sin_descarga), EXCEL_CIFRAS)}")
