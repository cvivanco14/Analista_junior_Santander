# -*- coding: utf-8 -*-
r"""
Encuestas del Banco Central (EEE y EOF) para el Chartpack de Expectativas
=========================================================================
Un solo script que:
1. Descarga las planillas de resultados de la EEE y la EOF que falten
   (Inputs\eee\raw\ e Inputs\eof\raw\).
2. EEE: extrae mediana, decil 1 y decil 9 de inflación en 11 y 23 meses
   -> Inputs\eee\eee_inflacion_deciles.xlsx
3. EOF: % de operadores con inflación a 2 años (13 a 24 meses) sobre 3,0%, sobre 3,2%
   y bajo 2,8%, por encuesta y promedio mensual
   -> Inputs\eof\eof_inflacion_2a_umbrales.xlsx (hojas "mensual" y "por_encuesta")

Ejecutar (no hay que editar nada):
    python encuestas_bcch.py                 # descarga lo que falte y actualiza EEE y EOF
    python encuestas_bcch.py --sin-descarga  # solo reprocesa las planillas ya descargadas
    python encuestas_bcch.py --solo eee      # o --solo eof

Carpetas (relativas a este script, sirven para cualquier usuario):
    <Chartpack de Expectativas>\Code\     este script y build_chartpack.py
    <Chartpack de Expectativas>\Inputs\   eee\ y eof\ (planillas raw y Excel resultante)
    <Chartpack de Expectativas>\Outputs\  chartpacks y datos finales

Requisitos (una vez):
    python -m pip install pandas openpyxl xlrd playwright
    python -m playwright install chromium
El sitio del Banco Central está detrás de Incapsula, que bloquea descargas con requests;
por eso las planillas se bajan con un navegador (Chromium vía Playwright).
"""
import argparse
import os
import re
import time
import urllib.parse
import warnings
from pathlib import Path

import pandas as pd

CODE = Path(__file__).resolve().parent
BASE = CODE.parent
INPUTS = BASE / "Inputs"
EEE_DIR, EOF_DIR = INPUTS / "eee", INPUTS / "eof"
EEE_XLSX = EEE_DIR / "eee_inflacion_deciles.xlsx"
EOF_XLSX = EOF_DIR / "eof_inflacion_2a_umbrales.xlsx"

MESES = {m: i for i, m in enumerate(
    ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
     "septiembre", "octubre", "noviembre", "diciembre"], start=1)}


# ─────────────────────────────────────────────────────────────────────────────
# 1. DESCARGA
# ─────────────────────────────────────────────────────────────────────────────

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36")

PAGINAS = {
    # resultado012014.xls, resultado_012021.xlsx, "Datos mensuales enero 2026.xlsx" y, desde
    # febrero de 2026, nombres sin extension: datos-mensuales-octubre-2026-xlsx
    "eee": ("https://www.bcentral.cl/encuestas-economicas/resultados-anteriores-eee",
            re.compile(r"(?i)^(resultado_?\d{5,6}\.xlsx?|datos[ _+-]mensuales[ _+-]\w+[ _+-]\d{4}"
                       r"(\.xlsx?|-xlsx?))$")),
    # EOFResultado20100115.xls, EOFRESULTADOS_201801q1.xlsx, EOF_RESULTADOS_Pre_RPM_201905.xlsx,
    # "EOF RESULTADOS Post RPM Enero 2026.xlsx" (y la misma variante sin extension, por si cambia)
    "eof": ("https://www.bcentral.cl/encuestas-economicas/resultados-anteriores-eof",
            re.compile(r"(?i)^eof.*(\.xls[xm]?|-xls[xm]?)$")),
}


def _nombre_archivo(href):
    return urllib.parse.unquote_plus(urllib.parse.urlparse(href).path.split("/")[4])


def _nombre_local(nombre):
    """Nombre con que se guarda la planilla: misma convencion para nombres viejos y nuevos."""
    m = re.match(r"(?i)^datos[ _+-]mensuales[ _+-](\w+)[ _+-](\d{4})", nombre)
    if m:
        return f"Datos mensuales {m.group(1).lower()} {m.group(2)}.xlsx"
    m = re.match(r"(?i)^(.*)-(xls[xm]?)$", nombre)   # "eof-...-xlsx" -> "eof ... .xlsx"
    if m:
        return m.group(1).replace("-", " ") + "." + m.group(2).lower()
    return nombre


def _anio(nombre):
    m = re.search(r"(20\d\d)", nombre)
    return int(m.group(1)) if m else 0


def descargar(encuesta, desde):
    """Baja a Inputs/<encuesta>/raw/ las planillas que falten; devuelve cuantas bajo.

    Corre en un hilo aparte: la API sincronica de Playwright no funciona dentro del loop de
    asyncio que tienen las consolas de Spyder, Jupyter e IPython.
    """
    from concurrent.futures import ThreadPoolExecutor
    with ThreadPoolExecutor(max_workers=1) as hilo:
        return hilo.submit(_descargar, encuesta, desde).result()


def _descargar(encuesta, desde):
    from playwright.sync_api import sync_playwright

    url, patron = PAGINAS[encuesta]
    raw = INPUTS / encuesta / "raw"
    raw.mkdir(parents=True, exist_ok=True)
    chromium = os.environ.get("CHROMIUM_PATH", "/opt/pw-browsers/chromium")
    nuevas = 0
    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path=chromium if Path(chromium).exists() else None,
                                    args=["--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(accept_downloads=True, user_agent=UA)
        page = ctx.new_page()
        page.goto(url, wait_until="networkidle", timeout=90_000)

        hrefs = page.eval_on_selector_all("a[href*='/documents/']", "els => els.map(e => e.href)")
        planillas = {}
        for h in hrefs:
            n = _nombre_archivo(h)
            if patron.match(n) and _anio(n) >= desde:
                planillas[_nombre_local(n)] = h
        faltan = {n: h for n, h in planillas.items() if not (raw / n).exists()}
        print(f"  {encuesta.upper()}: {len(planillas)} planillas en la página, {len(faltan)} por descargar")

        for n, h in sorted(faltan.items()):
            destino = raw / n
            try:
                with page.expect_download(timeout=60_000) as d:
                    page.evaluate("""(u) => { const a = document.createElement('a');
                        a.href = u; a.download = ''; document.body.appendChild(a); a.click(); a.remove(); }""", h)
                d.value.save_as(destino)
                if destino.stat().st_size < 10_000:  # pagina de bloqueo, no un Excel
                    destino.unlink()
                    print(f"    bloqueado: {n}")
                else:
                    nuevas += 1
                    print(f"    ok: {n}")
            except Exception as e:  # noqa: BLE001 - se informa y se sigue con la siguiente
                print(f"    error: {n} {e}")
            time.sleep(1.5)
        browser.close()
    return nuevas


def _leer_excel(path):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return pd.read_excel(path, sheet_name=None, header=None)


# ─────────────────────────────────────────────────────────────────────────────
# 2. EEE: mediana y deciles de inflación en 11 y 23 meses
# ─────────────────────────────────────────────────────────────────────────────

EEE_DESDE = "2014-01-01"
HORIZONTES = {"11m": r"en 11 meses", "23m": r"en 23 meses"}


def eee_fecha(path):
    """Mes de la encuesta a partir del nombre del archivo."""
    n = path.stem.lower()
    m = re.search(r"datos[ _+-]mensuales[ _+-](\w+)[ _+-](\d{4})", n)
    if m:
        return pd.Timestamp(int(m.group(2)), MESES[m.group(1)], 1)
    m = re.search(r"resultado_?(\d{1,2})(\d{4})$", n)
    if m:
        return pd.Timestamp(int(m.group(2)), int(m.group(1)), 1)
    return None


def eee_leer_planilla(path):
    """Devuelve {horizonte: (mediana, decil1, decil9)} de la hoja resumen."""
    res = {}
    # El resumen suele ser la primera hoja, pero hay archivos con hojas vacias antes.
    df = next((h for h in _leer_excel(path).values() if not h.empty), pd.DataFrame())
    for _, fila in df.iterrows():
        celdas = list(fila)
        for i, c in enumerate(celdas):
            if not isinstance(c, str):
                continue
            etiqueta = c.replace("\xa0", " ").lower()
            for h, patron in HORIZONTES.items():
                # La primera aparicion corresponde a la seccion "Inflacion" (IPC total).
                if h in res or not re.search(patron, etiqueta):
                    continue
                nums = [v for v in celdas[i + 1:] if isinstance(v, (int, float)) and not pd.isna(v)]
                if len(nums) >= 3:
                    res[h] = tuple(float(v) for v in nums[:3])
    return res


def eee_serie():
    filas = {}
    for path in sorted((EEE_DIR / "raw").glob("*.xls*")):
        fecha = eee_fecha(path)
        if fecha is None:
            continue
        datos = eee_leer_planilla(path)
        for h, (med, d1, d9) in datos.items():
            if not d1 <= med <= d9:
                print(f"  aviso: orden inesperado en {path.name} {h}: {med}, {d1}, {d9}")
            filas[(fecha, h)] = {"fecha": fecha, "horizonte": h, "mediana": med,
                                 "decil_1": d1, "decil_9": d9, "archivo": path.name}
        if len(datos) < len(HORIZONTES):
            print(f"  aviso: {path.name} sin datos para {set(HORIZONTES) - set(datos)}")
    serie = pd.DataFrame(filas.values(), columns=["fecha", "horizonte", "mediana", "decil_1", "decil_9", "archivo"])
    serie = serie[serie["fecha"] >= EEE_DESDE]
    # meses que ya estaban en el Excel y cuya planilla no esta en raw/ se conservan
    if EEE_XLSX.exists():
        previo = pd.read_excel(EEE_XLSX, sheet_name="deciles", parse_dates=["fecha"])
        previo = previo[["fecha", "horizonte", "mediana", "decil_1", "decil_9", "archivo"]]
        clave = set(zip(serie["fecha"], serie["horizonte"]))
        previo = previo[[k not in clave for k in zip(previo["fecha"], previo["horizonte"])]]
        serie = pd.concat([previo, serie], ignore_index=True)
    serie = serie.sort_values(["horizonte", "fecha"]).reset_index(drop=True)
    serie[["mediana", "decil_1", "decil_9"]] = serie[["mediana", "decil_1", "decil_9"]].round(2)
    serie["rango_9_1"] = (serie["decil_9"] - serie["decil_1"]).round(2)
    return serie


def actualizar_eee():
    serie = eee_serie()
    with pd.ExcelWriter(EEE_XLSX, datetime_format="yyyy-mm") as xl:
        serie.to_excel(xl, sheet_name="deciles", index=False)
    ult = serie["fecha"].max()
    print(f"  EEE: {len(serie)} filas, último mes {ult:%Y-%m} -> {EEE_XLSX.name}")


# ─────────────────────────────────────────────────────────────────────────────
# 3. EOF: operadores con inflación a 2 años sobre/bajo umbrales
# ─────────────────────────────────────────────────────────────────────────────

# "Siguientes 12 meses (13 al 24)" desde 2018; "IPC (13 a 24 Meses)" antes.
PREGUNTA = re.compile(r"13\s*al?\s*24", re.I)

UMBRALES = [  # (columna, funcion sobre el nivel en %)
    ("sobre_3_0", lambda x: x > 3.0),
    ("sobre_3_2", lambda x: x > 3.2),
    ("bajo_2_8", lambda x: x < 2.8),
]


def _es_numero(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not pd.isna(v)


def _texto(v):
    return str(v).replace("\xa0", " ").strip().lower() if isinstance(v, str) else ""


def eof_tipo(nombre):
    n = nombre.lower()
    if "post" in n:
        return "Post RPM"
    if "pre" in n:
        return "Pre RPM"
    return "Quincenal"


def eof_fecha_nombre(nombre):
    n = nombre.lower().replace("_", " ")
    m = re.search(r"(\d{4})(\d{2})q([12])", n)  # EOFRESULTADOS_201601q1
    if m:
        return pd.Timestamp(int(m.group(1)), int(m.group(2)), 8 if m.group(3) == "1" else 22)
    m = re.search(r"rpm (\d{4})(\d{2})\b", n)  # EOF_RESULTADOS_Pre_RPM_201905
    if m:
        return pd.Timestamp(int(m.group(1)), int(m.group(2)), 15)
    m = re.search(r"(" + "|".join(MESES) + r") (\d{4})", n)  # ... Pre RPM Octubre 2026
    if m:
        return pd.Timestamp(int(m.group(2)), MESES[m.group(1)], 15)
    m = re.search(r"(\d{2})(\d{2})(\d{4})", n)  # EOFRESULTADOS22122015
    if m:
        return pd.Timestamp(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    return None


def eof_fecha(hojas, nombre):
    """Fecha de cierre: la ultima fecha del encabezado de la hoja resumen."""
    resumen = next(iter(hojas.values()))
    fechas = [v for v in resumen.iloc[:7].to_numpy().ravel() if isinstance(v, (pd.Timestamp,))
              or hasattr(v, "year") and hasattr(v, "month") and not isinstance(v, str)]
    fechas = [pd.Timestamp(f) for f in fechas if 2000 < pd.Timestamp(f).year < 2100]
    return max(fechas) if fechas else eof_fecha_nombre(nombre)


def eof_distribucion(df):
    """Devuelve [(nivel en %, n de respuestas)] de la pregunta 13 al 24 meses."""
    for (r, c), v in df.stack().items():
        if PREGUNTA.search(_texto(v)):
            break
    else:
        return []
    if _texto(df.iat[r + 1, c]).startswith("nivel"):
        # Layout por pares de columnas: Nivel | N R bajo cada pregunta.
        col_nivel, col_n, fila = c, c + 1, r + 2
    else:
        # Layout con una columna "Nivel" compartida a la izquierda de cada bloque.
        col_nivel = max(j for j in range(c) if _texto(df.iat[r, j]).startswith("nivel"))
        col_n, fila = c, r + 1
    dist = []
    for i in range(fila, len(df)):
        if any("resp" in _texto(x) for x in df.iloc[i, :col_nivel + 1]):
            break
        nivel, n = df.iat[i, col_nivel], df.iat[i, col_n]
        if _es_numero(nivel) and _es_numero(n) and n > 0:
            dist.append((round(nivel * 100, 3), int(n)))
    return dist


def eof_serie():
    filas = []
    for path in sorted((EOF_DIR / "raw").glob("*.xls*")):
        hojas = _leer_excel(path)
        hoja = next((h for nombre, h in hojas.items() if "distrib" in nombre.lower()), None)
        dist = eof_distribucion(hoja) if hoja is not None else []
        if not dist:
            print(f"  aviso: {path.name} sin distribución 13-24 meses")
            continue
        fecha = eof_fecha(hojas, path.name)
        fila = {"fecha": fecha, "mes": fecha.to_period("M").to_timestamp(), "tipo": eof_tipo(path.name),
                "n": sum(n for _, n in dist), "archivo": path.name}
        for col, cond in UMBRALES:
            fila[col] = sum(n for nivel, n in dist if cond(nivel))
        filas.append(fila)
    serie = (pd.DataFrame(filas).sort_values("fecha")
             .drop_duplicates(["fecha", "tipo"], keep="last").reset_index(drop=True))
    for col, _ in UMBRALES:
        serie[f"pct_{col}"] = (100 * serie[col] / serie["n"]).round(1)
    return serie


def eof_mensual(serie):
    """Promedio simple de las encuestas que cerraron cada mes; meses sin encuesta quedan vacios."""
    cols = [f"pct_{col}" for col, _ in UMBRALES]
    mensual = serie.groupby("mes").agg(n_encuestas=("archivo", "size"), **{c: (c, "mean") for c in cols})
    meses = pd.date_range(mensual.index.min(), mensual.index.max(), freq="MS", name="mes")
    mensual = mensual.reindex(meses)
    mensual["n_encuestas"] = mensual["n_encuestas"].fillna(0).astype(int)
    return mensual.round(1).reset_index()


def actualizar_eof():
    serie = eof_serie()
    mensual = eof_mensual(serie)
    with pd.ExcelWriter(EOF_XLSX, datetime_format="yyyy-mm-dd") as xl:
        mensual.to_excel(xl, sheet_name="mensual", index=False)
        serie.to_excel(xl, sheet_name="por_encuesta", index=False)
    print(f"  EOF: {len(serie)} encuestas en {len(mensual)} meses, último mes "
          f"{mensual['mes'].max():%Y-%m} -> {EOF_XLSX.name}")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="Descarga y procesa la EEE y la EOF del Banco Central.")
    ap.add_argument("--sin-descarga", action="store_true", help="no descargar; solo reprocesar raw/")
    ap.add_argument("--solo", choices=["eee", "eof"], help="procesar solo una encuesta")
    ap.add_argument("--desde", type=int, default=2013, help="primer año a descargar")
    args = ap.parse_args()
    encuestas = [args.solo] if args.solo else ["eee", "eof"]

    for e in encuestas:
        (INPUTS / e / "raw").mkdir(parents=True, exist_ok=True)
    if not args.sin_descarga:
        print(">> Descargando planillas")
        for e in encuestas:
            descargar(e, args.desde)
    print(">> Procesando")
    if "eee" in encuestas:
        actualizar_eee()
    if "eof" in encuestas:
        actualizar_eof()


if __name__ == "__main__":
    main()
