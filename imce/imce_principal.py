# -*- coding: utf-8 -*-
r"""
IMCE — Proceso mensual completo
================================
1. Descarga datos IMCE desde la API del BCCh y actualiza IMCE.xlsx
2. Genera el chartbook Santander en PPTX (graficos editables) y PDF, con el formato de
   ../formato/formato_graficos.py

Ejecutar (no hay que editar nada):
    python imce_principal.py

Requisitos (una vez):
    python -m pip install bcchapi pandas openpyxl requests python-pptx matplotlib pywin32
  (pywin32 permite que PowerPoint exporte el PDF en Windows)

Credenciales (una vez por computador, en PowerShell, con comillas SIMPLES):
    setx BDE_TOKEN 'tu_token'        # API Key Token del BCCh (tambien acepta BCCH_API_token)
  o, si se usa usuario y clave de la API en vez de token:
    setx BCCH_USER 'tu_correo'
    setx BCCH_PASS 'tu_clave'
  Luego cerrar y abrir VS Code. Las credenciales nunca se escriben en este archivo.

Carpetas (sirven para cualquier usuario: no dependen de C:\Users\<usuario>):
  Este script vive en
    ...\Estudios - Documentos\General\Expectativas y Encuestas\IMCE\Code
  y el formato comun en
    ...\Estudios - Documentos\General\Administrativos\Manual de Estilo\formato_graficos.py
  (con su carpeta assets\). El script sube desde su propia carpeta hasta encontrar
  Administrativos\Manual de Estilo; si no, lo busca en la carpeta sincronizada
  "Santander Office 365" del usuario. Para forzar otra: setx FORMATO_GRAFICOS_DIR 'C:\ruta'

  IMCE.xlsx se busca primero dentro de la carpeta IMCE del script (IMCE\, IMCE\Code\ y sus
  subcarpetas, ej. IMCE\Base\) y luego en <OneDrive>/.../Bases - Chartpack/Bases/Local/IMCE/Base/.
  Para forzar otra ubicacion: setx IMCE_XLSX 'C:\ruta\IMCE.xlsx'
  El chartbook (PPTX y PDF) queda en tmp_chartbook_imce_santander\, junto a IMCE.xlsx.

Formato: este script no define formato propio; usa formato_graficos.py del Manual de Estilo.
"""

import math
import os
import re
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import requests
from openpyxl import load_workbook


# ─────────────────────────────────────────────────────────────────────────────
# 1. CONFIGURACIÓN AUTOMÁTICA (no editar)
# ─────────────────────────────────────────────────────────────────────────────

# Fecha de inicio de descarga
FECHA_INICIO = "2004-01-01"

# Ruta de IMCE.xlsx dentro de la carpeta compartida de Estudios (ubicacion anterior)
RUTA_RELATIVA = Path("Bases - Chartpack") / "Bases" / "Local" / "IMCE" / "Base" / "IMCE.xlsx"

# Biblioteca de Estudios en SharePoint y carpetas fijas dentro de ella
AQUI = Path(__file__).resolve().parent                     # ...\IMCE\Code
BIBLIOTECA = Path("Estudios - Documentos") / "General"
MANUAL_ESTILO = Path("Administrativos") / "Manual de Estilo"
CARPETA_IMCE = Path("Expectativas y Encuestas") / "IMCE"

# nombres alternativos con que pueden estar definidas las variables de entorno
_ALIAS_TOKEN = ["BDE_TOKEN", "BCCH_API_token", "BCCH_API_TOKEN"]


def _credenciales() -> Dict[str, str]:
    """Token del BDE (preferido) o usuario y clave, desde variables de entorno."""
    for nombre in _ALIAS_TOKEN:
        if os.environ.get(nombre):
            return {"token": os.environ[nombre]}
    usuario, clave = os.environ.get("BCCH_USER"), os.environ.get("BCCH_PASS")
    if usuario and clave:
        return {"user": usuario, "pass": clave}
    sys.exit("No encuentro credenciales del BCCh. En PowerShell ejecuta una vez:\n"
             "    setx BDE_TOKEN 'tu_token'\n"
             "y luego cierra y abre VS Code.")


def _raices_onedrive() -> List[Path]:
    """Carpetas raiz de OneDrive / SharePoint sincronizadas en este computador."""
    raices = []
    for var in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        if os.environ.get(var):
            raices.append(Path(os.environ[var]))
    casa = Path.home()
    if casa.is_dir():
        for p in casa.iterdir():
            if p.is_dir() and re.search(r"santander|onedrive", p.name, re.I):
                raices.append(p)
    vistas, unicas = set(), []
    for r in raices:
        clave = str(r).lower()
        if clave not in vistas and r.is_dir():
            vistas.add(clave)
            unicas.append(r)
    return unicas


def encontrar_imce_xlsx() -> Path:
    """Ubica IMCE.xlsx sin depender del nombre de usuario de Windows."""
    if os.environ.get("IMCE_XLSX"):
        p = Path(os.environ["IMCE_XLSX"])
        if p.is_file():
            return p
        sys.exit(f"IMCE_XLSX apunta a un archivo que no existe: {p}")

    # 1) dentro de la carpeta del proyecto IMCE (la carpeta padre de Code\)
    proyecto = AQUI.parent if AQUI.name.lower() == "code" else AQUI
    locales = [AQUI / "IMCE.xlsx", proyecto / "IMCE.xlsx"]
    locales += [sub / "IMCE.xlsx" for sub in proyecto.iterdir() if sub.is_dir()]
    encontrados = [c for c in locales if c.is_file()]
    if encontrados:
        return max(encontrados, key=lambda p: p.stat().st_mtime)

    # 2) carpetas sincronizadas: proyecto IMCE en la biblioteca y ubicacion anterior
    candidatos = []
    raices = _raices_onedrive()
    for raiz in raices:
        imce = raiz / BIBLIOTECA / CARPETA_IMCE
        if imce.is_dir():
            candidatos.append(imce / "IMCE.xlsx")
            candidatos += [sub / "IMCE.xlsx" for sub in imce.iterdir() if sub.is_dir()]
        # <raiz>/Bases - Chartpack/... o <raiz>/<biblioteca, ej. "General - Estudios">/Bases - Chartpack/...
        candidatos.append(raiz / RUTA_RELATIVA)
        candidatos += [sub / RUTA_RELATIVA for sub in raiz.iterdir() if sub.is_dir()]
    encontrados = [c for c in candidatos if c.is_file()]
    if not encontrados:
        buscado = "\n    ".join(str(r) for r in raices) or "(no hay carpetas de OneDrive)"
        sys.exit(f"No encontré IMCE.xlsx en {proyecto} ni en:\n    " + buscado +
                 f"\n(rutas esperadas dentro: {BIBLIOTECA / CARPETA_IMCE} o {RUTA_RELATIVA})\n"
                 "Solución: deja IMCE.xlsx en la carpeta IMCE del proyecto, o en PowerShell:\n"
                 "    setx IMCE_XLSX 'C:\\ruta\\IMCE.xlsx'")
    # si hay mas de uno (ej. dos bibliotecas sincronizadas), el modificado mas recientemente
    return max(encontrados, key=lambda p: p.stat().st_mtime)


# se completan en main()
IMCE_XLSX: Path = Path()
OUT_DIR: Path = Path()


# ─────────────────────────────────────────────────────────────────────────────
# 2. SERIES BCCh — mapeo nombre_columna_excel → código_api
# ─────────────────────────────────────────────────────────────────────────────

COL_TO_SERIES: Dict[str, str] = {
    # IMCE agregados
    "Comercio":        "G089.IME.IND.A1.M",
    "Construcción":    "G089.IME.IND.A3.M",
    "Industria":       "G089.IME.IND.A2.M",
    "Minería":         "G089.IME.IND.A4.M",
    "IMCE Total":      "G089.IME.IND.A0.M",
    "IMCE: Sin Minería": "G089.IME.IND.A04.M",

    # INDUSTRIA
    "IND_Situacion actual - Producción":                        "G089.IME.IND.A21.M",
    "IND_Situacion actual - Inventarios":                       "G089.IME.IND.A22.M",
    "IND_Situacion actual - Demanda":                           "G089.IME.IND.A23.M",
    "IND_Situacion actual - Uso de capacidad":                  "G089.IME.IND.A24.M",
    "IND_Situacion actual - Situación general de la empresa":   "G089.IME.IND.A25.M",
    "IND_Expectativas - Producción":                            "G089.IME.IND.A26.M",
    "IND_Expectativas - Precios":                               "G089.IME.IND.A27.M",
    "IND_Expectativas - Salarios":                              "G089.IME.IND.A28.M",
    "IND_Expectativas - Costos":                                "G089.IME.IND.A29.M",
    "IND_Expectativas - Empleo":                                "G089.IME.IND.A2A.M",
    "IND_Expectativas - Inversiones":                           "G089.IME.IND.A2B.M",
    "IND_Expectativas - Situación económica global del país":   "G089.IME.IND.A2C.M",
    "IND_Expectativas - Situación general del negocio":         "G089.IME.IND.A2D.M",
    "IND_Expectativas - Inflación esperada 12 meses":           "G089.IME.IND.A2E.M",
    "IND_Indice de difusión sectorial":                         "G089.IME.IND.A2.M",

    # COMERCIO
    "COM_Situacion actual - Inventarios":                       "G089.IME.IND.A11.M",
    "COM_Situacion actual - Demanda":                           "G089.IME.IND.A12.M",
    "COM_Situacion actual - Situación general de la empresa":   "G089.IME.IND.A13.M",
    "COM_Expectativas - Precios":                               "G089.IME.IND.A14.M",
    "COM_Expectativas - Costos":                                "G089.IME.IND.A15.M",
    "COM_Expectativas - Empleo":                                "G089.IME.IND.A16.M",
    "COM_Expectativas - Inversiones":                           "G089.IME.IND.A17.M",
    "COM_Expectativas - Situación económica global del país":   "G089.IME.IND.A18.M",
    "COM_Expectativas - Situación general del negocio":         "G089.IME.IND.A19.M",
    "COM_Expectativas  - Ventas":                               "G089.IME.IND.A1A.M",
    "COM_Expectativas - Situación financiera empresa":          "G089.IME.IND.A1B.M",
    "COM_Expectativas - Inflación esperada 12 meses":           "G089.IME.IND.A1C.M",
    "COM_Indice de difusión sectorial":                         "G089.IME.IND.A1.M",

    # CONSTRUCCIÓN
    "CONS_Situacion actual - Demanda":                          "G089.IME.IND.A31.M",
    "CONS_Situacion actual - Nivel de actividad (3 meses)":     "G089.IME.IND.A32.M",
    "CONS_Situacion actual - Situación general de la empresa":  "G089.IME.IND.A33.M",
    "CONS_Expectativas - Precios":                              "G089.IME.IND.A34.M",
    "CONS_Expectativas - Costos":                               "G089.IME.IND.A35.M",
    "CONS_Expectativas - Salarios":                             "G089.IME.IND.A36.M",
    "CONS_Expectativas - Empleo":                               "G089.IME.IND.A37.M",
    "CONS_Expectativas - Situación económica global del país":  "G089.IME.IND.A38.M",
    "CONS_Expectativas - Situación financiera empresa":         "G089.IME.IND.A39.M",
    "CONS_Indice de difusión sectorial":                        "G089.IME.IND.A3.M",
}


# ─────────────────────────────────────────────────────────────────────────────
# 3. API BCCh
# ─────────────────────────────────────────────────────────────────────────────

BDE_URL = "https://si3.bcentral.cl/SieteRestWS/SieteRestWS.ashx"
TIMEOUT = 40
_SIETE = None


def _month_start(d: date) -> date:
    return date(d.year, d.month, 1)


def _to_float(x) -> float:
    if x is None:
        return math.nan
    if isinstance(x, (int, float)) and not (isinstance(x, float) and math.isnan(x)):
        return float(x)
    s = str(x).strip().replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return math.nan


def descargar_serie(series_id: str, cred: Dict[str, str]) -> List[Tuple[date, float]]:
    if "token" in cred:
        # API Key Token: libreria oficial del BCCh
        global _SIETE
        if _SIETE is None:
            import bcchapi
            _SIETE = bcchapi.Siete(token=cred["token"])
        s = _SIETE.cuadro(series=[series_id], desde=FECHA_INICIO).iloc[:, 0]
        return sorted((_month_start(pd.Timestamp(d).date()), _to_float(v)) for d, v in s.items())

    params = {
        "user": cred["user"], "pass": cred["pass"],
        "function": "GetSeries", "timeseries": series_id,
        "firstdate": FECHA_INICIO,
    }
    r = requests.get(BDE_URL, params=params, timeout=TIMEOUT)
    r.raise_for_status()
    data = r.json()

    # Detectar error de credenciales
    if data.get("Codigo", 0) < 0:
        raise RuntimeError(f"BCCh API error {data['Codigo']}: {data.get('Descripcion')} "
                           "— revisa las variables BCCH_USER y BCCH_PASS")

    obs = (data.get("Series") or {}).get("Obs") or []
    out = []
    for o in obs:
        ds = o.get("indexDateString")
        if not ds:
            continue
        d = datetime.strptime(ds, "%d-%m-%Y").date()
        out.append((_month_start(d), _to_float(o.get("value"))))
    return sorted(out)


# ─────────────────────────────────────────────────────────────────────────────
# 4. ACTUALIZAR IMCE.xlsx
# ─────────────────────────────────────────────────────────────────────────────

HEADER_ROW = 3
DATA_START_ROW = 4
PERIODO_HEADER = "Periodo"


def _build_header_to_col(ws) -> Dict[str, int]:
    return {str(ws.cell(HEADER_ROW, c).value).strip(): c
            for c in range(1, ws.max_column + 1)
            if ws.cell(HEADER_ROW, c).value is not None}


def _build_date_to_row(ws, periodo_col: int) -> Dict[date, int]:
    m = {}
    for r in range(DATA_START_ROW, ws.max_row + 1):
        v = ws.cell(r, periodo_col).value
        if v is None:
            break
        try:
            d = v.date() if hasattr(v, "date") else pd.to_datetime(v).date()
            m[_month_start(d)] = r
        except (ValueError, TypeError):
            break
    return m


def _ensure_row(ws, periodo_col, date_to_row, d):
    d = _month_start(d)
    if d in date_to_row:
        return date_to_row[d]
    last_r = max(date_to_row.values()) if date_to_row else DATA_START_ROW - 1
    new_r = last_r + 1
    ws.cell(new_r, periodo_col).value = datetime(d.year, d.month, d.day)
    date_to_row[d] = new_r
    return new_r


def actualizar_excel(cred: Dict[str, str]):
    print(f"  📂 {IMCE_XLSX}")
    try:
        wb = load_workbook(IMCE_XLSX)
    except PermissionError:
        sys.exit("IMCE.xlsx está abierto en Excel o bloqueado por OneDrive. Ciérralo y vuelve a correr.")
    ws = wb["IMCE"] if "IMCE" in wb.sheetnames else wb.active

    h2c = _build_header_to_col(ws)
    if PERIODO_HEADER not in h2c:
        raise ValueError(f"No encuentro 'Periodo' en fila {HEADER_ROW}. Headers: {list(h2c.keys())[:10]}")

    periodo_col = h2c[PERIODO_HEADER]
    date_to_row = _build_date_to_row(ws, periodo_col)

    # Avisar columnas no encontradas en Excel
    missing = [c for c in COL_TO_SERIES if c not in h2c]
    if missing:
        print(f"  ⚠️ Columnas no encontradas en Excel (se ignorarán): {missing}")

    for col_name, series_id in COL_TO_SERIES.items():
        if col_name not in h2c:
            continue
        try:
            obs = descargar_serie(series_id, cred)
        except Exception as e:  # noqa: BLE001 - se informa y se sigue con la siguiente serie
            print(f"  ⚠️ {col_name}: {e}")
            continue

        c_idx = h2c[col_name]
        for d, val in obs:
            rr = _ensure_row(ws, periodo_col, date_to_row, d)
            ws.cell(rr, c_idx).value = None if math.isnan(val) else float(val)

        print(f"  ✅ {col_name}: {len(obs)} obs | último={obs[-1][0] if obs else '-'}")

    try:
        wb.save(IMCE_XLSX)
    except PermissionError:
        sys.exit("No pude guardar IMCE.xlsx: está abierto en Excel. Ciérralo y vuelve a correr.")
    print("\n  💾 IMCE.xlsx guardado")


# ─────────────────────────────────────────────────────────────────────────────
# 5. CHARTBOOK PPTX + PDF (formato de Estudios: ../formato/formato_graficos.py)
# ─────────────────────────────────────────────────────────────────────────────

def _carpeta_formato() -> Path:
    """Carpeta de formato_graficos.py: Administrativos/Manual de Estilo de la biblioteca de Estudios."""
    candidatos = []
    if os.environ.get("FORMATO_GRAFICOS_DIR"):
        candidatos.append(Path(os.environ["FORMATO_GRAFICOS_DIR"]))
    # 1) subiendo desde la carpeta de este script (...\General\Expectativas y Encuestas\IMCE\Code)
    candidatos += [p / MANUAL_ESTILO for p in [AQUI, *AQUI.parents]]
    # 2) en la biblioteca sincronizada del usuario que corre el script
    for raiz in _raices_onedrive():
        candidatos.append(raiz / BIBLIOTECA / MANUAL_ESTILO)
        candidatos += [sub / "General" / MANUAL_ESTILO for sub in raiz.iterdir() if sub.is_dir()]
    # 3) estructura del repositorio de GitHub (formato/ junto a imce/) o junto a este script
    candidatos += [AQUI.parent / "formato", AQUI]
    for carpeta in candidatos:
        if (carpeta / "formato_graficos.py").is_file():
            return carpeta
    sys.exit("No encuentro formato_graficos.py. Debe estar en\n"
             f"    ...\\{BIBLIOTECA}\\{MANUAL_ESTILO}\n"
             "(con su carpeta assets). Para indicar otra carpeta, en PowerShell:\n"
             "    setx FORMATO_GRAFICOS_DIR 'C:\\ruta\\Manual de Estilo'")


def _importar_formato():
    carpeta = _carpeta_formato()
    sys.path.insert(0, str(carpeta))
    import formato_graficos
    print(f"  Formato: {carpeta / 'formato_graficos.py'}")
    return formato_graficos


MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]
FUENTE = "Banco Central de Chile"
NOTA = "Sobre 50 indica optimismo; bajo 50, pesimismo."
MESES_ENTRE_ETIQUETAS = 24

# (titulo de la lamina, [(nombre en la leyenda, [columnas candidatas en IMCE.xlsx])], (min, max) o None)
CHARTS = [
    ("IMCE total y sin minería",
     [("IMCE total", ["IMCE Total", "IMCE: Total", "Total"]),
      ("IMCE sin minería", ["IMCE: Sin Minería", "Sin Minería", "IMCE: Sin Mineria"])], None),
    ("IMCE por sector económico",
     [("Comercio", ["Comercio"]), ("Construcción", ["Construcción", "Construccion"]),
      ("Industria", ["Industria"])], None),
    ("Expectativas de empleo",
     [("Comercio", ["COM_Expectativas - Empleo"]), ("Construcción", ["CONS_Expectativas - Empleo"]),
      ("Industria", ["IND_Expectativas - Empleo"])], (0, 100)),
    ("Inventarios (vs deseado)",
     [("Comercio", ["COM_Situacion actual - Inventarios"]),
      ("Industria", ["IND_Situacion actual - Inventarios"])], (0, 100)),
    ("Presiones salariales",
     [("Construcción", ["CONS_Expectativas - Salarios"]), ("Industria", ["IND_Expectativas - Salarios"])],
     (0, 100)),
    ("Nivel de demanda percibida",
     [("Comercio", ["COM_Situacion actual - Demanda"]), ("Construcción", ["CONS_Situacion actual - Demanda"]),
      ("Industria", ["IND_Situacion actual - Demanda"])], (0, 100)),
    ("Expectativas de precios",
     [("Comercio", ["COM_Expectativas - Precios"]), ("Construcción", ["CONS_Expectativas - Precios"]),
      ("Industria", ["IND_Expectativas - Precios"])], (0, 100)),
    ("Expectativas de costos",
     [("Comercio", ["COM_Expectativas - Costos"]), ("Construcción", ["CONS_Expectativas - Costos"]),
      ("Industria", ["IND_Expectativas - Costos"])], (0, 100)),
]


def _norm_col(s):
    s = unicodedata.normalize("NFD", str(s) if s else "")
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", s.strip().lower()).replace("–", "-").replace("—", "-")


def find_col(df, candidates):
    cmap = {_norm_col(c): c for c in df.columns}
    for cand in candidates:
        k = _norm_col(cand)
        if k in cmap:
            return cmap[k]
    for cand in candidates:
        k = _norm_col(cand)
        for kk, real in cmap.items():
            if k in kk:
                return real
    raise KeyError(f"No encontré columna para {candidates}. Disponibles: {list(df.columns)[:10]}")


def read_imce() -> pd.DataFrame:
    df = pd.read_excel(IMCE_XLSX, sheet_name="IMCE", header=2)
    df.columns = [str(c).strip() for c in df.columns]
    c_f = find_col(df, ["Periodo"])
    df = df.rename(columns={c_f: "Fecha"})
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    df = df.dropna(subset=["Fecha"]).sort_values("Fecha").drop_duplicates("Fecha", keep="last").reset_index(drop=True)
    for c in df.columns:
        if c != "Fecha":
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def _ylim(df, paso=10, incluir=50.0):
    """Limites en multiplos de 'paso' que contienen las series y la linea de 50."""
    y = df.to_numpy(dtype=float)
    y = y[np.isfinite(y)]
    if y.size == 0:
        return (0, 100)
    lo, hi = min(y.min(), incluir), max(y.max(), incluir)
    return (paso * math.floor(lo / paso), paso * math.ceil(hi / paso))


def generar_chartbook():
    from pptx.util import Cm
    fg = _importar_formato()
    df = read_imce().set_index("Fecha")
    ult = df.index.max()
    ts = pd.Timestamp.now().strftime("%Y%m%d")
    x, y, ancho, alto = Cm(1.3), Cm(3.6), Cm(31.2), Cm(fg.ALTO_CM)

    prs = fg.nueva_presentacion()
    fg.portada(prs, f"IMCE · {MESES[ult.month - 1]} de {ult.year}")
    for num, (titulo, series, ylims) in enumerate(CHARTS, start=2):
        datos = df[[find_col(df, cand) for _, cand in series]]
        datos.columns = [nombre for nombre, _ in series]
        datos = datos.dropna(how="all")
        y_min, y_max = ylims or _ylim(datos)
        lam = fg.lamina(prs, titulo, num)
        fg.agregar_grafico(lam, datos, x, y, ancho, alto, unidad="(Índice)", posicion_unidad="arriba",
                           y_min=y_min, y_max=y_max, y_paso=10, linea_ref=50,
                           meses_entre_etiquetas=MESES_ENTRE_ETIQUETAS)
        fg.pie(lam, x, y + alto + Cm(0.15), ancho, FUENTE, NOTA)
        print(f"  📊 {titulo}")

    out_pptx = OUT_DIR / f"IMCE_Chartbook_Santander_{ts}.pptx"
    try:
        prs.save(out_pptx)
    except PermissionError:
        sys.exit(f"No pude guardar {out_pptx.name}: está abierto en PowerPoint. Ciérralo y vuelve a correr.")
    print(f"\n  ✅ PPTX: {out_pptx}")
    pdf = fg.exportar_pdf(out_pptx)
    if pdf:
        print(f"  ✅ PDF:  {pdf}")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    global IMCE_XLSX, OUT_DIR
    print("=" * 60)
    print("  IMCE — proceso mensual")
    print("=" * 60)

    cred = _credenciales()
    IMCE_XLSX = encontrar_imce_xlsx()
    OUT_DIR = IMCE_XLSX.parent / "tmp_chartbook_imce_santander"
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"  Credenciales: {'token' if 'token' in cred else 'usuario y clave'} (variables de entorno)")

    print(f"\n{'─'*60}")
    print(">> 1. Actualizando IMCE.xlsx desde API BCCh")
    print(f"{'─'*60}")
    actualizar_excel(cred)

    print(f"\n{'─'*60}")
    print(">> 2. Generando chartbook (PPTX y PDF)")
    print(f"{'─'*60}")
    generar_chartbook()

    print("\n" + "="*60)
    print("  PROCESO COMPLETO")
    print("="*60)


if __name__ == "__main__":
    main()
