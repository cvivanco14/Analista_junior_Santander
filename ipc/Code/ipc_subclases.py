"""
IPC Subclases — Proceso mensual completo
=========================================
Reemplaza el notebook. Ejecutar una vez al mes desde la carpeta Code:

    cd ipc/Code
    python ipc_subclases.py

Carpetas (relativas a este archivo, funciona en cualquier computador):
    Code/     este script (y x13as.exe / x13as_ascii.exe para el paso X-13)
    Inputs/   archivos descargados del INE y Base_IPC_plantilla.xlsx (opcional)
    Outputs/  Base_IPC.xlsx (se actualiza en cada corrida), respaldo/ y los chartbooks PDF

Requisitos (Windows):
    pip install pandas openpyxl numpy matplotlib reportlab statsmodels
                requests beautifulsoup4 xlwings pywin32 holidays lxml

Arquitectura:
    SIN Excel (openpyxl + pandas):
        1. empalme          → extiende "Índices empalmados (2023)"
        2. ipc_sv           → escribe valores en "Analíticos" cols B y C
        3. indices          → genera "Índices_analíticos", "Var_m_m_analiticos", "Var_a_a_analiticos"
        4. variaciones_sub  → genera "Var_m_m", "Var_a_a"
        5. texto_flash      → rellena "Texto_Flash"
        6. subidas_bajadas  → genera "Subidas_Bajadas"
        7. chartbook_sub    → PDF chartbook subclases (paralelo)
        8. chartbook_ana    → PDF chartbook analíticos (paralelo)

    CON Excel COM (xlwings / win32com):
        9.  x13             → desestacionaliza IPC_SV → "Velocidad" col C
        10. perfiles        → tablas + gráficos en "Tablas_Perfil_SubClases"
        11. fuente          → fija fuente Santander en todos los gráficos
"""

# ─────────────────────────────────────────────────────────────────────────────
# IMPORTS
# ─────────────────────────────────────────────────────────────────────────────
import os, re, sys, unicodedata, datetime as dt
from pathlib import Path
import shutil, time, warnings, textwrap

import numpy as np
import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from concurrent.futures import ProcessPoolExecutor, as_completed

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas


# ─────────────────────────────────────────────────────────────────────────────
# 1. CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────────────────────
CODE     = Path(__file__).resolve().parent
PROYECTO = CODE.parent
INPUTS   = PROYECTO / "Inputs"
OUTPUTS  = PROYECTO / "Outputs"
BASE_XLSX = OUTPUTS / "Base_IPC.xlsx"
PLANTILLA = INPUTS / "Base_IPC_plantilla.xlsx"

# Ubicación anterior (SharePoint/OneDrive): solo se usa para copiar Base_IPC.xlsx la primera vez.
_RUTAS_ANTERIORES = [
    Path.home() / "OneDrive - Santander Office 365" / "General - Estudios",
    Path.home() / "Santander Office 365" / "Estudios - General",
    Path.home() / "Santander Office 365" / "General",
]

def _asegurar_base() -> Path:
    """Deja Outputs/Base_IPC.xlsx; si no existe, la copia de la plantilla o de la ubicación anterior."""
    if BASE_XLSX.exists(): return BASE_XLSX
    candidatos = [PLANTILLA] + [r / "Bases - Chartpack" / "Bases" / "Local" / "IPC" / "Base_IPC.xlsx"
                                for r in _RUTAS_ANTERIORES]
    for c in candidatos:
        if c.is_file():
            shutil.copy2(c, BASE_XLSX)
            print(f"  📄 Base_IPC.xlsx copiada desde {c}")
            return BASE_XLSX
    raise FileNotFoundError(
        f"No encuentro Base_IPC.xlsx. Déjala en\n    {BASE_XLSX}\n"
        f"o la plantilla en\n    {PLANTILLA}")

def build_config() -> dict:
    for d in (INPUTS, OUTPUTS): d.mkdir(parents=True, exist_ok=True)
    ruta_xlsx = _asegurar_base()
    cfg = dict(
        code_dir       = CODE,                 # x13as.exe
        out_dir        = OUTPUTS,              # chartbooks
        backup_dir     = OUTPUTS / "respaldo",
        ruta_xlsx      = ruta_xlsx,
        temp_dir       = INPUTS,               # descargas del INE
        archivo_ine    = INPUTS / "ipc-xls-ine-subclases.xlsx",
        hacer_backup   = True,
        # Nombres de hojas (según Base_IPC.xlsx real)
        hoja_empalme   = "Índices empalmados (2023)",
        hoja_pond      = "Ponderadores",
        hoja_matriz    = "Matriz",
        hoja_analiticos= "Analíticos",
        hoja_texto     = "Texto_Flash",
        hoja_velocidad = "Velocidad",
        hoja_var_mm    = "Var_m_m",
        hoja_var_aa    = "Var_a_a",
        hoja_mm_ana    = "Var_m_m_analiticos",
        hoja_aa_ana    = "Var_a_a_analiticos",
        hoja_idx_ana   = "Índices_analíticos",
        hoja_sub_baj   = "Subidas_Bajadas",
        hoja_perfiles  = "Tablas_Perfil_SubClases",
    )
    print(f"✅ Config | xlsx={cfg['ruta_xlsx']}")
    return cfg


# ─────────────────────────────────────────────────────────────────────────────
# 2. CACHÉ EN MEMORIA (evita re-leer Excel entre pasos)
# ─────────────────────────────────────────────────────────────────────────────
_CACHE: dict = {}


# ─────────────────────────────────────────────────────────────────────────────
# 3. HELPERS COMUNES
# ─────────────────────────────────────────────────────────────────────────────
_MESES_MAP = {
    "ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,
    "jul":7,"ago":8,"sep":9,"sept":9,"oct":10,"nov":11,"dic":12,
}
_MESES_ES = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]

ANALITICOS = [
    "IPC","IPC sin volátiles","Bienes sin volátiles","Alimentos no volátiles",
    "Bienes sv. ex. alim.","Servicios sin volátiles","Servicios_A&I_sv","Servicios_ Resto_sv",
    "IPC volátiles","Energía Volátiles","Electricidad_v","Combustible_v",
    "Alim. Volát.","Alim. Resto_v","FyV (proxy)_v","Resto volátiles",
    "Bienes volát.","Servicios volát.","Servicios_A&I_v","Servicios_Resto_v",
]

def _norm(s) -> str:
    """Normaliza a mayúsculas sin tildes."""
    if pd.isna(s): return ""
    s = str(s).strip().upper()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", s)

def _norm_lower(s) -> str:
    """Normaliza a minúsculas sin tildes ni paréntesis."""
    s = "" if s is None else str(s)
    s = re.sub(r"\.\d+$", "", s.strip()).strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"\([^)]*\)", "", s).replace("_", " ")
    return re.sub(r"\s+", " ", s).strip()

def _to_float(x) -> float:
    if x is None: return np.nan
    if isinstance(x, (int, float, np.number)):
        return np.nan if (isinstance(x, float) and np.isnan(x)) else float(x)
    s = str(x).strip()
    if not s or s.lower() in ("nan", "#n/a", "#n/d", "#div/0!", "#¡div/0!", "#value!"): return np.nan
    s = s.replace("\u2212", "-").replace("%", "").replace("=", "")
    if "," in s and "." in s: s = s.replace(".", "").replace(",", ".")
    else: s = s.replace(",", ".")
    try: return float(s)
    except: return np.nan

def _parse_month(v) -> pd.Timestamp:
    """Convierte cualquier representación de mes a Timestamp(año, mes, 1)."""
    if v is None: return pd.NaT
    if isinstance(v, pd.Timestamp): return v.to_period("M").to_timestamp()
    if hasattr(v, "year"):
        try: return pd.Timestamp(int(v.year), int(v.month), 1)
        except: return pd.NaT
    s = str(v).strip().lower()
    if not s: return pd.NaT
    dt_try = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.notna(dt_try): return pd.Timestamp(dt_try.year, dt_try.month, 1)
    s2 = re.sub(r"\s+", "-", s.replace(".", "").replace("_", "-").replace("/", "-"))
    m = re.match(r"^(ene|feb|mar|abr|may|jun|jul|ago|sep|sept|oct|nov|dic)-(\d{2}|\d{4})$", s2)
    if m:
        yy = int(m.group(2))
        return pd.Timestamp(2000 + yy if yy < 100 else yy, _MESES_MAP[m.group(1)], 1)
    return pd.NaT

def _fmt_mes(d) -> str:
    meses = {1:"ene",2:"feb",3:"mar",4:"abr",5:"may",6:"jun",
             7:"jul",8:"ago",9:"sept",10:"oct",11:"nov",12:"dic"}
    return f"{meses[d.month]}-{d.year%100:02d}"

def _necesita_descarga(path: Path, max_age_hours: float = 6) -> bool:
    if not path.exists(): return True
    edad = (dt.datetime.now() - dt.datetime.fromtimestamp(path.stat().st_mtime)).total_seconds()
    return edad > max_age_hours * 3600

def _safe_name(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "", re.sub(r"\s+", "", str(s).strip()))[:120]

def _write_sheet(wb, sheet_name: str, df: pd.DataFrame):
    """Escribe DataFrame en una hoja (crea o reemplaza)."""
    if sheet_name in wb.sheetnames: wb.remove(wb[sheet_name])
    ws = wb.create_sheet(sheet_name)
    for j, col in enumerate(df.columns, 1):
        ws.cell(1, j, str(col))
    for i, (_, row) in enumerate(df.iterrows(), 2):
        for j, col in enumerate(df.columns, 1):
            v = row[col]
            if pd.isna(v) or (isinstance(v, float) and np.isinf(v)): continue
            ws.cell(i, j, v.to_pydatetime() if hasattr(v, "to_pydatetime")
                          else float(v) if col != "Fecha" else v)

def _hacer_backup(ruta: Path):
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    carpeta = ruta.parent / "respaldo"; carpeta.mkdir(parents=True, exist_ok=True)
    bk = carpeta / f"{stamp}_{ruta.name}"
    shutil.copy2(ruta, bk)
    print(f"  💾 Backup: {bk.name}")


# ─────────────────────────────────────────────────────────────────────────────
# 4. PASO 1 — EMPALME
# ─────────────────────────────────────────────────────────────────────────────
_URL_PORTAL   = "https://www.ine.gob.cl/estadisticas/economia/indices-de-precio-e-inflacion/indice-de-precios-al-consumidor"
_URL_FALLBACK = "https://www.ine.gob.cl/docs/default-source/%C3%ADndice-de-precios-al-consumidor/cuadros-estadisticos/base-anual-2023_100/series-de-tiempo/ipc-xls.xlsx"

def empalme(cfg: dict):
    """Descarga XLS del INE y extiende 'Índices empalmados (2023)' con meses nuevos."""
    t0 = time.time()
    ruta = cfg["ruta_xlsx"]
    if not ruta.exists(): raise FileNotFoundError(f"No existe: {ruta}")

    # — Descarga con caché —
    with requests.Session() as s:
        if _necesita_descarga(cfg["archivo_ine"]):
            try:
                resp = s.get(_URL_PORTAL, timeout=20); resp.raise_for_status()
                soup = BeautifulSoup(resp.text, "lxml")
                hrefs = [a["href"] for a in soup.find_all("a", href=True)
                         if a["href"].lower().endswith((".xlsx", ".xls")) and "ipc" in a["href"].lower()]
                hrefs = list(dict.fromkeys(hrefs))
                url = next((u for u in hrefs if "ipc-xls" in u.lower()), hrefs[0] if hrefs else _URL_FALLBACK)
            except Exception:
                url = _URL_FALLBACK
            print(f"  🔗 {url}")
            try:
                r = s.get(url, timeout=60); r.raise_for_status()
                cfg["archivo_ine"].write_bytes(r.content)
            except Exception as e:
                print(f"  ⚠️ Fallback: {e}")
                r = s.get(_URL_FALLBACK, timeout=60); r.raise_for_status()
                cfg["archivo_ine"].write_bytes(r.content)
            print("  📥 XLS descargado.")
        else:
            print("  ⚡ Cache válido — sin descarga.")

    # — Leer estado actual del empalme —
    df_emp = pd.read_excel(ruta, sheet_name=cfg["hoja_empalme"], header=None)
    glosas  = df_emp.iloc[0, 1:].astype(str).tolist()
    fechas_emp = pd.to_datetime(df_emp.iloc[2:, 0], errors="coerce").dropna()
    ultima = fechas_emp.max().date()
    glosas_norm = [_norm(g) for g in glosas]

    # — Leer XLS INE —
    xls = pd.ExcelFile(cfg["archivo_ine"])
    df_raw = None
    for sheet in xls.sheet_names:
        for hr in range(10):
            try:
                df_try = pd.read_excel(cfg["archivo_ine"], sheet_name=sheet, header=hr)
                cols_n = [_norm(c) for c in df_try.columns]
                if all(any(k in c for c in cols_n) for k in ["ANO", "MES", "GLOSA", "INDICE"]):
                    df_raw = df_try; break
            except Exception: continue
        if df_raw is not None: break
    if df_raw is None: raise ValueError("No encontré tabla Año/Mes/Glosa/Índice en XLS INE.")

    cols_n = [_norm(c) for c in df_raw.columns]
    mapa = {}
    for i, cn in enumerate(cols_n):
        col = df_raw.columns[i]
        if cn in ("ANO", "AÑO") and "ano" not in mapa:     mapa["ano"] = col
        elif cn == "MES"  and "mes" not in mapa:            mapa["mes"] = col
        elif ("INDICE" in cn or "ÍNDICE" in cn) and "ind" not in mapa: mapa["ind"] = col
        elif "GLOSA" in cn and "glosa" not in mapa:         mapa["glosa"] = col
        elif "SUBCLASE" in cn: mapa["subclase"] = col
        elif "PRODUCTO" in cn: mapa["producto"] = col

    df = df_raw[[mapa[k] for k in ["ano","mes","glosa","ind"] if k in mapa]].copy()
    df.columns = ["ano","mes","glosa","ind"]
    df["ano"] = pd.to_numeric(df["ano"], errors="coerce").astype("Int64")
    df["mes"] = pd.to_numeric(df["mes"], errors="coerce").astype("Int64")
    df["ind"] = pd.to_numeric(df["ind"], errors="coerce")
    df = df.dropna(subset=["ano","mes","ind"])
    if "subclase" in mapa and "producto" in mapa:
        df = df_raw.copy()  # re-filter with both cols
        df = df_raw[df_raw[mapa.get("subclase","_x_")].notna() &
                    df_raw[mapa.get("producto","_x_")].isna()].copy()
        df = df[[mapa["ano"],mapa["mes"],mapa["glosa"],mapa["ind"]]]
        df.columns = ["ano","mes","glosa","ind"]
        df["ano"] = pd.to_numeric(df["ano"], errors="coerce").astype("Int64")
        df["mes"] = pd.to_numeric(df["mes"], errors="coerce").astype("Int64")
        df["ind"] = pd.to_numeric(df["ind"], errors="coerce")
        df = df.dropna(subset=["ano","mes","ind"])

    df["glosa_n"] = df["glosa"].apply(_norm)
    df["fecha"]   = df.apply(lambda r: dt.date(int(r["ano"]), int(r["mes"]), 1), axis=1)
    df_sub = df[["fecha","glosa_n","ind"]].sort_values(["fecha","glosa_n"])

    # — Extender empalme —
    df_nuevo = df_sub[df_sub["fecha"] > ultima]
    if df_nuevo.empty:
        print("  ℹ️ Sin meses nuevos."); return

    fechas_nuevas = sorted(df_nuevo["fecha"].unique())
    print("  🆕 Meses nuevos:", ", ".join(str(f) for f in fechas_nuevas))
    mapa_idx = dict(zip(zip(df_nuevo["fecha"], df_nuevo["glosa_n"]), df_nuevo["ind"]))

    if cfg["hacer_backup"]: _hacer_backup(ruta)

    wb = load_workbook(ruta)
    ws = wb[cfg["hoja_empalme"]]
    n_cols = ws.max_column
    for fecha in fechas_nuevas:
        fila = [fecha] + [mapa_idx.get((fecha, g), None) for g in glosas_norm]
        if len(fila) < n_cols: fila += [None] * (n_cols - len(fila))
        ws.append(fila)
    wb.save(ruta)
    print(f"  ✅ {len(fechas_nuevas)} fila(s) nuevas — {time.time()-t0:.1f}s")


# ─────────────────────────────────────────────────────────────────────────────
# 5. PASO 2b — IPC EMPALME BCCh → Analíticos cols D y E
# ─────────────────────────────────────────────────────────────────────────────
# CREDENCIALES BCCh — cámbialas aquí si rotan
_BCCH_USER = "jpuentes@fen.uchile.cl"
_BCCH_PASS = "Svlh574477."
_BCCH_IPC  = "G073.IPC.IND.2023.M"    # IPC empalme BCCh → col D
_BCCH_SV   = "G073.IPCSV.IND.2023.M"  # SV empalme BCCh  → col E
_BCCH_URL  = "https://si3.bcentral.cl/SieteRestWS/SieteRestWS.ashx"

def _bcch_get_serie(codigo: str) -> pd.Series:
    """Descarga una serie mensual del BCCh y retorna pd.Series con índice DatetimeIndex."""
    import urllib.request, json as _json
    url = (f"{_BCCH_URL}?user={_BCCH_USER}&pass={_BCCH_PASS}"
           f"&function=GetSeries&timeseries={codigo}")
    with urllib.request.urlopen(url, timeout=30) as r:
        data = _json.loads(r.read())
    obs = data.get("Series", {}).get("Obs", [])
    if not obs:
        raise ValueError(f"Sin datos para {codigo}")
    rows = []
    for o in obs:
        try:
            fecha = pd.Timestamp(o["indexDateString"])
            valor = float(o["value"])
            rows.append((fecha, valor))
        except Exception:
            continue
    s = pd.Series({f: v for f, v in rows}).sort_index()
    return s[~s.index.duplicated(keep="last")]

def bcch_analiticos(cfg: dict):
    """Descarga IPC y SV empalme BCCh y escribe en Analíticos cols D y E.
    Si la API no responde o el mes nuevo aún no está disponible, avisa y continúa."""
    t0 = time.time()
    ruta = cfg["ruta_xlsx"]

    print("  Descargando series BCCh…")
    try:
        s_ipc = _bcch_get_serie(_BCCH_IPC)
        s_sv  = _bcch_get_serie(_BCCH_SV)
        print(f"  ✅ BCCh OK | rango: {s_ipc.index.min():%Y-%m} → {s_ipc.index.max():%Y-%m}")
    except Exception as e:
        print(f"  ⚠️  BCCh no disponible ({e}) — continuando sin actualizar cols D/E")
        return

    wb = load_workbook(ruta)
    ws = wb[cfg["hoja_analiticos"]]
    COL_F, COL_IPC_B, COL_IPC_D, COL_SV_E = 1, 4, 4, 5

    # Construir mapa fecha → fila leyendo col A
    date_to_row = {}; last_row = 1
    for r in range(1, ws.max_row + 1):
        d = _parse_month(ws.cell(r, COL_F).value)
        if pd.notna(d):
            date_to_row[pd.Timestamp(d.year, d.month, 1)] = r
            last_row = max(last_row, r)

    fmt_d = ws.cell(last_row, COL_IPC_D).number_format or "0.00"
    fmt_e = ws.cell(last_row, COL_SV_E).number_format  or "0.00"

    filled = skipped = 0
    for fecha, val_ipc in s_ipc.items():
        dt_ = pd.Timestamp(fecha.year, fecha.month, 1)
        if dt_ not in date_to_row:
            skipped += 1; continue        # fecha no existe aún en Analíticos
        val_sv = s_sv.get(dt_, np.nan)
        if np.isnan(val_sv): continue
        r = date_to_row[dt_]
        # Solo escribir si celda vacía (no sobreescribir valores existentes)
        if ws.cell(r, COL_IPC_D).value in (None, ""):
            ws.cell(r, COL_IPC_D).value = float(val_ipc)
            ws.cell(r, COL_IPC_D).number_format = fmt_d
        if ws.cell(r, COL_SV_E).value in (None, ""):
            ws.cell(r, COL_SV_E).value = float(val_sv)
            ws.cell(r, COL_SV_E).number_format = fmt_e
        filled += 1

    wb.save(ruta)
    print(f"  ✅ BCCh: {filled} filas actualizadas | {skipped} fechas sin fila en Analíticos — {time.time()-t0:.1f}s")


# ─────────────────────────────────────────────────────────────────────────────
# 5. PASO 2 — IPC GENERAL + SIN VOLÁTILES → Analíticos cols B y C
# ─────────────────────────────────────────────────────────────────────────────
_URL_IPC_XLS = "https://www.ine.gob.cl/docs/default-source/%C3%ADndice-de-precios-al-consumidor/cuadros-estadisticos/base-anual-2023_100/series-de-tiempo/ipc-xls.xlsx?sfvrsn=5b901f39_64"
_URL_ANA_XLS = "https://www.ine.gob.cl/docs/default-source/%C3%ADndice-de-precios-al-consumidor/cuadros-estadisticos/base-anual-2023_100/series-de-tiempo/analiticos-xls.xlsx?sfvrsn=3a9b5207_55"

def _find_header_row(df_raw) -> int:
    must = [_norm_lower(x) for x in ("Año","Mes","Glosa","Índice")]
    for r in range(min(80, df_raw.shape[0])):
        row = [_norm_lower(x) for x in df_raw.iloc[r, :].tolist()]
        if all(any(m in cell for cell in row) for m in must): return r
    raise ValueError("Sin encabezados Año/Mes/Glosa/Índice")

def _read_ine_long(path: Path) -> pd.DataFrame:
    for sh in pd.ExcelFile(path).sheet_names:
        try:
            raw = pd.read_excel(path, sheet_name=sh, header=None, dtype=object)
            hr  = _find_header_row(raw)
            df  = pd.read_excel(path, sheet_name=sh, header=hr, dtype=object)
            nrm = {c: _norm_lower(c) for c in df.columns}
            def pick(k):
                for c, cn in nrm.items():
                    if cn == _norm_lower(k): return c
                for c, cn in nrm.items():
                    if _norm_lower(k) in cn: return c
                return None
            cy, cm_, cg, ci = pick("Año"), pick("Mes"), pick("Glosa"), pick("Índice")
            if not all([cy, cm_, cg, ci]): continue
            out = df[[cy, cm_, cg, ci]].copy()
            out.columns = ["Año","Mes","Glosa","Índice"]
            out["Año"]    = pd.to_numeric(out["Año"], errors="coerce")
            out["Mes"]    = pd.to_numeric(out["Mes"], errors="coerce")
            out["Índice"] = out["Índice"].apply(_to_float)
            out = out.dropna(subset=["Año","Mes"])
            out = out[(out["Año"] >= 1989) & (out["Año"] <= 2035) & out["Mes"].between(1,12)]
            out["Fecha"] = out.apply(lambda r: dt.datetime(int(r["Año"]), int(r["Mes"]), 1), axis=1)
            return out.dropna(subset=["Índice"])
        except Exception: continue
    raise ValueError(f"No pude leer {path.name}")

def _build_serie(df_long, glosa: str) -> pd.Series:
    tg = _norm_lower(glosa)
    tmp = df_long[df_long["Glosa"].apply(_norm_lower) == tg]
    if tmp.empty: tmp = df_long[df_long["Glosa"].apply(_norm_lower).str.contains(tg, na=False)]
    if tmp.empty: raise ValueError(f"Glosa '{glosa}' no encontrada.")
    s = tmp.set_index("Fecha")["Índice"].astype(float).sort_index()
    return s[~s.index.duplicated(keep="last")]

def ipc_sv(cfg: dict):
    """Escribe IPC general y sin volátiles en hoja Analíticos cols B y C."""
    t0 = time.time()
    f_ipc = cfg["temp_dir"] / "ipc-xls.xlsx"
    f_ana = cfg["temp_dir"] / "analiticos-xls.xlsx"

    for url, dest in [(_URL_IPC_XLS, f_ipc), (_URL_ANA_XLS, f_ana)]:
        if _necesita_descarga(dest):
            with requests.get(url, stream=True, timeout=60) as r:
                r.raise_for_status(); dest.write_bytes(r.content)
            print(f"  📥 {dest.name}")
        else:
            print(f"  ⚡ Cache — {dest.name}")

    s_ipc = _build_serie(_read_ine_long(f_ipc), "IPC General")
    s_sv  = _build_serie(_read_ine_long(f_ana),  "IPC Sin volátiles")
    common = sorted(set(s_ipc.index) & set(s_sv.index))
    s_ipc, s_sv = s_ipc.loc[common], s_sv.loc[common]

    ruta = cfg["ruta_xlsx"]
    wb = load_workbook(ruta); ws = wb[cfg["hoja_analiticos"]]
    COL_F, COL_IPC, COL_SV = 1, 2, 3

    date_to_row = {}; last_row = 1
    for r in range(1, ws.max_row + 1):
        d = _parse_month(ws.cell(r, COL_F).value)
        if pd.notna(d):
            date_to_row[pd.Timestamp(d.year, d.month, 1)] = r
            last_row = max(last_row, r)
    last_dt  = max(date_to_row) if date_to_row else None
    fmt_b    = ws.cell(last_row, COL_IPC).number_format or "0.00"
    fmt_c    = ws.cell(last_row, COL_SV).number_format  or "0.00"
    fmt_f    = ws.cell(last_row, COL_F).number_format   or "mmm-yy"

    filled = appended = 0
    for d in common:
        dt_ = pd.Timestamp(d.year, d.month, 1)
        if dt_ in date_to_row:
            r = date_to_row[dt_]
            if ws.cell(r, COL_IPC).value in (None, ""):
                ws.cell(r, COL_IPC).value = float(s_ipc.loc[d])
                ws.cell(r, COL_IPC).number_format = fmt_b; filled += 1
            if ws.cell(r, COL_SV).value in (None, ""):
                ws.cell(r, COL_SV).value = float(s_sv.loc[d])
                ws.cell(r, COL_SV).number_format = fmt_c
        elif last_dt is None or dt_ > last_dt:
            last_row += 1
            # Escribir fecha como datetime real (no texto) para que ESNUMERO() funcione
            ws.cell(last_row, COL_F).value         = dt.datetime(d.year, d.month, 1)
            ws.cell(last_row, COL_F).number_format = fmt_f
            ws.cell(last_row, COL_IPC).value        = float(s_ipc.loc[d])
            ws.cell(last_row, COL_IPC).number_format = fmt_b
            ws.cell(last_row, COL_SV).value         = float(s_sv.loc[d])
            ws.cell(last_row, COL_SV).number_format  = fmt_c
            date_to_row[dt_] = last_row; appended += 1

    wb.save(ruta)
    print(f"  ✅ Rellenados:{filled} | Nuevos:{appended} — {time.time()-t0:.1f}s")

    # — Extender col A de Velocidad para los meses nuevos —
    # Velocidad tiene fórmulas pre-construidas apuntando a Analíticos!C208, C209...
    # pero col A (fecha) queda vacía → X-13 no puede leer el periodo
    if appended > 0:
        import re as _re
        wb2 = load_workbook(ruta)
        ws_v = wb2[cfg["hoja_velocidad"]]
        fmt_v = "mmm\\.yyyy"
        # Leer mapa: fila Analíticos referenciada → fila Velocidad
        for rv in range(2, ws_v.max_row + 1):
            b_val = ws_v.cell(rv, 1).value
            b_frm = ws_v.cell(rv, 2).value
            if b_val is not None: continue          # ya tiene fecha
            if not isinstance(b_frm, str): continue
            m = _re.search(r"Analíticos!C(\d+)", b_frm)
            if not m: continue
            ana_row = int(m.group(1))
            # ¿Qué fecha tiene esa fila en Analíticos?
            ws_a2 = wb2[cfg["hoja_analiticos"]]
            fecha_ana = ws_a2.cell(ana_row, 1).value
            if fecha_ana is None: continue
            fecha_ts = _parse_month(fecha_ana)
            if pd.isna(fecha_ts): continue
            ws_v.cell(rv, 1).value         = dt.datetime(fecha_ts.year, fecha_ts.month, 1)
            ws_v.cell(rv, 1).number_format = fmt_v
            print(f"  📅 Velocidad fila {rv} → {fecha_ts:%b-%Y}")
        wb2.save(ruta)
        print(f"  ✅ Velocidad col A extendida")


# ─────────────────────────────────────────────────────────────────────────────
# 6. PASO 3 — ÍNDICES ANALÍTICOS (vectorizado) + variaciones analíticos
# ─────────────────────────────────────────────────────────────────────────────
def indices_analiticos(cfg: dict):
    """Construye Índices_analíticos vectorizado + Var_m_m_analiticos + Var_a_a_analiticos."""
    t0 = time.time()
    ruta = cfg["ruta_xlsx"]

    # Cargar empalme
    df_emp = pd.read_excel(ruta, sheet_name=cfg["hoja_empalme"], header=None)
    hdr    = df_emp.iloc[1, :]
    cols   = [i for i, v in enumerate(hdr) if isinstance(v, (int, float)) and not pd.isna(v)]
    fechas = pd.to_datetime(df_emp.iloc[2:, 0], errors="coerce"); ok = fechas.notna()
    idx_emp = df_emp.iloc[2:, cols].loc[ok].astype(float)
    idx_emp.index   = fechas[ok]
    idx_emp.columns = [int(hdr[i]) for i in cols]

    # Cargar ponderadores
    df_p = pd.read_excel(ruta, sheet_name=cfg["hoja_pond"])
    base_years = [int(v) for v in df_p.iloc[0, 2:].tolist()]
    df_ps = df_p.iloc[1:, :].copy()
    df_ps = df_ps[pd.to_numeric(df_ps.iloc[:, 0], errors="coerce").notna()]
    df_ps = df_ps.rename(columns={df_ps.columns[0]: "N", df_ps.columns[1]: "Glosa"})
    weights = df_ps[["N"] + list(df_p.columns[2:])].copy()
    weights["N"] = pd.to_numeric(weights["N"], errors="coerce").astype("Int64")
    weights.set_index("N", inplace=True); weights.columns = base_years
    weights = weights.apply(pd.to_numeric, errors="coerce")

    # Cargar flags (Matriz)
    df_mx = pd.read_excel(ruta, sheet_name=cfg["hoja_matriz"], header=None)
    n_cols = 3 + len(ANALITICOS)
    body = df_mx.iloc[2:, :n_cols].copy()
    body.columns = ["Índices","Subclase","2023"] + ANALITICOS
    df_fl = body[pd.to_numeric(body["Subclase"], errors="coerce").notna()].copy()
    df_fl["Subclase"] = pd.to_numeric(df_fl["Subclase"], errors="coerce").astype("Int64")
    flags = df_fl.set_index("Subclase")[ANALITICOS].apply(pd.to_numeric, errors="coerce").fillna(0).astype(int)

    # Construcción vectorizada (×15-20 vs loop)
    by_sorted = sorted(int(b) for b in base_years)
    n_dates   = len(idx_emp)
    results   = {an: np.full(n_dates, np.nan) for an in ANALITICOS}
    dates_arr = idx_emp.index

    for k, base in enumerate(by_sorted):
        nxt  = by_sorted[k+1] if k+1 < len(by_sorted) else None
        mask = (dates_arr.year >= base) & (dates_arr.year < nxt) if nxt else (dates_arr.year >= base)
        idxs = np.where(mask)[0]
        if not len(idxs): continue
        sub_idx = idx_emp.iloc[idxs].values
        sub_ids = idx_emp.columns
        w_col   = weights[base].reindex(sub_ids).fillna(0).values
        for an in ANALITICOS:
            fl  = flags[an].reindex(sub_ids).fillna(0).values.astype(bool)
            if not fl.any(): continue
            w_s = w_col[fl]; idx_s = sub_idx[:, fl]; den = w_s.sum()
            if den == 0: continue
            results[an][idxs] = (idx_s * w_s).sum(axis=1) / den

    df_indices = pd.DataFrame(results, index=dates_arr)
    df_indices.index.name = "Fecha"
    df_indices = df_indices.reset_index()

    # Variaciones
    df_set = df_indices.set_index("Fecha")[ANALITICOS].apply(pd.to_numeric, errors="coerce")
    mm_raw = (df_set / df_set.shift(1)  - 1) * 100
    aa_raw = (df_set / df_set.shift(12) - 1) * 100

    # Guardar en caché
    _CACHE.update({
        "df_indices": df_indices, "idx_emp": idx_emp,
        "weights": weights, "flags": flags,
        "df_mm_ana": mm_raw, "df_aa_ana": aa_raw,
    })

    d0 = pd.Timestamp("2010-01-01")
    mm_out = mm_raw[mm_raw.index >= d0].reset_index()
    aa_out = aa_raw[aa_raw.index >= d0].reset_index()

    # ── Modo ACTUALIZACIÓN para Índices_analíticos (no recrear desde cero) ──
    # Solo escribe desde ene-2023 en adelante. Si la hoja existe y ya tiene
    # datos, solo agrega meses nuevos sin tocar los existentes.
    D_IDX_INICIO = pd.Timestamp("2023-01-01")
    df_idx_out = df_indices[df_indices["Fecha"] >= D_IDX_INICIO].copy()

    wb = load_workbook(ruta)

    # Índices_analíticos — actualizar sin borrar
    if cfg["hoja_idx_ana"] in wb.sheetnames:
        ws_idx = wb[cfg["hoja_idx_ana"]]
        # Leer fechas ya existentes en la hoja
        fechas_existentes = set()
        for r in range(2, ws_idx.max_row + 1):
            v = ws_idx.cell(r, 1).value
            if v is not None:
                try: fechas_existentes.add(pd.Timestamp(v).to_period("M").to_timestamp())
                except: pass
        # Solo agregar filas nuevas
        nuevas = df_idx_out[~df_idx_out["Fecha"].apply(
            lambda x: pd.Timestamp(x).to_period("M").to_timestamp()
        ).isin(fechas_existentes)]
        agregadas = 0
        for _, row in nuevas.iterrows():
            nr = ws_idx.max_row + 1
            ws_idx.cell(nr, 1, row["Fecha"].to_pydatetime() if hasattr(row["Fecha"], "to_pydatetime") else row["Fecha"])
            for j, an in enumerate(ANALITICOS, 2):
                v = row[an]
                if pd.notna(v) and not np.isinf(v):
                    ws_idx.cell(nr, j, float(v))
            agregadas += 1
        print(f"  ✅ {cfg['hoja_idx_ana']}: {agregadas} fila(s) nueva(s) agregada(s)")
    else:
        # Primera vez — crear desde cero desde 2023
        _write_sheet(wb, cfg["hoja_idx_ana"], df_idx_out)
        print(f"  ✅ {cfg['hoja_idx_ana']}: creada desde {D_IDX_INICIO:%Y-%m}")

    # Var_m_m_analiticos y Var_a_a_analiticos — igual, modo actualización
    for hoja, df_var in [(cfg["hoja_mm_ana"], mm_out), (cfg["hoja_aa_ana"], aa_out)]:
        if hoja in wb.sheetnames:
            ws_v = wb[hoja]
            fechas_ex = set()
            for r in range(2, ws_v.max_row + 1):
                v = ws_v.cell(r, 1).value
                if v is not None:
                    try: fechas_ex.add(pd.Timestamp(v).to_period("M").to_timestamp())
                    except: pass
            df_nuevas = df_var[~df_var["Fecha"].apply(
                lambda x: pd.Timestamp(x).to_period("M").to_timestamp()
            ).isin(fechas_ex)]
            for _, row in df_nuevas.iterrows():
                nr = ws_v.max_row + 1
                ws_v.cell(nr, 1, row["Fecha"].to_pydatetime() if hasattr(row["Fecha"], "to_pydatetime") else row["Fecha"])
                for j, an in enumerate(ANALITICOS, 2):
                    v = row.get(an, np.nan)
                    if pd.notna(v) and not np.isinf(float(v) if not isinstance(v, float) else v):
                        ws_v.cell(nr, j, float(v))
            print(f"  ✅ {hoja}: {len(df_nuevas)} fila(s) nueva(s)")
        else:
            _write_sheet(wb, hoja, df_var)
            print(f"  ✅ {hoja}: creada")

    wb.save(ruta)
    print(f"  Índices + variaciones analíticos — {time.time()-t0:.1f}s")



# ─────────────────────────────────────────────────────────────────────────────
# PASO 3b — REPARAR #REF! EN ANALÍTICOS (cols F-Y)
# ─────────────────────────────────────────────────────────────────────────────
def reparar_ref_analiticos(cfg: dict):
    """Repara fórmulas #REF! en Analíticos cols F-Y que apuntan a Índices_analíticos.
    
    Causa: en runs anteriores _write_sheet() borraba y recreaba Índices_analíticos,
    rompiendo todas las referencias. Esta función reconstruye las fórmulas correctas
    basándose en el mapa fecha→fila actual de Índices_analíticos.
    También corrige referencias con número de fila incorrecto (off-by-one).
    """
    from openpyxl.utils import get_column_letter
    ruta = cfg["ruta_xlsx"]

    # 1. Construir mapa fecha → fila en Índices_analíticos
    wb_r = load_workbook(ruta, read_only=True)
    ws_idx = wb_r[cfg["hoja_idx_ana"]]
    idx_date_to_row: dict = {}
    for r in range(1, ws_idx.max_row + 1):
        v = ws_idx.cell(r, 1).value
        if v is not None:
            try:
                ts = pd.Timestamp(v).to_period("M").to_timestamp()
                idx_date_to_row[ts] = r
            except Exception:
                pass
    wb_r.close()

    if not idx_date_to_row:
        print("  ⚠️  Índices_analíticos vacío — sin reparación")
        return

    # 2. Cargar Analíticos con fórmulas (sin data_only)
    wb = load_workbook(ruta)
    ws = wb[cfg["hoja_analiticos"]]

    IDX_SHEET   = "Índices_analíticos"
    ANA_COL_F   = 6   # col F
    ANA_COL_Y   = 25  # col Y
    IDX_OFFSET  = 4   # Analíticos_col - 4 = Índices_col (F=6→B=2, G=7→C=3, …, Y=25→U=21)

    reparadas = 0
    for r in range(2, ws.max_row + 1):
        a_val = ws.cell(r, 1).value
        if a_val is None:
            continue
        try:
            ts = pd.Timestamp(a_val).to_period("M").to_timestamp()
        except Exception:
            continue

        idx_row = idx_date_to_row.get(ts)
        if idx_row is None:
            continue  # mes que no tiene fila en Índices_analíticos todavía

        for ana_col in range(ANA_COL_F, ANA_COL_Y + 1):
            cell = ws.cell(r, ana_col)
            if not isinstance(cell.value, str):
                continue
            if IDX_SHEET not in cell.value and "#REF!" not in cell.value:
                continue  # no apunta a Índices_analíticos

            idx_col_letter = get_column_letter(ana_col - IDX_OFFSET)
            expected = (f'=IF(ISNUMBER($B{r}),{IDX_SHEET}!{idx_col_letter}{idx_row},"")'  )
            if cell.value != expected:
                cell.value = expected
                reparadas += 1

    if reparadas > 0:
        wb.save(ruta)
        print(f"  ✅ Reparadas {reparadas} fórmulas #REF!/incorrectas en Analíticos cols F-Y")
    else:
        print("  ✅ Sin fórmulas #REF! que reparar en Analíticos")
    wb.close()


# ─────────────────────────────────────────────────────────────────────────────
# 7. PASO 4 — VARIACIONES SUBCLASES (Var_m_m y Var_a_a)
# ─────────────────────────────────────────────────────────────────────────────
def variaciones_subclases(cfg: dict):
    """Calcula y escribe Var_m_m y Var_a_a para todas las subclases."""
    t0 = time.time()
    ruta = cfg["ruta_xlsx"]
    df = pd.read_excel(ruta, sheet_name=cfg["hoja_empalme"], header=None)
    glosas  = df.iloc[0, 1:]
    sub_ids = df.iloc[1, 1:]
    fechas  = pd.to_datetime(df.iloc[2:, 0], errors="coerce"); ok = fechas.notna()
    vals    = df.iloc[2:, 1:].loc[ok].astype(float)
    vals.index = fechas[ok]; vals.columns = sub_ids.values

    mm = (vals / vals.shift(1)  - 1) * 100
    aa = (vals / vals.shift(12) - 1) * 100
    _CACHE.update({"df_mm_sub": mm, "df_aa_sub": aa,
                   "glosas_sub": glosas, "sub_ids_sub": sub_ids})

    def _escribir(sheet_name, df_var):
        wb = load_workbook(ruta)
        if sheet_name in wb.sheetnames: wb.remove(wb[sheet_name])
        ws = wb.create_sheet(sheet_name)
        ws.cell(1, 1, "Índices")
        for j, gl in enumerate(glosas, 2): ws.cell(1, j, gl)
        ws.cell(2, 1, "Subclase")
        for j, sid in enumerate(sub_ids, 2):
            if not pd.isna(sid): ws.cell(2, j, float(sid))
        for i, (fecha, row) in enumerate(df_var.iterrows(), 3):
            ws.cell(i, 1, fecha.to_pydatetime())
            for j, val in enumerate(row.values, 2):
                if not (pd.isna(val) or np.isinf(val)): ws.cell(i, j, float(val))
        wb.save(ruta)
        print(f"  ✅ {sheet_name}")

    _escribir(cfg["hoja_var_mm"], mm)
    _escribir(cfg["hoja_var_aa"], aa)
    print(f"  Variaciones subclases — {time.time()-t0:.1f}s")


# ─────────────────────────────────────────────────────────────────────────────
# 8. PASO 5 — TEXTO FLASH
# ─────────────────────────────────────────────────────────────────────────────
def texto_flash(cfg: dict):
    """Rellena Texto_Flash: peso, m/m, incidencia, a/a, YTD. Usa caché."""
    t0 = time.time()
    if "df_mm_ana" not in _CACHE:
        raise RuntimeError("Ejecuta primero indices_analiticos().")

    mm = _CACHE["df_mm_ana"]; aa = _CACHE["df_aa_ana"]
    df_idx = _CACHE["df_indices"].copy()
    df_idx["Fecha"] = pd.to_datetime(df_idx["Fecha"])
    lvl = df_idx.set_index("Fecha")

    fechas = sorted(set(mm.index) & set(aa.index))
    if len(fechas) < 2: raise ValueError("Pocos meses en caché.")
    t_curr, t_prev = fechas[-1], fechas[-2]

    # IPC total para incidencia
    ipc_col = next((c for c in lvl.columns if _norm_lower(c) in ("ipc","ipc general")), None)
    if ipc_col is None:
        ipc_col = next((c for c in lvl.columns if "ipc" in _norm_lower(c) and "sin" not in _norm_lower(c)), None)
    ipc_prev_val = float(lvl.at[t_prev, ipc_col]) if ipc_col and t_prev in lvl.index else np.nan

    # Pesos desde Matriz
    ruta = cfg["ruta_xlsx"]
    wb_m = load_workbook(ruta, data_only=True); ws_m = wb_m[cfg["hoja_matriz"]]
    last_r = next((r for r in range(200, 2, -1) if ws_m.cell(r, 1).value not in (None, "")), 174)
    w_sub  = np.array([_to_float(ws_m.cell(r, 3).value) for r in range(3, last_r+1)], dtype=float)
    w_by_norm = {}
    col = 4
    while True:
        h = ws_m.cell(1, col).value
        if h is None or str(h).strip() == "": break
        vec = np.nan_to_num([_to_float(ws_m.cell(r, col).value) for r in range(3, last_r+1)])
        w_by_norm[_norm_lower(h)] = float(np.sum(np.nan_to_num(w_sub) * vec))
        col += 1
    wb_m.close()

    wb = load_workbook(ruta); ws = wb[cfg["hoja_texto"]]

    # Detectar fila header — busca "Var. Mensual" o "Var. Anual" (robusto aunque falte "Peso")
    hdr = next((r for r in range(1, 401)
                if any("var" in _norm_lower(str(ws.cell(r, c).value or "")) and
                       ("mens" in _norm_lower(str(ws.cell(r, c).value or "")) or
                        "anual" in _norm_lower(str(ws.cell(r, c).value or "")))
                       for c in range(1, 20))), None)
    if hdr is None: raise ValueError("Sin header en Texto_Flash.")
    r0 = hdr + 2; r1 = r0
    while r1 <= 400 and ws.cell(r1, 1).value not in (None, ""):
        r1 += 1
    r1 -= 1

    COL_ANA, COL_PESO = 1, 2
    COL_MM_P, COL_MM_C, COL_INCID, COL_AA_P, COL_AA_C = 3, 4, 5, 6, 7
    COL_ACUM = next((c for r in [hdr, hdr+1] for c in range(1, 40)
                     if "acum" in _norm_lower(ws.cell(r, c).value or "")), 8)

    ws.cell(hdr, 2).value = t_prev.to_pydatetime(); ws.cell(hdr, 2).number_format = "mmm-yy"
    ws.cell(hdr, 4).value = t_curr.to_pydatetime(); ws.cell(hdr, 4).number_format = "mmm-yy"
    # FIX: actualizar también la fila de texto con el nombre del mes en español
    _MESES_FULL = ["enero","febrero","marzo","abril","mayo","junio",
                   "julio","agosto","septiembre","octubre","noviembre","diciembre"]
    # Fila hdr-2: celda B = mes previo (texto), celda D = mes actual (texto)
    _lbl_row = hdr - 2
    if _lbl_row >= 1:
        _prev_cell = ws.cell(_lbl_row, 2)
        _curr_cell = ws.cell(_lbl_row, 4)
        if _prev_cell.value is None or isinstance(_prev_cell.value, str):
            _prev_cell.value = _MESES_FULL[t_prev.month - 1]
        if _curr_cell.value is None or isinstance(_curr_cell.value, str):
            _curr_cell.value = _MESES_FULL[t_curr.month - 1]

    map_norm = {_norm_lower(c): c for c in mm.columns if c in aa.columns}
    lvl_map  = {_norm_lower(c): c for c in lvl.columns}
    updated  = 0

    for r in range(r0, r1+1):
        name = ws.cell(r, COL_ANA).value
        if not name: continue
        key = _norm_lower(name)
        w   = w_by_norm.get(key, np.nan)
        if np.isfinite(w):
            ws.cell(r, COL_PESO).value = w; ws.cell(r, COL_PESO).number_format = "0.0"
        else:
            w = _to_float(ws.cell(r, COL_PESO).value)
        if key not in map_norm: continue
        col = map_norm[key]
        mm_p = _to_float(mm.at[t_prev, col]) if t_prev in mm.index else np.nan
        mm_c = _to_float(mm.at[t_curr, col]) if t_curr in mm.index else np.nan
        aa_p = _to_float(aa.at[t_prev, col]) if t_prev in aa.index else np.nan
        aa_c = _to_float(aa.at[t_curr, col]) if t_curr in aa.index else np.nan
        inc  = np.nan
        if key in lvl_map:
            lc = lvl_map[key]
            xt = _to_float(lvl.at[t_curr, lc]) if t_curr in lvl.index else np.nan
            xp = _to_float(lvl.at[t_prev, lc]) if t_prev in lvl.index else np.nan
            if all(np.isfinite(v) for v in [w, xt, xp, ipc_prev_val]) and ipc_prev_val != 0:
                inc = w * (xt - xp) / ipc_prev_val
        # YTD
        s_mm = mm[col].dropna() if col in mm.columns else pd.Series(dtype=float)
        s_ytd = s_mm[(s_mm.index.year == t_curr.year) & (s_mm.index <= t_curr)]
        acum = float((np.prod(1 + s_ytd.values/100) - 1)*100) if len(s_ytd) else np.nan

        for cc, vv in [(COL_MM_P,mm_p),(COL_MM_C,mm_c),(COL_INCID,inc),(COL_AA_P,aa_p),(COL_AA_C,aa_c),(COL_ACUM,acum)]:
            ws.cell(r, cc).value = float(vv) if np.isfinite(vv) else None
            ws.cell(r, cc).number_format = "0.00"
        updated += 1

    wb.save(ruta); wb.close()
    print(f"  ✅ Texto Flash: {updated} filas — {time.time()-t0:.1f}s")


# ─────────────────────────────────────────────────────────────────────────────
# 9. PASO 6 — SUBIDAS Y BAJADAS
# ─────────────────────────────────────────────────────────────────────────────
def subidas_bajadas(cfg: dict):
    """Genera ranking Top 15 subidas/bajadas por mes desde ene-2025."""
    t0 = time.time()
    ruta   = cfg["ruta_xlsx"]
    TOP_N  = 15
    START  = pd.Timestamp(2025, 1, 1)

    thin   = Side(style="thin", color="D9D9D9")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    F_TITLE = PatternFill("solid", fgColor="1F4E79")
    F_SUB   = PatternFill("solid", fgColor="BDD7EE")
    F_HDR   = PatternFill("solid", fgColor="2F5597")
    FONT_TI = Font(bold=True, color="FFFFFF", size=13)
    FONT_SH = Font(bold=True, size=11)
    FONT_HD = Font(bold=True, color="FFFFFF", size=10)
    FONT_CE = Font(size=9)
    CENTER  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT    = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    def _ti(x):
        if x is None: return None
        if isinstance(x, (float, np.floating)) and np.isnan(x): return None
        if isinstance(x, (int, np.integer)): return int(x)
        s = re.sub(r"[^\d]+", "", str(x).strip())
        return int(s) if s else None

    def _read_twoheader(sheet):
        raw = pd.read_excel(ruta, sheet_name=sheet, header=None, dtype=object)
        ri  = next((i for i in range(min(60, raw.shape[0]))
                    if "indices" in _norm_lower(raw.iloc[i,0]) or "índices" in _norm_lower(raw.iloc[i,0])), None)
        if ri is None: raise ValueError(f"Sin two-header en {sheet}")
        names = raw.iloc[ri,:].tolist(); codes = raw.iloc[ri+1,:].tolist()
        col_pairs = []; c2n = {}
        for j in range(1, raw.shape[1]):
            c = _ti(codes[j])
            if c is None: continue
            nm = str(names[j]).strip() if names[j] else f"Subclase_{c}"
            col_pairs.append((j, c)); c2n[c] = nm
        data   = raw.iloc[ri+2:,:].copy().dropna(how="all")
        fechas = data.iloc[:,0].apply(_parse_month); ok = fechas.notna()
        # FIX PerformanceWarning: construir DataFrame en un solo pd.concat
        _dok = data.loc[ok]
        _cols2 = {"Fecha": fechas.loc[ok].values}
        _cols2.update({code: _dok.iloc[:, j].apply(_to_float).values for j, code in col_pairs})
        out = pd.DataFrame(_cols2)
        return out.sort_values("Fecha").drop_duplicates("Fecha", keep="last").reset_index(drop=True), c2n

    # Ponderadores 2023
    raw_p = pd.read_excel(ruta, sheet_name=cfg["hoja_pond"], header=None, dtype=object)
    def is23(v):
        if v is None or (isinstance(v, float) and np.isnan(v)): return False
        try: return abs(float(v)-2023) < 0.1
        except: return str(v).strip() == "2023"
    hdr_p = next((i for i in range(min(400, raw_p.shape[0]))
                  if any("subclase" in str(x).lower() for x in raw_p.iloc[i,:])
                  and any(is23(x) for x in raw_p.iloc[i,:])), None)
    if hdr_p is None: raise ValueError("Sin header en Ponderadores")
    headers_p = raw_p.iloc[hdr_p,:].tolist()
    cc_p = next((j for j,h in enumerate(headers_p) if any(k in str(h).lower() for k in ["nº","n°","numero"])), 0)
    cy_p = next((j for j,h in enumerate(headers_p) if is23(h)), None)
    if cy_p is None: raise ValueError("Sin col 2023 en Ponderadores")
    data_p = raw_p.iloc[hdr_p+1:,:].dropna(how="all")
    # FIX: filtrar por tipo numérico real (strings tipo 'Cobertura subclases 2023' extraen dígitos pero no son int válidos)
    w2023 = {}
    for _c, _w in zip(data_p.iloc[:, cc_p], data_p.iloc[:, cy_p]):
        if not isinstance(_c, (int, float, np.integer, np.floating)):
            continue
        if isinstance(_c, float) and (np.isnan(_c) or np.isinf(_c)):
            continue
        _wf = _to_float(_w)
        if np.isfinite(_wf):
            w2023[int(_c)] = _wf

    if "df_mm_sub" in _CACHE:
        mm_df = _CACHE["df_mm_sub"].reset_index().rename(columns={"index":"Fecha", None:"Fecha"})
        aa_df = _CACHE["df_aa_sub"].reset_index().rename(columns={"index":"Fecha", None:"Fecha"})
        _, ctn = _read_twoheader(cfg["hoja_var_mm"]); code_to_name = ctn
    else:
        mm_df, ctn_mm = _read_twoheader(cfg["hoja_var_mm"])
        aa_df, ctn_aa = _read_twoheader(cfg["hoja_var_aa"])
        code_to_name  = {**ctn_mm, **ctn_aa}

    if "Fecha" not in mm_df.columns: mm_df = mm_df.reset_index().rename(columns={"index":"Fecha"})
    if "Fecha" not in aa_df.columns: aa_df = aa_df.reset_index().rename(columns={"index":"Fecha"})
    mm_df = mm_df[mm_df["Fecha"] >= START]; aa_df = aa_df[aa_df["Fecha"] >= START]
    mm = mm_df.set_index("Fecha"); aa = aa_df.set_index("Fecha")
    fechas = sorted(set(mm.index) & set(aa.index))
    codes  = sorted(set(mm.columns) & set(aa.columns))

    def write_table(ws, sr, sc, title, rows):
        hdrs = ["Rank","Código","Subclase","Var m/m (%)","Pond 2023 (%)","Incidencia (pp)","Var a/a (%)"]
        r = sr
        ws.merge_cells(start_row=r, start_column=sc, end_row=r, end_column=sc+len(hdrs)-1)
        c = ws.cell(r, sc, title); c.fill = F_SUB; c.font = FONT_SH; c.alignment = LEFT; r += 1
        for j, h in enumerate(hdrs):
            cc = ws.cell(r, sc+j, h); cc.fill = F_HDR; cc.font = FONT_HD; cc.alignment = CENTER; cc.border = BORDER
        r += 1
        for row in rows:
            for j, h in enumerate(hdrs):
                v = row.get(h); cc = ws.cell(r, sc+j, v); cc.font = FONT_CE
                cc.alignment = LEFT if h == "Subclase" else CENTER; cc.border = BORDER
                if h in ["Var m/m (%)","Pond 2023 (%)","Var a/a (%)","Incidencia (pp)"]: cc.number_format = "0.00"
            r += 1
        widths = {"Rank":6,"Código":9,"Subclase":44,"Var m/m (%)":11,"Pond 2023 (%)":12,"Incidencia (pp)":14,"Var a/a (%)":11}
        for j, h in enumerate(hdrs): ws.column_dimensions[get_column_letter(sc+j)].width = widths[h]
        return r

    wb = load_workbook(ruta)
    if cfg["hoja_sub_baj"] in wb.sheetnames: wb.remove(wb[cfg["hoja_sub_baj"]])
    ws = wb.create_sheet(cfg["hoja_sub_baj"])
    ws.merge_cells("A1:G1"); t = ws["A1"]
    t.value = f"Top {TOP_N} Subclases — desde {START:%b-%Y}"; t.fill = F_TITLE; t.font = FONT_TI; t.alignment = LEFT
    ws.merge_cells("A2:G2"); ws["A2"].value = "Incidencia (pp) = Var m/m(%) × Pond 2023(%) / 100"
    row = 4

    for fecha in fechas:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        c = ws.cell(row, 1, f"Mes: {fecha:%b-%Y}")
        c.fill = F_TITLE; c.font = Font(bold=True, color="FFFFFF", size=12); c.alignment = LEFT; row += 1

        s_mm = pd.to_numeric(mm.loc[fecha].reindex(codes), errors="coerce").replace([np.inf,-np.inf], np.nan).dropna()
        s_aa = pd.to_numeric(aa.loc[fecha].reindex(codes), errors="coerce").replace([np.inf,-np.inf], np.nan)

        def build_rows(sorted_s):
            out = []
            for rk, (code, vmm) in enumerate(sorted_s.items(), 1):
                w   = w2023.get(int(code), np.nan)
                vaa = float(s_aa.get(code, np.nan))
                inc = (float(vmm)*float(w)/100) if np.isfinite(vmm) and np.isfinite(w) else np.nan
                out.append({"Rank":rk,"Código":int(code),
                            "Subclase":code_to_name.get(int(code), f"Subclase_{code}"),
                            "Var m/m (%)":float(vmm),"Pond 2023 (%)":w if np.isfinite(w) else np.nan,
                            "Incidencia (pp)":inc,"Var a/a (%)":vaa if np.isfinite(vaa) else np.nan})
            return out

        row = write_table(ws, row, 1, f"TOP {TOP_N} — SUBEN", build_rows(s_mm.sort_values(ascending=False).head(TOP_N))) + 1
        row = write_table(ws, row, 1, f"TOP {TOP_N} — BAJAN", build_rows(s_mm.sort_values(ascending=True).head(TOP_N)))  + 2

    ws.freeze_panes = "A4"; wb.save(ruta); wb.close()
    print(f"  ✅ Subidas_Bajadas lista — {time.time()-t0:.1f}s")


# ─────────────────────────────────────────────────────────────────────────────
# 10. PASO 7 — CHARTBOOK SUBCLASES (PDF paralelo)
# ─────────────────────────────────────────────────────────────────────────────
def _png_worker_sub(args):
    """Worker top-level para ProcessPoolExecutor (subclases)."""
    (code, nm, lvl_v, lvl_i, mm_v, mm_i, aa_v, aa_i,
     peso, inc_pp, last_mm, last_aa, report_month, out_png,
     PROFILE_YEARS, BAND_START, BAND_END, MM_YRS, AA_YRS, DPI, START_COMMON, MESES_ES) = args

    import matplotlib; matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import numpy as np, pandas as pd

    lvl = pd.Series(lvl_v, index=pd.DatetimeIndex(lvl_i))
    mm  = pd.Series(mm_v,  index=pd.DatetimeIndex(mm_i))
    aa  = pd.Series(aa_v,  index=pd.DatetimeIndex(aa_i))
    rpt = pd.Timestamp(report_month)

    def _b(s, y0, y1):
        s = s.dropna(); mn = np.arange(1,13)
        if s.empty: n = np.full(12,np.nan); return mn,n,n,n
        df = pd.DataFrame({"v":s,"y":s.index.year,"m":s.index.month})
        h  = df[(df.y>=y0)&(df.y<=y1)]
        if h.empty: n = np.full(12,np.nan); return mn,n,n,n
        g = h.groupby("m")["v"]
        return mn,g.min().reindex(range(1,13)).values,g.max().reindex(range(1,13)).values,g.mean().reindex(range(1,13)).values

    def _yl(s, years):
        s = s.dropna()
        return {y:s[s.index.year==y].groupby(s[s.index.year==y].index.month).last().reindex(range(1,13)).values for y in years}

    fig = plt.figure(figsize=(11.4,6.6), dpi=DPI)
    gs  = fig.add_gridspec(2,2, hspace=.35, wspace=.25)
    ax1,ax2,ax3,ax4 = [fig.add_subplot(gs[r,c]) for r in range(2) for c in range(2)]

    mn,vmin,vmax,vavg = _b(mm, BAND_START, BAND_END); yl = _yl(mm, PROFILE_YEARS)
    ax1.fill_between(mn, vmin, vmax, alpha=.18, linewidth=0)
    ax1.plot(mn, vavg, linewidth=2.5, label=f"Prom.{BAND_START}-{BAND_END}")
    for y in PROFILE_YEARS:
        yv = yl.get(y, np.full(12,np.nan))
        if np.isfinite(yv).sum() <= 1: ax1.plot(mn,yv,linestyle="None",marker="o",markersize=8,label=str(y))
        else: ax1.plot(mn,yv,linewidth=2.5,label=str(y))
    ax1.axhline(0,linewidth=1); ax1.set_xlim(1,12); ax1.set_xticks(range(1,13))
    ax1.set_xticklabels(MESES_ES, rotation=90); ax1.set_title("m/m profile", fontsize=10)
    ax1.grid(True,alpha=.2); ax1.tick_params(labelsize=8); ax1.legend(fontsize=7,ncol=2,frameon=False)
    ax1.text(.99,.02,f"Último:{rpt:%b-%y}",transform=ax1.transAxes,ha="right",va="bottom",fontsize=8)

    s2 = mm.dropna(); s2 = s2[s2.index >= rpt - pd.DateOffset(years=MM_YRS)]
    ax2.plot(s2.index, s2.values); ax2.axhline(0,linewidth=1)
    ax2.set_title(f"m/m — {MM_YRS}a", fontsize=10); ax2.grid(True,alpha=.2); ax2.tick_params(labelsize=8)

    s3 = lvl.dropna(); s3 = s3[s3.index >= pd.Timestamp(START_COMMON)]
    ax3.plot(s3.index, s3.values); ax3.set_title("Niveles (2023=100)",fontsize=10)
    if not s3.empty:
        lo,hi = float(s3.quantile(.02)), float(s3.quantile(.98))
        if np.isfinite(lo) and np.isfinite(hi) and hi > lo: ax3.set_ylim(lo-.1*(hi-lo), hi+.1*(hi-lo))
    ax3.grid(True,alpha=.2); ax3.tick_params(labelsize=8)

    s4 = aa.dropna(); s4 = s4[s4.index >= rpt - pd.DateOffset(years=AA_YRS)]
    ax4.plot(s4.index, s4.values); ax4.axhline(0,linewidth=1)
    ax4.set_title(f"a/a — {AA_YRS}a", fontsize=10); ax4.grid(True,alpha=.2); ax4.tick_params(labelsize=8)

    def fn(x, nd=2): return f"{x:.{nd}f}" if isinstance(x,float) and np.isfinite(x) else "n/a"
    fig.suptitle(f"{nm}  |  Peso:{fn(peso)}  |  m/m:{fn(last_mm)}  |  Inc:{fn(inc_pp,3)}  |  a/a:{fn(last_aa)}",
                 fontsize=11, y=.98)
    fig.savefig(out_png, bbox_inches="tight"); plt.close(fig)
    return str(out_png)

def chartbook_sub(cfg: dict):
    """Genera PDF chartbook de subclases con 4 paneles por subclase (paralelo)."""
    t0 = time.time()
    PROFILE_YEARS = [2023,2024,2025,2026]; BAND_START,BAND_END = 2013,2019
    MM_YRS = AA_YRS = 5; DPI = 150
    START_COMMON = "2009-01-01"
    N_WORKERS    = min(4, max(1, os.cpu_count()-1))
    OUT_DIR = cfg["out_dir"]; PNG_DIR = OUT_DIR / "PNG_subclases"
    OUT_DIR.mkdir(parents=True, exist_ok=True); PNG_DIR.mkdir(parents=True, exist_ok=True)
    ruta = cfg["ruta_xlsx"]

    def _read_two(sheet):
        raw = pd.read_excel(ruta, sheet_name=sheet, header=None)
        ri  = next((i for i in range(min(60,raw.shape[0]))
                    if "indices" in _norm_lower(str(raw.iloc[i,0])) or "índices" in _norm_lower(str(raw.iloc[i,0]))), None)
        if ri is None: raise ValueError(f"Sin two-header en {sheet}")
        names = raw.iloc[ri,:].tolist(); codes = raw.iloc[ri+1,:].tolist()
        col_pairs=[]; c2n={}
        for j in range(1,raw.shape[1]):
            def _ti2(x):
                if x is None: return None
                if isinstance(x,(float,np.floating)) and np.isnan(x): return None
                if isinstance(x,(int,np.integer)): return int(x)
                s=re.sub(r"[^\d]+","",str(x).strip()); return int(s) if s else None
            c = _ti2(codes[j])
            if c is None: continue
            col_pairs.append((j,c)); c2n[c] = str(names[j]).strip() if names[j] else f"Sub_{c}"
        data=raw.iloc[ri+2:,:].dropna(how="all")
        fechas=data.iloc[:,0].apply(_parse_month); ok=fechas.notna()
        # FIX PerformanceWarning: construir DataFrame en un solo pd.concat
        _data_ok = data.loc[ok]
        _cols = {"Fecha": fechas.loc[ok].values}
        _cols.update({code: _data_ok.iloc[:, j].apply(_to_float).values for j, code in col_pairs})
        out = pd.DataFrame(_cols)
        return out.sort_values("Fecha").drop_duplicates("Fecha",keep="last").reset_index(drop=True), c2n

    w2023 = {}
    raw_p = pd.read_excel(ruta, sheet_name=cfg["hoja_pond"], header=None)
    def is23(v):
        try: return abs(float(v)-2023)<0.1
        except: return False
    for i in range(min(400,raw_p.shape[0])):
        if any("subclase" in str(x).lower() for x in raw_p.iloc[i,:]) and any(is23(x) for x in raw_p.iloc[i,:]):
            hdr_p=i; break
    else: hdr_p=1
    hp=raw_p.iloc[hdr_p,:].tolist()
    cc_p=next((j for j,h in enumerate(hp) if any(k in str(h).lower() for k in ["nº","n°","numero"])),0)
    cy_p=next((j for j,h in enumerate(hp) if is23(h)),None)
    if cy_p:
        for _,r in raw_p.iloc[hdr_p+1:,:].iterrows():
            c=r.iloc[cc_p]; w=_to_float(r.iloc[cy_p])
            try:
                ci=int(float(str(c).strip()))
                if np.isfinite(w): w2023[ci]=w
            except: pass

    if "df_mm_sub" in _CACHE:
        mm_mat=_CACHE["df_mm_sub"]; aa_mat=_CACHE["df_aa_sub"]
        _,name_mm=_read_two(cfg["hoja_var_mm"])
    else:
        df_mm,name_mm=_read_two(cfg["hoja_var_mm"]); mm_mat=df_mm.set_index("Fecha")
        df_aa,_      =_read_two(cfg["hoja_var_aa"]); aa_mat=df_aa.set_index("Fecha")

    if "idx_emp" in _CACHE:
        lvl_mat=_CACHE["idx_emp"]
    else:
        df_lv,_=_read_two(cfg["hoja_empalme"]); lvl_mat=df_lv.set_index("Fecha")

    sc_ts = pd.Timestamp(START_COMMON)
    if hasattr(mm_mat,"loc"):  mm_mat  = mm_mat[mm_mat.index >= sc_ts]
    if hasattr(aa_mat,"loc"):  aa_mat  = aa_mat[aa_mat.index >= sc_ts]
    if hasattr(lvl_mat,"loc"): lvl_mat = lvl_mat[lvl_mat.index >= sc_ts]

    dates = sorted(set(mm_mat.index)&set(aa_mat.index)&set(lvl_mat.index))
    report_month = dates[-1]; prev_month = dates[-2]

    # IPC total para incidencia
    ipc_ana = pd.read_excel(ruta, sheet_name=cfg["hoja_analiticos"], usecols="A:B", header=0)
    ipc_ana.columns = ["Fecha","IPC"]
    ipc_ana["Fecha"] = ipc_ana["Fecha"].apply(_parse_month)
    ipc_ana = ipc_ana[ipc_ana["Fecha"].notna()].set_index("Fecha")["IPC"].apply(_to_float)
    ipc_prev = float(ipc_ana.loc[prev_month]) if prev_month in ipc_ana.index else np.nan

    mm_mat=mm_mat.loc[dates]; aa_mat=aa_mat.loc[dates]; lvl_mat=lvl_mat.loc[dates]
    codes = sorted(set(mm_mat.columns)&set(aa_mat.columns)&set(lvl_mat.columns))

    print(f"  {len(codes)} subclases | {report_month:%Y-%m} | workers={N_WORKERS}")
    tasks = []
    for code in codes:
        nm   = name_mm.get(code, f"Subclase_{code}")
        lv_s = lvl_mat[code]; mm_s = mm_mat[code]; aa_s = aa_mat[code]
        last_mm = float(mm_s.loc[report_month]) if pd.notna(mm_s.loc[report_month]) else np.nan
        last_aa = float(aa_s.loc[report_month]) if pd.notna(aa_s.loc[report_month]) else np.nan
        w = w2023.get(int(code), np.nan)
        xt = float(lv_s.loc[report_month]) if pd.notna(lv_s.loc[report_month]) else np.nan
        xp = float(lv_s.loc[prev_month])   if pd.notna(lv_s.loc[prev_month]) else np.nan
        inc = w*(xt-xp)/ipc_prev if all(np.isfinite(v) for v in [w,xt,xp,ipc_prev]) and ipc_prev!=0 else np.nan
        out_png = PNG_DIR / f"{int(code):03d}__{_safe_name(nm)}__4p.png"
        tasks.append((code,nm,
                      lv_s.values.tolist(),lv_s.index.astype(str).tolist(),
                      mm_s.values.tolist(),mm_s.index.astype(str).tolist(),
                      aa_s.values.tolist(),aa_s.index.astype(str).tolist(),
                      w,inc,last_mm,last_aa,str(report_month),str(out_png),
                      PROFILE_YEARS,BAND_START,BAND_END,MM_YRS,AA_YRS,DPI,START_COMMON,_MESES_ES))

    pngs = []
    with ProcessPoolExecutor(max_workers=N_WORKERS) as ex:
        futs = {ex.submit(_png_worker_sub, t): t[0] for t in tasks}
        for i, f in enumerate(as_completed(futs), 1):
            try:
                p = Path(f.result())
                if p.exists(): pngs.append(p)
            except Exception as e: print(f"  ⚠️ {futs[f]}: {e}")
            if i % 20 == 0: print(f"  {i}/{len(tasks)}…")
    pngs.sort()
    print(f"  PNGs: {len(pngs)}")

    out_pdf = OUT_DIR / "Chartbook_SubClases_4panel.pdf"
    cv = canvas.Canvas(str(out_pdf), pagesize=landscape(A4)); wp,hp = landscape(A4)
    cv.setFont("Helvetica-Bold", 18); cv.drawString(2*cm, hp-2.7*cm, "CPI Chartbook — Subclases")
    cv.setFont("Helvetica", 11);      cv.drawString(2*cm, hp-3.9*cm, f"Último: {report_month:%b-%Y}"); cv.showPage()
    mg = 1.2*cm
    for p in pngs:
        cv.drawImage(str(p), mg, mg, width=wp-2*mg, height=hp-2*mg, preserveAspectRatio=True, anchor="c")
        cv.setFont("Helvetica", 8); cv.drawRightString(wp-cm, .7*cm, p.stem[:90]); cv.showPage()
    cv.save()
    print(f"  ✅ PDF: {out_pdf} — {time.time()-t0:.1f}s")


# ─────────────────────────────────────────────────────────────────────────────
# 11. PASO 8 — CHARTBOOK ANALÍTICOS (PDF paralelo)
# ─────────────────────────────────────────────────────────────────────────────
def _png_worker_ana(args):
    """Worker top-level para ProcessPoolExecutor (analíticos)."""
    (name, lvl_v, lvl_i, mm_v, mm_i, aa_v, aa_i, peso, inc_pp, last_mm, last_aa,
     report_month, out_png, PROFILE_YEARS, BAND_START, BAND_END,
     MM_YRS, AA_YRS, DPI, START_COMMON, MESES_ES) = args

    import matplotlib; matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import numpy as np, pandas as pd

    lvl = pd.Series(lvl_v, index=pd.DatetimeIndex(lvl_i))
    mm  = pd.Series(mm_v,  index=pd.DatetimeIndex(mm_i))
    aa  = pd.Series(aa_v,  index=pd.DatetimeIndex(aa_i))
    rpt = pd.Timestamp(report_month); sc = pd.Timestamp(START_COMMON)

    def _b(s,y0,y1):
        s=s.dropna(); mn=np.arange(1,13)
        if s.empty: n=np.full(12,np.nan); return mn,n,n,n
        df=pd.DataFrame({"v":s,"y":s.index.year,"m":s.index.month})
        h=df[(df.y>=y0)&(df.y<=y1)]
        if h.empty: n=np.full(12,np.nan); return mn,n,n,n
        g=h.groupby("m")["v"]
        return mn,g.min().reindex(range(1,13)).values,g.max().reindex(range(1,13)).values,g.mean().reindex(range(1,13)).values

    def _yl(s,years):
        s=s.dropna()
        return {y:s[s.index.year==y].groupby(s[s.index.year==y].index.month).last().reindex(range(1,13)).values for y in years}

    fig=plt.figure(figsize=(11.4,6.6),dpi=DPI); gs=fig.add_gridspec(2,2,hspace=.35,wspace=.25)
    ax1,ax2,ax3,ax4=[fig.add_subplot(gs[r,c]) for r in range(2) for c in range(2)]

    mn,vmin,vmax,vavg=_b(mm,BAND_START,BAND_END); yl=_yl(mm,PROFILE_YEARS)
    ax1.fill_between(mn,vmin,vmax,alpha=.18,linewidth=0)
    ax1.plot(mn,vavg,linewidth=2.5,label=f"Prom.{BAND_START}-{BAND_END}")
    for y in PROFILE_YEARS:
        yv=yl.get(y,np.full(12,np.nan))
        if np.isfinite(yv).sum()<=1: ax1.plot(mn,yv,linestyle="None",marker="o",markersize=8,label=str(y))
        else: ax1.plot(mn,yv,linewidth=2.5,label=str(y))
    ax1.axhline(0,linewidth=1); ax1.set_xlim(1,12); ax1.set_xticks(range(1,13))
    ax1.set_xticklabels(MESES_ES,rotation=90); ax1.set_title("m/m profile",fontsize=10)
    ax1.grid(True,alpha=.2); ax1.tick_params(labelsize=8); ax1.legend(fontsize=7,ncol=2,frameon=False)

    s2=mm.dropna(); s2=s2[s2.index>=rpt-pd.DateOffset(years=MM_YRS)]
    ax2.plot(s2.index,s2.values); ax2.axhline(0,linewidth=1)
    ax2.set_title(f"m/m — {MM_YRS}a",fontsize=10); ax2.grid(True,alpha=.2); ax2.tick_params(labelsize=8)

    s3=lvl.dropna(); s3=s3[s3.index>=sc]; ax3.plot(s3.index,s3.values)
    if not s3.empty:
        lo,hi=float(s3.quantile(.02)),float(s3.quantile(.98))
        if np.isfinite(lo) and np.isfinite(hi) and hi>lo: ax3.set_ylim(lo-.1*(hi-lo),hi+.1*(hi-lo))
    ax3.set_title("Niveles (2023=100)",fontsize=10); ax3.grid(True,alpha=.2); ax3.tick_params(labelsize=8)

    s4=aa.dropna(); s4=s4[s4.index>=rpt-pd.DateOffset(years=AA_YRS)]
    ax4.plot(s4.index,s4.values); ax4.axhline(0,linewidth=1)
    ax4.set_title(f"a/a — {AA_YRS}a",fontsize=10); ax4.grid(True,alpha=.2); ax4.tick_params(labelsize=8)

    def fn(x,nd=2): return f"{x:.{nd}f}" if isinstance(x,float) and np.isfinite(x) else "n/a"
    fig.suptitle(f"{name}  |  Peso:{fn(peso,1)}  |  m/m:{fn(last_mm)}  |  Inc:{fn(inc_pp,3)}  |  a/a:{fn(last_aa)}",
                 fontsize=11,y=.98)
    fig.savefig(out_png,bbox_inches="tight"); plt.close(fig)
    return str(out_png)

def chartbook_ana(cfg: dict):
    """Genera PDF chartbook analíticos con 4 paneles + tabla resumen (paralelo)."""
    t0 = time.time()
    PROFILE_YEARS=[2023,2024,2025,2026]; BAND_START,BAND_END=2013,2019
    MM_YRS=AA_YRS=5; DPI=150; START_COMMON="2009-01-01"
    EXCLUDE = {"ipc empalme bcch","sv empalme bcch","ipc modelo","ipc sv modelo"}
    N_WORKERS = min(4,max(1,os.cpu_count()-1))
    OUT_DIR = cfg["out_dir"]; PNG_DIR = OUT_DIR / "PNG_analiticos"
    OUT_DIR.mkdir(parents=True,exist_ok=True); PNG_DIR.mkdir(parents=True,exist_ok=True)
    ruta = cfg["ruta_xlsx"]

    if "df_indices" in _CACHE:
        df_lvl = _CACHE["df_indices"].copy()
        df_lvl["Fecha"] = pd.to_datetime(df_lvl["Fecha"])
        lvl = df_lvl.set_index("Fecha")
        mm  = _CACHE.get("df_mm_ana", pd.DataFrame())
        aa  = _CACHE.get("df_aa_ana", pd.DataFrame())
    else:
        raise RuntimeError("Ejecuta primero indices_analiticos().")

    sc_ts = pd.Timestamp(START_COMMON)
    lvl = lvl[lvl.index >= sc_ts]; mm = mm[mm.index >= sc_ts]; aa = aa[aa.index >= sc_ts]
    dates = sorted(set(lvl.index)&set(mm.index)&set(aa.index))
    if len(dates) < 2: raise ValueError("Pocos meses comunes.")
    rpt, prv = dates[-1], dates[-2]

    ipc_col = next((c for c in lvl.columns if _norm_lower(c) in ("ipc","ipc general")), None)
    if ipc_col is None: ipc_col = next((c for c in lvl.columns if "ipc" in _norm_lower(c) and "sin" not in _norm_lower(c)), None)
    ipc_prev = float(lvl.at[prv, ipc_col]) if ipc_col and prv in lvl.index else np.nan

    # Pesos desde Matriz
    wb_m = load_workbook(ruta,data_only=True); ws_m=wb_m[cfg["hoja_matriz"]]
    last_r=next((r for r in range(200,2,-1) if ws_m.cell(r,1).value not in (None,"")),174)
    pesos={}; col=4
    while True:
        h=ws_m.cell(1,col).value
        if h is None or str(h).strip()=="": break
        w_sub=np.array([_to_float(ws_m.cell(r,3).value) for r in range(3,last_r+1)],dtype=float)
        vec=np.nan_to_num([_to_float(ws_m.cell(r,col).value) for r in range(3,last_r+1)])
        pesos[_norm_lower(h)]=float(np.sum(np.nan_to_num(w_sub)*vec)); col+=1
    wb_m.close()

    cols_to_plot = [c for c in lvl.columns if _norm_lower(c) not in EXCLUDE]
    print(f"  {len(cols_to_plot)} analíticos | {rpt:%Y-%m} | workers={N_WORKERS}")
    tasks=[]
    for c in cols_to_plot:
        k=_norm_lower(c); p=pesos.get(k,np.nan)
        mm_s=mm[c] if c in mm.columns else pd.Series(dtype=float)
        aa_s=aa[c] if c in aa.columns else pd.Series(dtype=float)
        lv_s=lvl[c] if c in lvl.columns else pd.Series(dtype=float)
        last_mm=_to_float(mm_s.loc[rpt]) if rpt in mm_s.index else np.nan
        last_aa=_to_float(aa_s.loc[rpt]) if rpt in aa_s.index else np.nan
        xt=_to_float(lv_s.loc[rpt]) if rpt in lv_s.index else np.nan
        xp=_to_float(lv_s.loc[prv]) if prv in lv_s.index else np.nan
        inc=p*(xt-xp)/ipc_prev if all(np.isfinite(v) for v in [p,xt,xp,ipc_prev]) and ipc_prev!=0 else np.nan
        out_png=PNG_DIR/f"{_safe_name(c)}__4p.png"
        tasks.append((c,lv_s.values.tolist(),lv_s.index.astype(str).tolist(),
                       mm_s.values.tolist(),mm_s.index.astype(str).tolist(),
                       aa_s.values.tolist(),aa_s.index.astype(str).tolist(),
                       p,inc,last_mm,last_aa,str(rpt),str(out_png),
                       PROFILE_YEARS,BAND_START,BAND_END,MM_YRS,AA_YRS,DPI,START_COMMON,_MESES_ES))

    pngs=[]
    with ProcessPoolExecutor(max_workers=N_WORKERS) as ex:
        futs={ex.submit(_png_worker_ana,t):t[0] for t in tasks}
        for f in as_completed(futs):
            try:
                p=Path(f.result())
                if p.exists(): pngs.append(p)
            except Exception as e: print(f"  ⚠️ {futs[f]}: {e}")
    pngs=sorted(pngs)

    out_pdf=OUT_DIR/"Chartbook_Analiticos_4panel.pdf"
    cv=canvas.Canvas(str(out_pdf),pagesize=landscape(A4)); wp,hp=landscape(A4)
    cv.setFont("Helvetica-Bold",18); cv.drawString(2*cm,hp-2.7*cm,"CPI Chartbook — Analíticos")
    cv.setFont("Helvetica",11);      cv.drawString(2*cm,hp-3.9*cm,f"Último: {rpt:%b-%Y}"); cv.showPage()
    mg=1.2*cm
    for p in pngs:
        cv.drawImage(str(p),mg,mg,width=wp-2*mg,height=hp-2*mg,preserveAspectRatio=True,anchor="c")
        cv.setFont("Helvetica",8); cv.drawRightString(wp-cm,.7*cm,p.stem[:90]); cv.showPage()
    cv.save()
    print(f"  ✅ PDF: {out_pdf} — {time.time()-t0:.1f}s")


# ─────────────────────────────────────────────────────────────────────────────

# 12. PASO 9 — X-13 DESESTACIONALIZACIÓN (necesita Windows + Excel COM)
# ─────────────────────────────────────────────────────────────────────────────
def x13_velocidad(cfg: dict):
    """Desestacionaliza IPC_SV con X-13 y escribe col C en hoja Velocidad."""
    try:
        import xlwings as xw
        from statsmodels.tsa.x13 import x13_arima_analysis
    except ImportError as e:
        print(f"  ⚠️ Dependencia no disponible: {e}"); return

    t0 = time.time()
    ruta     = cfg["ruta_xlsx"]
    base_dir = cfg["code_dir"]   # x13as.exe vive en Code/
    x13_ascii = base_dir / "x13as_ascii.exe"
    x13_alias = base_dir / "x13as.exe"
    os.environ["X13PATH"] = str(base_dir)
    if x13_ascii.exists() and not x13_alias.exists():
        try: shutil.copy2(x13_ascii, x13_alias)
        except Exception as e: print(f"  ⚠️ No pude copiar x13as.exe: {e}")

    print("  1) Recalculando fórmulas con xlwings…")
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False; app.screen_updating = False
    try:
        wb = app.books.open(str(ruta), update_links=False, read_only=False)
        try: app.api.CalculateFullRebuild()
        except:
            try: app.api.CalculateFull()
            except: app.calculate()
        wb.save(); wb.close()
    finally:
        app.quit()
    print("  ✅ Recalculo COM listo.")

    # Leer valores calculados
    MESES_V = {"ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,"jul":7,"ago":8,"sep":9,"sept":9,"oct":10,"nov":11,"dic":12}
    def _parse_p(v):
        if v is None: return pd.NaT
        if isinstance(v, pd.Timestamp): return v.to_period("M").to_timestamp()
        if hasattr(v,"year"): return pd.Timestamp(int(v.year),int(v.month),1)
        s=str(v).strip().lower()
        m=re.match(r"^(ene|feb|mar|abr|may|jun|jul|ago|sep|sept|oct|nov|dic)-(\d{4})$",
                   re.sub(r"[-/_. ]+","-",s).strip("-"))
        if m: return pd.Timestamp(int(m.group(2)),MESES_V[m.group(1)],1)
        dt_=pd.to_datetime(v,errors="coerce",dayfirst=True)
        return pd.Timestamp(dt_.year,dt_.month,1) if pd.notna(dt_) else pd.NaT

    wb_v = load_workbook(ruta, data_only=True); ws_v = wb_v[cfg["hoja_velocidad"]]
    hdr  = next((r for r in range(1,min(500,ws_v.max_row)+1)
                 if str(ws_v.cell(r,1).value or "").strip().lower()=="periodo"), None)
    if hdr is None: wb_v.close(); raise ValueError("Sin 'Periodo' en col A de Velocidad")
    rows=[]
    for r in range(hdr+1, ws_v.max_row+1):
        per=ws_v.cell(r,1).value
        if per is None or str(per).strip()=="": break
        rows.append((r, _parse_p(per), _to_float(ws_v.cell(r,2).value)))
    wb_v.close()
    df=pd.DataFrame(rows,columns=["row","Fecha","IPC_SV"])
    df=df[df["Fecha"].notna()].sort_values("Fecha").drop_duplicates("Fecha",keep="last")
    n_ok=int(np.isfinite(df["IPC_SV"]).sum())
    if n_ok < 36: raise ValueError(f"Pocos valores válidos ({n_ok}). Revisar fórmulas col B.")
    print(f"  Filas={len(df)} | válidas={n_ok} | {df.Fecha.min():%Y-%m}..{df.Fecha.max():%Y-%m}")

    # X-13
    s = pd.Series(df["IPC_SV"].values, index=df["Fecha"]).asfreq("MS")
    ylog = np.log(s[s>0].copy())
    print("  2) X-13 ARIMA…")
    res  = x13_arima_analysis(ylog, freq="M", x12path=str(base_dir), prefer_x13=True)
    sa   = np.exp(res.seasadj).to_dict()

    print("  3) Escribiendo col C…")
    wb_w = load_workbook(ruta, data_only=False); ws_w = wb_w[cfg["hoja_velocidad"]]
    for _, rr in df.iterrows():
        v = sa.get(rr["Fecha"], np.nan)
        c = ws_w.cell(int(rr["row"]), 3)
        c.value = None if np.isnan(v) else float(v); c.number_format = "0.00"
    wb_w.save(ruta); wb_w.close()
    print(f"  ✅ X-13 listo — {time.time()-t0:.1f}s")


# ─────────────────────────────────────────────────────────────────────────────


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("  IPC PRINCIPAL — proceso mensual")
    print("=" * 60)
    cfg = build_config()

    pasos = [
        ("1. Empalme",               empalme),
        ("2. IPC + SV -> Analiticos", ipc_sv),
        ("2b. BCCh Analiticos D/E",   bcch_analiticos),
        ("3. Indices analiticos",     indices_analiticos),
        ("3b. Reparar #REF Analiticos", reparar_ref_analiticos),
        ("4. Variaciones subclases",  variaciones_subclases),
        ("5. Texto Flash",            texto_flash),
        ("6. X-13 Velocidad",         x13_velocidad),
        ("7. Chartbook Subclases",    chartbook_sub),
        ("8. Chartbook Analiticos",   chartbook_ana),
        ("9. Subidas y Bajadas",      subidas_bajadas),
    ]

    for nombre, fn in pasos:
        print(f"\n{chr(9472)*60}")
        print(f">> {nombre}")
        print(f"{chr(9472)*60}")
        try:
            fn(cfg)
        except Exception as e:
            print(f"  ERROR en {nombre}: {e}")
            import traceback; traceback.print_exc()

    print("\n" + "="*60)
    print("  PROCESO COMPLETO")
    print("="*60)

if __name__ == "__main__":
    main()
    
    
    