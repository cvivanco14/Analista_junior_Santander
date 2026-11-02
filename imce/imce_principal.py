# -*- coding: utf-8 -*-
"""
IMCE — Proceso mensual completo
================================
1. Descarga datos IMCE desde la API del BCCh y actualiza IMCE.xlsx
2. Genera el PDF chartbook Santander

Ejecutar (no hay que editar nada):
    python imce_principal.py

Requisitos (una vez):
    python -m pip install bcchapi pandas openpyxl requests matplotlib reportlab

Credenciales (una vez por computador, en PowerShell, con comillas SIMPLES):
    setx BDE_TOKEN 'tu_token'        # API Key Token del BCCh (tambien acepta BCCH_API_token)
  o, si se usa usuario y clave de la API en vez de token:
    setx BCCH_USER 'tu_correo'
    setx BCCH_PASS 'tu_clave'
  Luego cerrar y abrir VS Code. Las credenciales nunca se escriben en este archivo.

Carpeta: el script busca solo IMCE.xlsx en OneDrive de Santander
    <OneDrive>/.../Bases - Chartpack/Bases/Local/IMCE/Base/IMCE.xlsx
  para cualquier usuario. Si no lo encuentra, usa un IMCE.xlsx junto a este script.
  Para forzar otra ubicacion: setx IMCE_XLSX 'C:\\ruta\\IMCE.xlsx'
"""

import math
import os
import re
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import requests
from openpyxl import load_workbook
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


# ─────────────────────────────────────────────────────────────────────────────
# 1. CONFIGURACIÓN AUTOMÁTICA (no editar)
# ─────────────────────────────────────────────────────────────────────────────

# Fecha de inicio de descarga
FECHA_INICIO = "2004-01-01"

# Ruta de IMCE.xlsx dentro de la carpeta compartida de Estudios
RUTA_RELATIVA = Path("Bases - Chartpack") / "Bases" / "Local" / "IMCE" / "Base" / "IMCE.xlsx"

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

    candidatos = []
    raices = _raices_onedrive()
    for raiz in raices:
        # <raiz>/Bases - Chartpack/... o <raiz>/<biblioteca, ej. "General - Estudios">/Bases - Chartpack/...
        candidatos.append(raiz / RUTA_RELATIVA)
        candidatos += [sub / RUTA_RELATIVA for sub in raiz.iterdir() if sub.is_dir()]
    candidatos.append(Path(__file__).resolve().parent / "IMCE.xlsx")

    encontrados = [c for c in candidatos if c.is_file()]
    if not encontrados:
        buscado = "\n    ".join(str(r) for r in raices) or "(no hay carpetas de OneDrive)"
        sys.exit("No encontré IMCE.xlsx. Busqué en:\n    " + buscado +
                 f"\n(ruta esperada dentro: {RUTA_RELATIVA})\n"
                 "Solución: copia IMCE.xlsx junto a este script, o en PowerShell:\n"
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
# 5. CHARTBOOK PDF
# ─────────────────────────────────────────────────────────────────────────────

SANTANDER_RED = (236/255, 0, 0)
SANTANDER_BLACK = (0, 0, 0)
SANTANDER_OCHRE = (255/255, 174/255, 66/255)
SANTANDER_GREY = (178/255, 178/255, 178/255)
GREEN = (0, 176/255, 80/255)

DPI = 200
LW_MAIN = 2.25          # pt, series mensuales (formato Estudios)
TOP_BAR_H = 18
PAGE_MARGIN = 18
FS_TITLE = 22; FS_LEGEND = 16; FS_TICKS = 14; FS_END = 16; FS_OPT = 16
FUENTE = "Fuente: Banco Central de Chile."
FS_FUENTE = 9           # pt en la pagina del PDF
MAX_ETIQUETAS_X = 10    # fechas horizontales: maximo de etiquetas para que no se encimen

MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


def _month_label(dt):
    return f"{MESES[dt.month-1]}-{str(dt.year)[-2:]}"


def _fmt(x):
    """Un decimal, coma decimal y sin ceros sobrantes: 52,3 / 50."""
    texto = f"{round(x, 1):.1f}".rstrip("0").rstrip(".")
    return texto.replace(".", ",")


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


def _ticks(dates):
    """Etiquetas en enero, cada tantos anios como haga falta para que quepan horizontales."""
    eneros = [i for i, dt in enumerate(dates) if pd.notna(dt) and dt.month == 1]
    paso = max(1, math.ceil(len(eneros) / MAX_ETIQUETAS_X))
    idx = eneros[::-1][::paso][::-1]  # se cuenta desde el ultimo enero hacia atras
    return idx, [_month_label(pd.Timestamp(dates.iloc[i])) for i in idx]


def _ylim(y_all, include=50.0, pad=0.14):
    y = y_all[np.isfinite(y_all)]
    if y.size == 0:
        return (0, 100)
    lo = min(float(np.quantile(y, 0.02)), float(y.min()), include)
    hi = max(float(np.quantile(y, 0.98)), float(y.max()), include)
    span = max(hi - lo, 1.0)
    step = 5.0
    return (step * np.floor((lo - pad*span)/step), step * np.ceil((hi + pad*span)/step))


def _opt_pes(ax):
    ax.text(0.01, 0.98, "(Índice)", transform=ax.transAxes, ha="left", va="top",
            fontsize=FS_OPT, color=SANTANDER_BLACK)
    ax.annotate("", xy=(0.08, 0.86), xytext=(0.08, 0.72),
                xycoords=ax.transAxes, textcoords=ax.transAxes,
                arrowprops=dict(arrowstyle="simple", color=GREEN, lw=0, mutation_scale=35),
                annotation_clip=False)
    ax.text(0.11, 0.80, "Optimista", transform=ax.transAxes,
            ha="left", va="center", fontsize=FS_OPT, color=SANTANDER_BLACK)
    ax.annotate("", xy=(0.08, 0.08), xytext=(0.08, 0.22),
                xycoords=ax.transAxes, textcoords=ax.transAxes,
                arrowprops=dict(arrowstyle="simple", color=SANTANDER_RED, lw=0, mutation_scale=35),
                annotation_clip=False)
    ax.text(0.11, 0.14, "Pesimista", transform=ax.transAxes,
            ha="left", va="center", fontsize=FS_OPT, color=SANTANDER_BLACK)


def make_chart(df, title, series, out_png, legend_ncol=None, force_ylims=None):
    d = df.copy().sort_values("Fecha").reset_index(drop=True)
    x = np.arange(len(d))
    tick_idx, tick_labels = _ticks(d["Fecha"])
    y_all = np.concatenate([d[col].astype(float).to_numpy() for _, col, _, _ in series])
    ylims = force_ylims if force_ylims is not None else _ylim(y_all)

    fig, ax = plt.subplots(figsize=(12.8, 7.2), dpi=DPI)
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    for label, col, color, lw in series:
        y = d[col].astype(float).to_numpy()
        ax.plot(x, y, color=color, linewidth=lw, label=label)
        # etiqueta del ultimo dato disponible de la serie
        ult = np.flatnonzero(np.isfinite(y))
        if ult.size:
            i = ult[-1]
            ax.text(x[i]+0.55, float(y[i]), _fmt(float(y[i])),
                    color=color, va="center", fontsize=FS_END, fontweight="bold")

    ax.hlines(50, 0, len(d)-1, colors=SANTANDER_GREY, linewidth=1.6, linestyles="--")
    ax.set_title(title, fontweight="bold", pad=18, fontsize=FS_TITLE)
    ax.set_xlim(0, len(d)-1+4)
    ax.set_ylim(*ylims)
    ax.set_xticks(tick_idx)
    ax.set_xticklabels(tick_labels, rotation=0, fontsize=FS_TICKS)   # fechas horizontales
    ax.yaxis.set_major_formatter(lambda v, _: _fmt(v))
    ax.tick_params(axis="x", length=0)
    ax.tick_params(axis="y", length=0, labelsize=FS_TICKS)
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.legend(loc="upper center", ncol=legend_ncol or len(series), frameon=False,
              bbox_to_anchor=(0.5, 1.03), fontsize=FS_LEGEND, handlelength=2.2, columnspacing=1.8)
    _opt_pes(ax)
    plt.tight_layout(rect=[0.03, 0.06, 0.97, 0.92])
    fig.savefig(out_png)
    plt.close(fig)


def _fuente_pdf(estilo):
    """Santander Text / Text Light si estan instaladas en Windows; si no, Helvetica."""
    nombre = {"normal": "Santander Text", "light": "Santander Text Light"}[estilo]
    clave = nombre.lower().replace(" ", "")
    carpetas = [Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts",
                Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Windows" / "Fonts"]
    for carpeta in carpetas:
        if not carpeta.is_dir():
            continue
        for f in carpeta.iterdir():
            base = f.stem.lower().replace(" ", "").replace("-", "")
            if f.suffix.lower() == ".ttf" and (base == clave or base == clave + "regular"):
                try:
                    pdfmetrics.registerFont(TTFont(nombre, str(f)))
                    return nombre
                except Exception:  # noqa: BLE001 - fuente no utilizable: se usa Helvetica
                    pass
    return "Helvetica"


def _top_bar(c, W, H):
    c.setFillColorRGB(236/255, 0, 0)
    c.rect(0, H-TOP_BAR_H, W, TOP_BAR_H, fill=1, stroke=0)


def build_pdf(cover_month_year, pngs, out_pdf):
    W, H = landscape(A4)
    cv = canvas.Canvas(str(out_pdf), pagesize=landscape(A4))
    font_rl = _fuente_pdf("normal")
    font_fuente = _fuente_pdf("light")

    # Portada
    _top_bar(cv, W, H)
    cv.setFillColorRGB(236/255, 0, 0)
    cv.setFont(font_rl, 44)
    cv.drawString(42, H-92, "Santander")
    cv.setFillColorRGB(0, 0, 0)
    cv.setFont(font_rl, 16)
    cv.drawString(44, H-122, "Santander Estudios Económicos")
    cv.setFont(font_rl, 34)
    cv.drawCentredString(W/2, H/2+40, "Índice Mensual de Confianza Empresarial (IMCE)")
    cv.setFont(font_rl, 30)
    cv.drawCentredString(W/2, H/2-5, cover_month_year)
    cv.setFont(font_rl, 18)
    cv.drawCentredString(W/2, H/2-45, f"Actualización: {pd.Timestamp.now():%d-%m-%Y}")
    cv.showPage()

    for i, p in enumerate(pngs, 1):
        _top_bar(cv, W, H)
        img = ImageReader(str(p))
        iw, ih = img.getSize()
        usable_h = H - TOP_BAR_H - 26
        usable_w = W - 2*PAGE_MARGIN
        scale = min(usable_w/iw, usable_h/ih)
        w_img, h_img = iw*scale, ih*scale
        cv.drawImage(img, (W-w_img)/2, (usable_h-h_img)/2+18, width=w_img, height=h_img, mask="auto")
        cv.setFillColorRGB(0, 0, 0)
        cv.setFont(font_fuente, FS_FUENTE)
        cv.drawString(PAGE_MARGIN + 10, 18, FUENTE)
        cv.setFont(font_rl, 12)
        cv.drawRightString(W-24, 18, str(i))
        cv.showPage()

    cv.save()


def generar_pdf():
    df = read_imce()
    last_date = df["Fecha"].max()
    cover = f"{MESES[last_date.month-1].upper()} {last_date.year}"
    ts = pd.Timestamp.now().strftime("%Y%m%d")
    pngs = []

    charts = [
        ("01_total_vs_nomi", "IMCE total y sin minería",
         [("IMCE total",       find_col(df, ["IMCE Total", "IMCE: Total", "Total"]),    SANTANDER_RED,   LW_MAIN),
          ("IMCE sin minería", find_col(df, ["IMCE: Sin Minería", "Sin Minería", "IMCE: Sin Mineria"]),
           SANTANDER_BLACK, LW_MAIN)],
         2, None),
        ("02_sectores", "IMCE por sector económico",
         [("Comercio",     find_col(df, ["Comercio"]),     SANTANDER_BLACK, LW_MAIN),
          ("Construcción", find_col(df, ["Construcción", "Construccion"]), SANTANDER_RED, LW_MAIN),
          ("Industria",    find_col(df, ["Industria"]),    SANTANDER_OCHRE, LW_MAIN)],
         3, None),
        ("03_empleo", "Expectativas de empleo",
         [("Comercio",     find_col(df, ["COM_Expectativas - Empleo"]),  SANTANDER_BLACK, LW_MAIN),
          ("Construcción", find_col(df, ["CONS_Expectativas - Empleo"]), SANTANDER_RED,   LW_MAIN),
          ("Industria",    find_col(df, ["IND_Expectativas - Empleo"]),  SANTANDER_OCHRE, LW_MAIN)],
         3, (0, 100)),
        ("04_inventarios", "Inventarios (vs deseado)",
         [("Comercio",  find_col(df, ["COM_Situacion actual - Inventarios"]), SANTANDER_BLACK, LW_MAIN),
          ("Industria", find_col(df, ["IND_Situacion actual - Inventarios"]), SANTANDER_RED,   LW_MAIN)],
         2, (0, 100)),
        ("05_salarios", "Presiones salariales",
         [("Construcción", find_col(df, ["CONS_Expectativas - Salarios"]), SANTANDER_RED,   LW_MAIN),
          ("Industria",    find_col(df, ["IND_Expectativas - Salarios"]),  SANTANDER_OCHRE, LW_MAIN)],
         2, (0, 100)),
        ("06_demanda", "Nivel de demanda percibida",
         [("Comercio",     find_col(df, ["COM_Situacion actual - Demanda"]),  SANTANDER_BLACK, LW_MAIN),
          ("Construcción", find_col(df, ["CONS_Situacion actual - Demanda"]), SANTANDER_RED,   LW_MAIN),
          ("Industria",    find_col(df, ["IND_Situacion actual - Demanda"]),  SANTANDER_OCHRE, LW_MAIN)],
         3, (0, 100)),
        ("07_precios", "Expectativas de precios",
         [("Comercio",     find_col(df, ["COM_Expectativas - Precios"]),  SANTANDER_BLACK, LW_MAIN),
          ("Construcción", find_col(df, ["CONS_Expectativas - Precios"]), SANTANDER_RED,   LW_MAIN),
          ("Industria",    find_col(df, ["IND_Expectativas - Precios"]),  SANTANDER_OCHRE, LW_MAIN)],
         3, (0, 100)),
        ("08_costos", "Expectativas de costos",
         [("Comercio",     find_col(df, ["COM_Expectativas - Costos"]),  SANTANDER_BLACK, LW_MAIN),
          ("Construcción", find_col(df, ["CONS_Expectativas - Costos"]), SANTANDER_RED,   LW_MAIN),
          ("Industria",    find_col(df, ["IND_Expectativas - Costos"]),  SANTANDER_OCHRE, LW_MAIN)],
         3, (0, 100)),
    ]

    for fname, title, series, ncol, ylims in charts:
        cols_needed = {col for _, col, _, _ in series}
        df_c = df[["Fecha"] + list(cols_needed)].copy()
        out_png = OUT_DIR / f"{fname}_{ts}.png"
        make_chart(df_c, title, series, out_png, legend_ncol=ncol, force_ylims=ylims)
        pngs.append(out_png)
        print(f"  📊 {title}")

    out_pdf = OUT_DIR / f"IMCE_Chartbook_Santander_{ts}.pdf"
    build_pdf(cover, pngs, out_pdf)
    print(f"\n  ✅ PDF: {out_pdf}")


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
    print(">> 2. Generando PDF chartbook")
    print(f"{'─'*60}")
    generar_pdf()

    print("\n" + "="*60)
    print("  PROCESO COMPLETO")
    print("="*60)


if __name__ == "__main__":
    main()
