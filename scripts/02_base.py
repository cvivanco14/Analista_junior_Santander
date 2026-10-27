"""Construye la base mensual que usan los modelos.

- Bloques hoja (11) de la descomposición IPC SV / volátiles (Carlomagno, Fornero y Sansone, 2021)
  tomados de la hoja `Historia` del Excel automatizado (solo datos efectivos, hasta el último IPC
  publicado) y validados contra los agregados del BCCh descargados de la BDE.
- Pesos base 2023 = suma de ponderadores de producto de la hoja `Matriz`.
- Variables macro mensuales de la BDE; el mes en curso del dólar y del petróleo se completa con el
  promedio de los datos diarios disponibles.
Salida: data/base_modelo.csv, data/pesos_bloques.csv
"""
import os

import numpy as np
import openpyxl
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLS = os.path.join(ROOT, "data", "insumos", "IPC_2023_Automatizado_v2.xlsx")

# bloque hoja -> (nombre, columna en Matriz, agregado padre)
BLOQUES = {
    "Alim_SV": ("Alimentos no volátiles", "Alimentos no volátiles", "Bienes SV"),
    "Bs_sv_ex_alim": ("Bienes SV excl. alimentos", "Bienes sv. ex. alim.", "Bienes SV"),
    "ss_sv_ai": ("Servicios SV administrados e indexados", "Servicios_A&I_sv", "Servicios SV"),
    "ss_sv_ex_ai": ("Servicios SV resto", "Servicios_ Resto_sv", "Servicios SV"),
    "elec": ("Electricidad", "Electricidad_v", "Energía volátil"),
    "comb": ("Combustibles", "Combustible_v", "Energía volátil"),
    "alim_rest_vol": ("Alimentos volátiles resto", "Alim. Resto_v", "Alimentos volátiles"),
    "fyv": ("Frutas y verduras", "FyV (proxy)_v", "Alimentos volátiles"),
    "bs_vol": ("Bienes volátiles", "Bienes volát.", "Resto volátiles"),
    "ss_v_ai": ("Servicios volátiles adm. e indexados", "Servicios_A&I_v", "Resto volátiles"),
    "ss_v_ex_ai": ("Servicios volátiles resto", "Servicios_Resto_v", "Resto volátiles"),
}
PADRES = {
    "Bienes SV": "IPC SV", "Servicios SV": "IPC SV", "Energía volátil": "IPC volátiles",
    "Alimentos volátiles": "IPC volátiles", "Resto volátiles": "IPC volátiles",
}


def leer_hoja(wb, nombre, fila_header, max_col):
    rows = list(wb[nombre].iter_rows(min_row=fila_header, max_col=max_col, values_only=True))
    return rows[0], rows[1:]


def pesos(wb):
    hdr, rows = leer_hoja(wb, "Matriz", 1, 23)
    m = pd.DataFrame([r for r in rows[1:] if r[0] is not None], columns=hdr)
    w = m["POND 2023=100"].astype(float)
    out = pd.DataFrame(
        [{"bloque": k, "nombre": v[0], "agregado": v[2], "agregado_sup": PADRES[v[2]],
          "peso_2023": (w * m[v[1]].astype(float)).sum(),
          "n_productos": int(m[v[1]].astype(float).sum())} for k, v in BLOQUES.items()]
    )
    assert abs(out.peso_2023.sum() - 100) < 1e-6
    return out


def historia(wb, ultimo):
    hdr, rows = leer_hoja(wb, "Historia", 3, 23)
    h = pd.DataFrame([r for r in rows if r[0] is not None], columns=hdr)
    h.index = pd.to_datetime(h.pop("mes").astype(str).str[:7])
    h = h.drop(columns=[c for c in h.columns if c is None]).apply(pd.to_numeric, errors="coerce")
    return h.loc[:ultimo]


def macro(bde_m, bde_d, mes_actual):
    x = bde_m.copy()
    d = bde_d.resample("MS").mean()
    # mes en curso: promedio de los diarios disponibles (dólar observado y WTI)
    for col_m, col_d in [("tcn", "tcn_d"), ("wti", "wti_d")]:
        x.loc[mes_actual, col_m] = d.loc[mes_actual, col_d]
    # brent del mes en curso: se mueve como el WTI
    x.loc[mes_actual, "brent"] = x.loc[mes_actual - pd.DateOffset(months=1), "brent"] * (
        x.loc[mes_actual, "wti"] / x.loc[mes_actual - pd.DateOffset(months=1), "wti"])
    for c in ["bei_swap_1a", "bei_swap_2a", "bei_fw_1x1"]:
        x[c] = d[c]
    x["petroleo_clp"] = x["wti"] * x["tcn"]
    x["fao_clp"] = x["fao"] * x["tcn"]
    return x.sort_index()


if __name__ == "__main__":
    wb = openpyxl.load_workbook(XLS, read_only=True, data_only=True)
    bde_m = pd.read_csv(os.path.join(ROOT, "data", "bde", "bde_mensual.csv"), index_col=0, parse_dates=True)
    bde_d = pd.read_csv(os.path.join(ROOT, "data", "bde", "bde_diario.csv"), index_col=0, parse_dates=True)
    ultimo = bde_m["ipc"].dropna().index.max()
    mes_actual = ultimo + pd.DateOffset(months=1)

    w = pesos(wb)
    h = historia(wb, ultimo)
    hojas = h[list(BLOQUES)]

    # validación: los agregados reconstruidos desde las hojas replican los del BCCh (BDE)
    ww = w.set_index("bloque").peso_2023
    rec = (hojas * ww).sum(axis=1) / 100
    err = (rec.pct_change() - bde_m["ipc"].pct_change()).loc["2024":].abs().max() * 100
    print(f"Máx. diferencia var. m/m IPC reconstruido vs BDE (2024-): {err:.4f} pp")
    for blo, cols in {"ipc_sv": ["Alim_SV", "Bs_sv_ex_alim", "ss_sv_ai", "ss_sv_ex_ai"],
                      "ene_vol": ["elec", "comb"]}.items():
        r = (hojas[cols] * ww[cols]).sum(axis=1)
        e = (r.pct_change() - bde_m[blo].pct_change()).loc["2024":].abs().max() * 100
        print(f"  {blo}: {e:.4f} pp")

    x = macro(bde_m, bde_d, mes_actual)
    base = hojas.add_prefix("idx_").join(x, how="outer").loc["1998-12":mes_actual]
    base.index.name = "periodo"
    base.to_csv(os.path.join(ROOT, "data", "base_modelo.csv"))
    w.to_csv(os.path.join(ROOT, "data", "pesos_bloques.csv"), index=False)
    print(w.round(3).to_string(index=False))
    print(f"Base: {base.shape}, último IPC {ultimo:%Y-%m}, mes en curso {mes_actual:%Y-%m}")
