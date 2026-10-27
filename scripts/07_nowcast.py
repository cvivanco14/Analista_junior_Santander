"""Nowcasting del IPC de corto plazo con datos de alta frecuencia.

1. Combustibles (ENAP): con las variaciones semanales de los informes de ENAP (data/nowcast/enap_informes.csv)
   y un precio ancla a público (CNE), se reconstruye el precio diario de gasolina, diésel y parafina. El
   promedio del mes se compara con el del mes anterior, como hace el INE con sus recolecciones semanales.
   - Backtest: jul-sep 2026 contra los micro-índices del INE.
   - Escenarios para el próximo ajuste del MEPCO (12-nov y 3-dic): repetir el último alza o sin cambio.
2. Supermercados: si hay fotos diarias de precios (scripts/06_scraper_supermercados.py), se calcula un
   índice Jevons por producto IPC (como el INE bajo el nivel de producto) y se agrega a los bloques de
   alimentos con pesos efectivos. Los productos sin cobertura usan la proyección del modelo.
3. Puente: se reemplaza la proyección de los bloques nowcasteados en oct-dic 2026 dentro de la proyección
   combinada (03_modelo.py); desde enero de 2027 se mantienen las tasas del modelo.
Salida: output/nowcast/*.csv y hojas nuevas en output/Base_Proyeccion_IPC.xlsx
"""
import glob
import os

import numpy as np
import openpyxl
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NC = os.path.join(ROOT, "data", "nowcast")
OUT = os.path.join(ROOT, "output", "nowcast")
XLS = os.path.join(ROOT, "data", "insumos", "IPC_2023_Automatizado_v2.xlsx")
os.makedirs(OUT, exist_ok=True)

PROD_ENAP = {"gasolina": 72221, "diesel": 72211, "kerosene": 44311}
GLOSA = {72221: "Gasolina", 72211: "Petróleo diésel", 44311: "Parafina", 44221: "Gas licuado", 44211: "Gas por red"}
# próximo ajuste MEPCO de gasolinas y diésel: cada tres semanas desde el 22-oct
AJUSTES_FUTUROS = ["2026-11-12", "2026-12-03", "2026-12-24"]
HORIZONTE = pd.period_range("2026-07", "2026-12", freq="M")

base = pd.read_csv(os.path.join(ROOT, "data", "base_modelo.csv"), index_col=0, parse_dates=True)
ULT = base["idx_Alim_SV"].dropna().index.max().to_period("M")


# --------------------------------------------------------------------------- micro-índices INE
def micro_indices():
    wb = openpyxl.load_workbook(XLS, read_only=True, data_only=True)
    rows = list(wb["Índice 2023"].iter_rows(min_row=1, max_row=284, max_col=110, values_only=True))
    cols = [c.strftime("%Y-%m") if hasattr(c, "strftime") else c for c in rows[0]]
    df = pd.DataFrame([r for r in rows[1:] if r[0]], columns=cols).set_index("Code")
    pond = df["POND 2023=100"].astype(float)
    idx = df.loc[:, "2018-01":ULT.strftime("%Y-%m")].apply(pd.to_numeric, errors="coerce").T
    idx.index = pd.PeriodIndex(idx.index, freq="M")
    return idx, pond


IDX, POND = micro_indices()


# --------------------------------------------------------------------------- ENAP
def precio_diario(cambios, ancla, desde, hasta, extra=None):
    """Precio diario a público: ancla ± variaciones ENAP según su fecha de vigencia."""
    dias = pd.date_range(desde, hasta, freq="D")
    c = cambios.copy()
    if extra is not None:
        c = pd.concat([c, extra])
    c = c.groupby("fecha_vigencia").variacion_clp_litro.sum()
    acum = pd.Series(0.0, index=dias)
    for f, v in c.items():
        acum[acum.index >= f] += v
    return acum - acum[pd.Timestamp(ancla.fecha)] + ancla.precio_clp_litro


def nowcast_enap():
    inf = pd.read_csv(os.path.join(NC, "enap_informes.csv"), parse_dates=["fecha_vigencia"])
    anc = pd.read_csv(os.path.join(NC, "precios_ancla.csv"), parse_dates=["fecha"]).set_index("producto")
    filas, diarios = [], {}
    for esc, mult in [("Base: repite último alza", 1.0), ("Sin nuevas alzas", 0.0)]:
        for prod, cod in PROD_ENAP.items():
            ci = inf[inf.producto == prod]
            extra = None
            if prod in ("gasolina", "diesel"):
                ult = ci.sort_values("fecha_vigencia").iloc[-1].variacion_clp_litro
                extra = pd.DataFrame({"fecha_vigencia": pd.to_datetime(AJUSTES_FUTUROS),
                                      "variacion_clp_litro": ult * mult})
            p = precio_diario(ci, anc.loc[prod], "2026-06-01", "2026-12-31", extra)
            diarios[(esc, prod)] = p
            mensual = p.groupby(p.index.to_period("M")).mean()
            var = mensual.pct_change() * 100
            for per in HORIZONTE:
                ine = IDX[cod].pct_change().get(per, np.nan) * 100 if per <= ULT else np.nan
                filas.append({"escenario": esc, "producto": prod, "codigo_ipc": cod, "periodo": str(per),
                              "precio_promedio": mensual[per], "var_mm_enap": var[per], "var_mm_ine": ine})
    res = pd.DataFrame(filas)
    d = pd.DataFrame({f"{p} | {e}": s for (e, p), s in diarios.items()})
    d.index.name = "fecha"
    return res, d


# --------------------------------------------------------------------------- supermercados
def nowcast_supermercados():
    archivos = sorted(glob.glob(os.path.join(NC, "supermercados", "*.csv")))
    if not archivos:
        return None, "Sin fotos de precios: el scraper no pudo conectarse a las tiendas (dominios bloqueados)."
    p = pd.concat(pd.read_csv(a, parse_dates=["fecha"]) for a in archivos)
    p["periodo"] = p.fecha.dt.to_period("M")
    # promedio mensual por SKU y tienda; relativo de cada SKU presente en ambos meses; Jevons por producto
    m = p.groupby(["tienda", "codigo_ipc", "sku", "periodo"]).precio.mean().unstack("periodo")
    if m.shape[1] < 2:
        return None, f"Solo hay datos de {m.columns[0]}: se necesitan al menos dos meses para una variación."
    rel = np.log(m).diff(axis=1)
    jev = (np.exp(rel.groupby(level="codigo_ipc").mean()) - 1) * 100
    cobertura = rel.notna().groupby(level="codigo_ipc").sum()
    return pd.concat({"var_mm": jev, "n_sku": cobertura}, axis=1), "OK"


# --------------------------------------------------------------------------- puente a la proyección
def bloque_comb(enap, escenario):
    """Variación mensual del bloque combustibles: productos ENAP + gas (tendencia de 12 meses)."""
    cods = [72221, 72211, 44311, 44221, 44211]
    w_ef = POND[cods] * IDX.loc[ULT, cods] / 100
    w_ef = w_ef / w_ef.sum()
    out = {}
    for per in pd.period_range(ULT + 1, "2026-12", freq="M"):
        v = {}
        for prod, cod in PROD_ENAP.items():
            v[cod] = enap[(enap.escenario == escenario) & (enap.producto == prod) & (enap.periodo == str(per))].var_mm_enap.iloc[0]
        for cod in (44221, 44211):
            v[cod] = IDX[cod].pct_change().loc[ULT - 11:ULT].mean() * 100
        out[per] = sum(w_ef[c] * v[c] for c in cods)
    return pd.Series(out), w_ef


def puente(comb_mm):
    pm = pd.read_csv(os.path.join(ROOT, "output", "proy_bloques_modelos.csv"), parse_dates=["periodo"])
    pw = pd.read_csv(os.path.join(ROOT, "data", "pesos_bloques.csv")).set_index("bloque").peso_2023
    q = pm[pm.modelo == "Combinado"].pivot(index="periodo", columns="bloque", values="log_indice")
    lv = np.exp(q / 100)
    lv.index = lv.index.to_period("M")
    nuevo = lv.copy()
    g = lv["comb"].pct_change()
    for per in lv.index[lv.index > ULT]:
        tasa = comb_mm[per] / 100 if per in comb_mm.index else g[per]
        nuevo.loc[per, "comb"] = nuevo.loc[per - 1, "comb"] * (1 + tasa)
    ipc = {k: (x[pw.index] * pw).sum(axis=1) / 100 for k, x in [("modelo", lv), ("nowcast", nuevo)]}
    sv = ["Alim_SV", "Bs_sv_ex_alim", "ss_sv_ai", "ss_sv_ex_ai"]
    vol = [b for b in pw.index if b not in sv]
    filas = []
    for per in lv.index[lv.index > ULT]:
        r = {"periodo": str(per)}
        for k, x in [("modelo", lv), ("nowcast", nuevo)]:
            r[f"IPC_mm_{k}"] = (ipc[k][per] / ipc[k][per - 1] - 1) * 100
            r[f"IPC_12m_{k}"] = (ipc[k][per] / ipc[k][per - 12] - 1) * 100 if per - 12 in ipc[k].index else np.nan
            r[f"comb_mm_{k}"] = (x.loc[per, "comb"] / x.loc[per - 1, "comb"] - 1) * 100
            r[f"incid_comb_{k}"] = (x.loc[per, "comb"] - x.loc[per - 1, "comb"]) * pw["comb"] / ipc[k][per - 1]
            vv = (x[vol] * pw[vol]).sum(axis=1)
            r[f"volatiles_12m_{k}"] = (vv[per] / vv[per - 12] - 1) * 100 if per - 12 in vv.index else np.nan
        filas.append(r)
    return pd.DataFrame(filas)


# --------------------------------------------------------------------------- principal
if __name__ == "__main__":
    enap, diarios = nowcast_enap()
    enap.to_csv(os.path.join(OUT, "enap_productos.csv"), index=False)
    diarios.to_csv(os.path.join(OUT, "enap_precio_diario.csv"))
    bt = enap[(enap.escenario == enap.escenario.iloc[0]) & enap.var_mm_ine.notna() & (enap.periodo >= "2026-07")]
    bt = bt.assign(error_pp=bt.var_mm_ine - bt.var_mm_enap)
    bt.to_csv(os.path.join(OUT, "enap_backtest.csv"), index=False)
    print(bt[["producto", "periodo", "var_mm_enap", "var_mm_ine", "error_pp"]].round(2).to_string(index=False))

    super_tab, estado = nowcast_supermercados()
    print("Supermercados:", estado)
    pd.DataFrame({"estado": [estado]}).to_csv(os.path.join(OUT, "supermercados_estado.csv"), index=False)

    resultados, comb_esc = [], {}
    for esc in enap.escenario.unique():
        comb, w_ef = bloque_comb(enap, esc)
        comb_esc[esc] = comb
        r = puente(comb)
        r.insert(0, "escenario", esc)
        resultados.append(r)
    res = pd.concat(resultados)
    res.to_csv(os.path.join(OUT, "puente_ipc.csv"), index=False)
    pd.DataFrame({"codigo_ipc": w_ef.index, "producto": [GLOSA[c] for c in w_ef.index], "peso_efectivo_en_bloque": w_ef.values}
                 ).to_csv(os.path.join(OUT, "pesos_comb.csv"), index=False)
    print(res[res.periodo <= "2027-01"].round(2).to_string(index=False))

    # hojas en la base Excel
    xl = os.path.join(ROOT, "output", "Base_Proyeccion_IPC.xlsx")
    with pd.ExcelWriter(xl, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        pd.read_csv(os.path.join(NC, "enap_informes.csv")).to_excel(w, sheet_name="NC_ENAP_informes", index=False)
        pd.read_csv(os.path.join(NC, "precios_ancla.csv")).to_excel(w, sheet_name="NC_ENAP_anclas", index=False)
        enap.to_excel(w, sheet_name="NC_ENAP_productos", index=False)
        bt.to_excel(w, sheet_name="NC_ENAP_backtest", index=False)
        res.to_excel(w, sheet_name="NC_puente_IPC", index=False)
        pd.read_csv(os.path.join(NC, "canasta_supermercado.csv")).to_excel(w, sheet_name="NC_canasta_super", index=False)
        pd.DataFrame({"estado": [estado]}).to_excel(w, sheet_name="NC_super_estado", index=False)
    print("Excel actualizado:", xl)
