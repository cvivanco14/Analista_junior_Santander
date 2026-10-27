"""Genera el reporte HTML de la proyección con el mismo formato que Apuntes_IPC_Chile.html.

El cuerpo se escribe en Markdown, se convierte con pandoc y se inserta en la plantilla del
documento de apuntes (R Markdown / Bootstrap), con los gráficos embebidos en base64.
Salida: output/Proyeccion_IPC_Chile.html
"""
import base64
import io
import os
import subprocess

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "output")
TAB = os.path.join(OUT, "tablas")
PLANTILLA = os.path.join(ROOT, "data", "insumos", "Apuntes_IPC_Chile.html")

# paleta de referencia (modo claro) y tinta
AZUL, NARANJA, AGUA, AMARILLO, VIOLETA = "#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#6250d6"
TINTA, TINTA2, GRILLA = "#0b0b0b", "#52514e", "#e4e3df"
plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": GRILLA, "axes.labelcolor": TINTA2,
    "xtick.color": TINTA2, "ytick.color": TINTA2, "axes.spines.top": False, "axes.spines.right": False,
    "axes.grid": True, "grid.color": GRILLA, "grid.linewidth": 0.6, "axes.axisbelow": True,
    "legend.frameon": False, "figure.dpi": 110,
})


def t(name):
    return pd.read_csv(os.path.join(TAB, f"{name}.csv"))


def f(x, d=1):
    """Número con coma decimal (formato de los apuntes)."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "–"
    return f"{x:,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def png(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=150)
    plt.close(fig)
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


def tabla_md(df, alin=None):
    cols = list(df.columns)
    alin = alin or ["l"] + ["r"] * (len(cols) - 1)
    sep = ["---:" if a == "r" else ":---" for a in alin]
    lineas = ["| " + " | ".join(cols) + " |", "| " + " | ".join(sep) + " |"]
    for _, r in df.iterrows():
        lineas.append("| " + " | ".join(str(v) for v in r.values) + " |")
    return "\n".join(lineas)


MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


def mes_txt(p):
    p = pd.Timestamp(p)
    return f"{MESES[p.month - 1]}-{str(p.year)[2:]}"


# ------------------------------------------------------------------ datos
proy, bandas, inc, comp = t("proyeccion"), t("bandas"), t("incidencias"), t("comparacion")
rmse, rmse_rec, dm = t("rmse"), t("rmse_pos2023"), t("dm")
dif, val, pcmp = t("difusion"), t("validacion"), t("pesos_cmp")
agr = t("agregados_central")
for x in (bandas, dif, agr):
    x["periodo"] = pd.to_datetime(x["periodo"])
agr = agr.set_index("periodo")
a12 = agr.pct_change(12) * 100
proy["periodo_dt"] = pd.to_datetime(proy.periodo)
base = pd.read_csv(os.path.join(ROOT, "data", "base_modelo.csv"), index_col=0, parse_dates=True)
cat = pd.read_csv(os.path.join(ROOT, "data", "bde", "catalogo_series.csv"))
grupos = pd.read_csv(os.path.join(OUT, "tendencias_comunes_grupos.csv"))
pares = pd.read_csv(os.path.join(OUT, "tendencias_comunes_pares.csv"))
pcomb = pd.read_csv(os.path.join(OUT, "pesos_combinacion.csv"))
sup = pd.read_csv(os.path.join(OUT, "supuestos_exogenos.csv"), index_col=0, parse_dates=True)
pw = pd.read_csv(os.path.join(ROOT, "data", "pesos_bloques.csv"))
NOMBRE = dict(zip(pw.bloque, pw.nombre))
ULT = base["idx_Alim_SV"].dropna().index.max()
P = proy.set_index("periodo")
B = bandas.set_index("periodo")
esc = {m: t(f"a12_{m}").assign(periodo=lambda x: pd.to_datetime(x.periodo)).set_index("periodo")
       for m in ["M0_estacional", "M1_AR_est", "M2_ARDL_macro", "M3_ECM_tc", "Combinado"]}

dic26, dic27 = P.loc["2026-12", "IPC_var_12m"], P.loc["2027-12", "IPC_var_12m"]
sv26, sv27 = P.loc["2026-12", "IPC sin volátiles_12m"], P.loc["2027-12", "IPC sin volátiles_12m"]
oct_mm, nov_mm, dic_mm = P.loc["2026-10", "IPC_var_mm"], P.loc["2026-11", "IPC_var_mm"], P.loc["2026-12", "IPC_var_mm"]
tcn_sep, tcn_oct = base.loc["2026-09-01", "tcn"], base.loc["2026-10-01", "tcn"]
wti_sep, wti_oct = base.loc["2026-09-01", "wti"], base.loc["2026-10-01", "wti"]
ipc_sep_12 = a12.loc[ULT, "IPC"]
sv_sep_12 = a12.loc[ULT, "IPC sin volátiles"]
eee = base.loc["2026-10-01"]
cmp_i = comp.set_index("indicador")

# ------------------------------------------------------------------ gráficos
figs = {}

# 1. IPC 12m con abanico
fig, ax = plt.subplots(figsize=(8.6, 4.2))
h = a12.loc["2023-01":ULT, "IPC"]
ax.axhspan(2, 4, color="#f1f0ec", zorder=0)
ax.axhline(3, color=TINTA2, lw=0.8, ls=(0, (4, 3)))
bb = bandas.set_index("periodo")
x0 = pd.Index([ULT]).append(bb.index)
def con_origen(s):
    return np.r_[h.iloc[-1], s.values]
ax.fill_between(x0, con_origen(bb.p10_pos2023), con_origen(bb.p90_pos2023), color=AZUL, alpha=0.14, lw=0,
                label="Banda 10-90 (errores fuera de muestra, jul-23 a sep-26)")
ax.fill_between(x0, con_origen(bb.p30_pos2023), con_origen(bb.p70_pos2023), color=AZUL, alpha=0.28, lw=0,
                label="Banda 30-70")
ax.plot(h.index, h.values, color=TINTA, lw=2, label="IPC efectivo")
ax.plot(x0, con_origen(bb.central), color=AZUL, lw=2, label="Proyección central")
ax.plot(x0, con_origen(esc["M3_ECM_tc"].loc[bb.index, "IPC"]), color=VIOLETA, lw=1.5, ls=(0, (3, 2)),
        label="Escenario M3 (tendencia común)")
for fecha, v, lab in [("2026-12-01", eee.eee_dic26, "EEE dic-26"), ("2027-12-01", eee.eee_dic27, "EEE dic-27")]:
    ax.plot(pd.Timestamp(fecha), v, "o", ms=8, color=NARANJA, mec="white", mew=2, zorder=5)
    ax.annotate(f"{lab}: {f(v)}%", (pd.Timestamp(fecha), v), xytext=(10, -26), textcoords="offset points",
                color=TINTA2, fontsize=8.5, ha="left")
ax.annotate(f"{f(dic26)}%", (pd.Timestamp("2026-12-01"), dic26), xytext=(-6, 9), textcoords="offset points",
            color=TINTA, fontsize=9, ha="right", weight="bold")
ax.annotate(f"{f(dic27)}%", (pd.Timestamp("2027-12-01"), dic27), xytext=(4, 8), textcoords="offset points",
            color=TINTA, fontsize=9, weight="bold")
ax.set_ylabel("Variación en 12 meses (%)")
ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
ax.set_xlim(pd.Timestamp("2023-01-01"), pd.Timestamp("2028-03-01"))
ax.set_ylim(1, 9)
ax.legend(loc="upper right", fontsize=8.5, ncol=2)
ax.set_title("IPC: inflación en 12 meses, efectiva y proyectada", loc="left", fontsize=11, color=TINTA)
figs["abanico"] = png(fig)

# 2. variación mensual proyectada vs EEE y proyección previa
fig, ax = plt.subplots(figsize=(8.6, 3.6))
xs = np.arange(len(proy))
ax.bar(xs, proy.IPC_var_mm, width=0.62, color=AZUL, label="Proyección central")
ax.plot([0, 1], [eee.eee_mm_mes, eee.eee_mm_sig], "o", ms=8, color=NARANJA, mec="white", mew=2, label="EEE (oct-26)")
prev = comp.set_index("indicador")["proyeccion_previa_excel"]
ax.plot([0, 1], [prev["IPC oct-26, var. m/m"], prev["IPC nov-26, var. m/m"]], "D", ms=7, color=AGUA, mec="white",
        mew=2, label="Proyección previa (Excel)")
for i, v in enumerate(proy.IPC_var_mm):
    ax.text(i, v + (0.02 if v >= 0 else -0.06), f(v, 2), ha="center", fontsize=7.5, color=TINTA2)
ax.set_xticks(xs, [mes_txt(p) for p in proy.periodo_dt], rotation=0, fontsize=8)
ax.axhline(0, color=TINTA2, lw=0.8)
ax.set_ylabel("Variación mensual (%)")
ax.set_ylim(0, 0.8)
ax.legend(fontsize=8.5, loc="upper center", ncol=3)
ax.set_title("IPC: variación mensual proyectada", loc="left", fontsize=11, color=TINTA)
figs["mensual"] = png(fig)

# 3. incidencias en la inflación de diciembre
fig, ax = plt.subplots(figsize=(8.6, 4.4))
ii = inc.iloc[::-1]
y = np.arange(len(ii))
ax.barh(y + 0.19, ii.incid_12m_dic26, height=0.36, color=AZUL, label="Dic-26")
ax.barh(y - 0.19, ii.incid_12m_dic27, height=0.36, color=NARANJA, label="Dic-27")
ax.set_yticks(y, ii.nombre, fontsize=8.5)
ax.axvline(0, color=TINTA2, lw=0.8)
ax.set_xlabel("Incidencia en la variación en 12 meses del IPC (pp)")
ax.legend(fontsize=8.5, loc="lower right")
ax.set_title("Aporte de cada bloque a la inflación de diciembre", loc="left", fontsize=11, color=TINTA)
figs["incidencias"] = png(fig)

# 4. SV vs volátiles
fig, ax = plt.subplots(figsize=(8.6, 3.8))
rng = slice("2022-01", "2027-12")
for col, color, lab in [("IPC", TINTA2, "IPC"), ("IPC sin volátiles", AZUL, "IPC sin volátiles"),
                        ("IPC volátiles", NARANJA, "IPC volátiles")]:
    s = a12.loc[rng, col]
    ax.plot(s.loc[:ULT].index, s.loc[:ULT], color=color, lw=2, label=lab)
    ax.plot(s.loc[ULT:].index, s.loc[ULT:], color=color, lw=2, ls=(0, (3, 2)))
ax.axvline(ULT, color=GRILLA, lw=1)
ax.axhline(3, color=TINTA2, lw=0.8, ls=(0, (4, 3)))
ax.set_ylabel("Variación en 12 meses (%)")
ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
ax.set_xlim(pd.Timestamp("2022-01-01"), pd.Timestamp("2028-01-01"))
ax.legend(fontsize=8.5, loc="upper right")
ax.set_title("Sin volátiles y volátiles (línea discontinua: proyección)", loc="left", fontsize=11, color=TINTA)
figs["sv_vol"] = png(fig)

# 5. RMSE por horizonte
fig, ax = plt.subplots(figsize=(8.6, 3.6))
r = rmse_rec[rmse_rec.h <= 12]
for m, color, lab in [("Combinado", AZUL, "Combinado (central)"), ("M3_ECM_tc", VIOLETA, "M3 tendencia común"),
                      ("D1_AR_est_agregado", NARANJA, "AR estacional agregado"),
                      ("D2_paseo_aleatorio_12m", AMARILLO, "Paseo aleatorio 12m")]:
    s = r[r.modelo == m].set_index("h").rmse_12m
    ax.plot(s.index, s.values, color=color, lw=2, marker="o", ms=4, label=lab)
ax.set_xlabel("Horizonte (meses)")
ax.set_ylabel("RMSE inflación 12m (pp)")
ax.set_xticks(range(1, 13))
ax.legend(fontsize=8.5, loc="upper left")
ax.set_title("Error fuera de muestra, orígenes jul-23 a ago-26", loc="left", fontsize=11, color=TINTA)
figs["rmse"] = png(fig)

# ------------------------------------------------------------------ tablas
filas = []
for _, rr in proy.iterrows():
    filas.append({"Mes": mes_txt(rr.periodo_dt), "IPC m/m": f(rr.IPC_var_mm, 2), "IPC 12m": f(rr.IPC_var_12m),
                  "Acum. año": f(rr.IPC_var_acum), "IPC SV m/m": f(rr["IPC sin volátiles_mm"], 2),
                  "IPC SV 12m": f(rr["IPC sin volátiles_12m"]), "Volátiles 12m": f(rr["IPC volátiles_12m"]),
                  "Banda 10-90 (12m)": f"{f(rr.banda_12m_p10_pos2023)} – {f(rr.banda_12m_p90_pos2023)}"})
t_proy = tabla_md(pd.DataFrame(filas))

agr_tab = []
for a in ["IPC sin volátiles", "Bienes SV", "Servicios SV", "IPC volátiles", "Energía volátil",
          "Alimentos volátiles", "Resto volátiles"]:
    agr_tab.append({"Agregado": a, "Sep-26 (efectivo)": f(a12.loc[ULT, a]), "Dic-26": f(a12.loc["2026-12-01", a]),
                    "Jun-27": f(a12.loc["2027-06-01", a]), "Dic-27": f(a12.loc["2027-12-01", a])})
agr_tab.insert(0, {"Agregado": "**IPC**", "Sep-26 (efectivo)": f(ipc_sep_12), "Dic-26": f"**{f(dic26)}**",
                   "Jun-27": f(a12.loc["2027-06-01", "IPC"]), "Dic-27": f"**{f(dic27)}**"})
t_agr = tabla_md(pd.DataFrame(agr_tab))

t_inc = tabla_md(pd.DataFrame({
    "Bloque": inc.nombre, "Peso 2023": inc.peso_2023.map(lambda v: f(v, 2)),
    "Peso efectivo oct-26": inc.peso_efectivo_oct26.map(lambda v: f(v, 2)),
    "Var. 12m dic-26": inc.var_12m_dic26.map(f), "Incid. dic-26 (pp)": inc.incid_12m_dic26.map(lambda v: f(v, 2)),
    "Var. 12m dic-27": inc.var_12m_dic27.map(f), "Incid. dic-27 (pp)": inc.incid_12m_dic27.map(lambda v: f(v, 2)),
}))

esc_rows = []
for m, lab in [("M0_estacional", "M0 Estacional anclado"), ("M1_AR_est", "M1 AR estacional"),
               ("M2_ARDL_macro", "M2 ARDL con impulsores macro"), ("M3_ECM_tc", "M3 ECM tendencia común"),
               ("Combinado", "**Combinado (central)**")]:
    e = esc[m]
    esc_rows.append({"Modelo": lab, "Oct-26 12m": f(e.loc["2026-10-01", "IPC"]), "Dic-26": f(e.loc["2026-12-01", "IPC"]),
                     "Jun-27": f(e.loc["2027-06-01", "IPC"]), "Dic-27": f(e.loc["2027-12-01", "IPC"]),
                     "IPC SV dic-27": f(e.loc["2027-12-01", "IPC sin volátiles"])})
t_esc = tabla_md(pd.DataFrame(esc_rows))

t_comp = tabla_md(pd.DataFrame({
    "Indicador": comp.indicador, "Esta proyección": comp.esta_proyeccion.map(lambda v: f(v, 2)),
    "EEE (oct-26)": comp.EEE_oct26.map(f), "Proyección previa del Excel": comp.proyeccion_previa_excel.map(lambda v: f(v, 2)),
    "Compensación swap 1 año": comp.get("compensacion_swap_1a", pd.Series([np.nan] * len(comp))).map(lambda v: f(v, 2)),
}))

def rm(df, m, h, c):
    s = df[(df.modelo == m) & (df.h == h)][c]
    return f(s.iloc[0], 2) if len(s) else "–"

ev_rows = []
for m, lab in [("Combinado", "**Combinado (central)**"), ("M3_ECM_tc", "M3 ECM tendencia común"),
               ("M2_ARDL_macro", "M2 ARDL macro"), ("M1_AR_est", "M1 AR estacional"), ("M0_estacional", "M0 Estacional"),
               ("D1_AR_est_agregado", "AR estacional sobre el IPC agregado"),
               ("D2_paseo_aleatorio_12m", "Paseo aleatorio (inflación 12m constante)"), ("EEE", "EEE (mediana)")]:
    ev_rows.append({"Modelo": lab, "m/m h=1 (20-26)": rm(rmse, m, 1, "rmse_mm"), "12m h=3 (20-26)": rm(rmse, m, 3, "rmse_12m"),
                    "12m h=12 (20-26)": rm(rmse, m, 12, "rmse_12m"), "m/m h=1 (23-26)": rm(rmse_rec, m, 1, "rmse_mm"),
                    "12m h=3 (23-26)": rm(rmse_rec, m, 3, "rmse_12m"), "12m h=12 (23-26)": rm(rmse_rec, m, 12, "rmse_12m")})
t_eval = tabla_md(pd.DataFrame(ev_rows))

dm_show = dm[dm.rival.isin(["D1_AR_est_agregado", "D2_paseo_aleatorio_12m", "M3_ECM_tc", "EEE"])].copy()
dm_show["rival"] = dm_show.rival.map({"D1_AR_est_agregado": "AR agregado", "D2_paseo_aleatorio_12m": "Paseo aleatorio",
                                      "M3_ECM_tc": "M3 tendencia común", "EEE": "EEE"})
t_dm = tabla_md(pd.DataFrame({"Muestra (orígenes)": dm_show.muestra_origenes, "Rival": dm_show.rival,
                              "Horizonte": dm_show.h, "Variable": dm_show.variable,
                              "DM": dm_show.estadistico_DM.map(lambda v: f(v, 2)),
                              "p-valor": dm_show.p_valor.map(lambda v: f(v, 3))}))

t_grupos = tabla_md(pd.DataFrame({
    "Bloque": grupos.bloque.map(NOMBRE),
    "Comparte tendencia con (ADF, 5%)": grupos.comparte_tendencia_con.fillna("").apply(
        lambda s: ", ".join(NOMBRE[x.strip()] for x in s.split(",") if x.strip()) or "— (modelo sin restricción)"),
}))

pc = pcomb.copy()
pc["bloque"] = pc.bloque.map(NOMBRE)
pc.loc[pc.bloque == "Electricidad", "modelo"] = "Regla"
piv = pc.pivot(index="bloque", columns="modelo", values="peso").reindex(columns=["M0_estacional", "M1_AR_est",
                                                                                   "M2_ARDL_macro", "M3_ECM_tc", "Regla"])
piv = piv.reindex([NOMBRE[b] for b in pw.bloque])
t_pc = tabla_md(pd.DataFrame({"Bloque": piv.index, **{c: piv[c].map(lambda v: "–" if np.isnan(v) else f(100 * v, 0) + "%")
                                                      for c in piv.columns}}))

t_pesos = tabla_md(pd.DataFrame({"Agregado": pcmp.agregado, "Canasta 2018": pcmp.peso_2018.map(lambda v: f(v, 2)),
                                 "Canasta 2023": pcmp.peso_2023.map(lambda v: f(v, 2)),
                                 "Cambio (pp)": (pcmp.peso_2023 - pcmp.peso_2018).map(lambda v: f(v, 2))}))

bl_rows = []
DRIV = {
    "Alim_SV": "FAO en pesos (rezagos 1-6), dólar (1-6), expectativas", "Bs_sv_ex_alim": "Dólar (0-2 y 3-8), expectativas",
    "ss_sv_ai": "IPC de los 12 meses previos (indexación), expectativas",
    "ss_sv_ex_ai": "Salarios nominales, brecha de Imacec, expectativas, IPC pasado",
    "elec": "Regla: ajuste en ene y jul por la inflación del semestre previo", "comb": "Petróleo WTI en pesos (0, 1 y 2)",
    "alim_rest_vol": "FAO en pesos (0-3 y 4-9), dólar (0-3)", "fyv": "Estacionalidad + rezago 12",
    "bs_vol": "Dólar (0-2 y 3-8)", "ss_v_ai": "IPC de los 12 meses previos", "ss_v_ex_ai": "Petróleo (1-3), dólar (0-2)",
}
for _, rr in pw.iterrows():
    bl_rows.append({"Bloque": rr.nombre, "Agregado": rr.agregado, "Peso 2023": f(rr.peso_2023, 2),
                    "Productos": rr.n_productos, "Impulsores en M2/M3": DRIV[rr.bloque]})
t_bloques = tabla_md(pd.DataFrame(bl_rows))

cat_show = cat.copy()
t_cat = tabla_md(pd.DataFrame({"Variable": cat_show.variable, "Código BDE": "`" + cat_show.codigo_bde + "`",
                               "Descripción": cat_show.descripcion, "Frecuencia": cat_show.frecuencia,
                               "Desde": cat_show.desde.str[:7], "Hasta": cat_show.hasta}))

t_val = tabla_md(pd.DataFrame({"Bloque": val.bloque, "Serie en historico_bcch_indices.csv": val.serie_historico_bcch,
                               "Correlación var. 12m 2010-23": val.corr_var12m_2010_2023.map(lambda v: f(v, 3)),
                               "Dif. media absoluta (pp)": val.dif_media_abs_pp.map(lambda v: f(v, 2))}))

dif_rec = dif.tail(9)
t_dif = tabla_md(pd.DataFrame({"Mes": dif_rec.periodo.map(mes_txt), "Difusión (% productos con alza)": dif_rec.difusion_IPC.map(f),
                               "Promedio histórico": dif_rec.hist_prom.map(f), "Mínimo": dif_rec.hist_min.map(f),
                               "Máximo": dif_rec.hist_max.map(f)}))

sup_m = sup.loc["2026-09-01":"2027-12-01"]
t_sup = tabla_md(pd.DataFrame({
    "Mes": [mes_txt(p) for p in sup_m.index[[0, 1, 2, 3, 6, 9, 15]]],
    "Dólar (pesos)": sup_m.tcn.iloc[[0, 1, 2, 3, 6, 9, 15]].map(lambda v: f(v, 0)),
    "EEE 11 meses (%)": sup_m.eee.iloc[[0, 1, 2, 3, 6, 9, 15]].map(lambda v: f(v, 2)),
    "Remuneraciones (var. 12m)": (sup.rem.pct_change(12) * 100).loc[sup_m.index].iloc[[0, 1, 2, 3, 6, 9, 15]].map(f),
    "Imacec desest. (var. 12m)": (sup.ima.pct_change(12) * 100).loc[sup_m.index].iloc[[0, 1, 2, 3, 6, 9, 15]].map(f),
}))

ms = pares.comparte_tendencia.sum()
fecha_hoy = pd.Timestamp.today()

# ------------------------------------------------------------------ texto
inc_i = inc.set_index("bloque")
md = f"""
# Proyección del IPC de Chile: octubre 2026 a diciembre 2027 (método de bloques à la Carlomagno)

> **Fuentes principales**: Base de Datos Estadísticos (BDE) del Banco Central de Chile, descargada vía API el
> {fecha_hoy:%d-%m-%Y}; Excel `IPC_2023_Automatizado_v2.xlsx` (índices por producto y bloques, matriz de la canasta 2023);
> insumos de la carpeta (canasta 2023=100, matriz de productos, pesos 2018, difusión histórica, índices históricos del BCCh).
> **Marco metodológico**: Carlomagno, Fornero y Sansone (2021, DT BCCh N°913), que define la partición sin volátiles /
> volátiles, y Carlomagno y Espasa (2015-2021), que proponen modelar muchos desagregados con ecuaciones únicas que imponen
> las tendencias comunes detectadas por pares.
>
> Último IPC efectivo: **{mes_txt(ULT)}** (IPC {f(ipc_sep_12)}% a/a; IPC SV {f(sv_sep_12)}% a/a). Es un ejercicio
> mecánico de modelos: no incorpora juicio sobre decretos tarifarios, impuestos ni información de precios de alta frecuencia.

------------------------------------------------------------------------

## 0. Resumen en una página

| Pregunta | Respuesta corta |
|:---|:---|
| ¿Qué se proyecta? | El IPC (base 2023=100) y sus agregados del BCCh, mes a mes entre {mes_txt(proy.periodo_dt.iloc[0])} y {mes_txt(proy.periodo_dt.iloc[-1])}. |
| ¿Cuánto da? | **{f(oct_mm, 2)}%** m/m en octubre, {f(nov_mm, 2)}% en noviembre y {f(dic_mm, 2)}% en diciembre. Cierre de **2026 en {f(dic26)}%** y de **2027 en {f(dic27)}%** a/a. |
| ¿Y el IPC sin volátiles? | {f(sv26)}% a/a a diciembre de 2026 y {f(sv27)}% a diciembre de 2027. |
| ¿Cómo se construye? | 11 bloques de la canasta (los de la descomposición SV/volátiles del BCCh), 4 modelos de ecuación única por bloque, combinación por desempeño fuera de muestra y agregación de Laspeyres (pesos efectivos). |
| ¿Qué aporta Carlomagno? | (i) La partición de la canasta (Carlomagno, Fornero y Sansone). (ii) Los tests de tendencia común **por pares** y su imposición como término de corrección de error (Carlomagno y Espasa). |
| ¿Qué tan bueno es? | Fuera de muestra (orígenes 2020-26) el combinado le gana al AR sobre el IPC agregado y al paseo aleatorio en el dato mensual (Diebold-Mariano, p < 0,05). Desde jul-23 empata con el modelo de tendencia común y le gana a ambos benchmarks. La EEE sigue siendo mejor en el dato del mes. |
| ¿Qué lo mueve? | El salto de **combustibles** (dólar de {f(tcn_sep, 0)} a {f(tcn_oct, 0)} pesos y WTI de {f(wti_sep, 0)} a {f(wti_oct, 0)} US$/bbl entre septiembre y octubre) y la persistencia de **servicios** indexados. |
| ¿Contra el mercado? | Sobre la EEE de octubre (4,6% a dic-26, 3,0% a dic-27). Para sep-27 ({f(P.loc['2027-09', 'IPC_var_12m'])}%) coincide con la compensación inflacionaria swap a 1 año ({f(comp.compensacion_swap_1a.dropna().iloc[0], 2)}%). |
| ¿Riesgos? | Sesgo al alza: en los episodios de shock (2021-22) todos los modelos subestimaron la inflación, y hoy hay un shock de petróleo en curso. |

------------------------------------------------------------------------

## 1. Datos utilizados

### 1.1 Variables macro de la BDE

Se descargaron {len(cat)} series vía la API REST de la BDE (`SieteRestWS`), {int((cat.frecuencia == 'Mensual').sum())} mensuales y
{int((cat.frecuencia == 'Diaria').sum())} diarias. Las diarias (dólar observado, WTI, cobre y compensaciones inflacionarias)
sirven para completar el **mes en curso** con el promedio de los días disponibles: octubre de 2026 tiene datos hasta el
{pd.read_csv(os.path.join(ROOT, 'data', 'bde', 'bde_diario.csv')).dropna(subset=['tcn_d']).fecha.iloc[-1]}.

<details><summary>Ver el catálogo completo de series (código BDE, rango)</summary>

{t_cat}

</details>

### 1.2 Insumos de la carpeta

| Archivo | Uso en este ejercicio |
|:---|:---|
| `IPC_2023_Automatizado_v2.xlsx` | Hoja `Historia`: índices de los 11 bloques desde 1998 (solo datos efectivos, hasta {mes_txt(ULT)}). Hoja `Matriz`: asignación de los 283 productos a bloques y pesos. Hoja `Índice 2023`: micro-índices por producto (difusión). Proyección previa del Excel como comparación. |
| `canasta-ipc-2023100.xlsx` / `matriz_productos.csv` | Estructura y ponderadores de la canasta base 2023. Los pesos de bloque suman exactamente 100. |
| `pesos_canasta_2018.csv` | Comparación de los pesos de los agregados entre canastas (§1.4). |
| `historico_bcch_indices.csv` | Validación de la historia de los bloques en 2010-2023 (§1.5). |
| `difusion_historica.csv` | Rango histórico mensual del índice de difusión (§5.4). |
| `ipc_base2018_alimentos.csv` | Contexto (división 1 y restaurantes en base 2018). No entra a los modelos. |
| `Apuntes_IPC_Chile.html` | Plantilla de formato de este documento y marco conceptual (pesos efectivos, arrastre, estacionalidad). |

### 1.3 Los 11 bloques

La canasta se divide en los 11 bloques "hoja" de la descomposición IPC sin volátiles / volátiles que usa el BCCh. La suma de
sus índices, ponderada por los pesos 2023, **replica el IPC del BCCh** (diferencia máxima en la variación mensual 2024-26:
0,0000 pp) y el IPC SV (0,0003 pp).

{t_bloques}

> ⚠️ Nota sobre los apuntes: en la tabla de índices analíticos (§8) del documento `Apuntes_IPC_Chile.html` los pesos del
> IPC SV y del IPC volátiles aparecen invertidos (37,95% y 62,05%). Según la matriz de la canasta 2023, el IPC SV pesa
> **62,05%** y los volátiles **37,95%**.

### 1.4 Pesos: canasta 2018 vs 2023

{t_pesos}

### 1.5 Validación de la historia de los bloques

{t_val}

------------------------------------------------------------------------

## 2. Metodología

### 2.1 La idea de Carlomagno

**Carlomagno, Fornero y Sansone (2021)** construyen la medida de inflación subyacente del BCCh (IPC sin volátiles)
eligiendo, mediante una función de pérdida, qué subclases excluir. El resultado es una partición de la canasta en bloques con
comportamientos distintos: unos inerciales y ligados a holguras, salarios y tipo de cambio, y otros dominados por shocks.
Esta proyección **usa esa misma partición** como unidad de modelación.

**Carlomagno y Espasa** muestran que, al proyectar un agregado desde muchos desagregados, conviene:

1. Detectar **tendencias comunes por pares**. Para cada par de componentes se testea si comparten tendencia, lo que da
   $N(N-1)/2$ tests. Luego se forman los subconjuntos que la comparten. No hace falta que la tendencia común sea pervasiva.
2. Modelar **cada componente con una ecuación única** que incorpore la restricción de tendencia común como término de
   corrección de error. Los componentes sin tendencia común se modelan sin restricción.
3. **Agregar** las proyecciones con los pesos de la canasta.

### 2.2 Implementación

**Paso 1: tendencias comunes.** Para los 11 bloques se testean los {len(pares)} pares. El test es un ADF con constante y
tendencia sobre el log del precio relativo $\\ln I_i - \\ln I_j$, en 2010-2026, con vector de cointegración (1, −1) impuesto.
Hay {ms} pares que comparten tendencia al 5%. La tendencia común del bloque $i$, $T_{{-i}}$, es el índice ponderado de los
bloques con los que comparte tendencia.

{t_grupos}

**Paso 2: modelos por bloque.** Todos modelan la variación mensual del bloque $\\pi_{{i,t}}$ como desvío respecto de un
**ancla** $a_{{i,t}} = \\pi^{{e,23m}}_{{t-1}}/12 + \\delta_i$. El primer término es la expectativa a 2 años de la EEE y
$\\delta_i$ es la deriva relativa del bloque frente al IPC en 2010-2019 (por ejemplo, los servicios suben más que los bienes de
forma persistente). Sin ancla, las constantes absorben la inflación alta de 2021-23 y el modelo proyecta cerca de 5% a dos años.
Las dummies estacionales están centradas, es decir, suman cero en el año.

$$\\pi_{{i,t}} - a_{{i,t}} = \\underbrace{{\\sum_{{m=1}}^{{11}} s_m (D_{{m,t}} - D_{{12,t}})}}_{{\\text{{M0}}}} + \\underbrace{{\\phi_1 y_{{i,t-1}} + \\phi_2 y_{{i,t-2}}}}_{{\\text{{M1}}}} + \\underbrace{{\\beta' x_{{i,t}}}}_{{\\text{{M2}}}} + \\underbrace{{\\gamma\\,(\\ln I_{{i}} - \\ln T_{{-i}} - c - d\\,t)_{{t-1}}}}_{{\\text{{M3}}}} + \\varepsilon_{{i,t}}$$

Los modelos son anidados:

- **M0, estacional anclado**: el ancla más la estacionalidad. Es el benchmark "meta + estacionalidad".
- **M1, AR estacional**: M0 más dos rezagos propios.
- **M2, ARDL con impulsores macro**: M1 más los impulsores del bloque (tabla de §1.3). El dólar y el petróleo entran en log
  diferencias, la FAO y el petróleo en pesos. La indexación se mide como la inflación del IPC de los 12 meses previos menos el
  ancla.
- **M3, ECM de tendencia común**: M2 más la corrección de error frente a $T_{{-i}}$ (Carlomagno y Espasa). Se usa solo si
  $\\hat\\gamma<0$, es decir, si el término efectivamente corrige.
- **Electricidad**: no se modela econométricamente, por ser una tarifa regulada por decreto. Se aplica una regla: ajuste en
  enero y julio igual a la inflación acumulada del semestre previo.

**Paso 3: combinación.** Para cada bloque, los modelos se ponderan por el inverso del error cuadrático medio fuera de muestra
del **cambio acumulado** del índice a 1-12 meses. En cada origen solo se usan los errores ya observados. La combinación se
hace dentro de la recursión: cada mes se combinan las proyecciones a un paso con el estado combinado, de modo que la
indexación al IPC pasado usa el IPC combinado.

{t_pc}

**Paso 4: agregación.** $IPC_t = \\sum_i w_i I_{{i,t}}$, con los pesos 2023. Equivale a usar los pesos efectivos
$\\hat w_{{i,t}} = w_i I_{{i,t-1}}/IPC_{{t-1}}$ de los apuntes (§7.2): por ejemplo, la electricidad pesa hoy
{f(inc_i.loc['elec', 'peso_efectivo_oct26'], 2)}% y no 2,21%, porque acumuló un alza de ~75% desde 2023.

### 2.3 Supuestos de las variables exógenas

- **Dólar, petróleo WTI y FAO.** Quedan en su último valor (paseo aleatorio). Para octubre se usa el promedio de los datos
  diarios disponibles.
- **Expectativas.** La expectativa a 11 meses de la EEE converge en 12 meses a la de 23 meses.
- **Remuneraciones e Imacec.** Crecen a su ritmo de los últimos 12 meses.

{t_sup}

------------------------------------------------------------------------

## 3. Proyección

### 3.1 IPC: trayectoria central

![]({figs['abanico']})

![]({figs['mensual']})

{t_proy}

*Bandas 10-90 y 30-70: percentiles empíricos de los errores fuera de muestra del modelo combinado para cada horizonte, con
orígenes entre jul-23 y ago-26 (régimen sin shocks grandes). Con los errores de 2020-26, que incluyen 2021-22, el percentil 90 a
dic-27 llega a {f(B.loc['2027-12-01', 'p90'])}%. Ese es el escenario de cola si el shock de petróleo se propaga como en
2021-22.*

### 3.2 Agregados (variación en 12 meses, %)

{t_agr}

![]({figs['sv_vol']})

**Lectura.**

- La inflación en 12 meses sube de {f(ipc_sep_12)}% en septiembre a {f(dic26)}% en diciembre, por dos razones. Los
  combustibles suben {f(inc_i.loc['comb', 'var_12m_dic26'])}% en 12 meses y aportan {f(inc_i.loc['comb', 'incid_12m_dic26'], 2)} pp.
  Además, dic-25 fue un mes negativo (−0,2%), lo que genera un **efecto base**.
- En el primer bimestre de 2027 el IPC se mantiene sobre 5% porque la variación de feb-26 fue nula (−0,0%).
- En abril de 2027 la inflación cae casi 1 pp: salen de la ventana los meses del shock petrolero de marzo y abril de 2026
  (+1,0% y +1,3% m/m).
- El IPC SV se mantiene entre {f(P['IPC sin volátiles_12m'].min())}% y {f(P['IPC sin volátiles_12m'].max())}%. Los servicios
  SV siguen sobre 4% por la indexación, y los bienes SV en torno a 2-3%.

### 3.3 Incidencias por bloque

![]({figs['incidencias']})

{t_inc}

### 3.4 Escenarios por modelo

{t_esc}

La dispersión entre modelos es en sí misma información. Los modelos con impulsores macro (M2) y con tendencia común (M3)
traspasan más el shock de dólar y petróleo, y fueron los mejores en 2020-22. El M3 funciona como **escenario alto**. Los
modelos puramente estadísticos (M0, M1) convergen más rápido a la meta y funcionan como **escenario bajo**.

------------------------------------------------------------------------

## 4. Evaluación fuera de muestra

Es un ejercicio recursivo con origen en cada mes entre ene-19 y ago-26. En cada origen se reestima todo con la información
disponible y se proyecta 1 a 15 meses. Las variables exógenas quedan constantes desde el origen: no se usa información
futura del dólar ni del petróleo. Los errores se miden contra el IPC efectivo.

![]({figs['rmse']})

**RMSE (pp)**, por muestra de orígenes:

{t_eval}

**Test de Diebold-Mariano** (combinado vs. rival; un estadístico negativo indica que el combinado es mejor):

{t_dm}

**Lectura.**

- **Agregar desde bloques le gana al modelo agregado** y al paseo aleatorio. La mejora es significativa a 1 mes en ambas
  muestras y a 3 meses desde 2023. Es el resultado central de Carlomagno y Espasa.
- En 2020-22, los modelos con impulsores y tendencia común (M2 y M3) dominan: capturaron mejor la propagación del dólar y de
  los commodities. Desde jul-23 la combinación es igual o mejor que M3, por eso es la proyección central.
- La **EEE** gana en el dato del mes. Es esperable: se levanta después del IPC anterior e incorpora información de alta
  frecuencia (precios de combustibles semanales, supermercados) que estos modelos no usan. Es la principal mejora pendiente
  (§6).
- Los errores tienen **sesgo positivo**: en 2020-26 la inflación efectiva superó la proyectada, en promedio, en
  {f(rmse[(rmse.modelo == 'Combinado') & (rmse.h == 12)].sesgo_12m.iloc[0])} pp a 12 meses, concentrado en 2021-22.

------------------------------------------------------------------------

## 5. Comparación con el mercado y contexto

### 5.1 Contra encuestas, precios de activos y la proyección previa

{t_comp}

- **Oct-26.** Proyectamos {f(oct_mm, 2)}% contra 0,4% de la EEE. La diferencia está casi toda en combustibles (+2,2% m/m,
  {f(inc_i.loc['comb', 'incid_12m_dic26'], 2)} pp de incidencia anual a diciembre). Esta proyección incorpora el dólar y el WTI
  de octubre completos hasta el día {pd.read_csv(os.path.join(ROOT, 'data', 'bde', 'bde_diario.csv')).dropna(subset=['tcn_d']).fecha.iloc[-1][-2:]};
  la EEE se levantó a inicios de mes, cuando parte de esa alza no ocurría.
- **2027.** Proyectamos {f(dic27)}%, sobre el 3,0% de la EEE, pero en línea con la compensación inflacionaria a un año. Según
  los apuntes, el IPoM de septiembre ubicaba el cierre de 2026 en torno a 4,3% y la convergencia a 3% en el 2T-2027.
- **Proyección previa del Excel.** Respecto de ella, esta proyección es más baja en noviembre y diciembre. La previa supone
  en noviembre alzas de 1,9% en combustibles y de 3,0% en servicios volátiles resto (pasajes y paquetes).

### 5.2 Variables macro recientes (BDE)

| Variable | Último dato | Valor |
|:---|:---|---:|
| TPM | {mes_txt(base['tpm'].dropna().index[-1])} | {f(base['tpm'].dropna().iloc[-1], 2)}% |
| Imacec desest., var. 12m | {mes_txt(base['imacec_sa'].dropna().index[-1])} | {f(base['imacec_sa'].pct_change(12, fill_method=None).dropna().iloc[-1] * 100)}% |
| Desocupación (desest.) | {mes_txt(base['desempleo_sa'].dropna().index[-1])} | {f(base['desempleo_sa'].dropna().iloc[-1])}% |
| Remuneraciones nominales, var. 12m | {mes_txt(base['rem_nom'].dropna().index[-1])} | {f(base['rem_nom'].pct_change(12, fill_method=None).dropna().iloc[-1] * 100)}% |
| Dólar observado (promedio mensual) | oct-26 (parcial) | {f(tcn_oct, 0)} pesos por dólar |
| WTI | oct-26 (parcial) | {f(wti_oct, 1)} US$/bbl |
| Brent | {mes_txt(ULT)} | {f(base.loc[ULT, 'brent'], 1)} US$/bbl |
| Cobre | {mes_txt(ULT)} | {f(base.loc[ULT, 'cobre'], 2)} US$/lb |
| Índice FAO | {mes_txt(ULT)} | {f(base.loc[ULT, 'fao'], 1)} |
| EEE inflación 11 meses / 23 meses | oct-26 | {f(eee.eee_11m)}% / {f(eee.eee_23m)}% |

### 5.3 Difusión

{t_dif}

La difusión de septiembre (porcentaje de los 283 productos con alza mensual, calculado desde la hoja `Índice 2023`) está
cerca de su promedio histórico. El alza reciente está concentrada en energía y no es generalizada, lo que es coherente con
un IPC SV estable.

------------------------------------------------------------------------

## 6. Riesgos, limitaciones y mejoras

1. **Combustibles y MEPCO.** El modelo traspasa el petróleo con dos rezagos. Si el MEPCO acumula un diferencial grande, el
   traspaso puede durar más. Si el petróleo revierte, el escenario baja rápido. Como cada 10% de petróleo en pesos se traspasa
   ~1,4% a combustibles en dos meses, eso equivale a ~0,1 pp de IPC.
2. **Tarifas eléctricas.** La regla semestral es un supuesto. Hay que reemplazarla por el calendario de decretos tarifarios.
3. **Nowcasting.** Para el dato del mes, conviene complementar con los precios semanales de combustibles (ENAP/CNE), el
   scraping de supermercados y la EEE/EOF (§10.5 de los apuntes). Esa es la brecha frente a la EEE en h=1.
4. **Tendencias comunes.** El test ADF por pares tiene poca potencia en muestras de ~200 meses. Además, en la evaluación los
   grupos se fijaron con la muestra completa (leve ventaja informativa). Se impone el vector (1, −1), es decir, precios
   relativos con tendencia determinística.
5. **Bandas.** Las bandas provienen de errores históricos y no condicionan en el shock actual. La cola superior relevante es
   la del período 2020-26.

------------------------------------------------------------------------

## 7. Cómo replicar y actualizar

```bash
export BCCH_USER=...  BCCH_PASS=...          # credenciales de la API de la BDE
python scripts/01_descarga_bde.py            # descarga 46 series mensuales y 7 diarias → data/bde/
python scripts/02_base.py                    # bloques + pesos + macro → data/base_modelo.csv
python scripts/03_modelo.py                  # tendencias comunes, evaluación recursiva y proyección
python scripts/04_resultados.py              # tablas + output/Base_Proyeccion_IPC.xlsx
python scripts/05_reporte_html.py            # este documento
```

Para actualizar con un nuevo IPC, primero hay que agregar el mes efectivo a la hoja `Historia` del Excel automatizado; el
resto se recalcula solo. El último IPC se detecta desde la BDE.

------------------------------------------------------------------------

## Bibliografía

- Carlomagno, G., Fornero, J. y Sansone, A. (2021). *Toward a general framework for constructing and evaluating core
  inflation measures*. Documento de Trabajo BCCh N°913. Versión 2023 en el *Latin American Journal of Central Banking*, 4(3).
- Carlomagno, G. y Espasa, A. (2015). *Forecasting a large set of disaggregates with common trends and outliers*. Working
  Paper 15-18, Universidad Carlos III de Madrid. (<a href="https://e-archivo.uc3m.es/entities/publication/18b7092c-659f-4284-8f86-6e40c06e70e4/full">e-Archivo UC3M</a>)
- Carlomagno, G. y Espasa, A. (2021). *Discovering specific common trends in a large set of disaggregates: statistical
  procedures, their properties and an empirical application*. (<a href="https://cinve.org.uy/discovering-specific-common-trends-in-a-large-set-of-disaggregates-statistical-procedures-their-properties-and-an-empirical-application">CINVE</a>)
- Espasa, A. y Mayo-Burgos, I. (2013). Forecasting aggregates and disaggregates with common features. *International
  Journal of Forecasting*, 29(4).
- Cobb, M. (2009). *Forecasting Chilean Inflation From Disaggregate Components*. Documento de Trabajo BCCh N°545.
- Diebold, F. y Mariano, R. (1995). Comparing Predictive Accuracy. *Journal of Business & Economic Statistics*, 13(3).
- INE (2024). *Manual Metodológico Índice de Precios al Consumidor, Base Anual 2023=100*, v2.0.
- Banco Central de Chile. Base de Datos Estadísticos (si3.bcentral.cl) y Encuesta de Expectativas Económicas.
"""

# ------------------------------------------------------------------ pandoc + plantilla
cuerpo = subprocess.run(["pandoc", "-f", "markdown+tex_math_dollars", "-t", "html", "--mathjax", "--section-divs"],
                        input=md, capture_output=True, text=True, check=True).stdout
plantilla = open(PLANTILLA, encoding="utf-8").read()
ini = plantilla.find('<div class="container-fluid main-container">')
fin = plantilla.find("\n<script>\n\n// add bootstrap table styles")
cabecera = plantilla[:ini].replace("<title>Apuntes_IPC_Chile</title>", "<title>Proyección IPC Chile</title>")
extra_css = """<style>
details{margin:0 0 1em} summary{cursor:pointer;color:#2a78d6}
table{font-size:13px} .main-container{max-width:1040px} img{display:block;margin:0.5em auto 1.2em}
</style>
"""
cabecera = cabecera.replace("</head>", extra_css + "</head>", 1)
html = (cabecera + '<div class="container-fluid main-container">\n<div id="header">\n</div>\n' + cuerpo
        + "\n</div>\n" + plantilla[fin:])
destino = os.path.join(OUT, "Proyeccion_IPC_Chile.html")
open(destino, "w", encoding="utf-8").write(html)
print("HTML:", destino, f"{len(html) / 1e6:.2f} MB")
