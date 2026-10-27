"""Proyección del IPC por agregación de bloques (bottom-up) con modelos de ecuación única.

Esquema (guiado por Carlomagno, Fornero y Sansone, 2021, y Carlomagno y Espasa, 2015-2021):
  1. La canasta se separa en 11 bloques hoja de la descomposición IPC sin volátiles / volátiles
     del BCCh (la que se obtiene minimizando la pérdida de Carlomagno-Fornero-Sansone).
  2. Para cada par de bloques se testea si comparten tendencia (precio relativo estacionario,
     ADF con constante y tendencia). El subconjunto con el que cada bloque comparte tendencia
     define su "tendencia común" (enfoque pairwise de Carlomagno y Espasa).
  3. Cada bloque se proyecta con modelos de ecuación única para su variación mensual:
       M0 Estacional : constante + dummies estacionales (últimos 60 meses).
       M1 AR_est     : M0 (muestra completa) + AR(2).
       M2 ARDL_macro : M1 + impulsores macro del bloque (dólar, petróleo, FAO, salarios,
                       brecha, expectativas, indexación al IPC pasado).
       M3 ECM_tc     : M2 + término de corrección de error respecto de la tendencia común
                       (restricción que implica compartir tendencia).
     La electricidad se proyecta con una regla de indexación semestral (tarifa regulada).
  4. Los modelos se combinan por bloque con pesos inversos al ECM fuera de muestra
     (calculados solo con errores ya observados en cada origen) y los bloques se agregan
     como índice de Laspeyres (equivale a usar pesos efectivos).
  5. Evaluación fuera de muestra recursiva 2019-2026 contra modelos agregados directos,
     paseo aleatorio y la EEE; bandas de incertidumbre desde los errores empíricos.
Salida: output/*.csv usados por 04_reporte.py
"""
import os
import warnings

import numpy as np
import pandas as pd
from statsmodels.tsa.stattools import adfuller

warnings.filterwarnings("ignore")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "output")
H = 15                 # oct-2026 a dic-2027
INICIO_EST = "2010-01"  # tras el quiebre de cobertura de 2009
INICIO_EVAL = "2019-01"
INICIO_COMB = "2020-01"  # primer origen con errores previos suficientes para ponderar

base = pd.read_csv(os.path.join(ROOT, "data", "base_modelo.csv"), index_col=0, parse_dates=True)
pw = pd.read_csv(os.path.join(ROOT, "data", "pesos_bloques.csv"))
HOJAS = list(pw.bloque)
W = pw.set_index("bloque").peso_2023.values / 100
ULT = base["idx_" + HOJAS[0]].dropna().index.max()          # último IPC publicado
fechas = pd.date_range("1998-12-01", ULT + pd.DateOffset(months=H), freq="MS")
T0 = fechas.get_loc(ULT)
mes = fechas.month.values

# --------------------------------------------------------------------------- datos en arrays
def arr(col):
    return base[col].reindex(fechas).values.astype(float)

LOG = {b: 100 * np.log(arr("idx_" + b)) for b in HOJAS}
EXO = {
    "tcn": 100 * np.log(arr("tcn")),
    "oil": 100 * np.log(arr("petroleo_clp")),
    "fao": 100 * np.log(arr("fao_clp")),
    "rem": 100 * np.log(arr("rem_nom")),
    "ima": 100 * np.log(arr("imacec_sa")),
    "eee": arr("eee_11m"),
    "e23": arr("eee_23m"),
}
EXO_RW = {"tcn", "oil", "fao", "e23"}        # se extrapolan planos (paseo aleatorio)
EXO_CONV = {"eee"}                           # converge linealmente a la expectativa a 2 años en 12 meses      # se extrapolan planos (paseo aleatorio)
EXO_TEND = {"rem", "ima"}                    # se extrapolan con su crecimiento de 12 meses


def lag(x, k):
    out = np.full_like(x, np.nan)
    out[k:] = x[:-k] if k else x
    return out


def d(x):
    return x - lag(x, 1)


def mlag(x, a, b):
    return np.mean(np.vstack([lag(x, k) for k in range(a, b + 1)]), axis=0)


def pct(x, t, k):
    """Variación porcentual a k meses desde un log-nivel ×100."""
    return 100 * (np.exp((x[t] - x[t - k]) / 100) - 1)


def ipc_log(L):
    lv = np.vstack([np.exp(L[b] / 100) for b in HOJAS])
    return 100 * np.log(W @ lv)


# --------------------------------------------------------------------------- tendencias comunes
def tendencias_comunes(L, hasta):
    sl = slice(fechas.get_loc(pd.Timestamp(INICIO_EST)), hasta + 1)
    pares = []
    for i, a in enumerate(HOJAS):
        for b in HOJAS[i + 1:]:
            rel = (L[a] - L[b])[sl] / 100
            p = adfuller(rel, regression="ct", autolag="AIC")[1]
            pares.append({"bloque_a": a, "bloque_b": b, "p_valor_adf": p, "comparte_tendencia": p < 0.05})
    pares = pd.DataFrame(pares)
    grupos = {}
    for b in HOJAS:
        s = pares[(pares.comparte_tendencia) & ((pares.bloque_a == b) | (pares.bloque_b == b))]
        grupos[b] = sorted((set(s.bloque_a) | set(s.bloque_b)) - {b})
    return pares, grupos


# --------------------------------------------------------------------------- regresores
# Dummies estacionales para los benchmarks agregados, y contrastes centrados (suman cero en el
# año) para los modelos por bloque, que no llevan constante: su media la fija el ancla.
SEAS = np.vstack([(mes == m).astype(float) for m in range(2, 13)])
SEAS_C = np.vstack([(mes == m).astype(float) - (mes == 12) for m in range(1, 12)])

# Deriva relativa de cada bloque frente al IPC (variación mensual promedio, 2010-2019):
# recoge, p. ej., que los servicios suben más que los bienes de forma persistente.
_sl = slice(fechas.get_loc(pd.Timestamp(INICIO_EST)), fechas.get_loc(pd.Timestamp("2019-12-01")) + 1)
_dipc = np.diff(np.r_[np.nan, 100 * np.log(W @ np.vstack([np.exp(LOG[b] / 100) for b in HOJAS]))])
DELTA = {b: np.nanmean((np.diff(np.r_[np.nan, LOG[b]]) - _dipc)[_sl]) for b in HOJAS}

DRIVERS = {
    "Alim_SV": ["fao_1_6", "tcn_1_6", "eee"],
    "Bs_sv_ex_alim": ["tcn_0_2", "tcn_3_8", "eee"],
    "ss_sv_ai": ["ipc_1_12", "eee"],
    "ss_sv_ex_ai": ["rem", "gap", "eee", "ipc_1_12"],
    "comb": ["oil_0", "oil_1", "oil_2"],
    "alim_rest_vol": ["fao_0_3", "fao_4_9", "tcn_0_3"],
    "fyv": ["ar12"],
    "bs_vol": ["tcn_0_2", "tcn_3_8"],
    "ss_v_ai": ["ipc_1_12"],
    "ss_v_ex_ai": ["oil_1_3", "tcn_0_2"],
}
MODELOS = ["M0_estacional", "M1_AR_est", "M2_ARDL_macro", "M3_ECM_tc"]


def driver(nombre, b, L, E, ipc):
    pi = d(L[b])
    if nombre == "ar12":
        return lag(pi, 12)
    if nombre == "eee":        # expectativa a 11 meses respecto de la de largo plazo
        return lag(E["eee"] - E["e23"], 1) / 12
    if nombre == "rem":        # salarios nominales respecto del ancla
        return lag(E["rem"] - lag(E["rem"], 12) - E["e23"], 2) / 12
    if nombre == "gap":
        return lag(E["ima"] - mlag(E["ima"], 1, 24), 2)
    v, a, *bb = nombre.split("_")
    a = int(a)
    bb = int(bb[0]) if bb else a
    x = d(ipc) - lag(E["e23"], 1) / 12 if v == "ipc" else d(E[v])   # indexación: IPC pasado vs ancla
    return mlag(x, a, bb)


def ancla(b, E):
    """Variación mensual de largo plazo del bloque: expectativa a 2 años (EEE) + deriva relativa."""
    return lag(E["e23"], 1) / 12 + DELTA[b]


def regresores(b, modelo, L, E, ipc, grupos):
    """y = variación mensual del bloque menos su ancla; X = regresores del modelo."""
    pi = d(L[b])
    y = pi - ancla(b, E)
    cols = [*SEAS_C]
    if modelo != "M0_estacional":
        cols += [lag(y, 1), lag(y, 2)]
    if modelo in ("M2_ARDL_macro", "M3_ECM_tc"):
        cols += [driver(n, b, L, E, ipc) for n in DRIVERS[b]]
    if modelo == "M3_ECM_tc":
        # desvío del precio relativo frente a la tendencia común del subconjunto: residuo de
        # Engle-Granger con vector (1,-1) y tendencia determinística, estimado hasta el origen
        g = grupos[b]
        tend = 100 * np.log(np.vstack([np.exp(L[j] / 100) for j in g]).T @ W[[HOJAS.index(j) for j in g]])
        rel = L[b] - tend
        tt = np.arange(len(pi), dtype=float)
        sl = slice(fechas.get_loc(pd.Timestamp(INICIO_EST)), E["_t_orig"] + 1)
        a, c = np.polyfit(tt[sl], rel[sl], 1)
        cols += [lag(rel - (a * tt + c), 1)]
    return y, np.vstack(cols).T


def disponible(b, modelo, grupos):
    if b == "elec":
        return modelo == "M1_AR_est"          # se reemplaza por la regla de indexación
    if modelo == "M3_ECM_tc":
        return len(grupos[b]) > 0
    return True


def estimar(b, modelo, L, E, ipc, grupos, t_fin):
    y, X = regresores(b, modelo, L, E, ipc, grupos)
    t_ini = fechas.get_loc(pd.Timestamp(INICIO_EST))
    sl = slice(t_ini, t_fin + 1)
    yy, XX = y[sl], X[sl]
    ok = np.isfinite(yy) & np.isfinite(XX).all(axis=1)
    beta, *_ = np.linalg.lstsq(XX[ok], yy[ok], rcond=None)
    res = yy[ok] - XX[ok] @ beta
    return beta, res.std(ddof=XX.shape[1])


def regla_elec(t, ipc):
    """Tarifa regulada: ajustes en enero y julio por la inflación acumulada del semestre previo."""
    if mes[t] in (1, 7):
        return ipc[t - 1] - ipc[t - 7]
    return 0.0


# --------------------------------------------------------------------------- simulación recursiva
def preparar_estado(t_orig, t_exo):
    """Copia de los datos con información hasta t_orig (endógenas) y t_exo (exógenas)."""
    L = {b: v.copy() for b, v in LOG.items()}
    for b in L:
        L[b][t_orig + 1:] = np.nan
    E = {k: v.copy() for k, v in EXO.items()}
    for k in E:
        x = E[k]
        last = np.max(np.where(np.isfinite(x[: t_exo + 1]))[0])
        if k in EXO_RW:
            x[last + 1:] = x[last]
        elif k in EXO_CONV:
            meta = E["e23"][np.max(np.where(np.isfinite(E["e23"][: t_exo + 1]))[0])]
            pasos = np.minimum(np.arange(1, len(x) - last), 12) / 12
            x[last + 1:] = x[last] + (meta - x[last]) * pasos
        else:
            g = (x[last] - x[last - 12]) / 12
            x[last + 1:] = x[last] + g * np.arange(1, len(x) - last)
    E["_t_orig"] = t_orig
    return L, E


def proyectar(t_orig, t_exo, grupos, pesos_comb=None, h=H):
    L0, E = preparar_estado(t_orig, t_exo)
    ipc0 = ipc_log(L0)
    coefs = {(b, m): estimar(b, m, L0, E, ipc0, grupos, t_orig)
             for b in HOJAS for m in MODELOS if disponible(b, m, grupos)}
    # la restricción de tendencia común solo se usa si el ajuste es corrector (γ < 0)
    for b in HOJAS:
        if (b, "M3_ECM_tc") in coefs and coefs[(b, "M3_ECM_tc")][0][-1] >= 0:
            coefs.pop((b, "M3_ECM_tc"))
    trayectorias = {}
    for modo in MODELOS + ["Combinado"]:
        L = {b: v.copy() for b, v in L0.items()}
        for t in range(t_orig + 1, t_orig + 1 + h):
            ipc = ipc_log(L)
            for b in HOJAS:
                if b == "elec":
                    p = regla_elec(t, ipc)
                else:
                    preds = {}
                    for m in MODELOS:
                        if (b, m) not in coefs or (modo != "Combinado" and m != modo):
                            continue
                        _, X = regresores(b, m, L, E, ipc, grupos)
                        preds[m] = ancla(b, E)[t] + X[t] @ coefs[(b, m)][0]
                    if not preds:   # bloque sin ese modelo: usa M2
                        _, X = regresores(b, "M2_ARDL_macro", L, E, ipc, grupos)
                        preds["M2_ARDL_macro"] = ancla(b, E)[t] + X[t] @ coefs[(b, "M2_ARDL_macro")][0]
                    if modo == "Combinado":
                        w = np.array([pesos_comb.get((b, m), 1.0) if pesos_comb else 1.0 for m in preds])
                        p = np.dot(w / w.sum(), list(preds.values()))
                    else:
                        p = next(iter(preds.values()))
                L[b][t] = L[b][t - 1] + p
        trayectorias[modo] = L
    return trayectorias, coefs


def directos(t_orig, h=H):
    """Benchmarks agregados: AR estacional sobre el IPC y paseo aleatorio en 12 meses."""
    ipc = ipc_log(LOG)
    pi = d(ipc)
    X = np.vstack([np.ones_like(pi), *SEAS, lag(pi, 1), lag(pi, 2)]).T
    t_ini = fechas.get_loc(pd.Timestamp(INICIO_EST))
    sl = slice(t_ini, t_orig + 1)
    beta, *_ = np.linalg.lstsq(X[sl], pi[sl], rcond=None)
    ar = ipc.copy(); ar[t_orig + 1:] = np.nan
    for t in range(t_orig + 1, t_orig + 1 + h):
        p = d(ar)
        x = np.r_[1, SEAS[:, t], p[t - 1], p[t - 2]]
        ar[t] = ar[t - 1] + x @ beta
    rw = ipc.copy(); rw[t_orig + 1:] = np.nan
    g = (ipc[t_orig] - ipc[t_orig - 12]) / 12
    rw[t_orig + 1: t_orig + 1 + h] = ipc[t_orig] + g * np.arange(1, h + 1)
    return {"D1_AR_est_agregado": ar, "D2_paseo_aleatorio_12m": rw}


# --------------------------------------------------------------------------- evaluación
def evaluar(grupos):
    ipc_obs = ipc_log(LOG)
    pi_obs = {b: d(LOG[b]) for b in HOJAS}
    t_ini, t_fin = fechas.get_loc(pd.Timestamp(INICIO_EVAL)), T0 - 1
    reg_bloque, reg_ipc = [], []
    eee_mes = base["eee_mm_mes"].reindex(fechas).values
    eee_sig = base["eee_mm_sig"].reindex(fechas).values
    for t in range(t_ini, t_fin + 1):
        pesos = None
        if fechas[t] >= pd.Timestamp(INICIO_COMB) and reg_bloque:
            r = pd.DataFrame(reg_bloque)
            r = r[(r.t_obj <= t) & (r.h <= 12)]
            mse = r.groupby(["bloque", "modelo"]).err_acum.apply(lambda e: np.mean(e ** 2))
            pesos = (1 / mse).to_dict()
        tray, _ = proyectar(t, t, grupos, pesos)
        tray_ipc = {k: ipc_log(v) for k, v in tray.items()}
        if pesos is None:
            tray_ipc.pop("Combinado"); tray.pop("Combinado")
        tray_ipc.update(directos(t))
        for hh in range(1, H + 1):
            to = t + hh
            if to > T0:
                break
            for modo, L in tray.items():
                for b in HOJAS:
                    reg_bloque.append({"origen": fechas[t], "t_obj": to, "h": hh, "bloque": b, "modelo": modo,
                                       "err": pi_obs[b][to] - (L[b][to] - L[b][to - 1]),
                                       "err_acum": (LOG[b][to] - LOG[b][t]) - (L[b][to] - L[b][t])})
            for modo, x in tray_ipc.items():
                reg_ipc.append({"origen": fechas[t], "objetivo": fechas[to], "h": hh, "modelo": modo,
                                "err_mm": pct(ipc_obs, to, 1) - pct(x, to, 1),
                                "err_12m": pct(ipc_obs, to, 12) - pct(x, to, 12)})
        # EEE: encuesta levantada en t+1 con el IPC de t ya publicado
        if t + 1 <= T0 and np.isfinite(eee_mes[t + 1]):
            reg_ipc.append({"origen": fechas[t], "objetivo": fechas[t + 1], "h": 1, "modelo": "EEE",
                            "err_mm": pct(ipc_obs, t + 1, 1) - eee_mes[t + 1],
                            "err_12m": np.nan})
        if t + 2 <= T0 and np.isfinite(eee_sig[t + 1]):
            reg_ipc.append({"origen": fechas[t], "objetivo": fechas[t + 2], "h": 2, "modelo": "EEE",
                            "err_mm": pct(ipc_obs, t + 2, 1) - eee_sig[t + 1],
                            "err_12m": np.nan})
        if t % 12 == 0:
            print(f"  origen {fechas[t]:%Y-%m}")
    return pd.DataFrame(reg_bloque), pd.DataFrame(reg_ipc)


def dm_test(e1, e2, h):
    """Diebold-Mariano con varianza HAC (h-1 rezagos). Negativo: modelo 1 mejor."""
    dd = e1 ** 2 - e2 ** 2
    n = len(dd)
    g = [np.mean((dd[k:] - dd.mean()) * (dd[: n - k] - dd.mean())) for k in range(h)]
    v = (g[0] + 2 * sum(g[1:])) / n
    from scipy.stats import norm
    stat = dd.mean() / np.sqrt(v)
    return stat, 2 * (1 - norm.cdf(abs(stat)))


# --------------------------------------------------------------------------- principal
if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    pares, grupos = tendencias_comunes(LOG, T0)
    pares.to_csv(os.path.join(OUT, "tendencias_comunes_pares.csv"), index=False)
    pd.DataFrame([{"bloque": b, "comparte_tendencia_con": ", ".join(g), "n": len(g)} for b, g in grupos.items()]
                 ).to_csv(os.path.join(OUT, "tendencias_comunes_grupos.csv"), index=False)
    print("Grupos de tendencia común:", {b: len(g) for b, g in grupos.items()})

    print("Evaluación fuera de muestra…")
    eb, ei = evaluar(grupos)
    eb.to_csv(os.path.join(OUT, "eval_errores_bloques.csv"), index=False)
    ei.to_csv(os.path.join(OUT, "eval_errores_ipc.csv"), index=False)

    # pesos finales de combinación: todos los errores observados (h<=12)
    r = eb[(eb.h <= 12) & (eb.modelo != "Combinado")]
    mse = r.groupby(["bloque", "modelo"]).err_acum.apply(lambda e: np.mean(e ** 2))
    pesos = (1 / mse).to_dict()
    tabla_pesos = (1 / mse).rename("inv_mse").reset_index()
    tabla_pesos = tabla_pesos[[disponible(b, m, grupos) for b, m in zip(tabla_pesos.bloque, tabla_pesos.modelo)]]
    tabla_pesos["peso"] = tabla_pesos.groupby("bloque").inv_mse.transform(lambda x: x / x.sum())
    tabla_pesos["rmse_error_acumulado"] = np.sqrt(1 / tabla_pesos.inv_mse)
    tabla_pesos.to_csv(os.path.join(OUT, "pesos_combinacion.csv"), index=False)

    # proyección final: origen = último IPC; dólar y petróleo del mes en curso ya conocidos
    tray, coefs = proyectar(T0, T0 + 1, grupos, pesos)
    filas = []
    for modo, L in tray.items():
        for b in HOJAS:
            for t in range(T0 - 24, T0 + H + 1):
                filas.append({"periodo": fechas[t], "modelo": modo, "bloque": b, "log_indice": L[b][t]})
    pd.DataFrame(filas).to_csv(os.path.join(OUT, "proy_bloques_modelos.csv"), index=False)
    dirs = directos(T0)
    pd.DataFrame({k: v[T0 - 24:] for k, v in dirs.items()}, index=fechas[T0 - 24:]).to_csv(
        os.path.join(OUT, "proy_directos.csv"))

    coef_rows = []
    for (b, m), (beta, s) in coefs.items():
        names = [f"estacional_{k}" for k in range(1, 12)]
        if m != "M0_estacional":
            names += ["ar1_desvio", "ar2_desvio"]
        if m in ("M2_ARDL_macro", "M3_ECM_tc"):
            names += DRIVERS[b]
        if m == "M3_ECM_tc":
            names += ["ecm_tend_comun"]
        for n, v in zip(names, beta):
            coef_rows.append({"bloque": b, "modelo": m, "variable": n, "coef": v})
        coef_rows.append({"bloque": b, "modelo": m, "variable": "sigma_residuo", "coef": s})
    pd.DataFrame(coef_rows).to_csv(os.path.join(OUT, "coeficientes.csv"), index=False)

    # supuestos exógenos usados
    _, E = preparar_estado(T0, T0 + 1)
    pd.DataFrame({k: v if k in ("eee", "e23") else np.exp(v / 100)
                  for k, v in E.items() if not k.startswith("_")}, index=fechas
                 ).loc[fechas[T0 - 12]:].to_csv(os.path.join(OUT, "supuestos_exogenos.csv"))
    print("OK")
