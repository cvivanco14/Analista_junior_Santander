"""Descarga las variables macro desde la BDE del Banco Central de Chile.

Uso:  BCCH_USER=... BCCH_PASS=... python scripts/01_descarga_bde.py
Salida: data/bde/bde_mensual.csv, data/bde/bde_diario.csv, data/bde/catalogo_series.csv
"""
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
import bde  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "data", "bde")
DESDE = "1998-12-01"

# (nombre corto, código BDE, descripción, grupo)
MENSUAL = [
    # IPC y agregados analíticos (empalme BCCh, base 2023=100)
    ("ipc", "G073.IPC.IND.2023.M", "IPC general, empalme BCCh", "Precios"),
    ("ipc_sv", "G073.IPCSV.IND.2023.M", "IPC sin volátiles", "Precios"),
    ("bs_sv", "G073.IPCBSV.IND.2023.M", "Bienes sin volátiles", "Precios"),
    ("ss_sv", "G073.IPCSSV.IND.2023.M", "Servicios sin volátiles", "Precios"),
    ("ipc_vol", "G073.IPCV.IND.2023.M", "IPC volátiles", "Precios"),
    ("alim_vol", "G073.IPCAV.IND.2023.M", "Alimentos volátiles", "Precios"),
    ("ene_vol", "G073.IPCEN.IND.2023.M", "Energía volátiles", "Precios"),
    ("resto_vol", "G073.IPCRV.IND.2023.M", "Resto volátiles", "Precios"),
    ("ipc_ine", "F074.IPC.IND.Z.EP23.C.M", "IPC INE empalmado dic-2009, base 2023", "Precios"),
    ("ipc_sae", "F074.IPCSAE.IND.Z.EP23.Z.M", "IPC SAE (INE)", "Precios"),
    ("ipc_alim", "F074.IPCA.IND.Z.EP23.Z.M", "IPC Alimentos (INE)", "Precios"),
    ("ipc_ene", "F074.IPCE.IND.Z.EP23.Z.M", "IPC Energía (INE)", "Precios"),
    ("ipc_fyv", "F074.IPCFV.IND.Z.EP23.Z.M", "IPC Frutas y verduras frescas (INE)", "Precios"),
    ("ipc_bienes", "F074.IPCB.IND.Z.EP23.Z.M", "IPC Bienes (INE)", "Precios"),
    ("ipc_serv", "F074.IPCS.IND.Z.EP23.Z.M", "IPC Servicios (INE)", "Precios"),
    ("ipc_trans", "F074.IPCT.IND.Z.EP23.Z.M", "IPC Transables (INE)", "Precios"),
    ("ipc_notrans", "F074.IPCN.IND.Z.EP23.Z.M", "IPC No transables (INE)", "Precios"),
    ("ipp", "F075.IPP.IND.P0551.2019.Z.M", "IPP general industrias (2019=100)", "Precios"),
    # Actividad y mercado laboral
    ("imacec", "F032.IMC.IND.Z.Z.EP18.Z.Z.0.M", "Imacec, serie original (2018=100)", "Actividad"),
    ("imacec_sa", "F032.IMC.IND.Z.Z.EP18.Z.Z.1.M", "Imacec desestacionalizado (2018=100)", "Actividad"),
    ("imacec_nm_sa", "F032.IMC.IND.Z.Z.EP18.N03.Z.1.M", "Imacec no minero desestacionalizado", "Actividad"),
    ("desempleo_sa", "F049.DES.TAS.INE.10.M", "Tasa de desocupación, ajustada estacionalmente (%)", "Mercado laboral"),
    ("desempleo_hist", "F049.DES.TAS.HIST.10.M", "Tasa de desocupación, serie histórica (%)", "Mercado laboral"),
    ("rem_nom", "G049.RMM.IND.INE23.NE.M", "Índice nominal de remuneraciones, empalmado (2023=100)", "Mercado laboral"),
    # Política monetaria, tipo de cambio y precios externos
    ("tpm", "F022.TPM.TIN.D001.NO.Z.M", "Tasa de política monetaria (%)", "Monetario"),
    ("tcn", "F073.TCO.PRE.HIST.M", "Dólar observado, promedio mensual ($/US$)", "Externo"),
    ("tcr", "F073.TCR.IND.199101.M", "Tipo de cambio real (1986=100)", "Externo"),
    ("ipe", "F073.IPE.IND.198601.M", "Índice de precios externos (1986=100)", "Externo"),
    ("brent", "F019.PPB.PRE.41AB.M", "Petróleo Brent (US$/barril)", "Externo"),
    ("wti", "F019.PPB.PRE.41B.M", "Petróleo WTI (US$/barril)", "Externo"),
    ("gasolina_us", "F019.PPB.PRE.42.M", "Gasolina EE.UU. (US$/m3)", "Externo"),
    ("diesel_us", "F019.PPB.PRE.35B.M", "Diésel (US$/galón)", "Externo"),
    ("cobre", "F019.PPB.PRE.40.M", "Cobre BML (US$/libra)", "Externo"),
    ("fao", "F019.PPBA.INP.46T.M", "Índice FAO de alimentos (2014-16=100)", "Externo"),
    ("core_us_a", "F019.IPS.V12.10.M", "Inflación subyacente EE.UU. (% a/a)", "Externo"),
    # Expectativas (Encuesta de Expectativas Económicas)
    ("eee_mm_mes", "F089.IPC.VAR.11.M", "EEE: inflación del mes (% m/m)", "Expectativas"),
    ("eee_mm_sig", "F089.IPC.VAR.24.M", "EEE: inflación del mes siguiente (% m/m)", "Expectativas"),
    ("eee_11m", "F089.IPC.V12.14.M", "EEE: inflación en 11 meses (% a/a)", "Expectativas"),
    ("eee_23m", "F089.IPC.V12.15.M", "EEE: inflación en 23 meses (% a/a)", "Expectativas"),
    ("eee_dic26", "F089.IPC.V12.2026.M", "EEE: inflación diciembre 2026 (% a/a)", "Expectativas"),
    ("eee_dic27", "F089.IPC.V12.2027.M", "EEE: inflación diciembre 2027 (% a/a)", "Expectativas"),
    ("eee_sv_mes", "F089.IPCSV.VAR.11.M", "EEE: IPC SV del mes (% m/m)", "Expectativas"),
    ("eee_sv_sig", "F089.IPCSV.VAR.24.M", "EEE: IPC SV del mes siguiente (% m/m)", "Expectativas"),
    ("eee_sv_dic26", "F089.IPCSV.V12.2026.M", "EEE: IPC SV diciembre 2026 (% a/a)", "Expectativas"),
    ("eee_sv_dic27", "F089.IPCSV.V12.2027.M", "EEE: IPC SV diciembre 2027 (% a/a)", "Expectativas"),
    ("eee_tpm_11m", "F089.TPM.TAS.14.M", "EEE: TPM en 11 meses (%)", "Expectativas"),
]

DIARIO = [
    ("tcn_d", "F073.TCO.PRE.Z.D", "Dólar observado ($/US$)", "Externo"),
    ("wti_d", "F019.PPB.PRE.41B.D", "Petróleo WTI (US$/barril)", "Externo"),
    ("cobre_d", "F019.PPB.PRE.100.D", "Cobre BML (US$/libra)", "Externo"),
    ("bei_swap_1a", "F022.SWSP.TAS.AN01.Z.Z.D", "Compensación inflacionaria swaps, spot 1 año (%)", "Expectativas"),
    ("bei_swap_2a", "F022.SWSP.TAS.AN02.Z.Z.D", "Compensación inflacionaria swaps, spot 2 años (%)", "Expectativas"),
    ("bei_bono_2a", "F022.BOSP.TAS.AN02.Z.Z.D", "Compensación inflacionaria bonos BCCh, spot 2 años (%)", "Expectativas"),
    ("bei_fw_1x1", "F022.SWFW.TAS.E1X1.Z.Z.D", "Compensación inflacionaria swaps, forward 1 en 1 (%)", "Expectativas"),
]


def _a_serie(res, nombre):
    obs = res["Series"]["Obs"] or []
    s = pd.Series(
        {pd.to_datetime(o["indexDateString"], format="%d-%m-%Y"): float(o["value"]) for o in obs},
        name=nombre, dtype=float,
    )
    return s.dropna()


def descargar(lista, desde):
    series, cat = [], []
    for nombre, cod, desc, grupo in lista:
        res = bde.get_series(cod, desde)
        if res.get("Codigo") != 0 or not res["Series"].get("Obs"):
            print(f"  ! {cod}: {res.get('Descripcion')}")
            continue
        s = _a_serie(res, nombre)
        series.append(s)
        cat.append({"variable": nombre, "codigo_bde": cod, "descripcion": desc, "grupo": grupo,
                    "titulo_bde": res["Series"]["descripEsp"], "desde": s.index.min().date(),
                    "hasta": s.index.max().date(), "n_obs": len(s)})
        print(f"  {nombre:<14} {cod:<34} {s.index.min():%Y-%m} → {s.index.max():%Y-%m-%d}")
    return pd.concat(series, axis=1, sort=True), pd.DataFrame(cat)


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    print("Mensuales:")
    men, cat_m = descargar(MENSUAL, DESDE)
    men.index = men.index.to_period("M").to_timestamp()
    men.index.name = "periodo"
    print("Diarias:")
    dia, cat_d = descargar(DIARIO, "2022-01-01")
    dia.index.name = "fecha"
    cat_m["frecuencia"], cat_d["frecuencia"] = "Mensual", "Diaria"
    men.to_csv(os.path.join(OUT, "bde_mensual.csv"))
    dia.to_csv(os.path.join(OUT, "bde_diario.csv"))
    pd.concat([cat_m, cat_d]).to_csv(os.path.join(OUT, "catalogo_series.csv"), index=False)
    print(f"OK: {men.shape[1]} series mensuales, {dia.shape[1]} diarias")
