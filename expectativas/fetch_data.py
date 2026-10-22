"""Descarga las series del Chartpack de Expectativas.

Fuentes:
  - Banco Central de Chile, API BDE (credenciales en BCCH_USER / BCCH_PASS).
  - Reserva Federal de Nueva York, Global Supply Chain Pressure Index (GSCPI).

Guarda cada serie en data/<codigo>.csv (fecha, valor).
"""
import io
import os
import sys
from pathlib import Path

import pandas as pd
import requests

DATA = Path(__file__).parent / "data"
BCCH_URL = "https://si3.bcentral.cl/SieteRestWS/SieteRestWS.ashx"
GSCPI_URL = "https://www.newyorkfed.org/medialibrary/research/interactives/gscpi/downloads/gscpi_data.xlsx"

# codigo BCCh -> fecha de inicio de descarga
SERIES_BCCH = {
    # IMCE: precios y costos esperados por sector
    "G089.IME.IND.A14.M": "2017-01-01",  # Comercio, precios
    "G089.IME.IND.A34.M": "2017-01-01",  # Construccion, precios
    "G089.IME.IND.A27.M": "2017-01-01",  # Industria, precios
    "G089.IME.IND.A15.M": "2017-01-01",  # Comercio, costos
    "G089.IME.IND.A35.M": "2017-01-01",  # Construccion, costos
    "G089.IME.IND.A29.M": "2017-01-01",  # Industria, costos
    # IMCE: expectativas de inversion (comercio e industria)
    "G089.IME.IND.A17.M": "2009-01-01",
    "G089.IME.IND.A2B.M": "2009-01-01",
    # EEE: inflacion a 11 y 23 meses
    "F089.IPC.V12.14.M": "2014-01-01",
    "F089.IPC.V12.15.M": "2014-01-01",
    # EOF: inflacion 12 meses siguientes y 13-24 meses
    "F089.EOF.VII.12MS.D": "2014-01-01",
    "F089.EOF.VII.S12M.D": "2014-01-01",
    # Compensacion inflacionaria swaps promedio camara
    "F022.SWSP.TAS.AN01.Z.Z.D": "2021-06-01",
    "F022.SWSP.TAS.AN02.Z.Z.D": "2021-06-01",
    "F022.SWFW.TAS.E1X1.Z.Z.D": "2021-06-01",
    # EDEP: evolucion de costos, ultimos 3 meses
    "F089.EDEP.IND.88CL.M": "2021-01-01",
    "F089.EDEP.IND.88PI.M": "2021-01-01",
    "F089.EDEP.IND.88CMT.M": "2021-01-01",
    "F089.EDEP.IND.88CF.M": "2021-01-01",
    # EDEP: influencia sobre precio de venta
    "F089.EDEP.IND.84IFNV.M": "2021-01-01",
    "F089.EDEP.IND.84IFPD.M": "2021-01-01",
    "F089.EDEP.IND.84IFMG.M": "2021-01-01",
    "F089.EDEP.IND.84IFDI.M": "2021-01-01",
    "F089.EDEP.IND.84IFC.M": "2021-01-01",
    # EDEP: evolucion de costos, proximos 3 meses (lamina "EPN" del original)
    "F089.EDEP.IND.90CL.M": "2021-01-01",
    "F089.EDEP.IND.90PI.M": "2021-01-01",
    "F089.EDEP.IND.90CMT.M": "2021-01-01",
    "F089.EDEP.IND.90CF.M": "2021-01-01",
    # EDEP: expectativas de inflacion
    "F089.EDEP.POR.91EI12.M": "2021-01-01",
    "F089.EDEP.POR.91EI24.M": "2021-01-01",
    # Tipo de cambio
    "F073.TCO.PRE.Z.D": "2017-06-01",
    "F073.TCM.IND.199502.D": "2017-06-01",
    # Swap promedio camara
    "F022.SPC.TPR.D360.NO.Z.D": "2024-06-01",
    "F022.SPC.TIN.AN02.NO.Z.D": "2024-06-01",
    "F022.SPC.TIN.AN01.UF.Z.D": "2024-06-01",
    "F022.SPC.TIN.AN02.UF.Z.D": "2024-06-01",
    # EEE: PIB, consumo y FBCF por anio calendario
    "F089.PIB.V12.2026.M": "2024-01-01",
    "F089.PIB.V12.2027.M": "2024-01-01",
    "F089.CHOG.V12.2026.M": "2024-01-01",
    "F089.CHOG.V12.2027.M": "2024-01-01",
    "F089.FBKF.V12.2026.M": "2024-01-01",
    "F089.FBKF.V12.2027.M": "2024-01-01",
    # Imacec no minero (serie original empalmada 2018=100)
    "F032.IMC.IND.Z.Z.EP18.N03.Z.0.M": "2008-01-01",
}


def get_bcch(code, first, user, password):
    params = {"user": user, "pass": password, "function": "GetSeries",
              "timeseries": code, "firstdate": first}
    r = requests.get(BCCH_URL, params=params, timeout=120)
    r.raise_for_status()
    j = r.json()
    if j.get("Codigo") != 0:
        raise RuntimeError(f"{code}: {j.get('Descripcion')}")
    obs = j["Series"]["Obs"] or []
    s = pd.Series({pd.to_datetime(o["indexDateString"], format="%d-%m-%Y"): float(o["value"])
                   for o in obs}, name=code, dtype=float)
    return s.dropna()


def get_gscpi():
    r = requests.get(GSCPI_URL, timeout=120)
    r.raise_for_status()
    df = pd.read_excel(io.BytesIO(r.content), sheet_name="GSCPI Monthly Data", header=None)
    df = df.iloc[:, :2].dropna()
    df.columns = ["fecha", "valor"]
    df["fecha"] = pd.to_datetime(df["fecha"], format="%d-%b-%Y", errors="coerce")
    df = df.dropna()
    s = pd.Series(df["valor"].astype(float).values, index=df["fecha"].dt.to_period("M").dt.to_timestamp(),
                  name="GSCPI")
    return s


def save(s):
    s.rename_axis("fecha").rename("valor").to_csv(DATA / f"{s.name}.csv")


def main():
    user, password = os.environ.get("BCCH_USER"), os.environ.get("BCCH_PASS")
    if not user or not password:
        sys.exit("Faltan BCCH_USER / BCCH_PASS en el entorno")
    DATA.mkdir(exist_ok=True)
    for code, first in SERIES_BCCH.items():
        s = get_bcch(code, first, user, password)
        save(s)
        print(f"{code:36s} {len(s):5d} obs  ultimo {s.index.max():%Y-%m-%d} = {s.iloc[-1]:.2f}")
    s = get_gscpi()
    save(s)
    print(f"{'GSCPI':36s} {len(s):5d} obs  ultimo {s.index.max():%Y-%m-%d} = {s.iloc[-1]:.2f}")


if __name__ == "__main__":
    main()
