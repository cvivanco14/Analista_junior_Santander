# -*- coding: utf-8 -*-
r"""Proyección del IPC de Chile (base 2023=100): modelos, proyecciones, nowcast, Excel y reporte.

Un solo script que hace todo el flujo:
  1. Descargas  : BDE del Banco Central (macro y expectativas) e INE (IPC por producto y subclase).
  2. Base       : 283 productos y 172 subclases (2018 en adelante), 11 bloques sin volátiles / volátiles
                  (desde 1998), agregados analíticos y difusión. Se valida contra lo publicado por el INE.
  3. Nowcast    : combustibles con los informes semanales de ENAP (tabla incluida en este archivo).
  4. Modelo A   : agregados por bloques (Carlomagno-Fornero-Sansone y Carlomagno-Espasa).
  5. Modelo B   : 172 subclases ancladas a expectativas, con corrección de error de tendencia común.
  6. Evaluación : recursiva fuera de muestra de ambos modelos, su combinación, benchmarks y la EEE.
  7. Salidas    : Excel autocontenido con fórmulas, reporte HTML y tablas CSV.

Carpetas (relativas a este script, sirven en cualquier computador):
    <proyeccion_ipc>\Code\       proyeccion_ipc.py (este archivo)
    <proyeccion_ipc>\Inputs\     insumos fijos: IPC_2023_Automatizado_v2.xlsx (historia 2018-2022 de productos
                                 y 1998-2017 de bloques), matriz_productos.csv, difusion_historica.csv y
                                 Apuntes_IPC_Chile.html (plantilla del reporte)
    <proyeccion_ipc>\Inputs\INE\ Excel descargados del INE (se reemplazan en cada corrida)
    <proyeccion_ipc>\Inputs\BDE\ series descargadas de la BDE (se reemplazan en cada corrida)
    <proyeccion_ipc>\Outputs\    IPC_Chile_Proyecciones.xlsx, Proyeccion_IPC_Chile_modelos.html y tablas\

Uso (no hay que editar nada; en Spyder basta F5):
    python proyeccion_ipc.py                 # descarga BDE e INE, modela y genera las salidas
    python proyeccion_ipc.py --sin-descarga  # usa lo ya descargado en Inputs\INE y Inputs\BDE
    python proyeccion_ipc.py --sin-reporte   # sin el HTML (no requiere pandoc)

Credenciales del BCCh (una vez por computador; nunca se escriben en este archivo). API Key Token:
  macOS (consola de Spyder; queda en el Llavero):
    import keyring, getpass; keyring.set_password('estudios_santander', 'BDE_TOKEN', getpass.getpass('Token BCCh: '))
  Windows, Anaconda Prompt o cmd (comillas DOBLES):   setx BDE_TOKEN "tu_token"
  Windows, PowerShell (comillas SIMPLES):             setx BDE_TOKEN 'tu_token'
  (con setx, luego cerrar y abrir Spyder / VS Code). También acepta la variable BCCH_API_token,
  o usuario y clave de la API en BCCH_USER y BCCH_PASS.

Requisitos: pandas, numpy, scipy, statsmodels, openpyxl, matplotlib, requests, bcchapi, keyring; pandoc para el
reporte HTML y, opcional, LibreOffice para dejar el Excel con los valores ya calculados.
"""
import argparse
import base64
import io
import json
import os
import subprocess
import sys
import time
import urllib.parse
import urllib.request
import warnings
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd

warnings.filterwarnings("ignore")

# Carpetas del proyecto: Code/ (este archivo), Inputs/ y Outputs/
CODE = Path(__file__).resolve().parent
PROYECTO = CODE.parent
INPUTS = PROYECTO / "Inputs"
D_INE = INPUTS / "INE"                     # Excel descargados del INE
D_BDE = INPUTS / "BDE"                     # series descargadas de la BDE
OUTPUTS = PROYECTO / "Outputs"
TABLAS = OUTPUTS / "tablas"
XLS_AUTOMATIZADO = INPUTS / "IPC_2023_Automatizado_v2.xlsx"
H_FIN = pd.Period("2027-12", "M")            # último mes proyectado

# =====================================================================================================
# 0. PARÁMETROS Y DATOS INCLUIDOS EN EL SCRIPT
# =====================================================================================================
BDE_MENSUAL = [
    ("ipc", "G073.IPC.IND.2023.M", "IPC general, empalme BCCh"),
    ("ipc_sv", "G073.IPCSV.IND.2023.M", "IPC sin volátiles"),
    ("ipc_vol", "G073.IPCV.IND.2023.M", "IPC volátiles"),
    ("ipc_sae_ine", "F074.IPCSAE.IND.Z.EP23.Z.M", "IPC SAE (INE), empalmado base 2023"),
    ("ipc_ene_ine", "F074.IPCE.IND.Z.EP23.Z.M", "IPC Energía (INE), empalmado base 2023"),
    ("imacec_sa", "F032.IMC.IND.Z.Z.EP18.Z.Z.1.M", "Imacec desestacionalizado (2018=100)"),
    ("desempleo_sa", "F049.DES.TAS.INE.10.M", "Tasa de desocupación desestacionalizada (%)"),
    ("rem_nom", "G049.RMM.IND.INE23.NE.M", "Índice nominal de remuneraciones (2023=100)"),
    ("tpm", "F022.TPM.TIN.D001.NO.Z.M", "Tasa de política monetaria (%)"),
    ("tcn", "F073.TCO.PRE.HIST.M", "Dólar observado, promedio mensual"),
    ("tcr", "F073.TCR.IND.199101.M", "Tipo de cambio real (1986=100)"),
    ("brent", "F019.PPB.PRE.41AB.M", "Petróleo Brent (US$/bbl)"),
    ("wti", "F019.PPB.PRE.41B.M", "Petróleo WTI (US$/bbl)"),
    ("cobre", "F019.PPB.PRE.40.M", "Cobre BML (US$/lb)"),
    ("fao", "F019.PPBA.INP.46T.M", "Índice FAO de alimentos (2014-16=100)"),
    ("eee_mm_mes", "F089.IPC.VAR.11.M", "EEE: inflación del mes (% m/m)"),
    ("eee_mm_sig", "F089.IPC.VAR.24.M", "EEE: inflación del mes siguiente (% m/m)"),
    ("eee_11m", "F089.IPC.V12.14.M", "EEE: inflación en 11 meses (% a/a)"),
    ("eee_23m", "F089.IPC.V12.15.M", "EEE: inflación en 23 meses (% a/a)"),
    ("eee_dic26", "F089.IPC.V12.2026.M", "EEE: inflación diciembre 2026 (% a/a)"),
    ("eee_dic27", "F089.IPC.V12.2027.M", "EEE: inflación diciembre 2027 (% a/a)"),
    ("eee_sv_dic26", "F089.IPCSV.V12.2026.M", "EEE: IPC SV diciembre 2026 (% a/a)"),
    ("eee_sv_dic27", "F089.IPCSV.V12.2027.M", "EEE: IPC SV diciembre 2027 (% a/a)"),
]
BDE_DIARIO = [
    ("tcn_d", "F073.TCO.PRE.Z.D", "Dólar observado diario"),
    ("wti_d", "F019.PPB.PRE.41B.D", "Petróleo WTI diario"),
    ("bei_swap_1a", "F022.SWSP.TAS.AN01.Z.Z.D", "Compensación inflacionaria swaps 1 año (%)"),
    ("bei_swap_2a", "F022.SWSP.TAS.AN02.Z.Z.D", "Compensación inflacionaria swaps 2 años (%)"),
]

# Archivos del INE (carpeta "Cuadros estadísticos" del IPC); se buscan por título en su API de archivos.
INE_API = "https://www.ine.gob.cl/estadisticas-por-tema/precios-e-inflacion/indice-de-precios-al-consumidor/getArchivos/"
INE_CARPETA = "ddd4c6b9-bc01-4fb7-8b28-ebd908780fab"
INE_ARCHIVOS = {"IPC XLS": "ipc_base2023.xlsx", "IPC Referencial Base 2023 XLS": "ipc_referencial_2023.xlsx",
                "ANALITICOS XLS": "analiticos_base2023.xlsx"}

# Informes semanales de ENAP (variación del precio mayorista, $/litro, por fecha de vigencia).
# Recopilados desde la prensa porque enap.cl no es accesible desde el entorno; la fuente de cada fecha está
# en ENAP_FUENTES. Gasolina = 93 octanos (la de 97 varía igual o casi igual).
ENAP_INFORMES = [  # (vigencia, gasolina, diésel, parafina)
    ("2026-05-28", -0.3, -65.0, 0.0), ("2026-06-18", -95.0, -117.5, 0.0), ("2026-07-02", 0.0, 0.0, 0.0),
    ("2026-07-09", -100.3, -155.4, 0.0), ("2026-07-30", 32.9, 28.5, 0.0), ("2026-08-06", 0.0, 0.0, 0.0),
    ("2026-08-20", 34.8, 97.4, 0.0), ("2026-09-10", 35.0, 89.0, 0.0), ("2026-10-01", 39.4, 95.0, 281.8),
    ("2026-10-08", 0.0, 0.0, -32.7), ("2026-10-15", 0.0, 0.0, 5.1), ("2026-10-22", 41.3, 91.5, 24.8),
    ("2026-11-05", 0.0, 0.0, 0.0),
]
ENAP_FUENTES = {
    "2026-05-28": "https://www.meganoticias.cl/nacional/523134-baja-precio-bencina-petroleo-diesel-combustubles-27-05-2026.html",
    "2026-06-18": "https://www.t13.cl/noticia/nacional/negocios/enap-confirma-baja-precio-las-bencinas-17-6-2026",
    "2026-07-02": "https://www.concierto.cl/2026/07/enap-confirma-el-nuevo-precio-de-la-bencina-en-chile-desde-el-jueves-2-de-julio/",
    "2026-07-09": "https://www.biobiochile.cl/noticias/economia/actualidad-economica/2026/07/08/confirmado-bencinas-caen-100-por-litro-segun-informe-de-la-enap.shtml",
    "2026-07-30": "https://radio.uchile.cl/2026/07/30/enap-bencinas-de-93-y-97-aumentaran-329-pesos-por-litro/",
    "2026-08-06": "https://www.concierto.cl/2026/08/enap-confirma-el-nuevo-precio-de-la-bencina-en-chile-desde-el-jueves-6-de-agosto/",
    "2026-08-20": "https://www.meganoticias.cl/nacional/529754-sube-valor-bencina-precio-gasolina-combustibles-petroleo-diesel-jueves-20-agosto-parafina-19-08-2026.html",
    "2026-09-10": "https://www.epicentrochile.com/2026/09/09/nuevo-golpe-al-bolsillo-enap-anuncia-alzas-de-35-en-bencinas-y-casi-90-en-el-diesel/",
    "2026-10-01": "https://www.emol.com/noticias/Economia/2026/09/30/1212882/alza-precios-combustibles-diesel-enap.html",
    "2026-10-08": "https://www.redimin.cl/enap-anuncia-baja-de-32-7-en-la-parafina-y-mantiene-sin-variaciones-las-bencinas-y-el-diesel",
    "2026-10-15": "https://www.concierto.cl/2026/10/enap-confirma-el-nuevo-precio-de-las-bencinas-en-chile-desde-el-jueves-15-de-octubre/",
    "2026-10-22": "https://www.biobiochile.cl/noticias/economia/tu-bolsillo/2026/10/21/enap-confirma-fuerte-alza-de-combustibles-bencinas-subiran-413-y-diesel-se-dispara-915-por-litro.shtml",
    "2026-11-05": "https://www.diarioconcepcion.cl/ciudad/2026/11/05/enap-informe-semanal-precios-combustibles.html",
}
# Precio promedio a público (CNE, vía agregadores) en una fecha: ancla de la reconstrucción del precio diario.
ENAP_ANCLAS = {"gasolina": ("2026-09-29", 1459.0), "diesel": ("2026-10-25", 1586.0)}
# Próximos ajustes del MEPCO (cada tres semanas). El 12-nov el Gobierno anunció +$40 en bencinas y diésel.
ENAP_ESCENARIOS = {
    "Base": [("2026-11-12", 40.0, 40.0), ("2026-12-03", 40.0, 40.0), ("2026-12-24", 40.0, 40.0)],
    "Solo 12-nov": [("2026-11-12", 40.0, 40.0)],
}
ENAP_ESCENARIO_CENTRAL = "Base"
ENAP_FUENTES_ESCENARIO = [
    "https://www.eldinamo.cl/economia/dinero/2026/11/09/kast-confirma-alza-de-combustibles-y-anuncia-proyecto/",
    "https://www.emol.com/noticias/Economia/2026/11/08/1216478/bencinas-alza-transportistas-camiones.html",
]
PROD_ENAP = {"gasolina": 72221, "diesel": 72211}
PARAFINA = 44311      # el INE la deja sin variación fuera de la temporada de calefacción (abril-agosto)

# Bloques hoja de la descomposición IPC SV / volátiles (columna de matriz_productos.csv).
BLOQUES = {
    "Alim_SV": ("Alimentos no volátiles", "Alimentos no volátiles"),
    "Bs_sv_ex_alim": ("Bienes SV excl. alimentos", "Bienes sv. ex. alim."),
    "ss_sv_ai": ("Servicios SV administrados e indexados", "Servicios_A&I_sv"),
    "ss_sv_ex_ai": ("Servicios SV resto", "Servicios_ Resto_sv"),
    "elec": ("Electricidad", "Electricidad_v"),
    "comb": ("Combustibles", "Combustible_v"),
    "alim_rest_vol": ("Alimentos volátiles resto", "Alim. Resto_v"),
    "fyv": ("Frutas y verduras", "FyV (proxy)_v"),
    "bs_vol": ("Bienes volátiles", "Bienes volát."),
    "ss_v_ai": ("Servicios volátiles adm. e indexados", "Servicios_A&I_v"),
    "ss_v_ex_ai": ("Servicios volátiles resto", "Servicios_Resto_v"),
}
HOJAS = list(BLOQUES)
SV = ["Alim_SV", "Bs_sv_ex_alim", "ss_sv_ai", "ss_sv_ex_ai"]
VOL = [b for b in HOJAS if b not in SV]
ENERGIA_INE = [44111, 44211, 44221, 44311, 44411, 72211, 72221]
# Agregados que se reportan: (nombre, definición). Definición = lista de bloques o una regla por producto.
AGREGADOS = {
    "IPC": HOJAS,
    "IPC sin volátiles": SV,
    "Bienes SV": ["Alim_SV", "Bs_sv_ex_alim"],
    "Servicios SV": ["ss_sv_ai", "ss_sv_ex_ai"],
    "IPC volátiles": VOL,
    "Energía volátil": ["elec", "comb"],
    "Alimentos volátiles": ["alim_rest_vol", "fyv"],
    "Resto volátiles": ["bs_vol", "ss_v_ai", "ss_v_ex_ai"],
    "Volátiles excl. energía": ["alim_rest_vol", "fyv", "bs_vol", "ss_v_ai", "ss_v_ex_ai"],
    "Alimentos": ["Alim_SV", "alim_rest_vol", "fyv"],
    "Energía (INE)": "energia_ine",
    "IPC SAE (INE)": "sae_ine",
    # bloques hoja y servicios volátiles (clases del cuadro "Resumen Proyecciones" del Excel automatizado)
    "Alimentos no volátiles": ["Alim_SV"],
    "Bienes SV excl. alimentos": ["Bs_sv_ex_alim"],
    "Servicios SV adm. e indexados": ["ss_sv_ai"],
    "Servicios SV resto": ["ss_sv_ex_ai"],
    "Electricidad": ["elec"],
    "Combustibles": ["comb"],
    "Alimentos volátiles resto": ["alim_rest_vol"],
    "Frutas y verduras": ["fyv"],
    "Bienes volátiles": ["bs_vol"],
    "Servicios volátiles": ["ss_v_ai", "ss_v_ex_ai"],
    "Servicios volátiles adm. e indexados": ["ss_v_ai"],
    "Servicios volátiles resto": ["ss_v_ex_ai"],
}
# Columna de la hoja Historia del Excel automatizado con la serie empalmada oficial de cada clase (perfiles hasta 2023)
HISTORIA_COL = {
    "IPC": "IPC_ref", "IPC sin volátiles": "IPC_SV", "Bienes SV": "BS_SV", "Alimentos no volátiles": "Alim_SV",
    "Bienes SV excl. alimentos": "Bs_sv_ex_alim", "Servicios SV": "SS_SV", "Servicios SV adm. e indexados": "ss_sv_ai",
    "Servicios SV resto": "ss_sv_ex_ai", "IPC volátiles": "IPC_Vol", "Energía volátil": "Ene_Vol", "Electricidad": "elec",
    "Combustibles": "comb", "Alimentos volátiles": "Alim_Vol", "Alimentos volátiles resto": "alim_rest_vol",
    "Frutas y verduras": "fyv", "Resto volátiles": "Resto_Vol", "Bienes volátiles": "bs_vol", "Servicios volátiles": "ss_vol",
    "Servicios volátiles adm. e indexados": "ss_v_ai", "Servicios volátiles resto": "ss_v_ex_ai",
}
# Clases del resumen y de los perfiles, en el orden jerárquico del Excel automatizado (nombre, nivel de sangría)
CLASES_RESUMEN = [
    ("IPC", 0), ("IPC sin volátiles", 1), ("Bienes SV", 2), ("Alimentos no volátiles", 3),
    ("Bienes SV excl. alimentos", 3), ("Servicios SV", 2), ("Servicios SV adm. e indexados", 3),
    ("Servicios SV resto", 3), ("IPC volátiles", 1), ("Energía volátil", 2), ("Electricidad", 3), ("Combustibles", 3),
    ("Alimentos volátiles", 2), ("Alimentos volátiles resto", 3), ("Frutas y verduras", 3), ("Resto volátiles", 2),
    ("Bienes volátiles", 3), ("Servicios volátiles", 3), ("Servicios volátiles adm. e indexados", 4),
    ("Servicios volátiles resto", 4), ("Volátiles excl. energía", 1), ("Alimentos", 1), ("Energía (INE)", 1),
    ("IPC SAE (INE)", 1),
]
AGREG_PRINCIPALES = ["IPC", "IPC sin volátiles", "Bienes SV", "Servicios SV", "IPC volátiles", "Alimentos volátiles",
                     "Energía volátil", "Resto volátiles", "Volátiles excl. energía", "Alimentos", "Energía (INE)",
                     "IPC SAE (INE)"]
MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


def log(msg):
    print(time.strftime("%H:%M:%S"), msg, flush=True)


def mes_txt(p):
    p = pd.Period(p, "M")
    return f"{MESES[p.month - 1]}-{str(p.year)[2:]}"


# =====================================================================================================
# 1. DESCARGAS
# =====================================================================================================
# Credenciales del BCCh: token (variable de entorno o llavero) o usuario y clave de la API.
_ALIAS_TOKEN = ["BDE_TOKEN", "BCCH_API_token", "BCCH_API_TOKEN"]
LLAVERO = "estudios_santander"   # servicio en el Llavero de macOS / Administrador de credenciales
BDE_URL = "https://si3.bcentral.cl/SieteRestWS/SieteRestWS.ashx"


def _token_llavero():
    """Token guardado con keyring (Llavero de macOS o Administrador de credenciales de Windows)."""
    try:
        import keyring
        return keyring.get_password(LLAVERO, "BDE_TOKEN")
    except Exception:  # noqa: BLE001 - sin keyring o sin llavero disponible
        return None


def _como_guardar_token(variable="BDE_TOKEN"):
    if sys.platform == "darwin":
        return ("En la consola de Spyder (una vez; pide el token sin mostrarlo):\n"
                "    import keyring, getpass; keyring.set_password('estudios_santander', 'BDE_TOKEN', "
                "getpass.getpass('Token BCCh: '))\n"
                "Queda en el Llavero de macOS; no hace falta reiniciar Spyder.")
    return ("Ejecuta una vez y luego cierra y abre Spyder / VS Code:\n"
            f'    Anaconda Prompt o cmd:  setx {variable} "tu_token"\n'
            f"    PowerShell:             setx {variable} 'tu_token'\n"
            "(en PowerShell las comillas dobles borran las partes del token que empiezan con $)")


def _credenciales():
    """Token del BDE (variable de entorno o llavero) o usuario y clave de la API."""
    for nombre in _ALIAS_TOKEN:
        if os.environ.get(nombre):
            return {"token": os.environ[nombre].strip().strip("'\"").strip(), "origen": nombre}
    token = _token_llavero()
    if token:
        return {"token": token.strip().strip("'\"").strip(), "origen": "llavero"}
    usuario, clave = os.environ.get("BCCH_USER"), os.environ.get("BCCH_PASS")
    if usuario and clave:
        return {"user": usuario.strip().strip("'\""), "pass": clave.strip("'\""), "origen": "BCCH_USER / BCCH_PASS"}
    sys.exit("No encuentro el token del BCCh.\n" + _como_guardar_token())


def _diagnostico_credenciales(cred):
    """Pistas para corregir la credencial, sin mostrarla."""
    if "token" not in cred:
        return ("Se usaron BCCH_USER y BCCH_PASS. Revisa que sean el usuario y la clave de la API "
                "del BDE (si3.bcentral.cl), no los de otra cuenta.")
    t = cred["token"]
    donde = "del Llavero" if cred["origen"] == "llavero" else f"de la variable {cred['origen']}"
    lineas = [f"Se usó el token {donde}:",
              f"  - largo: {len(t)} caracteres (un token del BCCh suele tener 60)",
              f"  - empieza con '$2a$': {'sí' if t.startswith('$2a$') else 'NO'}"]
    if any(c in t for c in " '\""):
        lineas.append("  - contiene espacios o comillas: se copió con caracteres de más")
    variable = cred["origen"] if cred["origen"] in _ALIAS_TOKEN else "BDE_TOKEN"
    lineas += ["Vuelve a guardarlo. " + _como_guardar_token(variable),
               "Si sigue fallando, genera un token nuevo en si3.bcentral.cl."]
    return "\n".join(lineas)


_SIETE = None


def _bde_serie(codigo, desde, cred):
    """Una serie de la BDE como pd.Series (fechas → valor), con token (bcchapi) o usuario y clave (API REST)."""
    global _SIETE
    if "token" in cred:
        if _SIETE is None:
            import bcchapi
            _SIETE = bcchapi.Siete(token=cred["token"])
        s = _SIETE.cuadro(series=[codigo], desde=desde).iloc[:, 0]
        s.index = pd.to_datetime(s.index)
        return pd.to_numeric(s, errors="coerce").dropna()
    p = {"user": cred["user"], "pass": cred["pass"], "function": "GetSeries", "timeseries": codigo, "firstdate": desde}
    with urllib.request.urlopen(BDE_URL + "?" + urllib.parse.urlencode(p), timeout=60) as r:
        data = json.loads(r.read().decode("latin-1"))
    if data.get("Codigo", 0) < 0:
        raise RuntimeError(f"BCCh API error {data['Codigo']}: {data.get('Descripcion')}")
    obs = (data.get("Series") or {}).get("Obs") or []
    return pd.Series({pd.to_datetime(o["indexDateString"], format="%d-%m-%Y"): float(o["value"]) for o in obs},
                     dtype=float).dropna()


def descargar_bde():
    """Descarga las series macro y de expectativas a Inputs\\BDE (si una serie falla, se avisa y se sigue)."""
    D_BDE.mkdir(parents=True, exist_ok=True)
    cred = _credenciales()
    log(f"BDE: credencial {cred['origen']}")
    out = {}
    for frec, lista, desde in [("mensual", BDE_MENSUAL, "1998-12-01"), ("diario", BDE_DIARIO, "2025-01-01")]:
        series, cat = [], []
        for nombre, cod, desc in lista:
            try:
                s = _bde_serie(cod, desde, cred).rename(nombre)
            except Exception as e:  # noqa: BLE001
                texto = f"{type(e).__name__} {e}".lower()
                if "invalidcredentials" in texto or "invalid username or password" in texto:
                    sys.exit("El BCCh rechazó la credencial.\n" + _diagnostico_credenciales(cred))
                log(f"  ❌ BDE {cod}: {e}")
                continue
            if s.empty:
                log(f"  ⚠️ BDE sin datos: {cod}")
                continue
            if frec == "mensual":
                s.index = s.index.to_period("M").to_timestamp()
            series.append(s)
            cat.append({"variable": nombre, "codigo_bde": cod, "descripcion": desc, "frecuencia": frec,
                        "desde": s.index.min().date(), "hasta": s.index.max().date()})
        df = pd.concat(series, axis=1, sort=True)
        df.index.name = "periodo" if frec == "mensual" else "fecha"
        df.to_csv(D_BDE / f"bde_{frec}.csv")
        out[frec] = pd.DataFrame(cat)
    pd.concat(out.values()).to_csv(D_BDE / "catalogo_bde.csv", index=False)
    log(f"  ✅ BDE: {len(out['mensual'])} series mensuales y {len(out['diario'])} diarias en {D_BDE}")


def descargar_ine():
    """Descarga los cuadros del IPC base 2023 a Inputs\\INE. Si una descarga falla, conserva el archivo anterior."""
    D_INE.mkdir(parents=True, exist_ok=True)
    req = urllib.request.Request(INE_API, data=urllib.parse.urlencode({"idFolder": INE_CARPETA}).encode(),
                                 headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        docs = json.loads(r.read().decode("utf-8"))["documento"]
    for d in docs:
        nombre = INE_ARCHIVOS.get(d["Titulo"].strip())
        if not nombre or "2023" not in d["Url"]:
            continue
        url = urllib.parse.quote(d["Url"].replace("http:", "https:"), safe=":/?=&_%.-")
        destino = D_INE / nombre
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=120) as r:
                contenido = r.read()
            if not contenido.startswith(b"PK"):          # un .xlsx es un zip; si no, es una página de error
                raise ValueError("la respuesta no es un archivo Excel")
            destino.write_bytes(contenido)
            log(f"  ✅ INE: {nombre}")
        except Exception as e:  # noqa: BLE001
            aviso = "se usa el archivo anterior" if destino.exists() else "no hay archivo anterior"
            log(f"  ❌ INE: no pude descargar {nombre} ({e}); {aviso}.")


# =====================================================================================================
# 2. BASE
# =====================================================================================================
def _leer_ine_ipc():
    """IPC por producto/subclase del INE: serie referencial 2023 + oficial desde 2024."""
    x = pd.concat([pd.read_excel(D_INE / f, header=3)
                   for f in ("ipc_referencial_2023.xlsx", "ipc_base2023.xlsx")])
    x = x[x["Año"].notna()].copy()
    x["periodo"] = pd.PeriodIndex([pd.Period(f"{int(a)}-{int(m):02d}", "M") for a, m in zip(x["Año"], x["Mes"])])
    niveles = ["División", "Grupo", "Clase", "Subclase", "Producto"]
    x["nivel"] = x[niveles].notna().sum(axis=1)
    x["codigo"] = x[niveles].apply(lambda r: "".join(str(int(v)) for v in r if pd.notna(v)), axis=1)
    x["cod_subclase"] = x[niveles[:4]].apply(lambda r: "".join(str(int(v)) for v in r if pd.notna(v)), axis=1)
    x["cod_division"] = x["División"].map(lambda v: str(int(v)) if pd.notna(v) else "")
    return x.drop_duplicates(["periodo", "nivel", "codigo"])


def _leer_excel_automatizado():
    wb = openpyxl.load_workbook(XLS_AUTOMATIZADO, read_only=True, data_only=True)
    rows = list(wb["Índice 2023"].iter_rows(min_row=1, max_row=284, max_col=110, values_only=True))
    cols = [c.strftime("%Y-%m") if hasattr(c, "strftime") else c for c in rows[0]]
    p = pd.DataFrame([r for r in rows[1:] if r[0]], columns=cols).set_index("Code")
    p = p.loc[:, "2018-01":"2022-12"].apply(pd.to_numeric, errors="coerce").T
    p.index = pd.PeriodIndex(p.index, freq="M")
    rows = list(wb["Historia"].iter_rows(min_row=3, max_col=23, values_only=True))
    h = pd.DataFrame([r for r in rows[1:] if r[0] is not None], columns=rows[0])
    h.index = pd.PeriodIndex(pd.to_datetime(h.pop("mes").astype(str).str[:7]), freq="M")
    h = h[[c for c in h.columns if c]].apply(pd.to_numeric, errors="coerce")
    return p, h


class Base:
    """Todos los datos que usan los modelos, el Excel y el reporte."""

    def __init__(self):
        ine = _leer_ine_ipc()
        self.ine = ine
        prod = ine[ine.nivel == 5]
        idx_ine = prod.pivot(index="periodo", columns="codigo", values="Índice")
        idx_ine.columns = idx_ine.columns.astype(int)
        self.T = idx_ine.index.max()                                   # último IPC publicado
        hist_p, hist_b = _leer_excel_automatizado()
        self.historia = hist_b.loc[:"2023-12"]           # agregados empalmados del Excel automatizado (oficiales)
        # productos: historia 2018-2022 del Excel (empalmada a base 2023) + INE oficial desde 2023
        self.prod_idx = pd.concat([hist_p[idx_ine.columns], idx_ine]).sort_index()
        m = pd.read_csv(INPUTS / "matriz_productos.csv").set_index("codigo")
        info = prod.drop_duplicates("codigo").set_index("codigo")[["Glosa", "Ponderación", "cod_subclase", "cod_division"]]
        info.index = info.index.astype(int)
        info.columns = ["glosa", "peso", "subclase", "division"]
        info["bloque"] = m.loc[info.index, [BLOQUES[b][1] for b in HOJAS]].values.argmax(axis=1)
        info["bloque"] = [HOJAS[i] for i in info.bloque]
        info["sv"] = info.bloque.isin(SV).astype(int)
        self.prod_info = info
        # subclases: glosa y peso del INE; índice = Laspeyres de sus productos (idéntico al INE desde 2023)
        sc = ine[ine.nivel == 4].drop_duplicates("codigo").set_index("codigo")[["Glosa", "Ponderación"]]
        sc.columns = ["glosa", "peso"]
        sc["bloque"] = info.groupby("subclase").apply(lambda g: g.groupby("bloque").peso.sum().idxmax())
        sc["division"] = [c[:-3] for c in sc.index]
        self.sc_info = sc
        self.sc_idx = self.agregar_por(self.prod_idx, info.subclase)
        # agregados por producto (pertenencia 0/1)
        self.M = pd.DataFrame({a: self._miembros(d) for a, d in AGREGADOS.items()}, index=info.index).astype(float)
        self.agr_idx = self.agregar_matriz(self.prod_idx)
        # bloques de largo plazo: Historia (1998-2017) empalmada con los productos desde 2018
        bl_prod = pd.DataFrame({b: self._indice(self.prod_idx, info.index[info.bloque == b]) for b in HOJAS})
        bl = hist_b[HOJAS].loc[:"2017-12"]
        bl = bl * (bl_prod.loc["2018-01"] / hist_b[HOJAS].loc["2018-01"])
        self.bloques = pd.concat([bl, bl_prod]).sort_index()
        self.validar()
        self.macro = self._macro()
        self.difusion = self._difusion()
        self.agr_largo = self._agregados_largos()

    def _miembros(self, d):
        info = self.prod_info
        if d == "energia_ine":
            return info.index.isin(ENERGIA_INE)
        if d == "sae_ine":
            return ~info.index.isin(ENERGIA_INE) & (info.division != "1")
        return info.bloque.isin(d)

    def _indice(self, idx, codigos):
        w = self.prod_info.peso.loc[codigos]
        return (idx[codigos] * w).sum(axis=1) / w.sum()

    def agregar_por(self, idx, grupo):
        w = self.prod_info.peso
        return pd.DataFrame({g: (idx[c.index] * w[c.index]).sum(axis=1) / w[c.index].sum()
                             for g, c in grupo.groupby(grupo)})

    def agregar_matriz(self, idx):
        """Índices de los agregados desde índices por producto (Laspeyres base 2023)."""
        w = self.prod_info.peso
        return pd.DataFrame({a: (idx * (w * self.M[a])).sum(axis=1) / (w * self.M[a]).sum() for a in self.M})

    def validar(self):
        ine = self.ine
        gen = ine[ine.nivel == 0].set_index("periodo")["Índice"]
        sc_ine = ine[ine.nivel == 4].pivot(index="periodo", columns="codigo", values="Índice")
        an = pd.read_excel(D_INE / "analiticos_base2023.xlsx", header=3).dropna(subset=["Año"])
        an["periodo"] = pd.PeriodIndex([pd.Period(f"{int(a)}-{int(m):02d}", "M") for a, m in zip(an["Año"], an["Mes"])])
        an = an.pivot(index="periodo", columns="Glosa", values="Índice")
        ref = {"IPC": gen, "Alimentos": an["Alimentos"], "Energía (INE)": an["Energía"],
               "IPC SAE (INE)": an["IPC General menos alimentos y energía"],
               "IPC sin volátiles": an["IPC Sin volátiles"], "IPC volátiles": an["IPC volátiles"]}
        filas = []
        for k, s in ref.items():
            a = self.agr_idx[k].loc[s.index]
            filas.append({"serie": k, "fuente": "INE", "dif_max_var_mm_pp": (a.pct_change() - s.pct_change()).abs().max() * 100,
                          "dif_max_nivel_%": ((a / s - 1).abs().max()) * 100})
        d = (self.sc_idx.loc["2023-01":, sc_ine.columns] / sc_ine - 1).abs().max().max() * 100
        filas.append({"serie": "172 subclases", "fuente": "INE", "dif_max_var_mm_pp": np.nan, "dif_max_nivel_%": d})
        self.validacion = pd.DataFrame(filas)

    def _macro(self):
        m = pd.read_csv(D_BDE / "bde_mensual.csv", index_col=0, parse_dates=True)
        m.index = m.index.to_period("M")
        d = pd.read_csv(D_BDE / "bde_diario.csv", index_col=0, parse_dates=True)
        dm = d.groupby(d.index.to_period("M")).mean()
        ult = dm["tcn_d"].dropna().index.max()
        m = m.reindex(pd.period_range(m.index.min(), max(m.index.max(), ult), freq="M"))
        for c, cd in [("tcn", "tcn_d"), ("wti", "wti_d")]:            # mes en curso: promedio de los diarios
            for p in dm.index:
                if pd.isna(m.loc[p, c]) if p in m.index else False:
                    m.loc[p, c] = dm.loc[p, cd]
        m["brent"] = m["brent"].fillna(m["brent"].shift(1) * m["wti"] / m["wti"].shift(1))
        for c in ["bei_swap_1a", "bei_swap_2a"]:
            m[c] = dm[c]
        m["petroleo_clp"] = m["wti"] * m["tcn"]
        m["fao_clp"] = m["fao"] * m["tcn"]
        self.diario = d
        return m

    def _agregados_largos(self):
        """Agregados desde 1998: por bloques (exacto, los bloques salen de los productos desde 2018) y, para los del
        INE (energía y SAE), la serie empalmada de la BDE antes de 2018."""
        w = self.prod_info.groupby("bloque").peso.sum()
        out = {}
        for a, d in AGREGADOS.items():
            if isinstance(d, list):
                out[a] = (self.bloques[d] * w[d]).sum(axis=1) / w[d].sum()
            else:
                col = {"energia_ine": "ipc_ene_ine", "sae_ine": "ipc_sae_ine"}[d]
                corto = self.agr_idx[a]
                largo = self.macro[col] if col in self.macro else pd.Series(dtype=float)
                if largo.dropna().empty:
                    out[a] = corto
                    continue
                largo = largo.loc[:"2017-12"] * corto.loc["2018-01"] / largo.loc["2018-01"]
                out[a] = pd.concat([largo, corto]).sort_index()
        return pd.DataFrame(out)

    def _difusion(self):
        mm = self.sc_idx.pct_change()
        dif = pd.DataFrame({"difusion_subclases": (mm > 0).mean(axis=1) * 100,
                            "difusion_productos": (self.prod_idx.pct_change() > 0).mean(axis=1) * 100}).iloc[1:]
        h = pd.read_csv(INPUTS / "difusion_historica.csv")
        h = h[h.agregado == "IPC"].pivot(index="mes", columns="estadistico", values="valor")
        for c in ["min", "prom", "max"]:
            dif[f"hist_{c}"] = dif.index.month.map(h[c])
        return dif


# =====================================================================================================
# 3. NOWCAST DE COMBUSTIBLES (ENAP)
# =====================================================================================================
def enap_tabla():
    t = pd.DataFrame(ENAP_INFORMES, columns=["vigencia", "gasolina", "diesel", "parafina"])
    t["vigencia"] = pd.to_datetime(t.vigencia)
    t["fuente"] = t.vigencia.dt.strftime("%Y-%m-%d").map(ENAP_FUENTES)
    return t


def enap_precio_diario(escenario, hasta="2026-12-31"):
    """Precio diario a público reconstruido: ancla CNE ± variaciones de ENAP por fecha de vigencia."""
    t = enap_tabla()
    fut = pd.DataFrame(ENAP_ESCENARIOS[escenario], columns=["vigencia", "gasolina", "diesel"])
    fut["vigencia"] = pd.to_datetime(fut.vigencia)
    t = pd.concat([t, fut], ignore_index=True)
    dias = pd.date_range("2026-05-01", hasta, freq="D")
    out = {}
    for prod, (fecha, precio) in ENAP_ANCLAS.items():
        acum = pd.Series(0.0, index=dias)
        for f, v in zip(t.vigencia, t[prod]):
            acum[acum.index >= f] += v
        out[prod] = acum - acum[pd.Timestamp(fecha)] + precio
    return pd.DataFrame(out)


def enap_mensual(escenario):
    """Variación % del promedio mensual del precio (el INE recolecta combustibles todas las semanas)."""
    p = enap_precio_diario(escenario)
    m = p.groupby(p.index.to_period("M")).mean()
    return m, m.pct_change() * 100


def enap_backtest(base):
    _, v = enap_mensual(ENAP_ESCENARIO_CENTRAL)
    filas = []
    for prod, cod in PROD_ENAP.items():
        ine = base.prod_idx[cod].pct_change() * 100
        for per in pd.period_range("2026-07", base.T, freq="M"):
            filas.append({"producto": prod, "periodo": str(per), "nowcast_enap": v.loc[per, prod],
                          "ine_efectivo": ine.loc[per], "error_pp": ine.loc[per] - v.loc[per, prod]})
    return pd.DataFrame(filas)


def nowcast_fijos_productos(base, escenario):
    """Variaciones m/m impuestas por el nowcast para los meses posteriores al último IPC, por producto."""
    _, v = enap_mensual(escenario)
    fijos = {}
    for per in pd.period_range(base.T + 1, "2026-12", freq="M"):
        fijos[per] = {PROD_ENAP[p]: v.loc[per, p] for p in PROD_ENAP}
        if per.month not in range(4, 9):
            fijos[per][PARAFINA] = 0.0
    return fijos


def nowcast_bloque_comb(base, fijos, mm_modelo_resto):
    """m/m del bloque combustibles con pesos efectivos: ENAP + parafina por regla + resto con el modelo."""
    cods = base.prod_info.index[base.prod_info.bloque == "comb"]
    wef = base.prod_info.peso[cods] * base.prod_idx.loc[base.T, cods]
    wef = wef / wef.sum()
    out = {}
    for per, f in fijos.items():
        out[per] = sum(wef[c] * (f[c] if c in f else mm_modelo_resto) for c in cods)
    return out, wef


# =====================================================================================================
# 4. MODELO A: AGREGADOS POR BLOQUES (Carlomagno-Fornero-Sansone + Carlomagno-Espasa)
# =====================================================================================================
def _lag(x, k):
    out = np.full_like(x, np.nan, dtype=float)
    out[k:] = x[:-k] if k else x
    return out


def _d(x):
    return x - _lag(x, 1)


def _mlag(x, a, b):
    return np.mean(np.vstack([_lag(x, k) for k in range(a, b + 1)]), axis=0)


def _pct(x, t, k):
    return 100 * (np.exp((x[t] - x[t - k]) / 100) - 1)


class ModeloBloques:
    """11 bloques, 4 ecuaciones por bloque (M0-M3), combinación por ECM fuera de muestra y agregación Laspeyres."""
    MODELOS = ["M0_estacional", "M1_AR_est", "M2_ARDL_macro", "M3_ECM_tc"]
    INICIO_EST, INICIO_EVAL, INICIO_COMB = "2010-01", "2019-01", "2020-01"
    DRIVERS = {
        "Alim_SV": ["fao_1_6", "tcn_1_6", "eee"], "Bs_sv_ex_alim": ["tcn_0_2", "tcn_3_8", "eee"],
        "ss_sv_ai": ["ipc_1_12", "eee"], "ss_sv_ex_ai": ["rem", "gap", "eee", "ipc_1_12"],
        "comb": ["oil_0", "oil_1", "oil_2"], "alim_rest_vol": ["fao_0_3", "fao_4_9", "tcn_0_3"], "fyv": ["ar12"],
        "bs_vol": ["tcn_0_2", "tcn_3_8"], "ss_v_ai": ["ipc_1_12"], "ss_v_ex_ai": ["oil_1_3", "tcn_0_2"],
    }

    def __init__(self, base):
        self.base = base
        self.H = (H_FIN - base.T).n
        self.per = pd.period_range("1998-12", H_FIN, freq="M")
        self.T0 = self.per.get_loc(base.T)
        self.mes = self.per.month.values
        w = base.prod_info.groupby("bloque").peso.sum()
        self.W = (w[HOJAS] / 100).values
        self.LOG = {b: 100 * np.log(base.bloques[b].reindex(self.per).values) for b in HOJAS}
        mac = base.macro.reindex(self.per)
        self.EXO = {"tcn": 100 * np.log(mac.tcn.values), "oil": 100 * np.log(mac.petroleo_clp.values),
                    "fao": 100 * np.log(mac.fao_clp.values), "rem": 100 * np.log(mac.rem_nom.values),
                    "ima": 100 * np.log(mac.imacec_sa.values), "eee": mac.eee_11m.values, "e23": mac.eee_23m.values}
        self.SEAS = np.vstack([(self.mes == m).astype(float) for m in range(2, 13)])
        self.SEAS_C = np.vstack([(self.mes == m).astype(float) - (self.mes == 12) for m in range(1, 12)])
        self._sl = slice(self.per.get_loc(pd.Period(self.INICIO_EST, "M")), self.per.get_loc(pd.Period("2019-12", "M")) + 1)
        dipc = _d(self.ipc_log(self.LOG))
        self.DELTA = {b: np.nanmean((_d(self.LOG[b]) - dipc)[self._sl]) for b in HOJAS}
        self.pares, self.grupos = self.tendencias_comunes()

    def ipc_log(self, L, cuales=HOJAS):
        idx = [HOJAS.index(b) for b in cuales]
        lv = np.vstack([np.exp(L[b] / 100) for b in cuales])
        return 100 * np.log(self.W[idx] @ lv / self.W[idx].sum())

    def tendencias_comunes(self):
        from statsmodels.tsa.stattools import adfuller
        sl = slice(self.per.get_loc(pd.Period(self.INICIO_EST, "M")), self.T0 + 1)
        pares = []
        for i, a in enumerate(HOJAS):
            for b in HOJAS[i + 1:]:
                p = adfuller((self.LOG[a] - self.LOG[b])[sl] / 100, regression="ct", autolag="AIC")[1]
                pares.append({"bloque_a": a, "bloque_b": b, "p_valor_adf": p, "comparte_tendencia": p < 0.05})
        pares = pd.DataFrame(pares)
        grupos = {}
        for b in HOJAS:
            s = pares[pares.comparte_tendencia & ((pares.bloque_a == b) | (pares.bloque_b == b))]
            grupos[b] = sorted((set(s.bloque_a) | set(s.bloque_b)) - {b})
        return pares, grupos

    def driver(self, nombre, b, L, E, ipc):
        pi = _d(L[b])
        if nombre == "ar12":
            return _lag(pi, 12)
        if nombre == "eee":
            return _lag(E["eee"] - E["e23"], 1) / 12
        if nombre == "rem":
            return _lag(E["rem"] - _lag(E["rem"], 12) - E["e23"], 2) / 12
        if nombre == "gap":
            return _lag(E["ima"] - _mlag(E["ima"], 1, 24), 2)
        v, a, *bb = nombre.split("_")
        a, bb = int(a), int(bb[0]) if bb else int(a)
        x = _d(ipc) - _lag(E["e23"], 1) / 12 if v == "ipc" else _d(E[v])
        return _mlag(x, a, bb)

    def ancla(self, b, E):
        return _lag(E["e23"], 1) / 12 + self.DELTA[b]

    def regresores(self, b, modelo, L, E, ipc):
        y = _d(L[b]) - self.ancla(b, E)
        cols = [*self.SEAS_C]
        if modelo != "M0_estacional":
            cols += [_lag(y, 1), _lag(y, 2)]
        if modelo in ("M2_ARDL_macro", "M3_ECM_tc"):
            cols += [self.driver(n, b, L, E, ipc) for n in self.DRIVERS[b]]
        if modelo == "M3_ECM_tc":
            g = self.grupos[b]
            tend = 100 * np.log(np.vstack([np.exp(L[j] / 100) for j in g]).T @ self.W[[HOJAS.index(j) for j in g]])
            rel, tt = L[b] - tend, np.arange(len(y), dtype=float)
            sl = slice(self.per.get_loc(pd.Period(self.INICIO_EST, "M")), E["_t_orig"] + 1)
            a, c = np.polyfit(tt[sl], rel[sl], 1)
            cols += [_lag(rel - (a * tt + c), 1)]
        return y, np.vstack(cols).T

    def disponible(self, b, modelo):
        if b == "elec":
            return modelo == "M1_AR_est"
        return modelo != "M3_ECM_tc" or len(self.grupos[b]) > 0

    def estimar(self, b, modelo, L, E, ipc, t_fin):
        y, X = self.regresores(b, modelo, L, E, ipc)
        sl = slice(self.per.get_loc(pd.Period(self.INICIO_EST, "M")), t_fin + 1)
        yy, XX = y[sl], X[sl]
        ok = np.isfinite(yy) & np.isfinite(XX).all(axis=1)
        beta, *_ = np.linalg.lstsq(XX[ok], yy[ok], rcond=None)
        return beta, (yy[ok] - XX[ok] @ beta).std(ddof=XX.shape[1])

    def regla_elec(self, t, ipc):
        return ipc[t - 1] - ipc[t - 7] if self.mes[t] in (1, 7) else 0.0

    def estado(self, t_orig, t_exo):
        L = {b: v.copy() for b, v in self.LOG.items()}
        for b in L:
            L[b][t_orig + 1:] = np.nan
        E = {k: v.copy() for k, v in self.EXO.items()}
        for k, x in E.items():
            last = np.max(np.where(np.isfinite(x[: t_exo + 1]))[0])
            if k in ("tcn", "oil", "fao", "e23"):
                x[last + 1:] = x[last]
            elif k == "eee":
                meta = E["e23"][np.max(np.where(np.isfinite(E["e23"][: t_exo + 1]))[0])]
                x[last + 1:] = x[last] + (meta - x[last]) * np.minimum(np.arange(1, len(x) - last), 12) / 12
            else:
                x[last + 1:] = x[last] + (x[last] - x[last - 12]) / 12 * np.arange(1, len(x) - last)
        E["_t_orig"] = t_orig
        return L, E

    def proyectar(self, t_orig, t_exo, pesos=None, h=None, fijos=None):
        """fijos: {t: {bloque: m/m en %}} impone variaciones (nowcast) en meses dados."""
        h = h or self.H
        L0, E = self.estado(t_orig, t_exo)
        ipc0 = self.ipc_log(L0)
        coefs = {(b, m): self.estimar(b, m, L0, E, ipc0, t_orig)
                 for b in HOJAS for m in self.MODELOS if self.disponible(b, m)}
        for b in HOJAS:
            if (b, "M3_ECM_tc") in coefs and coefs[(b, "M3_ECM_tc")][0][-1] >= 0:
                coefs.pop((b, "M3_ECM_tc"))
        tray = {}
        for modo in self.MODELOS + ["Combinado"]:
            L = {b: v.copy() for b, v in L0.items()}
            for t in range(t_orig + 1, t_orig + 1 + h):
                ipc = self.ipc_log(L)
                for b in HOJAS:
                    if b == "elec":
                        p = self.regla_elec(t, ipc)
                    else:
                        preds = {}
                        for m in self.MODELOS:
                            if (b, m) in coefs and (modo == "Combinado" or m == modo):
                                _, X = self.regresores(b, m, L, E, ipc)
                                preds[m] = self.ancla(b, E)[t] + X[t] @ coefs[(b, m)][0]
                        if not preds:
                            _, X = self.regresores(b, "M2_ARDL_macro", L, E, ipc)
                            preds["M2"] = self.ancla(b, E)[t] + X[t] @ coefs[(b, "M2_ARDL_macro")][0]
                        w = np.array([pesos.get((b, m), 1.0) if pesos else 1.0 for m in preds])
                        p = np.dot(w / w.sum(), list(preds.values())) if modo == "Combinado" else next(iter(preds.values()))
                    if fijos and t in fijos and b in fijos[t]:
                        p = fijos[t][b](p) if callable(fijos[t][b]) else 100 * np.log1p(fijos[t][b] / 100)
                    L[b][t] = L[b][t - 1] + p
            tray[modo] = L
        return tray, coefs, E

    def directos(self, t_orig, h=None):
        h = h or self.H
        ipc = self.ipc_log(self.LOG)
        pi = _d(ipc)
        X = np.vstack([np.ones_like(pi), *self.SEAS, _lag(pi, 1), _lag(pi, 2)]).T
        sl = slice(self.per.get_loc(pd.Period(self.INICIO_EST, "M")), t_orig + 1)
        beta, *_ = np.linalg.lstsq(X[sl], pi[sl], rcond=None)
        ar = ipc.copy()
        ar[t_orig + 1:] = np.nan
        for t in range(t_orig + 1, min(t_orig + 1 + h, len(ar))):
            p = _d(ar)
            ar[t] = ar[t - 1] + np.r_[1, self.SEAS[:, t], p[t - 1], p[t - 2]] @ beta
        rw = ipc.copy()
        rw[t_orig + 1:] = np.nan
        n = min(h, len(rw) - t_orig - 1)
        rw[t_orig + 1: t_orig + 1 + n] = ipc[t_orig] + (ipc[t_orig] - ipc[t_orig - 12]) / 12 * np.arange(1, n + 1)
        return {"AR estacional agregado": ar, "Paseo aleatorio 12m": rw}

    def evaluar(self, h=15):
        """Evaluación recursiva: en cada origen se reestima todo y se proyecta h meses."""
        reg_b, reg_t = [], []
        t_ini = self.per.get_loc(pd.Period(self.INICIO_EVAL, "M"))
        for t in range(t_ini, self.T0):
            pesos = None
            if self.per[t] >= pd.Period(self.INICIO_COMB, "M") and reg_b:
                r = pd.DataFrame(reg_b)
                r = r[(r.t_obj <= t) & (r.h <= 12)]
                pesos = (1 / r.groupby(["bloque", "modelo"]).err_acum.apply(lambda e: np.mean(e ** 2))).to_dict()
            tray, _, _ = self.proyectar(t, t, pesos, h=h)
            if pesos is None:
                tray.pop("Combinado")
            for to in range(t + 1, min(t + h, self.T0) + 1):
                for modo, L in tray.items():
                    for b in HOJAS:
                        reg_b.append({"origen": self.per[t], "t_obj": to, "h": to - t, "bloque": b, "modelo": modo,
                                      "err_acum": (self.LOG[b][to] - self.LOG[b][t]) - (L[b][to] - L[b][t])})
            trays = {f"A_{k}": v for k, v in tray.items()}
            reg_t.append((t, trays, self.directos(t, h)))
        self.eval_bloques = pd.DataFrame(reg_b)
        return reg_t

    def correr(self, fijos_productos=None):
        log("Modelo A: evaluación fuera de muestra…")
        self.eval_trays = self.evaluar()
        r = self.eval_bloques[(self.eval_bloques.h <= 12) & (self.eval_bloques.modelo != "Combinado")]
        mse = r.groupby(["bloque", "modelo"]).err_acum.apply(lambda e: np.mean(e ** 2))
        self.pesos = (1 / mse).to_dict()
        tp = (1 / mse).rename("inv_mse").reset_index()
        tp = tp[[self.disponible(b, m) for b, m in zip(tp.bloque, tp.modelo)]]
        tp["peso"] = tp.groupby("bloque").inv_mse.transform(lambda x: x / x.sum())
        self.tabla_pesos = tp
        # proyección final sin y con nowcast
        tray, self.coefs, self.E = self.proyectar(self.T0, self.T0 + 1, self.pesos)
        self.tray = tray
        if fijos_productos:
            fijos = {}
            for per, f in fijos_productos.items():
                t = self.per.get_loc(per)
                fijos[t] = {"comb": (lambda p, f=f: 100 * np.log1p(
                    nowcast_bloque_comb(self.base, {per: f}, 100 * np.expm1(p / 100))[0][per] / 100))}
            tray_nc, _, _ = self.proyectar(self.T0, self.T0 + 1, self.pesos, fijos=fijos)
            self.tray_nc = tray_nc
        return self

    def niveles(self, L):
        """Índices de bloques (base 2023=100) desde log-niveles, como DataFrame."""
        return pd.DataFrame({b: np.exp(L[b] / 100) for b in HOJAS}, index=self.per)

    def agregados(self, L):
        lv = self.niveles(L)
        w = self.base.prod_info.groupby("bloque").peso.sum()
        out = {}
        for a, d in AGREGADOS.items():
            if isinstance(d, list):
                out[a] = (lv[d] * w[d]).sum(axis=1) / w[d].sum()
        return pd.DataFrame(out)


# =====================================================================================================
# 5. MODELO B: 172 SUBCLASES
# =====================================================================================================
class ModeloSubclases:
    """Una ecuación por subclase para su variación mensual, anclada a las expectativas (EEE 23 meses):

        π_s,t − a_s,t = Σ_m s_m (D_m − D_12) + φ (π_s,t−1 − a_s,t−1) + γ z_s,t−1 + ε

    a_s,t = π^e_23m/12 + δ_s, con δ_s = λ · promedio(π_s − π_IPC) (deriva relativa contraída, λ = 0,5).
    Por estabilidad, |φ| ≤ 0,8 (si la estimación lo excede, se fija en el límite y se reestima el resto).
    z_s = ln I_s − ln I_bloque(−s) desviado de su media: corrección de error frente a la tendencia común del
    bloque al que pertenece (Carlomagno-Espasa), incluida solo si el ADF del precio relativo rechaza raíz
    unitaria al 10% y γ̂ < 0. La electricidad usa la regla de indexación semestral de las tarifas reguladas.
    """
    INICIO_EVAL = "2021-01"
    LAMBDA = 0.5
    PHI_MAX = 0.8

    def __init__(self, base):
        self.base = base
        self.H = (H_FIN - base.T).n
        self.per = pd.period_range("2018-01", H_FIN, freq="M")
        self.T0 = self.per.get_loc(base.T)
        self.mes = self.per.month.values
        self.sc = list(base.sc_info.index)
        self.w = base.sc_info.peso.loc[self.sc].values
        self.bloque = base.sc_info.bloque.loc[self.sc].values
        self.L = 100 * np.log(base.sc_idx[self.sc].reindex(self.per).values)          # (T, N)
        e23 = base.macro.eee_23m.reindex(self.per).ffill().values
        self.a0 = np.r_[np.nan, e23[:-1]] / 12                                       # ancla común
        self.SEAS_C = np.vstack([(self.mes == m).astype(float) - (self.mes == 12) for m in range(1, 12)]).T
        self.elec = [i for i, c in enumerate(self.sc) if self.bloque[i] == "elec"]
        self.ecm = self._test_tendencias()

    def _bloque_excl(self, L):
        """ln índice del bloque de cada subclase, excluyéndola (T, N)."""
        lv = np.exp(L / 100) * self.w
        out = np.full_like(L, np.nan)
        for b in set(self.bloque):
            j = np.where(self.bloque == b)[0]
            tot, wt = lv[:, j].sum(axis=1, keepdims=True), self.w[j].sum()
            if len(j) > 1:
                out[:, j] = 100 * np.log((tot - lv[:, j]) / (wt - self.w[j]))
        return out

    def _test_tendencias(self):
        from statsmodels.tsa.stattools import adfuller
        rel = (self.L - self._bloque_excl(self.L))[: self.T0 + 1]
        ok = np.zeros(len(self.sc), bool)
        self.p_adf = np.full(len(self.sc), np.nan)
        for i in range(len(self.sc)):
            x = rel[:, i]
            if np.isfinite(x).sum() > 36 and np.nanstd(x) > 0:
                self.p_adf[i] = adfuller(x[np.isfinite(x)] / 100, regression="c", autolag="AIC")[1]
                ok[i] = self.p_adf[i] < 0.10
        return ok

    def ipc(self, L):
        return 100 * np.log((np.exp(L / 100) * self.w).sum(axis=1) / self.w.sum())

    def proyectar(self, t0, h=None, fijos=None):
        """fijos: {t: {índice de subclase: m/m %}} impone variaciones (nowcast)."""
        h = h or self.H
        L = self.L.copy()
        L[t0 + 1:] = np.nan
        pi = np.diff(np.vstack([np.full((1, L.shape[1]), np.nan), L]), axis=0)
        pipc = np.r_[np.nan, np.diff(self.ipc(L))]
        sl = slice(2, t0 + 1)
        delta = self.LAMBDA * np.nanmean(pi[sl] - pipc[sl, None], axis=0)
        a = self.a0[:, None] + delta[None, :]
        y = pi - a
        zraw = L - self._bloque_excl(L)
        zmed = np.nanmean(zraw[sl], axis=0)
        n, k = len(self.sc), 13
        B = np.zeros((n, k))
        for i in range(n):
            cols = [self.SEAS_C, np.r_[np.nan, y[:-1, i]][:, None]]
            if self.ecm[i]:
                cols.append(np.r_[np.nan, zraw[:-1, i] - zmed[i]][:, None])
            X = np.hstack(cols)[sl]
            yy = y[sl, i]
            ok = np.isfinite(yy) & np.isfinite(X).all(axis=1)
            X, yy = X[ok], yy[ok]
            if self.ecm[i]:
                b, *_ = np.linalg.lstsq(X, yy, rcond=None)
                if b[-1] >= 0:                      # sin corrección efectiva: se estima sin ECM
                    X = X[:, :-1]
            b, *_ = np.linalg.lstsq(X, yy, rcond=None)
            if abs(b[11]) > self.PHI_MAX:           # estabilidad: |φ| acotado y se reestima el resto
                phi = np.sign(b[11]) * self.PHI_MAX
                otros = [c for c in range(X.shape[1]) if c != 11]
                r, *_ = np.linalg.lstsq(X[:, otros], yy - phi * X[:, 11], rcond=None)
                b = np.insert(r, 11, phi)
            B[i, : len(b)] = b
        self.coef = B
        for t in range(t0 + 1, min(t0 + 1 + h, len(self.per))):
            z = (L[t - 1] - self._bloque_excl(L[t - 1: t])[0]) - zmed
            yhat = self.SEAS_C[t] @ B[:, :11].T + B[:, 11] * y[t - 1] + B[:, 12] * np.where(self.ecm, z, 0)
            p = a[t] + yhat
            ipcl = self.ipc(L[: t])
            for i in self.elec:
                p[i] = ipcl[t - 1] - ipcl[t - 7] if self.mes[t] in (1, 7) else 0.0
            if fijos and t in fijos:
                for i, v in fijos[t].items():
                    p[i] = 100 * np.log1p(v / 100)
            L[t] = L[t - 1] + p
            y[t] = p - a[t]
        return L

    def fijos_desde_productos(self, fijos_prod):
        """Traduce el nowcast por producto a subclases (promedio ponderado; el resto queda con el modelo)."""
        out = {}
        info = self.base.prod_info
        for per, f in fijos_prod.items():
            t = self.per.get_loc(per)
            out[t] = {}
            for cod, v in f.items():
                sc = info.subclase[cod]
                prods = info.index[info.subclase == sc]
                if len(prods) == 1 or all(c in f for c in prods):
                    w = info.peso[prods]
                    out[t][self.sc.index(sc)] = sum(w[c] * f[c] for c in prods) / w.sum()
        return out

    def niveles(self, L):
        return pd.DataFrame(np.exp(L / 100), index=self.per, columns=self.sc)

    def agregados(self, L):
        """Agregados por producto: cada producto hereda la variación proyectada de su subclase."""
        sc_lv = self.niveles(L)
        info = self.base.prod_info
        base_p = self.base.prod_idx.reindex(self.per)
        rel = sc_lv / sc_lv.loc[self.base.T]
        prod = base_p.copy()
        fut = self.per[self.per > self.base.T]
        sub = info.subclase.reindex(base_p.columns).values          # mismo orden de columnas que el panel
        prod.loc[fut] = rel.loc[fut, sub].values * base_p.loc[self.base.T].values
        return self.base.agregar_matriz(prod), prod

    def evaluar(self, h=15):
        trays = []
        for t in range(self.per.get_loc(pd.Period(self.INICIO_EVAL, "M")), self.T0):
            trays.append((t, self.proyectar(t, h=h)))
        return trays

    def correr(self, fijos_productos=None):
        log("Modelo B: evaluación fuera de muestra…")
        self.eval_trays = self.evaluar()
        self.Lp = self.proyectar(self.T0)
        self.coef_final = self.coef.copy()
        self.Lp_nc = self.proyectar(self.T0, fijos=self.fijos_desde_productos(fijos_productos)) if fijos_productos else self.Lp
        return self


# =====================================================================================================
# 6. EVALUACIÓN COMPARADA
# =====================================================================================================
def _agr_log_bloques(A, L, cuales):
    return A.ipc_log(L, cuales)


def _agr_log_sub(B, L, cuales):
    j = np.isin(B.bloque, cuales)
    return 100 * np.log((np.exp(L[:, j] / 100) * B.w[j]).sum(axis=1) / B.w[j].sum())


def evaluacion_comparada(A, B, base, desde="2021-01", h=12):
    """Errores fuera de muestra de ambos modelos y su combinación sobre los mismos orígenes y objetivos."""
    obj = {"IPC": HOJAS, "IPC sin volátiles": SV, "IPC volátiles": VOL}
    obs = {k: 100 * np.log(base.agr_idx[k]) for k in obj}
    eee_m = base.macro.eee_mm_mes
    eee_s = base.macro.eee_mm_sig
    trays_b = dict(B.eval_trays)
    filas = []
    for t, tray_a, dirs in A.eval_trays:
        per_t = A.per[t]
        if per_t < pd.Period(desde, "M") or "A_Combinado" not in tray_a:
            continue
        tb = B.per.get_loc(per_t)
        if tb not in trays_b:
            continue
        for ag, cuales in obj.items():
            series = {}
            for k in ["A_Combinado", "A_M1_AR_est", "A_M3_ECM_tc"]:
                x = _agr_log_bloques(A, tray_a[k], cuales)
                series[k.replace("A_", "A: ").replace("_", " ")] = pd.Series(x, index=A.per)
            series["B: subclases"] = pd.Series(_agr_log_sub(B, trays_b[tb], cuales), index=B.per)
            series["Combinación A+B"] = (series["A: Combinado"].reindex(B.per) + series["B: subclases"]) / 2
            if ag == "IPC":
                for k, v in dirs.items():
                    series[k] = pd.Series(v, index=A.per)
            o = obs[ag]
            for hh in range(1, h + 1):
                per_o = per_t + hh
                if per_o > base.T:
                    break
                for k, x in series.items():
                    x = x.copy()
                    x.loc[:per_t] = o.reindex(x.index).loc[:per_t]
                    f_mm = 100 * np.expm1((x[per_o] - x[per_o - 1]) / 100)
                    f_12 = 100 * np.expm1((x[per_o] - x[per_o - 12]) / 100)
                    filas.append({"origen": str(per_t), "objetivo": str(per_o), "h": hh, "agregado": ag, "modelo": k,
                                  "err_mm": 100 * np.expm1((o[per_o] - o[per_o - 1]) / 100) - f_mm,
                                  "err_12m": 100 * np.expm1((o[per_o] - o[per_o - 12]) / 100) - f_12})
            if ag == "IPC":
                for hh, e in [(1, eee_m.get(per_t + 1)), (2, eee_s.get(per_t + 1))]:
                    per_o = per_t + hh
                    if per_o <= base.T and e is not None and np.isfinite(e):
                        filas.append({"origen": str(per_t), "objetivo": str(per_o), "h": hh, "agregado": ag,
                                      "modelo": "EEE (mediana)", "err_mm": 100 * np.expm1((o[per_o] - o[per_o - 1]) / 100) - e,
                                      "err_12m": np.nan})
    return pd.DataFrame(filas)


def resumen_rmse(ev):
    r = ev.groupby(["agregado", "modelo", "h"]).agg(
        rmse_mm=("err_mm", lambda e: np.sqrt(np.mean(e ** 2))), rmse_12m=("err_12m", lambda e: np.sqrt(np.mean(e ** 2))),
        sesgo_12m=("err_12m", "mean"), n=("err_mm", "size")).reset_index()
    return r


def dm_test(ev, m1, m2, h, col, agregado="IPC"):
    from scipy.stats import norm
    a = ev[(ev.modelo == m1) & (ev.h == h) & (ev.agregado == agregado)].set_index("objetivo")[col]
    b = ev[(ev.modelo == m2) & (ev.h == h) & (ev.agregado == agregado)].set_index("objetivo")[col]
    j = pd.concat([a, b], axis=1, keys=["a", "b"]).dropna()
    dd = (j.a ** 2 - j.b ** 2).values
    n = len(dd)
    if n < 10:
        return np.nan, np.nan, n
    g = [np.mean((dd[k:] - dd.mean()) * (dd[: n - k] - dd.mean())) for k in range(h)]
    v = (g[0] + 2 * sum((1 - k / h) * g[k] for k in range(1, h))) / n
    st = dd.mean() / np.sqrt(v)
    return st, 2 * (1 - norm.cdf(abs(st))), n


# =====================================================================================================
# 7. RESULTADOS
# =====================================================================================================
def _tabla_proy(agr, base, desde="2025-01"):
    """m/m, 12m, dic/dic y promedio anual de cada agregado."""
    agr = agr.loc[:H_FIN]
    mm, a12 = agr.pct_change() * 100, agr.pct_change(12) * 100
    anual = {}
    for y in (2025, 2026, 2027):
        prom = agr.loc[f"{y}-01":f"{y}-12"].mean() / agr.loc[f"{y - 1}-01":f"{y - 1}-12"].mean() * 100 - 100
        anual[f"dic-{str(y)[2:]} (12m)"] = a12.loc[f"{y}-12"]
        anual[f"promedio {y}"] = prom
    return mm.loc[desde:], a12.loc[desde:], pd.DataFrame(anual)


def resultados(base, A, B):
    R = {"base": base, "A": A, "B": B, "T": base.T}
    R["agr_A"] = A.agregados(A.tray_nc["Combinado"]).loc["2018-01":]
    R["agr_A_sin_nc"] = A.agregados(A.tray["Combinado"]).loc["2018-01":]
    R["agr_A_mod"] = {m: A.agregados(A.tray_nc[m] if hasattr(A, "tray_nc") else A.tray[m]) for m in A.MODELOS}
    R["agr_B"], R["prod_B"] = B.agregados(B.Lp_nc)
    R["agr_B_sin_nc"], _ = B.agregados(B.Lp)
    for k in ["A", "B", "A_sin_nc", "B_sin_nc"]:
        R[f"mm_{k}"], R[f"a12_{k}"], R[f"anual_{k}"] = _tabla_proy(R[f"agr_{k}"], base)
    # combinación A+B (promedio de log-niveles) para los agregados que ambos proyectan
    comunes = [a for a in R["agr_A"].columns if a in R["agr_B"].columns]
    R["agr_AB"] = np.exp((np.log(R["agr_A"][comunes]) + np.log(R["agr_B"].loc[R["agr_A"].index, comunes])) / 2)
    R["mm_AB"], R["a12_AB"], R["anual_AB"] = _tabla_proy(R["agr_AB"], base)
    # subclases: m/m proyectado (con nowcast) e incidencias
    sc_lv = B.niveles(B.Lp_nc)
    R["sc_lv"] = sc_lv
    R["sc_mm"] = (sc_lv.pct_change() * 100).loc[base.T + 1:]
    w = base.sc_info.peso
    ipc = (sc_lv * w).sum(axis=1) / 100
    inc = lambda a, b: (sc_lv.loc[a] - sc_lv.loc[b]) * w / 100 / ipc.loc[b] * 100
    R["sc_inc"] = pd.DataFrame({"glosa": base.sc_info.glosa, "peso": w, "bloque": base.sc_info.bloque,
                                "var_12m_dic26": (sc_lv.loc["2026-12"] / sc_lv.loc["2025-12"] - 1) * 100,
                                "incid_dic26": inc(pd.Period("2026-12", "M"), pd.Period("2025-12", "M")),
                                "var_12m_dic27": (sc_lv.loc["2027-12"] / sc_lv.loc["2026-12"] - 1) * 100,
                                "incid_dic27": inc(pd.Period("2027-12", "M"), pd.Period("2026-12", "M"))})
    # bloques (modelo A) e incidencias
    lvA = A.niveles(A.tray_nc["Combinado"])
    wb = base.prod_info.groupby("bloque").peso.sum()
    ipcA = (lvA * wb).sum(axis=1) / 100
    R["bloques_inc"] = pd.DataFrame({
        "bloque": [BLOQUES[b][0] for b in HOJAS], "peso_2023": wb[HOJAS].values,
        "var_12m_dic26": (lvA.loc["2026-12", HOJAS] / lvA.loc["2025-12", HOJAS] - 1).values * 100,
        "incid_dic26": ((lvA.loc["2026-12", HOJAS] - lvA.loc["2025-12", HOJAS]) * wb[HOJAS] / 100 / ipcA["2025-12"]).values,
        "var_12m_dic27": (lvA.loc["2027-12", HOJAS] / lvA.loc["2026-12", HOJAS] - 1).values * 100,
        "incid_dic27": ((lvA.loc["2027-12", HOJAS] - lvA.loc["2026-12", HOJAS]) * wb[HOJAS] / 100 / ipcA["2026-12"]).values})
    # evaluación
    log("Evaluación comparada…")
    R["ev"] = evaluacion_comparada(A, B, base)
    R["rmse"] = resumen_rmse(R["ev"])
    R["rmse_pos23"] = resumen_rmse(R["ev"][R["ev"].origen >= "2023-07"])
    dm = []
    for muestra, ev in [("2021-2026", R["ev"]), ("jul.2023-2026", R["ev"][R["ev"].origen >= "2023-07"])]:
        for m1, m2 in [("A: Combinado", "AR estacional agregado"), ("A: Combinado", "Paseo aleatorio 12m"),
                       ("A: M3 ECM tc", "A: Combinado"), ("B: subclases", "A: Combinado"),
                       ("Combinación A+B", "A: Combinado"), ("A: Combinado", "EEE (mediana)")]:
            for h, col in [(1, "err_mm"), (3, "err_12m"), (12, "err_12m")]:
                if "EEE" in m2 and h != 1:
                    continue
                st, pv, n = dm_test(ev, m1, m2, h, col)
                dm.append({"muestra": muestra, "modelo": m1, "rival": m2, "h": h,
                           "variable": "m/m" if col == "err_mm" else "12m", "DM": st, "p_valor": pv, "n": n})
    R["dm"] = pd.DataFrame(dm)
    # nowcast
    R["enap_bt"] = enap_backtest(base)
    R["enap_mensual"] = {e: enap_mensual(e) for e in ENAP_ESCENARIOS}
    R["enap_diario"] = {e: enap_precio_diario(e) for e in ENAP_ESCENARIOS}
    # pronóstico del IPC de octubre hecho con información a septiembre (sin y con ENAP): validación
    t_sep = A.per.get_loc(base.T - 1)
    tray_sep, _, _ = A.proyectar(t_sep, t_sep + 1, A.pesos, h=1)
    R["oct_modelo_A"] = 100 * np.expm1((A.ipc_log(tray_sep["Combinado"])[t_sep + 1] - A.ipc_log(tray_sep["Combinado"])[t_sep]) / 100)
    Lb = B.proyectar(B.per.get_loc(base.T - 1), h=1)
    R["oct_modelo_B"] = 100 * np.expm1((B.ipc(Lb)[B.per.get_loc(base.T)] - B.ipc(Lb)[B.per.get_loc(base.T - 1)]) / 100)
    R["oct_efectivo"] = base.agr_idx["IPC"].pct_change().loc[base.T] * 100
    eb = A.eval_bloques
    x = eb[(eb.bloque == "comb") & (eb.h == 1) & (eb.modelo == "Combinado")].err_acum
    R["rmse_comb_h1"] = float(np.sqrt(np.mean(x ** 2)))
    return R


def guardar_csv(R):
    TABLAS.mkdir(parents=True, exist_ok=True)
    for k in ["a12_A", "a12_B", "a12_AB", "mm_A", "mm_B", "anual_A", "anual_B", "anual_AB", "anual_A_sin_nc",
              "anual_B_sin_nc", "sc_mm", "rmse", "rmse_pos23", "dm", "enap_bt", "sc_inc", "bloques_inc"]:
        R[k].to_csv(TABLAS / f"{k}.csv")
    R["base"].validacion.to_csv(TABLAS / "validacion.csv", index=False)
    R["ev"].to_csv(TABLAS / "eval_errores.csv", index=False)


# =====================================================================================================
# 8. EXCEL AUTOCONTENIDO (con fórmulas)
# =====================================================================================================
def construir_excel(R, destino=None):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter as L_
    base = R["base"]
    destino = destino or OUTPUTS / "IPC_Chile_Proyecciones.xlsx"
    wb = openpyxl.Workbook()
    F_HEAD = PatternFill("solid", fgColor="DDE7F3")
    F_IN = PatternFill("solid", fgColor="FFF2CC")
    F_PROY = PatternFill("solid", fgColor="F3F3F3")
    B_ = Font(bold=True)
    meses = pd.period_range("2018-01", H_FIN, freq="M")
    fut = meses[meses > base.T]
    C0 = 3                                                   # primera columna de meses en hojas de series
    col_mes = {p: C0 + i for i, p in enumerate(meses)}
    sc = list(base.sc_info.index)
    n_sc = len(sc)
    r_sc = {c: 2 + i for i, c in enumerate(sc)}
    agr = list(AGREGADOS)
    flag_col = {a: 7 + i for i, a in enumerate(agr)}         # columnas de pertenencia en Matriz (G...)
    share_col = {a: 6 + i for i, a in enumerate(agr)}        # columnas de participación en Subclases (F...)
    last_sc = 1 + n_sc

    def hoja(nombre, encabezado, anchos=None):
        ws = wb.create_sheet(nombre)
        for j, h in enumerate(encabezado, start=1):
            c = ws.cell(1, j, h)
            c.font, c.fill = B_, F_HEAD
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for j, w in (anchos or {}).items():
            ws.column_dimensions[L_(j)].width = w
        return ws

    def cab_meses(ws, primera_col=C0, desde=meses):
        for p in desde:
            c = ws.cell(1, col_mes[p] if primera_col == C0 else primera_col + list(desde).index(p), p.strftime("%Y-%m"))
            c.font, c.fill = B_, F_PROY if p > base.T else F_HEAD

    # ---------------------------------------------------------------------- LEEME
    ws = wb.active
    ws.title = "LEEME"
    texto = [
        ("IPC Chile: base, proyecciones y nowcast (libro autocontenido)", True),
        (f"Último IPC efectivo: {mes_txt(base.T)}. Proyección: {mes_txt(base.T + 1)} a {mes_txt(H_FIN)}. "
         f"Generado por proyeccion_ipc.py el {pd.Timestamp.today():%d-%m-%Y}.", False),
        ("", False),
        ("Cómo funciona", True),
        ("1. Matriz: 283 productos del IPC 2023=100 con su subclase, peso y pertenencia (0/1) a cada agregado.", False),
        ("2. Subclases: 172 subclases; el peso y la participación en cada agregado se calculan con fórmulas desde Matriz.", False),
        ("3. Indices: historia efectiva (INE, y empalme del Excel automatizado para 2018-2022) y proyección por fórmula:", False),
        ("   índice(t) = índice(t-1) × (1 + Proy_mm(t)/100).", False),
        ("4. Proy_mm (celdas amarillas = supuestos editables): variación mensual proyectada por subclase (Modelo B).", False),
        ("   Gasolina y diésel de nov-dic están enlazados a Nowcast_ENAP: al cambiar un informe de ENAP se recalcula todo.", False),
        ("5. Agregados: índices de IPC, sin volátiles, volátiles, alimentos, energía, SAE, etc. con SUMPRODUCT.", False),
        ("6. Var_mm y Var_12m: variaciones de los agregados. Incidencias: aporte de cada componente al IPC.", False),
        ("7. Var_mm_subclases y Difusion: % de subclases con alza en cada mes, contra el rango histórico 2001-2023.", False),
        ("8. Resumen: dic/dic y promedio anual 2026-2027 de este libro (Modelo B editable) y del Modelo A (bloques).", False),
        ("9. Resumen_Proyecciones: var. mensual y en 12 meses 2025-2027 de las clases del BCCh (como en el Excel automatizado),", False),
        ("   con el Modelo A y la EEE como memo, y dic/dic y promedios anuales.", False),
        ("10. Perfiles_Analiticos: m/m por mes calendario 2010-2027 de cada clase, con promedio 2010-2019 y 2013-2019,", False),
        ("    máximo, mínimo y rango; desde 2018 enlazado a Var_mm (la proyección se actualiza sola).", False),
        ("11. Productos: 283 productos; historia del INE y proyección = índice del último mes × variación de su subclase", False),
        ("    (hoja Indices). La última fila recalcula el IPC desde los productos como control.", False),
        ("12. Modelo_agregados, Evaluacion, Macro_BDE, Fuentes: resultados y datos de respaldo (valores).", False),
        ("", False),
        ("Colores: amarillo = supuesto editable; gris = mes proyectado; azul = encabezado.", False),
        ("Para actualizar con un nuevo IPC: correr `python proyeccion_ipc.py` (carpeta Code), que descarga BDE e INE y regenera este libro.", False),
    ]
    for i, (t, negrita) in enumerate(texto, start=1):
        c = ws.cell(i, 1, t)
        c.font = Font(bold=negrita, size=13 if i == 1 else 11)
    ws.column_dimensions["A"].width = 130

    # ---------------------------------------------------------------------- Matriz
    info = base.prod_info.sort_values("subclase")
    ws = hoja("Matriz", ["Código producto", "Glosa", "Subclase", "División", "Peso 2023", "Bloque"] + agr,
              {2: 42, 6: 16})
    for i, (cod, r) in enumerate(info.iterrows(), start=2):
        ws.cell(i, 1, int(cod)); ws.cell(i, 2, r.glosa); ws.cell(i, 3, str(r.subclase)); ws.cell(i, 4, str(r.division))
        ws.cell(i, 5, float(r.peso)); ws.cell(i, 6, BLOQUES[r.bloque][0])
        for a in agr:
            ws.cell(i, flag_col[a], int(base.M.loc[cod, a]))
    ws.freeze_panes = "C2"
    n_p = len(info) + 1
    rng = lambda col: f"Matriz!${L_(col)}$2:${L_(col)}${n_p}"

    # ---------------------------------------------------------------------- Subclases
    ws = hoja("Subclases", ["Código subclase", "Glosa", "División", "Bloque principal", "Peso 2023 (fórmula)"] +
              [f"Participación en {a}" for a in agr], {2: 40, 4: 22})
    for c in sc:
        r = r_sc[c]
        ws.cell(r, 1, str(c)); ws.cell(r, 2, base.sc_info.glosa[c]); ws.cell(r, 3, base.sc_info.division[c])
        ws.cell(r, 4, BLOQUES[base.sc_info.bloque[c]][0])
        ws.cell(r, 5, f"=SUMIF({rng(3)},$A{r},{rng(5)})")
        for a in agr:
            ws.cell(r, share_col[a], f"=IFERROR(SUMIFS({rng(5)},{rng(3)},$A{r},{rng(flag_col[a])},1)/$E{r},0)")
    ws.freeze_panes = "C2"

    # ---------------------------------------------------------------------- Proy_mm (supuestos)
    ws = hoja("Proy_mm", ["Código subclase", "Glosa"], {2: 40})
    pcol = {p: 3 + i for i, p in enumerate(fut)}
    for p in fut:
        c = ws.cell(1, pcol[p], p.strftime("%Y-%m")); c.font, c.fill = B_, F_PROY
    for c in sc:
        r = r_sc[c]
        ws.cell(r, 1, str(c)); ws.cell(r, 2, base.sc_info.glosa[c])
        for p in fut:
            cell = ws.cell(r, pcol[p], round(float(R["sc_mm"].loc[p, c]), 4))
            cell.fill = F_IN
    ws.freeze_panes = "C2"

    # ---------------------------------------------------------------------- Nowcast_ENAP
    wn = hoja("Nowcast_ENAP", ["Vigencia", "Gasolina ($/l)", "Diésel ($/l)", "Fuente / escenario"], {1: 12, 4: 60})
    tab = enap_tabla()
    filas = [(r.vigencia, r.gasolina, r.diesel, r.fuente) for r in tab.itertuples()]
    filas += [(pd.Timestamp(f), g, d, f"Escenario {ENAP_ESCENARIO_CENTRAL} (editable)")
              for f, g, d in ENAP_ESCENARIOS[ENAP_ESCENARIO_CENTRAL]]
    n_inf = len(filas) + 1
    for i, (f, g, d, src) in enumerate(filas, start=2):
        wn.cell(i, 1, f.to_pydatetime()).number_format = "yyyy-mm-dd"
        wn.cell(i, 2, g); wn.cell(i, 3, d); wn.cell(i, 4, src)
        if "Escenario" in src:
            for j in (1, 2, 3):
                wn.cell(i, j).fill = F_IN
    for j, h in enumerate(["Producto", "Fecha ancla", "Precio ancla ($/l)"], start=6):
        c = wn.cell(1, j, h); c.font, c.fill = B_, F_HEAD
    for i, (prod, (f, v)) in enumerate(ENAP_ANCLAS.items(), start=2):
        wn.cell(i, 6, prod); wn.cell(i, 7, pd.Timestamp(f).to_pydatetime()).number_format = "yyyy-mm-dd"
        wn.cell(i, 8, v).fill = F_IN
    for j, h in enumerate(["Fecha", "Mes (aaaamm)", "Gasolina diario", "Diésel diario"], start=10):
        c = wn.cell(1, j, h); c.font, c.fill = B_, F_HEAD
    dias = pd.date_range("2026-05-01", "2026-12-31", freq="D")
    inf_rng = lambda col: f"${col}$2:${col}${n_inf}"
    for i, dday in enumerate(dias, start=2):
        wn.cell(i, 10, dday.to_pydatetime()).number_format = "yyyy-mm-dd"
        wn.cell(i, 11, f"=YEAR(J{i})*100+MONTH(J{i})")
        wn.cell(i, 12, f"=$H$2+SUMIFS({inf_rng('B')},{inf_rng('A')},\"<=\"&J{i})-SUMIFS({inf_rng('B')},{inf_rng('A')},\"<=\"&$G$2)")
        wn.cell(i, 13, f"=$H$3+SUMIFS({inf_rng('C')},{inf_rng('A')},\"<=\"&J{i})-SUMIFS({inf_rng('C')},{inf_rng('A')},\"<=\"&$G$3)")
    n_d = len(dias) + 1
    for j, h in enumerate(["Mes (aaaamm)", "Gasolina promedio", "Diésel promedio", "Gasolina % m/m", "Diésel % m/m"], start=15):
        c = wn.cell(1, j, h); c.font, c.fill = B_, F_HEAD
    meses_enap = pd.period_range("2026-05", "2026-12", freq="M")
    fila_mes_enap = {}
    for i, p in enumerate(meses_enap, start=2):
        fila_mes_enap[p] = i
        wn.cell(i, 15, p.year * 100 + p.month)
        wn.cell(i, 16, f"=AVERAGEIFS($L$2:$L${n_d},$K$2:$K${n_d},O{i})")
        wn.cell(i, 17, f"=AVERAGEIFS($M$2:$M${n_d},$K$2:$K${n_d},O{i})")
        if i > 2:
            wn.cell(i, 18, f"=(P{i}/P{i - 1}-1)*100"); wn.cell(i, 19, f"=(Q{i}/Q{i - 1}-1)*100")
    wn.freeze_panes = "A2"
    # gasolina y diésel en Proy_mm enlazados al nowcast
    wsp = wb["Proy_mm"]
    for prod, colx in [("gasolina", "R"), ("diesel", "S")]:
        c_sc = base.prod_info.subclase[PROD_ENAP[prod]]
        for p in fut:
            if p in fila_mes_enap:
                cell = wsp.cell(r_sc[c_sc], pcol[p], f"=Nowcast_ENAP!{colx}{fila_mes_enap[p]}")
                cell.fill = PatternFill("solid", fgColor="E2EFDA")

    # ---------------------------------------------------------------------- Indices (historia + fórmulas)
    ws = hoja("Indices", ["Código subclase", "Glosa"], {2: 40})
    cab_meses(ws)
    for c in sc:
        r = r_sc[c]
        ws.cell(r, 1, str(c)); ws.cell(r, 2, base.sc_info.glosa[c])
        for p in meses:
            j = col_mes[p]
            if p <= base.T:
                ws.cell(r, j, round(float(base.sc_idx.loc[p, c]), 4))
            else:
                ws.cell(r, j, f"={L_(j - 1)}{r}*(1+Proy_mm!{L_(pcol[p])}{r}/100)").fill = F_PROY
    ws.freeze_panes = "C2"

    # ---------------------------------------------------------------------- Var_mm_subclases
    ws = hoja("Var_mm_subclases", ["Código subclase", "Glosa"], {2: 40})
    cab_meses(ws)
    for c in sc:
        r = r_sc[c]
        ws.cell(r, 1, str(c)); ws.cell(r, 2, base.sc_info.glosa[c])
        for p in meses[1:]:
            j = col_mes[p]
            ws.cell(r, j, f"=(Indices!{L_(j)}{r}/Indices!{L_(j - 1)}{r}-1)*100")
    ws.freeze_panes = "C2"

    # ---------------------------------------------------------------------- Agregados, Var_mm, Var_12m
    w_rng = f"Subclases!$E$2:$E${last_sc}"
    sh_rng = lambda a: f"Subclases!${L_(share_col[a])}$2:${L_(share_col[a])}${last_sc}"
    r_ag = {a: 2 + i for i, a in enumerate(agr)}
    ws = hoja("Agregados", ["Agregado", "Peso 2023"], {1: 26})
    cab_meses(ws)
    for a in agr:
        r = r_ag[a]
        ws.cell(r, 1, a).font = B_ if a == "IPC" else Font()
        ws.cell(r, 2, f"=SUMPRODUCT({w_rng},{sh_rng(a)})")
        for p in meses:
            j = col_mes[p]
            ws.cell(r, j, f"=SUMPRODUCT({w_rng}*{sh_rng(a)},Indices!{L_(j)}$2:{L_(j)}${last_sc})/$B{r}")
    ws.freeze_panes = "C2"
    for nombre, k in [("Var_mm", 1), ("Var_12m", 12)]:
        ws = hoja(nombre, ["Agregado", ""], {1: 26})
        cab_meses(ws)
        for a in agr:
            r = r_ag[a]
            ws.cell(r, 1, a)
            for p in meses[k:]:
                j = col_mes[p]
                ws.cell(r, j, f"=(Agregados!{L_(j)}{r}/Agregados!{L_(j - k)}{r}-1)*100").number_format = "0.00"
        ws.freeze_panes = "C2"

    # ---------------------------------------------------------------------- Incidencias
    comp = ["Bienes SV", "Servicios SV", "Energía volátil", "Alimentos volátiles", "Resto volátiles"]
    ws = hoja("Incidencias", ["Componente", "Medida"], {1: 24, 2: 14})
    cab_meses(ws)
    fila = 2
    for medida, k in [("m/m (pp)", 1), ("12m (pp)", 12)]:
        for a in comp + ["Suma"]:
            ws.cell(fila, 1, a); ws.cell(fila, 2, medida)
            for p in meses[k:]:
                j = col_mes[p]
                if a == "Suma":
                    ws.cell(fila, j, f"=SUM({L_(j)}{fila - len(comp)}:{L_(j)}{fila - 1})").font = B_
                else:
                    ra, ri = r_ag[a], r_ag["IPC"]
                    ws.cell(fila, j, f"=Agregados!$B{ra}*(Agregados!{L_(j)}{ra}-Agregados!{L_(j - k)}{ra})/Agregados!{L_(j - k)}{ri}")
            fila += 1
        fila += 1
    ws.freeze_panes = "C2"

    # ---------------------------------------------------------------------- Difusion
    ws = hoja("Difusion", ["Mes", "Tipo", "% subclases con alza", "Mes del año", "Mín. histórico", "Promedio histórico",
                           "Máx. histórico"], {1: 10, 2: 12})
    dh = base.difusion[["hist_min", "hist_prom", "hist_max"]].groupby(base.difusion.index.month).first()
    ws.cell(1, 10, "Mes"); ws.cell(1, 11, "Mín"); ws.cell(1, 12, "Prom"); ws.cell(1, 13, "Máx")
    for mth in range(1, 13):
        ws.cell(1 + mth, 10, mth)
        for j, col in enumerate(["hist_min", "hist_prom", "hist_max"], start=11):
            ws.cell(1 + mth, j, round(float(dh.loc[mth, col]), 2))
    for i, p in enumerate(meses[1:], start=2):
        j = L_(col_mes[p])
        ws.cell(i, 1, p.strftime("%Y-%m")); ws.cell(i, 2, "Proyección" if p > base.T else "Efectivo")
        ws.cell(i, 3, f"=COUNTIF(Var_mm_subclases!{j}$2:{j}${last_sc},\">0\")/COUNT(Var_mm_subclases!{j}$2:{j}${last_sc})*100")
        ws.cell(i, 4, p.month)
        for k, col in enumerate(["K", "L", "M"]):
            ws.cell(i, 5 + k, f"=INDEX(${col}$2:${col}$13,D{i})")
    ws.freeze_panes = "A2"

    # ---------------------------------------------------------------------- Resumen
    ws = hoja("Resumen", ["Agregado", "Dic-26 12m (libro, Modelo B)", "Dic-27 12m (libro)", "Promedio 2026 (libro)",
                          "Promedio 2027 (libro)", "Dic-26 12m Modelo A", "Dic-27 12m Modelo A",
                          "Dic-26 12m combinación A+B", "Dic-27 12m combinación A+B"], {1: 26})
    j26, j27 = col_mes[pd.Period("2026-12", "M")], col_mes[pd.Period("2027-12", "M")]
    rango_anio = lambda y: f"{L_(col_mes[pd.Period(f'{y}-01', 'M')])}{{r}}:{L_(col_mes[pd.Period(f'{y}-12', 'M')])}{{r}}"
    for i, a in enumerate(agr, start=2):
        r = r_ag[a]
        ws.cell(i, 1, a)
        ws.cell(i, 2, f"=Var_12m!{L_(j26)}{r}"); ws.cell(i, 3, f"=Var_12m!{L_(j27)}{r}")
        for k, y in [(4, 2026), (5, 2027)]:
            ws.cell(i, k, f"=(AVERAGE(Agregados!{rango_anio(y).format(r=r)})/AVERAGE(Agregados!{rango_anio(y - 1).format(r=r)})-1)*100")
        for k, key, y in [(6, "anual_A", 2026), (7, "anual_A", 2027), (8, "anual_AB", 2026), (9, "anual_AB", 2027)]:
            t = R[key]
            col = f"dic-{str(y)[2:]} (12m)"
            if a in t.index:
                ws.cell(i, k, round(float(t.loc[a, col]), 3))
        for k in range(2, 10):
            ws.cell(i, k).number_format = "0.00"
    ws.cell(len(agr) + 3, 1, f"Último IPC efectivo: {mes_txt(base.T)}. Las columnas B-E se recalculan si se editan Proy_mm o Nowcast_ENAP.")

    # ---------------------------------------------------------------------- Productos (índices desde subclases)
    PC0 = 5                                                  # primera columna de meses en Productos
    pcol_mes = {p: PC0 + i for i, p in enumerate(meses)}
    ws = hoja("Productos", ["Código producto", "Glosa", "Subclase", "Peso 2023"], {2: 40})
    for p in meses:
        c = ws.cell(1, pcol_mes[p], p.strftime("%Y-%m"))
        c.font, c.fill = B_, F_PROY if p > base.T else F_HEAD
    jT_ind, jT_prod = L_(col_mes[base.T]), L_(pcol_mes[base.T])
    for i, (cod, r_) in enumerate(info.iterrows(), start=2):
        ws.cell(i, 1, int(cod)); ws.cell(i, 2, r_.glosa); ws.cell(i, 3, str(r_.subclase)); ws.cell(i, 4, float(r_.peso))
        rs = r_sc[str(r_.subclase)]
        for p in meses:
            j = pcol_mes[p]
            if p <= base.T:
                v = base.prod_idx.loc[p, cod]
                ws.cell(i, j, None if pd.isna(v) else round(float(v), 4))
            else:
                ws.cell(i, j, f"=${jT_prod}{i}*Indices!{L_(col_mes[p])}${rs}/Indices!${jT_ind}${rs}").fill = F_PROY
    n_prod = len(info) + 1
    fila_ctrl = n_prod + 2
    ws.cell(fila_ctrl, 2, "IPC desde productos (control = Agregados!IPC)").font = B_
    for p in meses:
        j = L_(pcol_mes[p])
        ws.cell(fila_ctrl, pcol_mes[p], f"=SUMPRODUCT($D$2:$D${n_prod},{j}2:{j}{n_prod})/SUM($D$2:$D${n_prod})")
    ws.freeze_panes = "E2"

    # ---------------------------------------------------------------------- Resumen_Proyecciones
    ws = wb.create_sheet("Resumen_Proyecciones")
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 9
    meses_res = [p for p in meses if p >= pd.Period("2025-01", "M")]
    rc0 = 3
    rcol = {p: rc0 + i for i, p in enumerate(meses_res)}
    ws.cell(1, 1, f"Resumen de proyecciones: variación mensual y en 12 meses por clase (%). Efectivo hasta "
                  f"{mes_txt(base.T)}; desde {mes_txt(base.T + 1)}, proyección del libro (Proy_mm y Nowcast_ENAP).").font = B_
    fila = 3
    eee_mes = base.macro["eee_mm_mes"]
    for titulo, hoja_var, memo in [("Var. mensual", "Var_mm", "mm_A"), ("Var. 12 meses", "Var_12m", "a12_A")]:
        ws.cell(fila, 1, titulo).font = Font(bold=True, size=12)
        fila += 1
        for j, h in [(1, "Clase"), (2, "Peso")]:
            c = ws.cell(fila, j, h); c.font, c.fill = B_, F_HEAD
        for p in meses_res:
            c = ws.cell(fila, rcol[p], p.strftime("%Y-%m")); c.font, c.fill = B_, F_PROY if p > base.T else F_HEAD
            ws.column_dimensions[L_(rcol[p])].width = 8
        fila += 1
        for nombre, nivel in CLASES_RESUMEN:
            ra = r_ag[nombre]
            c = ws.cell(fila, 1, nombre)
            c.alignment = Alignment(indent=nivel)
            c.font = B_ if nivel <= 1 else Font()
            ws.cell(fila, 2, f"=Agregados!$B{ra}").number_format = "0.00"
            for p in meses_res:
                c = ws.cell(fila, rcol[p], f"={hoja_var}!{L_(col_mes[p])}{ra}")
                c.number_format = "0.00"
                if p > base.T:
                    c.fill = F_PROY
            fila += 1
        ws.cell(fila, 1, "Memo: IPC Modelo A (bloques)").font = Font(italic=True)
        for p in meses_res:
            v = R[memo]["IPC"].get(p)
            if v is not None and np.isfinite(v):
                ws.cell(fila, rcol[p], round(float(v), 3)).number_format = "0.00"
        fila += 1
        if hoja_var == "Var_mm":
            ws.cell(fila, 1, "Memo: EEE, inflación esperada del mes (mediana)").font = Font(italic=True)
            for p in meses_res:
                v = eee_mes.get(p)
                if v is not None and np.isfinite(v):
                    ws.cell(fila, rcol[p], float(v)).number_format = "0.00"
            fila += 1
        fila += 2
    # dic/dic y promedios anuales
    ws.cell(fila, 1, "Dic/dic y promedio anual").font = Font(bold=True, size=12)
    fila += 1
    cab = ["Clase", "Peso", "Dic-25", "Prom. 2025", "Dic-26", "Prom. 2026", "Dic-27", "Prom. 2027"]
    for j, h in enumerate(cab, start=1):
        c = ws.cell(fila, j, h); c.font, c.fill = B_, F_HEAD
    fila += 1
    rango = lambda y, r: (f"Agregados!{L_(col_mes[pd.Period(f'{y}-01', 'M')])}{r}:"
                          f"{L_(col_mes[pd.Period(f'{y}-12', 'M')])}{r}")
    for nombre, nivel in CLASES_RESUMEN:
        ra = r_ag[nombre]
        c = ws.cell(fila, 1, nombre); c.alignment = Alignment(indent=nivel); c.font = B_ if nivel <= 1 else Font()
        ws.cell(fila, 2, f"=Agregados!$B{ra}")
        for k, y in enumerate((2025, 2026, 2027)):
            ws.cell(fila, 3 + 2 * k, f"=Var_12m!{L_(col_mes[pd.Period(f'{y}-12', 'M')])}{ra}")
            ws.cell(fila, 4 + 2 * k, f"=(AVERAGE({rango(y, ra)})/AVERAGE({rango(y - 1, ra)})-1)*100")
        for j in range(2, 9):
            ws.cell(fila, j).number_format = "0.00"
        fila += 1
    ws.freeze_panes = "C5"

    # ---------------------------------------------------------------------- Perfiles_Analiticos
    ws = wb.create_sheet("Perfiles_Analiticos")
    anios = list(range(2010, H_FIN.year + 1))
    Y0 = 7                                                   # columna del primer año
    ws.column_dimensions["A"].width = 10
    for j in range(2, 7):
        ws.column_dimensions[L_(j)].width = 11
    ws.cell(1, 1, "Perfiles analíticos: variación mensual (%) por mes calendario. Hasta dic-2023, series empalmadas "
                  "oficiales (hoja Historia del Excel automatizado; energía y SAE: INE empalmado; alimentos y volátiles "
                  "excl. energía: agregación de bloques). Desde ene-2024, enlazada a Var_mm, que incluye la proyección (gris).").font = B_
    mm_largo = base.agr_largo.pct_change() * 100
    mm_hist = base.historia.pct_change() * 100
    fila = 3
    for nombre, _ in CLASES_RESUMEN:
        ra = r_ag[nombre]
        ws.cell(fila, 1, nombre).font = Font(bold=True, size=12)
        fila += 1
        cab = ["Mes", "Prom. 2010-2019", "Prom. 2013-2019", "Máx. 2010-2019", "Mín. 2010-2019", "Máx-Mín"] + anios
        for j, h in enumerate(cab, start=1):
            c = ws.cell(fila, j, h); c.font = B_
            c.fill = F_PROY if isinstance(h, int) and h > base.T.year else F_HEAD
        fila += 1
        g, p_, j_ = L_(Y0), L_(Y0 + 9), L_(Y0 + 3)
        for mth in range(1, 13):
            r = fila
            ws.cell(r, 1, MESES[mth - 1].capitalize())
            ws.cell(r, 2, f"=AVERAGE({g}{r}:{p_}{r})"); ws.cell(r, 3, f"=AVERAGE({j_}{r}:{p_}{r})")
            ws.cell(r, 4, f"=MAX({g}{r}:{p_}{r})"); ws.cell(r, 5, f"=MIN({g}{r}:{p_}{r})"); ws.cell(r, 6, f"=D{r}-E{r}")
            for k, y in enumerate(anios):
                per = pd.Period(f"{y}-{mth:02d}", "M")
                cell = ws.cell(r, Y0 + k)
                if per > H_FIN:
                    continue
                if per < pd.Period("2024-01", "M"):
                    serie = mm_hist[HISTORIA_COL[nombre]] if nombre in HISTORIA_COL else mm_largo[nombre]
                    v = serie.get(per)
                    cell.value = None if v is None or not np.isfinite(v) else round(float(v), 3)
                else:
                    cell.value = f"=Var_mm!{L_(col_mes[per])}{ra}"
                    if per > base.T:
                        cell.fill = F_PROY
                cell.number_format = "0.00"
            for j in range(2, 7):
                ws.cell(r, j).number_format = "0.00"
            fila += 1
        fila += 2
    ws.freeze_panes = "B3"

    # ---------------------------------------------------------------------- hojas de valores
    def valores(nombre, df, index=True, ancho=14):
        df = df.copy()
        ws = wb.create_sheet(nombre)
        cols = ([df.index.name or ""] if index else []) + [str(c) for c in df.columns]
        for j, h in enumerate(cols, start=1):
            c = ws.cell(1, j, h); c.font, c.fill = B_, F_HEAD
            ws.column_dimensions[L_(j)].width = ancho
        for i, (ix, row) in enumerate(df.iterrows(), start=2):
            vals = ([str(ix)] if index else []) + list(row.values)
            for j, v in enumerate(vals, start=1):
                if isinstance(v, (float, np.floating)):
                    v = None if not np.isfinite(v) else round(float(v), 4)
                elif isinstance(v, (np.integer,)):
                    v = int(v)
                elif not isinstance(v, (int, str)) and v is not None:
                    v = str(v)
                ws.cell(i, j, v)
        ws.freeze_panes = "B2"
        return ws

    ma = pd.concat({"m/m Modelo A": R["mm_A"].T, "12m Modelo A": R["a12_A"].T,
                    "12m Combinación A+B": R["a12_AB"].T}).loc[:, lambda d: d.columns >= base.T - 9]
    ma.columns = [str(c) for c in ma.columns]
    ma.index = [f"{a} | {b}" for a, b in ma.index]
    ma.index.name = "Medida | agregado"
    valores("Modelo_agregados", ma, ancho=12).column_dimensions["A"].width = 44
    valores("Incidencias_bloques_A", R["bloques_inc"].set_index("bloque"), ancho=16)
    ev = R["rmse"][R["rmse"].h.isin([1, 3, 6, 12])].set_index("agregado")
    valores("Evaluacion", ev, ancho=14)
    valores("Test_DM", R["dm"], index=False)
    valores("Nowcast_backtest", R["enap_bt"], index=False)
    mac = base.macro.copy()
    mac.index.name = "periodo"
    valores("Macro_BDE", mac.loc["2015-01":])
    fuentes = pd.DataFrame({"Fuente": [
        "INE, IPC base anual 2023=100 (productos, subclases, analíticos): www.ine.gob.cl, cuadros estadísticos.",
        "INE, serie referencial 2023 (ene-dic 2023).",
        "Historia 2018-2022 de productos y 1998-2017 de bloques: IPC_2023_Automatizado_v2.xlsx.",
        "Banco Central de Chile, BDE (API SieteRestWS): macro, expectativas (EEE) y compensaciones.",
        "ENAP, informes semanales de precios (vía prensa; URL por fecha en hoja Nowcast_ENAP).",
        "Precios ancla de combustibles: CNE/Bencina en Línea vía agregadores (preciocombustible.cl).",
        "Clasificación sin volátiles / volátiles: matriz_productos.csv (BCCh, canasta 2023).",
    ]})
    valores("Fuentes", fuentes, index=False, ancho=120)
    orden = ["LEEME", "Resumen", "Resumen_Proyecciones", "Perfiles_Analiticos", "Proy_mm", "Nowcast_ENAP", "Agregados",
             "Var_mm", "Var_12m", "Incidencias", "Difusion", "Indices", "Productos", "Var_mm_subclases", "Subclases",
             "Matriz", "Modelo_agregados", "Incidencias_bloques_A", "Evaluacion", "Test_DM", "Nowcast_backtest",
             "Macro_BDE", "Fuentes"]
    wb._sheets = [wb[n] for n in orden]
    wb.save(destino)
    _recalcular(destino)
    log(f"Excel: {destino}")
    return destino


def _recalcular(xlsx):
    """Recalcula con LibreOffice (si está instalado) para que el archivo traiga los valores calculados.
    Excel recalcula igual al abrir; esto sirve para visores y para leer el libro desde Python."""
    import shutil
    import tempfile
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return
    with tempfile.TemporaryDirectory() as tmp:
        r = subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", tmp, str(xlsx)],
                           capture_output=True, timeout=600)
        out = Path(tmp) / Path(xlsx).name
        if r.returncode == 0 and out.exists():
            shutil.move(str(out), str(xlsx))


# =====================================================================================================
# ==== MAIN (al final del archivo)
# =====================================================================================================

# =====================================================================================================
# 9. REPORTE HTML (formato de Apuntes_IPC_Chile.html)
# =====================================================================================================
def _f(x, d=1):
    if x is None or (isinstance(x, (float, np.floating)) and not np.isfinite(x)):
        return "–"
    return f"{x:,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _tabla_md(df):
    cols = list(df.columns)
    lin = ["| " + " | ".join(str(c) for c in cols) + " |", "|" + "|".join([":---"] + ["---:"] * (len(cols) - 1)) + "|"]
    lin += ["| " + " | ".join(str(v) for v in r) + " |" for r in df.values]
    return "\n".join(lin)


def _graficos(R):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    AZ, NA, AG, AM, VI, TI, T2, GR = "#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#6250d6", "#0b0b0b", "#52514e", "#e4e3df"
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": GR, "axes.labelcolor": T2,
                         "xtick.color": T2, "ytick.color": T2, "axes.spines.top": False, "axes.spines.right": False,
                         "axes.grid": True, "grid.color": GR, "grid.linewidth": 0.6, "axes.axisbelow": True,
                         "legend.frameon": False})
    base, T = R["base"], R["T"]
    figs = {}

    def png(fig):
        buf = io.BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight", dpi=150)
        plt.close(fig)
        return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()

    def xt(ax, idx):
        ticks = [p for p in idx if p.month in (1, 7)]
        ax.set_xticks([p.to_timestamp() for p in ticks], [mes_txt(p) for p in ticks], fontsize=8)

    # 1. RMSE por horizonte
    fig, axs = plt.subplots(1, 2, figsize=(9.6, 3.6), sharey=False)
    for ax, (tit, r) in zip(axs, [("Orígenes ene-21 a sep-26", R["rmse"]), ("Orígenes jul-23 a sep-26", R["rmse_pos23"])]):
        r = r[(r.agregado == "IPC") & (r.h <= 12)]
        for m, c, lab in [("A: Combinado", AZ, "A: bloques combinado"), ("A: M3 ECM tc", VI, "A: M3 tendencia común"),
                          ("B: subclases", AG, "B: subclases"), ("AR estacional agregado", NA, "AR sobre el IPC"),
                          ("Paseo aleatorio 12m", AM, "Paseo aleatorio")]:
            x = r[r.modelo == m].set_index("h").rmse_12m
            ax.plot(x.index, x.values, color=c, lw=2, marker="o", ms=3.5, label=lab)
        ax.set_title(tit, loc="left", fontsize=10, color=TI)
        ax.set_xlabel("Horizonte (meses)")
        ax.set_xticks(range(1, 13))
    axs[0].set_ylabel("RMSE inflación 12m (pp)")
    axs[0].legend(fontsize=8, loc="upper left")
    figs["rmse"] = png(fig)

    # 2. IPC 12m: historia y proyecciones
    fig, ax = plt.subplots(figsize=(9.2, 4.0))
    a12 = {k: R[f"a12_{k}"]["IPC"] for k in ["A", "B", "AB"]}
    h = base.agr_idx["IPC"].pct_change(12).loc["2023-01":T] * 100
    ax.axhspan(2, 4, color="#f1f0ec", zorder=0)
    ax.axhline(3, color=T2, lw=0.8, ls=(0, (4, 3)))
    ax.plot(h.index.to_timestamp(), h.values, color=TI, lw=2, label="IPC efectivo (INE)")
    for k, c, lab, ls in [("A", AZ, "Modelo A: agregados por bloques", "-"), ("B", AG, "Modelo B: subclases", "-"),
                          ("AB", VI, "Combinación A+B", (0, (3, 2)))]:
        x = a12[k].loc[T:]
        ax.plot(x.index.to_timestamp(), x.values, color=c, lw=2, ls=ls, label=lab)
    m = base.macro
    for col, fecha, lab in [("eee_dic26", "2026-12", "EEE dic-26"), ("eee_dic27", "2027-12", "EEE dic-27")]:
        v = m[col].dropna().iloc[-1]
        ax.plot(pd.Period(fecha, "M").to_timestamp(), v, "o", ms=8, color=NA, mec="white", mew=2, zorder=5)
        ax.annotate(f"{lab}: {_f(v)}%", (pd.Period(fecha, "M").to_timestamp(), v), xytext=(10, -26),
                    textcoords="offset points", fontsize=8, color=T2)
    ax.set_ylabel("Variación en 12 meses (%)")
    ax.set_xlim(pd.Timestamp("2023-01-01"), pd.Timestamp("2028-03-01"))
    ax.legend(fontsize=8.5, loc="upper right", ncol=2)
    ax.set_title("IPC: inflación en 12 meses, efectiva y proyectada (con nowcast de combustibles)", loc="left", fontsize=10.5, color=TI)
    figs["ipc"] = png(fig)

    # 3. agregados principales: A vs B
    fig, axs = plt.subplots(2, 3, figsize=(10.5, 5.6), sharex=True)
    for ax, ag in zip(axs.flat, ["IPC sin volátiles", "IPC volátiles", "Alimentos", "Volátiles excl. energía",
                                 "Energía volátil", "Servicios SV"]):
        hh = base.agr_idx[ag].pct_change(12).loc["2024-01":T] * 100
        ax.plot(hh.index.to_timestamp(), hh.values, color=TI, lw=1.8)
        for k, c in [("A", AZ), ("B", AG)]:
            x = R[f"a12_{k}"][ag].loc[T:]
            ax.plot(x.index.to_timestamp(), x.values, color=c, lw=1.8)
        ax.axhline(3, color=T2, lw=0.7, ls=(0, (4, 3)))
        ax.set_title(ag, loc="left", fontsize=9.5, color=TI)
        ax.xaxis.set_major_locator(matplotlib.dates.YearLocator())
        ax.xaxis.set_major_formatter(matplotlib.dates.DateFormatter("%Y"))
        ax.tick_params(labelsize=8)
    hs = [matplotlib.lines.Line2D([], [], color=c, lw=2) for c in (TI, AZ, AG)]
    fig.supylabel("Variación en 12 meses (%)", fontsize=9, color=T2)
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    fig.legend(hs, ["Efectivo", "Modelo A (bloques)", "Modelo B (subclases)"], loc="upper center", ncol=3, fontsize=9)
    figs["agregados"] = png(fig)

    # 4. incidencias por bloque (modelo A)
    fig, ax = plt.subplots(figsize=(9.0, 4.2))
    ii = R["bloques_inc"].iloc[::-1]
    y = np.arange(len(ii))
    ax.barh(y + 0.19, ii.incid_dic26, height=0.36, color=AZ, label="Dic-26")
    ax.barh(y - 0.19, ii.incid_dic27, height=0.36, color=NA, label="Dic-27")
    ax.set_yticks(y, ii.bloque, fontsize=8.5)
    ax.axvline(0, color=T2, lw=0.8)
    ax.set_xlabel("Incidencia en la inflación de 12 meses (pp)")
    ax.legend(fontsize=8.5, loc="lower right")
    ax.set_title("Modelo A: aporte de cada bloque a la inflación de diciembre", loc="left", fontsize=10.5, color=TI)
    figs["incid"] = png(fig)

    # 5. ENAP: precio diario
    fig, ax = plt.subplots(figsize=(9.0, 3.7))
    hoy = pd.Timestamp.today().normalize()
    for prod, c, lab in [("gasolina", AZ, "Gasolina 93"), ("diesel", NA, "Diésel")]:
        sb = R["enap_diario"]["Base"][prod]
        so = R["enap_diario"]["Solo 12-nov"][prod]
        ax.plot(sb.loc[:hoy].index, sb.loc[:hoy], color=c, lw=2, label=lab)
        ax.plot(sb.loc[hoy:].index, sb.loc[hoy:], color=c, lw=2, ls=(0, (3, 2)))
        ax.plot(so.loc[hoy:].index, so.loc[hoy:], color=c, lw=1.2, ls=(0, (1, 2)))
    ax.axvline(hoy, color=T2, lw=0.8)
    mt = pd.date_range("2026-05-01", "2027-01-01", freq="MS")
    ax.set_xticks(mt, [MESES[d.month - 1] for d in mt])
    ax.set_ylabel("Pesos por litro")
    ax.legend(fontsize=8.5, loc="upper left")
    ax.set_title("Precio a público reconstruido con ENAP (discontinua: escenario base; punteada: solo 12-nov)",
                 loc="left", fontsize=10, color=TI)
    figs["enap"] = png(fig)

    # 6. difusión
    fig, ax = plt.subplots(figsize=(9.0, 3.4))
    dif = base.difusion.loc["2024-01":]
    x = dif.index.to_timestamp()
    ax.fill_between(x, dif.hist_min, dif.hist_max, color=AZ, alpha=0.15, lw=0, label="Rango histórico 2001-2023 (mín-máx)")
    ax.plot(x, dif.hist_prom, color=AZ, lw=1, ls=(0, (3, 2)), label="Promedio histórico")
    ax.plot(x, dif.difusion_productos, color=TI, lw=2, label="% de productos con alza")
    ax.set_ylabel("%")
    ax.xaxis.set_major_formatter(matplotlib.dates.DateFormatter("%b-%y"))
    ax.legend(fontsize=8, loc="lower left", ncol=3)
    ax.set_title("Difusión del IPC", loc="left", fontsize=10.5, color=TI)
    figs["difusion"] = png(fig)
    return figs


def construir_reporte(R, destino=None):
    base, A, B, T = R["base"], R["A"], R["B"], R["T"]
    figs = _graficos(R)
    destino = destino or OUTPUTS / "Proyeccion_IPC_Chile_modelos.html"
    m = base.macro
    ultimo = lambda c: (m[c].dropna().iloc[-1], m[c].dropna().index[-1])
    eee26, eee26_m = ultimo("eee_dic26")
    eee27, eee27_m = ultimo("eee_dic27")
    swap1 = base.diario["bei_swap_1a"].dropna().iloc[-1]
    swap1_f = pd.Timestamp(base.diario["bei_swap_1a"].dropna().index[-1]).strftime("%d-%m-%Y")
    oct_ef = R["oct_efectivo"]
    an = {k: R[f"anual_{k}"] for k in ["A", "B", "AB", "A_sin_nc", "B_sin_nc"]}
    g = lambda k, ag, col: an[k].loc[ag, col] if ag in an[k].index else np.nan

    # ---- tablas de evaluación
    def tab_rmse(r):
        filas = []
        for mod, lab in [("A: Combinado", "**A: bloques, combinado**"), ("A: M3 ECM tc", "A: M3 tendencia común"),
                         ("A: M1 AR est", "A: M1 AR estacional"), ("B: subclases", "B: subclases"),
                         ("Combinación A+B", "Combinación A+B"), ("AR estacional agregado", "AR estacional sobre el IPC"),
                         ("Paseo aleatorio 12m", "Paseo aleatorio (12m constante)"), ("EEE (mediana)", "EEE (mediana)")]:
            x = r[(r.agregado == "IPC") & (r.modelo == mod)].set_index("h")
            fila = {"Modelo": lab}
            for h, c, nom in [(1, "rmse_mm", "m/m h=1"), (3, "rmse_12m", "12m h=3"), (6, "rmse_12m", "12m h=6"),
                              (12, "rmse_12m", "12m h=12")]:
                fila[nom] = _f(x[c].get(h, np.nan), 2) if h in x.index else "–"
            filas.append(fila)
        return _tabla_md(pd.DataFrame(filas))

    def tab_rmse_ag(r):
        filas = []
        for ag in ["IPC sin volátiles", "IPC volátiles"]:
            for mod in ["A: Combinado", "A: M3 ECM tc", "B: subclases", "Combinación A+B"]:
                x = r[(r.agregado == ag) & (r.modelo == mod)].set_index("h")
                filas.append({"Agregado": ag, "Modelo": mod, "12m h=3": _f(x.rmse_12m[3], 2), "12m h=6": _f(x.rmse_12m[6], 2),
                              "12m h=12": _f(x.rmse_12m[12], 2)})
        return _tabla_md(pd.DataFrame(filas))

    dm = R["dm"].copy()
    t_dm = _tabla_md(pd.DataFrame({"Muestra": dm.muestra, "Modelo": dm.modelo, "vs.": dm.rival, "h": dm.h,
                                   "Variable": dm.variable, "DM": dm.DM.map(lambda v: _f(v, 2)),
                                   "p-valor": dm.p_valor.map(lambda v: _f(v, 3))}))
    rp = R["rmse_pos23"]
    rr = R["rmse"]
    val = lambda r, mod, h, c: r[(r.agregado == "IPC") & (r.modelo == mod) & (r.h == h)][c].iloc[0]

    # ---- tablas de proyección
    def tab_anual(k, ags):
        filas = []
        for ag in ags:
            if ag not in an[k].index:
                continue
            ef = base.agr_idx[ag].pct_change(12).loc[T] * 100
            filas.append({"Agregado": ag, f"{mes_txt(T)} efectivo": _f(ef), "Dic-26": _f(g(k, ag, "dic-26 (12m)")),
                          "Promedio 2026": _f(g(k, ag, "promedio 2026")), "Dic-27": _f(g(k, ag, "dic-27 (12m)")),
                          "Promedio 2027": _f(g(k, ag, "promedio 2027"))})
        return _tabla_md(pd.DataFrame(filas))

    def tab_mm(k, ags):
        mm = R[f"mm_{k}"].loc[T + 1:]
        filas = []
        for p in mm.index:
            fila = {"Mes": mes_txt(p)}
            for ag in ags:
                fila[ag] = _f(mm.loc[p, ag], 2)
            filas.append(fila)
        return _tabla_md(pd.DataFrame(filas))

    sc = R["sc_inc"].copy()
    sc["bloque"] = sc.bloque.map(lambda b: BLOQUES[b][0])
    top = sc.sort_values("incid_dic26", ascending=False).head(12)
    t_top = _tabla_md(pd.DataFrame({"Subclase": top.glosa.str.capitalize(), "Bloque": top.bloque, "Peso": top.peso.map(lambda v: _f(v, 2)),
                                    "Var. 12m dic-26": top.var_12m_dic26.map(_f), "Incid. dic-26 (pp)": top.incid_dic26.map(lambda v: _f(v, 2)),
                                    "Var. 12m dic-27": top.var_12m_dic27.map(_f), "Incid. dic-27 (pp)": top.incid_dic27.map(lambda v: _f(v, 2))}))
    bot = sc.sort_values("incid_dic26").head(6)
    t_bot = _tabla_md(pd.DataFrame({"Subclase": bot.glosa.str.capitalize(), "Bloque": bot.bloque,
                                    "Var. 12m dic-26": bot.var_12m_dic26.map(_f), "Incid. dic-26 (pp)": bot.incid_dic26.map(lambda v: _f(v, 2))}))
    bi = R["bloques_inc"]
    t_binc = _tabla_md(pd.DataFrame({"Bloque": bi.bloque, "Peso 2023": bi.peso_2023.map(lambda v: _f(v, 2)),
                                     "Var. 12m dic-26": bi.var_12m_dic26.map(_f), "Incid. dic-26": bi.incid_dic26.map(lambda v: _f(v, 2)),
                                     "Var. 12m dic-27": bi.var_12m_dic27.map(_f), "Incid. dic-27": bi.incid_dic27.map(lambda v: _f(v, 2))}))
    esc = []
    for mod in A.MODELOS:
        ag = R["agr_A_mod"][mod]
        y = ag["IPC"]
        esc.append({"Modelo A": mod.replace("_", " "), "Dic-26": _f((y["2026-12"] / y["2025-12"] - 1) * 100),
                    "Dic-27": _f((y["2027-12"] / y["2026-12"] - 1) * 100),
                    "IPC SV dic-27": _f((ag["IPC sin volátiles"]["2027-12"] / ag["IPC sin volátiles"]["2026-12"] - 1) * 100)})
    t_esc = _tabla_md(pd.DataFrame(esc))
    comp = []
    for ag in ["IPC", "IPC sin volátiles"]:
        comp.append({"Indicador": f"{ag} dic-26 (12m)", "Modelo A": _f(g("A", ag, "dic-26 (12m)")), "Modelo B": _f(g("B", ag, "dic-26 (12m)")),
                     "A+B": _f(g("AB", ag, "dic-26 (12m)")),
                     "EEE": _f(eee26 if ag == "IPC" else m["eee_sv_dic26"].dropna().iloc[-1])})
        comp.append({"Indicador": f"{ag} dic-27 (12m)", "Modelo A": _f(g("A", ag, "dic-27 (12m)")), "Modelo B": _f(g("B", ag, "dic-27 (12m)")),
                     "A+B": _f(g("AB", ag, "dic-27 (12m)")),
                     "EEE": _f(eee27 if ag == "IPC" else m["eee_sv_dic27"].dropna().iloc[-1])})
    mmA, mmB = R["mm_A"]["IPC"], R["mm_B"]["IPC"]
    comp.append({"Indicador": f"IPC {mes_txt(T + 1)} (m/m)", "Modelo A": _f(mmA[T + 1], 2), "Modelo B": _f(mmB[T + 1], 2),
                 "A+B": _f(R["mm_AB"]["IPC"][T + 1], 2), "EEE": _f(m["eee_mm_mes"].dropna().iloc[-1], 2)})
    t_comp = _tabla_md(pd.DataFrame(comp))
    # nowcast
    bt = R["enap_bt"]
    t_bt = _tabla_md(pd.DataFrame({"Producto": bt.producto.map({"gasolina": "Gasolina", "diesel": "Diésel"}),
                                   "Mes": bt.periodo.map(mes_txt), "Nowcast ENAP": bt.nowcast_enap.map(lambda v: _f(v, 2)),
                                   "INE efectivo": bt.ine_efectivo.map(lambda v: _f(v, 2)), "Error (pp)": bt.error_pp.map(lambda v: _f(v, 2))}))
    nc_rows = []
    for e, (mm_, var_) in R["enap_mensual"].items():
        for p in pd.period_range(T + 1, "2026-12", freq="M"):
            nc_rows.append({"Escenario": e, "Mes": mes_txt(p), "Gasolina (pesos/l)": _f(mm_.loc[p, "gasolina"], 0),
                            "Gasolina m/m": _f(var_.loc[p, "gasolina"], 2), "Diésel (pesos/l)": _f(mm_.loc[p, "diesel"], 0),
                            "Diésel m/m": _f(var_.loc[p, "diesel"], 2)})
    t_nc = _tabla_md(pd.DataFrame(nc_rows))
    imp = []
    for k, lab in [("A", "Modelo A"), ("B", "Modelo B")]:
        for p in pd.period_range(T + 1, "2026-12", freq="M"):
            imp.append({"Modelo": lab, "Mes": mes_txt(p), "IPC m/m sin nowcast": _f(R[f"mm_{k}_sin_nc"]["IPC"][p], 2),
                        "IPC m/m con nowcast": _f(R[f"mm_{k}"]["IPC"][p], 2)})
        imp.append({"Modelo": lab, "Mes": "**Dic-26 (12m)**", "IPC m/m sin nowcast": _f(g(f"{k}_sin_nc", "IPC", "dic-26 (12m)")),
                    "IPC m/m con nowcast": _f(g(k, "IPC", "dic-26 (12m)"))})
    t_imp = _tabla_md(pd.DataFrame(imp))
    pesos_A = A.tabla_pesos.copy()
    pesos_A["bloque"] = pesos_A.bloque.map(lambda b: BLOQUES[b][0])
    pv = pesos_A.pivot(index="bloque", columns="modelo", values="peso").reindex([BLOQUES[b][0] for b in HOJAS])
    pv.columns = [{"M0_estacional": "M0 estacional", "M1_AR_est": "M1 AR", "M2_ARDL_macro": "M2 macro",
                   "M3_ECM_tc": "M3 tendencia común"}[c] for c in pv.columns]
    t_pesosA = _tabla_md(pd.DataFrame({"Bloque": pv.index, **{c: pv[c].map(lambda v: "–" if not np.isfinite(v) else _f(100 * v, 0) + "%")
                                                                for c in pv.columns}}))
    grupos = A.grupos
    t_tc = _tabla_md(pd.DataFrame({"Bloque": [BLOQUES[b][0] for b in HOJAS],
                                   "Comparte tendencia con (ADF por pares, 5%)": [", ".join(BLOQUES[x][0] for x in grupos[b]) or "—" for b in HOJAS]}))
    v = base.validacion
    t_val = _tabla_md(pd.DataFrame({"Serie": v.serie, "Dif. máx. var. m/m (pp)": v.dif_max_var_mm_pp.map(lambda x: _f(x, 3)),
                                    "Dif. máx. nivel (%)": v.dif_max_nivel_pct.map(lambda x: _f(x, 3)) if "dif_max_nivel_pct" in v
                                    else v["dif_max_nivel_%"].map(lambda x: _f(x, 3))}))
    ecm_b = int(B.ecm.sum())
    principales = ["IPC", "IPC sin volátiles", "Bienes SV", "Servicios SV", "IPC volátiles", "Alimentos volátiles",
                   "Energía volátil", "Resto volátiles", "Volátiles excl. energía", "Alimentos", "Energía (INE)", "IPC SAE (INE)"]
    a_mm_cols = ["IPC", "IPC sin volátiles", "IPC volátiles", "Energía volátil", "Alimentos"]
    hoja_excel = [
        ("LEEME", "Instrucciones, flujo y convenciones de color.", "Texto"),
        ("Resumen", "Dic/dic y promedio anual 2026-2027 de cada agregado (libro editable) y del Modelo A.", "Fórmulas"),
        ("Resumen_Proyecciones", "Var. mensual y en 12 meses 2025-2027 de las 20 clases del BCCh (formato del Excel automatizado), memo con Modelo A y EEE, y dic/dic y promedios anuales.", "Fórmulas"),
        ("Perfiles_Analiticos", "Perfil m/m por mes calendario 2010-2027 de cada clase, con promedios 2010-19 y 2013-19, máx., mín. y rango.", "Valores + fórmulas"),
        ("Proy_mm", "**Supuestos**: variación mensual proyectada por subclase (Modelo B). Gasolina y diésel enlazados al nowcast.", "Entradas (amarillo)"),
        ("Nowcast_ENAP", "Informes de ENAP, escenario de próximos ajustes, ancla CNE, precio diario (SUMIFS) y promedio mensual (AVERAGEIFS).", "Entradas + fórmulas"),
        ("Agregados", "Índices de 12 agregados (IPC, SV, volátiles, alimentos, energía, SAE…) con SUMPRODUCT sobre subclases.", "Fórmulas"),
        ("Var_mm / Var_12m", "Variaciones mensuales y en 12 meses de los agregados.", "Fórmulas"),
        ("Incidencias", "Aporte m/m y 12m de bienes SV, servicios SV, energía, alimentos volátiles y resto volátiles (suman el IPC).", "Fórmulas"),
        ("Difusion", "% de subclases con alza cada mes contra el mínimo, promedio y máximo histórico del mes.", "Fórmulas"),
        ("Indices", "172 subclases, 2018-01 a 2027-12: historia (valores) y proyección = índice(t−1)×(1+Proy_mm).", "Valores + fórmulas"),
        ("Var_mm_subclases", "Variación mensual de cada subclase (insumo de la difusión).", "Fórmulas"),
        ("Subclases", "Peso (SUMIF) y participación en cada agregado (SUMIFS) desde la matriz.", "Fórmulas"),
        ("Matriz", "283 productos: subclase, división, peso, bloque y pertenencia 0/1 a cada agregado.", "Valores (maestro)"),
        ("Modelo_agregados", "Proyección del Modelo A (bloques) y de la combinación A+B.", "Valores"),
        ("Incidencias_bloques_A", "Incidencias por bloque del Modelo A.", "Valores"),
        ("Evaluacion / Test_DM", "RMSE fuera de muestra por horizonte y tests de Diebold-Mariano.", "Valores"),
        ("Nowcast_backtest", "Nowcast ENAP vs INE, jul-oct 2026.", "Valores"),
        ("Macro_BDE", "Variables macro y expectativas de la BDE (2015 en adelante).", "Valores"),
        ("Productos", "283 productos: historia (valores) y proyección = índice del último mes × variación de su subclase (Indices); fila de control con el IPC.", "Valores + fórmulas"),
        ("Fuentes", "Origen de cada dato.", "Texto"),
    ]
    t_excel = _tabla_md(pd.DataFrame(hoja_excel, columns=["Hoja", "Contenido", "Tipo"]))

    md = f"""
# IPC Chile: modelos de proyección, proyecciones 2026-2027 y herramientas

> **Datos**: INE, IPC base anual 2023=100 por producto y subclase, descargado de www.ine.gob.cl, con el último dato de
> **{mes_txt(T)}** (IPC {_f(base.agr_idx['IPC'].pct_change().loc[T] * 100, 2)}% m/m y {_f(base.agr_idx['IPC'].pct_change(12).loc[T] * 100)}% a/a).
> BDE del Banco Central (API) para macro y expectativas, e informes semanales de ENAP para combustibles. La historia
> 2018-2022 de productos y 1998-2017 de bloques viene del Excel `IPC_2023_Automatizado_v2.xlsx`.
>
> **Validación**: los agregados reconstruidos desde los productos replican los publicados por el INE con un error máximo
> de {_f(v.dif_max_var_mm_pp.max(), 3)} pp en la variación mensual.

------------------------------------------------------------------------

## Resumen

| Pregunta | Respuesta corta |
|:---|:---|
| ¿Qué modelo conviene para Chile? | Un **sistema en capas**. Para 0-2 meses, nowcast con alta frecuencia (ENAP/CNE, supermercados) contrastado con la EEE. Para 1-12 meses, el **modelo de bloques à la Carlomagno** (partición SV/volátiles, ecuaciones únicas con tendencias comunes y combinación). Para más de 12 meses, anclaje a expectativas y meta. |
| ¿Qué evidencia lo respalda? | En 2021-2026, el modelo de bloques le gana al AR sobre el IPC y al paseo aleatorio (Diebold-Mariano, p < 0,05 a 1 mes). Su versión con tendencias comunes (M3) es la mejor en el período de shocks. Desde jul-23, el combinado tiene el menor error a 12 meses ({_f(val(rp, 'A: Combinado', 12, 'rmse_12m'), 2)} pp contra {_f(val(rp, 'Paseo aleatorio 12m', 12, 'rmse_12m'), 2)} pp del paseo aleatorio). |
| ¿Y el modelo por subclases? | Es peor a 1-3 meses: 172 ecuaciones con ~8 años de datos acumulan ruido. Pero es la herramienta para el detalle: incidencias, difusión y juicio experto por producto. Por eso es la base del Excel editable. |
| IPC dic-26 | Modelo A **{_f(g('A', 'IPC', 'dic-26 (12m)'))}%**, Modelo B {_f(g('B', 'IPC', 'dic-26 (12m)'))}%, combinación {_f(g('AB', 'IPC', 'dic-26 (12m)'))}%. EEE ({mes_txt(eee26_m)}): {_f(eee26)}%. |
| IPC dic-27 | Modelo A **{_f(g('A', 'IPC', 'dic-27 (12m)'))}%**, Modelo B {_f(g('B', 'IPC', 'dic-27 (12m)'))}%, combinación {_f(g('AB', 'IPC', 'dic-27 (12m)'))}%. EEE ({mes_txt(eee27_m)}): {_f(eee27)}%. Swap a 1 año: {_f(swap1, 2)}%. |
| IPC sin volátiles | Dic-26: A {_f(g('A', 'IPC sin volátiles', 'dic-26 (12m)'))}% / B {_f(g('B', 'IPC sin volátiles', 'dic-26 (12m)'))}%. Dic-27: A {_f(g('A', 'IPC sin volátiles', 'dic-27 (12m)'))}% / B {_f(g('B', 'IPC sin volátiles', 'dic-27 (12m)'))}%. |
| Nowcast | Con la información de septiembre, el nowcast de ENAP acertó la gasolina de octubre (4,4% vs 4,5% del INE). El IPC de octubre fue {_f(oct_ef, 2)}%; el Modelo B había proyectado {_f(R['oct_modelo_B'], 2)}% y el A {_f(R['oct_modelo_A'], 2)}%. Para nov-dic se incorpora el alza de 40 pesos del 12-nov anunciada por el Gobierno. |
| Herramientas | `proyeccion_ipc.py`: un solo script que descarga, modela, evalúa y genera el Excel y este reporte (~30 s). `IPC_Chile_Proyecciones.xlsx`: libro con fórmulas donde se editan los supuestos por subclase y por informe de ENAP, y todo se recalcula. |

------------------------------------------------------------------------

## 1. Análisis de los modelos de proyección del IPC

### 1.1 Familias de modelos

| Familia | Idea | Fortalezas | Debilidades | Evidencia / uso en Chile |
|:---|:---|:---|:---|:---|
| Paseo aleatorio o meta | La inflación futura es la actual, o el 3% | Difícil de batir a 12-24 meses; base de comparación | No ve estacionalidad, shocks ni el calendario de precios | Benchmark estándar en el BCCh y en la literatura |
| ARIMA/SARIMA agregado | Dinámica propia del IPC | Simple; capta estacionalidad | Mezcla componentes con dinámicas distintas | En nuestra evaluación pierde contra los bloques a 1-3 meses |
| Curva de Phillips | Holguras, expectativas y tipo de cambio | Interpretación económica; escenarios | Relación inestable; holguras inobservables | Pincheira y Rubio (DT BCCh 559) encuentran poca capacidad predictiva en versiones simples |
| VAR / BVAR | Sistema con dólar, petróleo, TPM y actividad | Interacciones; escenarios condicionados | Muchos parámetros; sensible a quiebres | Herramienta habitual en bancos centrales para el mediano plazo |
| DSGE / semiestructurales | Modelos de equilibrio con política monetaria | Coherencia y narrativa para el IPoM | Débiles en el corto plazo mensual | Uso institucional en el BCCh |
| Desagregados (bottom-up) | Proyectar componentes y agregar con pesos | Cada componente tiene sus impulsores; calendario | Errores que se acumulan si los componentes son ruidosos | Cobb (DT BCCh 545): la desagregación ayuda cuando la inflación no es baja y estable |
| **Carlomagno-Espasa** | Desagregados con **tendencias comunes por pares** como restricciones en ecuaciones únicas | Disciplina el bottom-up: cada componente "vuelve" a su grupo; escala a cientos de series | Tests de cointegración con poca potencia en muestras cortas | Probado en el CPI de EE.UU.; aquí, sobre la partición de Carlomagno, Fornero y Sansone (DT BCCh 913) |
| Nowcasting | Precios de alta frecuencia: ENAP/CNE, scraping, ODEPA | Domina en el mes en curso | Solo cubre parte de la canasta; hay que mantenerlo | La EEE le gana a todos los modelos en h=1 porque incorpora esta información |
| Encuestas y mercado | EEE, EOF, compensaciones | Rápidas; información de muchos agentes | Sesgos de anclaje; primas | Referencia obligada para comparar |
| Combinación | Promediar modelos por desempeño | Robustez frente a cambios de régimen | Puede diluir al mejor modelo | Se usa aquí dentro de cada bloque |

### 1.2 El enfoque de Carlomagno, como se implementa aquí

**Partición (Carlomagno, Fornero y Sansone, 2021).** La canasta se divide en 11 bloques: alimentos, bienes y servicios sin
volátiles (con administrados e indexados por separado), y energía, alimentos y resto de volátiles. Son los bloques que usa el
BCCh tras minimizar su función de pérdida.

**Tendencias comunes por pares (Carlomagno y Espasa).** Para cada par de bloques se testea si el precio relativo es
estacionario: es un ADF con tendencia, 55 pares. Si un bloque comparte tendencia con otros, su ecuación incluye la corrección
de error frente a ese subconjunto.

{t_tc}

**Cuatro ecuaciones por bloque**, todas ancladas a la expectativa de inflación a 2 años de la EEE más la deriva relativa del
bloque:

- **M0**: estacionalidad.
- **M1**: AR(2).
- **M2**: impulsores macro (dólar, petróleo en pesos, FAO, salarios, brecha e indexación).
- **M3**: M2 más la corrección de error de tendencia común.

Se combinan por el inverso del error fuera de muestra. La electricidad sigue una regla de indexación semestral.

{t_pesosA}

**Modelo B (subclases).** Es la misma lógica aplicada a las 172 subclases:

- una ecuación por subclase anclada a expectativas, con deriva relativa contraída a la mitad, estacionalidad centrada y AR(1);
- corrección de error frente a su bloque cuando el ADF lo acepta y el coeficiente corrige. Eso ocurre en **{ecm_b}** subclases.

Los productos heredan la variación de su subclase. Los agregados (SAE, alimentos y energía del INE incluidos) se arman con la
matriz de productos.

### 1.3 Carrera de modelos (evaluación recursiva fuera de muestra)

Cada mes entre ene-21 y sep-26 se reestiman todos los modelos con la información disponible y se proyectan 1-12 meses. Las
variables exógenas quedan planas, sin información futura. El error se mide contra el IPC efectivo.

![]({figs['rmse']})

**IPC, orígenes ene-21 a sep-26** (RMSE en pp):

{tab_rmse(rr)}

**IPC, orígenes jul-23 a sep-26** (régimen post-shock):

{tab_rmse(rp)}

**Sin volátiles y volátiles** (12m, orígenes ene-21 a sep-26):

{tab_rmse_ag(rr)}

**Diebold-Mariano** (un valor negativo indica que el primer modelo es mejor):

<details><summary>Ver todos los tests</summary>

{t_dm}

</details>

**Lectura.**

1. **Los bloques à la Carlomagno le ganan a los modelos agregados.** Contra el AR estacional sobre el IPC el resultado es
   significativo a 1 y 3 meses en ambas muestras; contra el paseo aleatorio, a 1 mes en ambas y a 3 y 12 meses desde jul-23.
2. **La restricción de tendencia común (M3) paga en los shocks.** En 2021-2026 es la mejor ecuación a todos los horizontes
   (DM vs. combinado, p = {_f(dm[(dm.muestra == '2021-2026') & (dm.modelo == 'A: M3 ECM tc') & (dm.h == 1)].p_valor.iloc[0], 3)} a 1 mes).
   Desde jul-23 la combinación iguala a M3 a corto plazo y la supera a 12 meses: diversificar ayuda en tiempos normales.
3. **Más desagregación no es mejor.** El modelo de 172 subclases y la combinación A+B pierden contra el de bloques
   (significativo a 1-3 meses en 2021-26). Es el *trade-off* de Cobb y de Carlomagno-Espasa: desagregar ayuda si se imponen
   restricciones comunes; si no, se suma ruido.
4. **La EEE gana a 1 mes.** Incorpora información de alta frecuencia y el nowcast debe cerrar esa brecha (§2.4).

### 1.4 Propuesta: el mejor modelo para proyectar el IPC en Chile

**Sistema de tres capas, con el modelo de bloques de Carlomagno como núcleo:**

| Horizonte | Herramienta | Por qué |
|:---|:---|:---|
| Mes en curso y siguiente | **Nowcast**: ENAP semanal y CNE (combustibles), scraping de supermercados (alimentos), ODEPA (frutas y verduras), calendario de regulados (electricidad, educación, transporte). Se compara contra EEE/EOF. | La información de precios ya ocurrida domina (backtest de combustibles: error de 0,1-1,3 pp vs {_f(R['rmse_comb_h1'], 1)} pp de RMSE del modelo de bloques a un mes). |
| 1 a 12 meses | **Modelo A**: 11 bloques SV/volátiles, ecuaciones ancladas con impulsores macro y tendencias comunes, combinadas por desempeño. En episodios de shock conviene cargar más peso a M3. | Es el mejor fuera de muestra y significativamente mejor que los agregados. |
| Más de 12 meses | Anclaje a la expectativa a 2 años y meta; contraste con el IPoM y con un modelo semiestructural o BVAR para escenarios. | Los errores de todos los modelos crecen con el horizonte y sus proyecciones convergen al ancla de expectativas; ahí pesan más el juicio y los escenarios. |
| Detalle y juicio | **Modelo B** (172 subclases) en el Excel: incidencias, difusión y ajustes por producto (decretos, impuestos, separatas del INE). | Hace transparente qué hay detrás del número; no es el central. |

Recomendaciones prácticas:

- Reevaluar los pesos cada mes con los errores nuevos.
- Reportar el rango entre M1 (escenario bajo) y M3 (escenario alto).
- Aplicar ajustes discretos (tarifas, IVA, subsidios) en el mes en que rigen.

------------------------------------------------------------------------

## 2. Proyecciones de inflación 2026-2027

Ambos modelos incorporan el IPC de {mes_txt(T)} y el nowcast de combustibles para nov-dic (escenario base de ENAP, §2.4).

![]({figs['ipc']})

### 2.1 Modelo B: por subclases

{tab_anual('B', principales)}

**Subclases que más aportan a la inflación de diciembre de 2026:**

{t_top}

**Las que más restan:**

{t_bot}

<details><summary>Variación mensual proyectada (Modelo B)</summary>

{tab_mm('B', a_mm_cols + ['IPC SAE (INE)'])}

</details>

### 2.2 Modelo A: agregados principales (bloques)

{tab_anual('A', [a for a in principales if a in an['A'].index])}

El Modelo A no proyecta el SAE ni la energía del INE, porque sus bloques no coinciden con esas definiciones. Esos agregados
se toman del Modelo B.

![]({figs['agregados']})

![]({figs['incid']})

{t_binc}

**Escenarios por ecuación del Modelo A:**

{t_esc}

<details><summary>Variación mensual proyectada (Modelo A)</summary>

{tab_mm('A', a_mm_cols)}

</details>

### 2.3 Comparación entre modelos y con el mercado

{t_comp}

**Lectura.**

- Ambos modelos ven el peak de inflación entre dic-26 y feb-27, con energía en torno a 17-18% a/a, y la convergencia desde
  abril de 2027, cuando salen de la ventana los meses del shock petrolero de mar-abr 2026.
- El **Modelo A** termina 2027 más arriba ({_f(g('A', 'IPC', 'dic-27 (12m)'))}%): traspasa más el dólar y el petróleo, y la
  indexación mantiene a los servicios SV sobre 4%.
- El **Modelo B** converge más rápido a la meta ({_f(g('B', 'IPC', 'dic-27 (12m)'))}%).
- **Contra el mercado.** La EEE para dic-27 ({_f(eee27)}%) está más cerca del Modelo {'B' if abs(eee27 - g('B', 'IPC', 'dic-27 (12m)')) < abs(eee27 - g('A', 'IPC', 'dic-27 (12m)')) else 'A'}.
  La compensación swap a 1 año ({_f(swap1, 2)}% al {swap1_f}) es comparable con la inflación a 12 meses de
  {mes_txt(T + 13)}: A {_f(R['a12_A']['IPC'][T + 13])}%, B {_f(R['a12_B']['IPC'][T + 13])}%. Bajó desde ~4,2% a fines de
  octubre, tras un IPC de octubre menor al esperado.

### 2.4 Nowcast de corto plazo

**Validación con el IPC de octubre** (pronóstico hecho con la información a septiembre):

| | Proyección para oct-26 (m/m) |
|:---|---:|
| Modelo A (bloques) | {_f(R['oct_modelo_A'], 2)}% |
| Modelo B (subclases) | {_f(R['oct_modelo_B'], 2)}% |
| EEE de octubre | {_f(m.loc['2026-10', 'eee_mm_mes'], 2)}% |
| **INE efectivo** | **{_f(oct_ef, 2)}%** |

El Modelo B anticipó bien el mes. El A erró en dos direcciones:

- sobrestimó los servicios SV administrados e indexados (−0,09 pp de incidencia; ahí está internet) y los bienes (−0,14 pp);
- subestimó los combustibles (2,0% vs 4,3%; +0,14 pp), que el nowcast de ENAP sí capturaba.

La sorpresa de octubre vino de productos puntuales: conexión a internet
(−14,7%), pasajes aéreos nacionales (−11,8%) y paltas. Además, el INE no movió la parafina pese al alza de ENAP, porque fuera
de la temporada de calefacción queda sin variación. Esa regla ya está incorporada.

{t_bt}

**Nov-dic 2026.** El precio diario a público se reconstruye con los informes de ENAP (ancla CNE) y se promedia por mes. El
12-nov se incorpora el alza de 40 pesos en bencinas y diésel anunciada por el Gobierno. El escenario base supone alzas
iguales el 3-dic y el 24-dic (ciclo de tres semanas del MEPCO); el alternativo, solo la del 12-nov.

![]({figs['enap']})

{t_nc}

**Efecto del nowcast en la proyección:**

{t_imp}

**Supermercados.** El scraper (VTEX de Jumbo, Santa Isabel y Unimarc, 35 productos de alimentos) sigue sin poder correr:
el entorno solo permite los dominios del INE y del BCCh. Al habilitarlos se incorpora igual que ENAP.

![]({figs['difusion']})

------------------------------------------------------------------------

## 3. Propuesta: Excel autocontenido

`proyeccion_ipc/Outputs/IPC_Chile_Proyecciones.xlsx`. Contiene todos los datos y las fórmulas, de modo que no depende de
archivos externos ni de macros. Al editar un supuesto se recalculan los índices, los agregados, las incidencias, la difusión
y el resumen.

{t_excel}

**Flujo de cálculo:**

```
Matriz (283 productos, pesos, 0/1 por agregado)
   └─► Subclases: peso = SUMIF ; participación en agregado = SUMIFS / peso
Nowcast_ENAP: informes + escenario ─► precio diario (SUMIFS) ─► promedio mensual (AVERAGEIFS) ─► % m/m
   └─► Proy_mm (gasolina y diésel nov-dic) ; resto de subclases = Modelo B (editable)
Indices: historia INE │ proyección = índice(t−1) × (1 + Proy_mm/100)
   ├─► Var_mm_subclases ─► Difusion (COUNTIF > 0) vs rango histórico
   └─► Agregados (SUMPRODUCT peso × participación × índice) ─► Var_mm, Var_12m, Incidencias ─► Resumen
```

**Convenciones:** amarillo = supuesto editable; verde = celda enlazada al nowcast; gris = mes proyectado; azul =
encabezado.

**Verificación:** el libro se recalculó con LibreOffice sin errores de fórmula. Los resultados del Resumen coinciden con los
de Python al centésimo, y las incidencias suman exactamente la inflación del IPC.

**Para actualizar:** correr `python proyeccion_ipc.py`, que regenera el libro con el último IPC. Para un escenario propio, editar
Proy_mm (por ejemplo, un alza de tarifas eléctricas en enero) o las filas de escenario de Nowcast_ENAP.

------------------------------------------------------------------------

## 4. Propuesta: script Python autocontenido

`proyeccion_ipc/Code/proyeccion_ipc.py` es un solo archivo, con la misma estructura de carpetas que las minutas
(rutas relativas al script, sirve en cualquier computador y con F5 en Spyder):

| Carpeta | Contenido |
|:---|:---|
| `Code/` | `proyeccion_ipc.py` |
| `Inputs/` | Insumos fijos: `IPC_2023_Automatizado_v2.xlsx`, `matriz_productos.csv`, `difusion_historica.csv` y `Apuntes_IPC_Chile.html` (plantilla del reporte). |
| `Inputs/INE/`, `Inputs/BDE/` | Descargas del INE y de la BDE; se reemplazan en cada corrida. |
| `Outputs/` | `IPC_Chile_Proyecciones.xlsx`, este HTML y `tablas/` (CSV de proyecciones, evaluación, incidencias y nowcast). |

| Sección del script | Qué hace |
|:---|:---|
| 0. Parámetros | Series BDE, archivos INE, informes de ENAP con sus fuentes, anclas, escenarios, bloques y agregados. Todo está en el propio archivo. |
| 1. Descargas | BDE con el token del BCCh (`BDE_TOKEN` o `BCCH_API_token`, o llavero `estudios_santander`; si no, usuario y clave en `BCCH_USER`/`BCCH_PASS`) y API de archivos del INE (IPC por producto, serie referencial 2023, analíticos). |
| 2. Base | Panel de productos y subclases 2018-2026, bloques 1998-2026, agregados, macro y difusión. Valida contra el INE. |
| 3. Nowcast | Precio diario de combustibles desde ENAP, backtest, escenarios y regla de la parafina. |
| 4. Modelo A | Bloques Carlomagno: tendencias comunes, M0-M3, combinación y evaluación recursiva. |
| 5. Modelo B | 172 subclases con anclaje, AR acotado y ECM de tendencia común. |
| 6. Evaluación | RMSE por horizonte y Diebold-Mariano en orígenes comunes, con benchmarks y EEE. |
| 7-9. Salidas | CSV, Excel con fórmulas y este reporte HTML. |

```bash
cd proyeccion_ipc/Code
python proyeccion_ipc.py                 # descarga + modelos + Excel + reporte (~2 min con descargas)
python proyeccion_ipc.py --sin-descarga  # reusa Inputs/INE e Inputs/BDE (~30 s)
```

La credencial del BCCh se guarda una vez por computador y nunca se escribe en el script. En Mac, desde la consola de
Spyder: `import keyring, getpass; keyring.set_password('estudios_santander', 'BDE_TOKEN', getpass.getpass('Token BCCh: '))`.
En Windows: `setx BDE_TOKEN "tu_token"`.

------------------------------------------------------------------------

## Anexo: validación de la base

{t_val}

------------------------------------------------------------------------

## Bibliografía y fuentes

- Carlomagno, G., Fornero, J. y Sansone, A. (2021). *Toward a general framework for constructing and evaluating core
  inflation measures*. Documento de Trabajo BCCh N°913.
- Carlomagno, G. y Espasa, A. (2015). *Forecasting a large set of disaggregates with common trends and outliers*. UC3M.
  (<a href="https://e-archivo.uc3m.es/entities/publication/18b7092c-659f-4284-8f86-6e40c06e70e4/full">e-Archivo</a>)
- Carlomagno, G. y Espasa, A. (2021). *Discovering specific common trends in a large set of disaggregates*.
  (<a href="https://cinve.org.uy/discovering-specific-common-trends-in-a-large-set-of-disaggregates-statistical-procedures-their-properties-and-an-empirical-application">CINVE</a>)
- Cobb, M. (2009). *Forecasting Chilean Inflation From Disaggregate Components*. Documento de Trabajo BCCh N°545.
  (<a href="https://www.bcentral.cl/en/content/-/detalle/documento-de-trabajo-n-545">BCCh</a>)
- Pincheira, P. y Rubio, H. *The low predictive power of simple Phillips curves in Chile*. Documento de Trabajo BCCh N°559.
- Diebold, F. y Mariano, R. (1995). Comparing Predictive Accuracy. *JBES*, 13(3).
- INE. IPC base anual 2023=100, cuadros estadísticos (www.ine.gob.cl). Banco Central de Chile, BDE y EEE.
- ENAP, informes semanales de precios (vía prensa, una URL por fecha en el script). Anuncio del alza del 12-nov:
  <a href="{ENAP_FUENTES_ESCENARIO[0]}">El Dínamo</a>.
"""
    cuerpo = subprocess.run(["pandoc", "-f", "markdown+tex_math_dollars", "-t", "html", "--mathjax", "--section-divs"],
                            input=md, capture_output=True, text=True, check=True).stdout
    plantilla = (INPUTS / "Apuntes_IPC_Chile.html").read_text(encoding="utf-8")
    ini = plantilla.find('<div class="container-fluid main-container">')
    fin = plantilla.find("\n<script>\n\n// add bootstrap table styles")
    cab = plantilla[:ini].replace("<title>Apuntes_IPC_Chile</title>", "<title>IPC Chile modelos</title>")
    cab = cab.replace("</head>", "<style>details{margin:0 0 1em} summary{cursor:pointer;color:#2a78d6} table{font-size:13px}"
                                 " .main-container{max-width:1080px} img{display:block;margin:.5em auto 1.2em}</style></head>", 1)
    html = cab + '<div class="container-fluid main-container">\n<div id="header"></div>\n' + cuerpo + "\n</div>\n" + plantilla[fin:]
    open(destino, "w", encoding="utf-8").write(html)
    log(f"Reporte: {destino}")
    return destino


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0], formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--sin-descarga", action="store_true", help="usa los datos ya descargados")
    ap.add_argument("--sin-reporte", action="store_true", help="no genera el HTML")
    args = ap.parse_args()
    OUTPUTS.mkdir(parents=True, exist_ok=True)
    if not args.sin_descarga:
        descargar_bde()
        descargar_ine()
    log("Construyendo la base…")
    base = Base()
    log(f"Último IPC: {base.T}; {base.prod_idx.shape[1]} productos, {base.sc_idx.shape[1]} subclases")
    fijos = nowcast_fijos_productos(base, ENAP_ESCENARIO_CENTRAL)
    A = ModeloBloques(base).correr(fijos)
    B = ModeloSubclases(base).correr(fijos)
    R = resultados(base, A, B)
    guardar_csv(R)
    construir_excel(R)
    if not args.sin_reporte:
        construir_reporte(R)
    log("Listo.")


if __name__ == "__main__":
    main()
