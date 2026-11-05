# -*- coding: utf-8 -*-
r"""Datos de la Minuta Remuneraciones: tabulados del INE de remuneraciones (IR-ICL, base anual 2023=100).

1. Descarga los tabulados del INE (IR, ICL, IR real, desestacionalizados y horas) y los deja en
   Inputs\ (se reemplazan en cada corrida, porque el INE publica siempre con el mismo nombre).
2. cargar() los lee para minuta_remuneraciones.py.

Carpetas (relativas a este script; sirven para cualquier usuario):
    <Minuta Remuneraciones>\Code\     datos_remuneraciones.py y minuta_remuneraciones.py
    <Minuta Remuneraciones>\Inputs\   tabulados del INE (los deja este script)
    <Minuta Remuneraciones>\Outputs\  AAAAMMDD_Minuta_Remuneraciones.pdf y figuras\

Uso (no hay que editar nada):
    python datos_remuneraciones.py           # descarga los tabulados del INE a Inputs
    python minuta_remuneraciones.py          # arma la minuta en Outputs
"""
import sys
from pathlib import Path

import pandas as pd

CODE = Path(__file__).resolve().parent
PROYECTO = CODE.parent
INPUTS = PROYECTO / "Inputs"
OUT = PROYECTO / "Outputs"

BASE = ("https://www.ine.gob.cl/docs/default-source/sueldos-y-salarios/cuadros-estadisticos/"
        "ir-icl-base-anual-2023-100/")
ARCHIVOS = {
    "tabulado_ir_empalmado.xlsx": "series-empalmadas/",
    "tabulado_icl_empalmado.xlsx": "series-empalmadas/",
    "tabulado_ir_real_empalmado.xlsx": "series-empalmadas/",
    "tabulado_ir_icl_nominal_general_empalmado_desestacionalizado.xlsx": "series-empalmadas-desestacionalizadas/",
    "tabulado_ir_real_general_empalmado_desestacionalizado.xlsx": "series-empalmadas-desestacionalizadas/",
    "tabulado_estimaciones_ho_he_ht.xlsx": "series-base-2023/",
}
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36")

ACTIVIDADES = {
    "B": "Minería", "C": "Industria manufacturera", "D": "Electricidad, gas, vapor y aire acondicionado",
    "E": "Suministro de agua y gestión de desechos", "F": "Construcción", "G": "Comercio",
    "H": "Transporte y almacenamiento", "I": "Alojamiento y servicio de comidas",
    "J": "Información y comunicaciones", "K": "Actividades financieras y de seguros",
    "L": "Actividades inmobiliarias", "M": "Actividades profesionales, científicas y técnicas",
    "N": "Servicios administrativos y de apoyo", "O": "Administración pública", "P": "Enseñanza",
    "Q": "Salud humana y asistencia social", "R": "Actividades artísticas y recreativas",
}
GRUPOS = {
    "1": "Directores, gerentes y administradores", "2": "Profesionales, científicos e intelectuales",
    "3": "Técnicos y profesionales de nivel medio", "4": "Personal de apoyo administrativo",
    "5.1": "Trabajadores de los servicios", "5.2": "Vendedores", "7": "Artesanos y operarios de oficios",
    "8": "Operadores de instalaciones y máquinas", "9": "Ocupaciones elementales",
}


# ─────────────────────────────────────────────────────────────────────────────
# 1. DESCARGA
# ─────────────────────────────────────────────────────────────────────────────

def descargar(carpeta: Path = INPUTS) -> int:
    """Baja a Inputs/ los tabulados del INE; devuelve cuántos bajó. Si uno falla, se informa y se
    sigue con el siguiente (queda la versión anterior en Inputs, si existe)."""
    import requests

    carpeta.mkdir(parents=True, exist_ok=True)
    s = requests.Session()
    s.headers["User-Agent"] = UA
    nuevas = 0
    for nombre, sub in ARCHIVOS.items():
        try:
            r = s.get(BASE + sub + nombre, timeout=120)
            r.raise_for_status()
            if not r.content.startswith(b"PK"):   # página de error o de bloqueo, no un Excel
                raise RuntimeError("la respuesta no es un Excel válido")
            (carpeta / nombre).write_bytes(r.content)
            nuevas += 1
            print(f"    ok: {nombre} ({len(r.content) / 1e3:.0f} KB)")
        except Exception as e:  # noqa: BLE001 - se informa y se sigue con el siguiente
            print(f"    error: {nombre} {e}")
    return nuevas


# ─────────────────────────────────────────────────────────────────────────────
# 2. LECTURA
# ─────────────────────────────────────────────────────────────────────────────

def _input(nombre, carpeta):
    ruta = carpeta / nombre
    if not ruta.is_file():
        sys.exit(f"No encuentro {ruta}. Corre primero: python datos_remuneraciones.py")
    return ruta


def _hoja(ruta: Path, hoja=1) -> pd.DataFrame:
    """Lee una hoja de datos (la primera después del menú, salvo que se indique otra)."""
    d = pd.read_excel(ruta, sheet_name=hoja)
    d.columns = [str(c).strip().lower() for c in d.columns]
    d = d[pd.to_numeric(d["año"], errors="coerce").notna()].copy()
    d.index = pd.PeriodIndex.from_fields(year=d["año"].astype(int), month=d["mes"].astype(int), freq="M")
    return d


def cargar(carpeta: Path = INPUTS) -> dict:
    f = lambda nombre: _input(nombre, carpeta)  # noqa: E731
    nsa = pd.DataFrame({
        "ir": _hoja(f("tabulado_ir_empalmado.xlsx"))["índice"],
        "icl": _hoja(f("tabulado_icl_empalmado.xlsx"))["índice"],
        "ir_real": _hoja(f("tabulado_ir_real_empalmado.xlsx"))["índice"],
    }).astype(float)
    sa_n = _hoja(f("tabulado_ir_icl_nominal_general_empalmado_desestacionalizado.xlsx"))
    col = "indicador desestacionalizado (id)"
    sa = pd.DataFrame({
        "ir": sa_n[sa_n["indicador"] == "IR"][col],
        "icl": sa_n[sa_n["indicador"] == "ICL"][col],
        "ir_real": _hoja(f("tabulado_ir_real_general_empalmado_desestacionalizado.xlsx"))[
            "índice desestacionalizado (id)"],
    }).astype(float)

    horas = f("tabulado_estimaciones_ho_he_ht.xlsx")
    general = _hoja(horas, "General")
    act = _hoja(horas, "Actividad")
    act["nombre"] = act["actividad"].map(ACTIVIDADES)
    grp = _hoja(horas, "Grupo")
    grp["nombre"] = grp["grupo"].map(lambda g: GRUPOS[f"{float(g):g}"])
    return {"nsa": nsa, "sa": sa, "horas": general, "horas_act": act, "horas_grp": grp}


def main():
    print("=" * 60 + "\n  Minuta Remuneraciones: datos del INE\n" + "=" * 60)
    nuevas = descargar(INPUTS)
    print(f"  {nuevas} de {len(ARCHIVOS)} tabulados descargados en {INPUTS}")
    if nuevas < len(ARCHIVOS):
        faltan = [n for n in ARCHIVOS if not (INPUTS / n).is_file()]
        if faltan:
            sys.exit("Faltan tabulados en Inputs: " + ", ".join(faltan))
        print("  Aviso: algunos tabulados no se actualizaron; se usarán los que ya estaban en Inputs.")


if __name__ == "__main__":
    main()
