# -*- coding: utf-8 -*-
"""Corre la Minuta IMCE: run minuta_imce.py

Va en la carpeta IMCE, junto a Code, Inputs y Outputs. Ejecuta Code/minuta_imce.py, que descarga los datos
del BCCh, actualiza Outputs/Base_IMCE.xlsx y arma la presentación Presentación_IMCE_<mes>_<año> (PPTX y PDF)
en Outputs.

Uso: abre este archivo en Spyder y presiona F5 (o, en la terminal: python run_minuta_imce.py).
"""
import runpy
import sys
from pathlib import Path

SIN_DESCARGA = False   # True: usa el Inputs/imce.csv que ya existe (no descarga)

SCRIPT = Path(__file__).resolve().parent / "Code" / "minuta_imce.py"

if __name__ == "__main__":
    print(f"run {SCRIPT.name}")
    sys.argv = [str(SCRIPT)] + (["--sin-descarga"] if SIN_DESCARGA else [])
    sys.path.insert(0, str(SCRIPT.parent))
    runpy.run_path(str(SCRIPT), run_name="__main__")
