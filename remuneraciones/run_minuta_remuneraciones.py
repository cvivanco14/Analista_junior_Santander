# -*- coding: utf-8 -*-
"""Corre la Minuta de Remuneraciones: run minuta_remuneraciones.py

Va en la carpeta Remuneraciones, junto a Code, Inputs y Outputs. Ejecuta Code/minuta_remuneraciones.py,
que descarga los datos del INE y arma el documento, la presentación e IRL_Base.xlsx en Outputs.

Uso: abre este archivo en Spyder y presiona F5 (o, en la terminal: python run_minuta_remuneraciones.py).
"""
import runpy
import sys
from pathlib import Path

SIN_DESCARGA = False   # True: usa los archivos del INE que ya están en Inputs (no descarga)

SCRIPT = Path(__file__).resolve().parent / "Code" / "minuta_remuneraciones.py"

if __name__ == "__main__":
    print(f"run {SCRIPT.name}")
    sys.argv = [str(SCRIPT)] + (["--sin-descarga"] if SIN_DESCARGA else [])
    sys.path.insert(0, str(SCRIPT.parent))
    runpy.run_path(str(SCRIPT), run_name="__main__")
