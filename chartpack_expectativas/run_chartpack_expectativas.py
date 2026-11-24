# -*- coding: utf-8 -*-
"""Corre el Chartpack de Expectativas: run encuestas_bcch.py y run build_chartpack.py

Va en la carpeta Expectativas, junto a Code, Inputs y Outputs. Ejecuta, en orden:
  1. Code/encuestas_bcch.py: descarga las planillas de la EEE y la EOF que falten y actualiza los Excel
     de Inputs/eee e Inputs/eof.
  2. Code/build_chartpack.py: descarga las series del BDE, arma el chartpack (PPTX y PDF) en Outputs y
     actualiza Chartpack_Expectativas.xlsx.

Uso: abre este archivo en Spyder o VS Code y presiona F5 / Run (o, en la terminal:
python run_chartpack_expectativas.py).
"""
import runpy
import sys
from pathlib import Path

SIN_DESCARGA_ENCUESTAS = False   # True: no descarga planillas EEE/EOF; usa las que ya están en Inputs

CODE = Path(__file__).resolve().parent / "Code"
ENCUESTAS = CODE / "encuestas_bcch.py"
CHARTPACK = CODE / "build_chartpack.py"


def correr(script, argumentos=()):
    print(f"\nrun {script.name}")
    sys.argv = [str(script), *argumentos]
    runpy.run_path(str(script), run_name="__main__")


if __name__ == "__main__":
    sys.path.insert(0, str(CODE))
    try:
        correr(ENCUESTAS, ["--sin-descarga"] if SIN_DESCARGA_ENCUESTAS else [])
    except SystemExit as e:   # si las encuestas fallan, el chartpack se arma con los Excel que ya hay en Inputs
        if e.code not in (None, 0):
            print(f"⚠️ encuestas_bcch.py terminó con un problema ({e}); sigo con los datos de Inputs.")
    correr(CHARTPACK)
