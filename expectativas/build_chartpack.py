"""Chartpack de Expectativas - Gerencia de Estudios (Santander Chile).

Define el contenido (series, transformaciones, laminas). Los datos, el formato y los
graficos nativos de PowerPoint vienen de macro.py.

Uso:
    python build_chartpack.py
Genera output/AAAAMMDD_Chartpack_Expectativas.pptx, el .pdf (si hay LibreOffice) y un
.xlsx con los datos de cada grafico.
"""
import datetime as dt
import shutil
import subprocess
from pathlib import Path

import pandas as pd
from pptx.util import Cm

import macro

OUT = Path(__file__).parent / "output"
BCCH = "Banco Central de Chile"
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]

# posiciones de graficos (cm): dos por lamina, tres por lamina y uno ancho
Y0, ALTO = Cm(3.6), Cm(macro.ALTO_CM)
DOS = [Cm(1.3), Cm(17.5)]
ANCHO_DOS = Cm(macro.ANCHO_CM)
TRES = [Cm(1.0), Cm(11.8), Cm(22.6)]
ANCHO_TRES = Cm(10.2)
TAM_TRES = 10   # graficos de 10 cm de ancho: fuente 10 para que leyenda y ejes no se superpongan
ANCHO_UNO = Cm(31.2)

DATOS = {}   # hoja del Excel de respaldo -> DataFrame


def graficar(lam, df, x, w, hoja, fuente=BCCH, nota=None, **kw):
    """Grafico + pie (nota y fuente) bajo el grafico; guarda los datos para el Excel."""
    macro.agregar_grafico(lam, df, x, Y0, w, ALTO, **kw)
    macro.pie(lam, x, Y0 + ALTO + Cm(0.15), w, fuente, nota, h=Cm(1.6))
    DATOS[hoja] = df


def main(hoy=None):
    hoy = pd.Timestamp(hoy or dt.date.today())
    a0, a1 = hoy.year, hoy.year + 1
    OUT.mkdir(exist_ok=True)

    prs = macro.nueva_presentacion()
    macro.portada(prs, f"{MESES[hoy.month - 1]} de {hoy.year}")

    # --- 2. IMCE: precios y costos esperados por sector
    sectores = ["Comercio", "Construcción", "Industria"]
    precios = macro.bde(["G089.IME.IND.A14.M", "G089.IME.IND.A34.M", "G089.IME.IND.A27.M"],
                        nombres=sectores, desde="2018-01-01")
    costos = macro.bde(["G089.IME.IND.A15.M", "G089.IME.IND.A35.M", "G089.IME.IND.A29.M"],
                       nombres=sectores, desde="2018-01-01")
    lam = macro.lamina(prs, "IMCE", 2)
    for x, df, tit, hoja in [(DOS[0], precios, "IMCE: Indicador de Precios", "02 IMCE precios"),
                             (DOS[1], costos, "IMCE: Indicador de Costos", "02 IMCE costos")]:
        graficar(lam, df, x, ANCHO_DOS, hoja, titulo=tit, unidad="(Índice)", y_min=0, y_max=100,
                 y_paso=10, linea_ref=50, meses_entre_etiquetas=12)

    # --- 3. Expectativas de inflacion EEE y EOF
    eee = macro.bde(["F089.IPC.V12.14.M", "F089.IPC.V12.15.M"], nombres=["EEE 1 año", "EEE 2 años"],
                    desde="2015-01-01")
    eof = macro.bde(["F089.EOF.VII.12MS.D", "F089.EOF.VII.S12M.D"], nombres=["EOF 1 año", "EOF 2 años"],
                    desde="2015-01-01")
    lam = macro.lamina(prs, "Expectativas de Inflación", 3)
    graficar(lam, eee, DOS[0], ANCHO_DOS, "03 EEE inflacion", titulo="Mediana EEE 1 y 2 años",
             unidad="(Var. anual, %)", posicion_unidad="arriba", y_min=1, y_max=9, y_paso=1,
             meses_entre_etiquetas=12, formato_fecha="yyyy",
             nota="Inflación esperada a 11 meses (1 año) y a 23 meses (2 años).")
    graficar(lam, eof, DOS[1], ANCHO_DOS, "03 EOF inflacion", titulo="Mediana EOF 1 y 2 años",
             unidad="(Var. anual, %)", posicion_unidad="arriba", y_min=1, y_max=9, y_paso=1,
             meses_entre_etiquetas=12, formato_fecha="yyyy",
             nota="Inflación esperada para los 12 meses siguientes (1 año) y de 13 a 24 meses (2 años).")

    # --- 4. Compensacion inflacionaria (swaps promedio camara)
    comp = macro.bde(["F022.SWSP.TAS.AN01.Z.Z.D", "F022.SWSP.TAS.AN02.Z.Z.D"],
                     nombres=["1 año", "2 años"], desde="2022-01-01")
    c11 = macro.bde(["F022.SWFW.TAS.E1X1.Z.Z.D"], nombres=["Swap 1 en 1"], desde="2022-01-01")
    lam = macro.lamina(prs, "Compensación Inflacionaria", 4)
    graficar(lam, comp, DOS[0], ANCHO_DOS, "04 Compensaciones", titulo="Compensaciones inflacionarias",
             unidad="(Var. anual, %)", posicion_unidad="arriba", y_min=0, y_paso=2,
             meses_entre_etiquetas=6, nota="Compensación spot a partir de swaps promedio cámara.")
    graficar(lam, c11, DOS[1], ANCHO_DOS, "04 Swap 1 en 1", titulo="Compensación inflacionaria 1 en 1: Swap",
             unidad="(%)", posicion_unidad="arriba", y_min=0, y_paso=1, meses_entre_etiquetas=6,
             nota="Compensación forward 1 año en 1 año a partir de swaps promedio cámara.")

    # --- 5. EDEP: costos y determinantes del precio de venta
    ini = "2021-11-01"
    costos_edep = ["Costos laborales", "Precios de insumos", "Costo de maquinarias o tecnología",
                   "Costos fijos"]
    edep_c = macro.bde([f"F089.EDEP.IND.88{k}.M" for k in ["CL", "PI", "CMT", "CF"]],
                       nombres=costos_edep, desde=ini)
    edep_i = macro.bde([f"F089.EDEP.IND.84IF{k}.M" for k in ["NV", "PD", "MG", "DI", "C"]],
                       nombres=["Niveles de venta", "Precio del dólar", "Márgenes de ganancias",
                                "Disponibilidad de insumos", "Costos"], desde=ini)
    edep_p = macro.bde([f"F089.EDEP.IND.90{k}.M" for k in ["CL", "PI", "CMT", "CF"]],
                       nombres=costos_edep, desde=ini)
    lam = macro.lamina(prs, "EDEP", 5)
    comun = dict(unidad="(Índice de difusión)", linea_ref=50, meses_entre_etiquetas=6, tam=TAM_TRES)
    graficar(lam, edep_c, TRES[0], ANCHO_TRES, "05 EDEP costos", titulo="EDEP: Evolución de costos",
             y_min=40, y_paso=5, nota="Últimos 3 meses; datos de trimestre móvil.", **comun)
    graficar(lam, edep_i, TRES[1], ANCHO_TRES, "05 EDEP precio venta",
             titulo="EDEP: Influencia en precio de venta", y_min=0, y_paso=10,
             nota="Últimos 3 meses; datos de trimestre móvil.", **comun)
    graficar(lam, edep_p, TRES[2], ANCHO_TRES, "05 EDEP costos esperados",
             titulo="EDEP: Costos esperados", y_min=40, y_paso=5,
             nota="Próximos 3 meses; datos de trimestre móvil.", **comun)

    # --- 6. EDEP expectativas de inflacion y costos de transporte
    edep_inf = macro.bde(["F089.EDEP.POR.91EI12.M", "F089.EDEP.POR.91EI24.M"],
                         nombres=["EDEP a 1 año", "EDEP a 2 años"], desde=ini)
    gscpi = macro.nyfed_gscpi(desde="1998-01-01", nombre="GSCPI")
    lam = macro.lamina(prs, "IMCE, EDEP y Costos de transporte", 6)
    graficar(lam, edep_inf, DOS[0], ANCHO_DOS, "06 EDEP inflacion", titulo="EDEP: Expectativas de Inflación",
             unidad="(Var. anual, %)", y_min=0, y_paso=2, meses_entre_etiquetas=6,
             nota="Mediana de inflación esperada por las empresas, 1 a 12 y 13 a 24 meses.")
    graficar(lam, gscpi, DOS[1], ANCHO_DOS, "06 GSCPI", fuente=gscpi.attrs["fuente"],
             titulo="Costos de transporte", unidad="(Desv. estándar del promedio)",
             posicion_unidad="arriba", y_paso=1, meses_entre_etiquetas=36, formato_fecha="yyyy",
             nota="Global Supply Chain Pressure Index (GSCPI).")

    # --- 7. Tipo de cambio nominal y multilateral
    tc = macro.bde(["F073.TCO.PRE.Z.D", "F073.TCM.IND.199502.D"], nombres=["TCN", "TCM (der.)"],
                   desde="2018-01-01")
    lam = macro.lamina(prs, "TCN TCM", 7)
    graficar(lam, tc, DOS[0], ANCHO_UNO, "07 TCN TCM",
             titulo="Tipo de cambio nominal y tipo de cambio multilateral", decimales=0,
             unidad="(Pesos chilenos)", unidad_derecha="(Índice)", posicion_unidad="arriba",
             eje_derecho=["TCM (der.)"], y_min=500, y_paso=100, y2_min=90, y2_paso=10,
             meses_entre_etiquetas=12, formato_fecha="yyyy",
             nota="TCM: tipo de cambio multilateral, índice 2 de enero de 1998=100.")

    # --- 8. Tasas swap promedio camara
    ini8 = f"{a0 - 1}-01-01"
    spc_clp = macro.bde(["F022.SPC.TPR.D360.NO.Z.D", "F022.SPC.TIN.AN02.NO.Z.D"],
                        nombres=["SPC pesos 1 año", "SPC pesos 2 años"], desde=ini8)
    spc_uf = macro.bde(["F022.SPC.TIN.AN01.UF.Z.D", "F022.SPC.TIN.AN02.UF.Z.D"],
                       nombres=["SPC UF 1 año", "SPC UF 2 años"], desde=ini8)
    lam = macro.lamina(prs, "Tasas SPC", 8)
    graficar(lam, spc_clp, DOS[0], ANCHO_DOS, "08 SPC pesos", titulo="SPC pesos: 1 y 2 años",
             unidad="(%)", y_paso=0.5, meses_entre_etiquetas=3, decimales=1,
             nota="SPC pesos 1 año: tasa swap promedio cámara nominal a 360 días.")
    graficar(lam, spc_uf, DOS[1], ANCHO_DOS, "08 SPC UF", titulo="SPC UF: 1 y 2 años",
             unidad="(%)", y_paso=0.5, meses_entre_etiquetas=3, decimales=1)

    # --- 9. EEE: actividad, anio en curso y siguiente
    ini9 = f"{a0 - 1}-01-01"
    lam = macro.lamina(prs, "EEE", 9)
    for x, (cod, tit, paso) in zip(TRES, [("PIB", "Expectativas de PIB", 0.5),
                                          ("CHOG", "Expectativas de Consumo hogares", 0.5),
                                          ("FBKF", "Expectativas de FBCF", 1)]):
        df = macro.bde([f"F089.{cod}.V12.{a0}.M", f"F089.{cod}.V12.{a1}.M"],
                       nombres=[str(a0), str(a1)], desde=ini9)
        graficar(lam, df, x, ANCHO_TRES, f"09 EEE {cod}", titulo=tit, unidad="(Var. anual, %)", y_paso=paso,
                 meses_entre_etiquetas=6, tam=TAM_TRES, nota="Mediana de la EEE para cada año calendario.")

    # --- 10. IMCE expectativas de inversion vs Imacec no minero
    imce = macro.bde(["G089.IME.IND.A17.M", "G089.IME.IND.A2B.M"],
                     nombres=["IMCE Comercio", "IMCE Industria"], desde="2010-01-01")
    imce_inv, nota_imce = macro.promedio(imce, "IMCE expectativas de inversión (der.)")
    imacec = macro.bde(["F032.IMC.IND.Z.Z.EP18.N03.Z.0.M"], nombres=["Imacec No Minero"],
                       desde="2008-01-01")
    imacec, nota_mm = macro.media_movil(imacec, 3, "meses")
    imacec, nota_va = macro.var_anual(imacec, 12)
    df10 = pd.concat([imacec, imce_inv], axis=1)
    df10 = df10[df10.index >= "2010-01-01"].dropna(how="all")
    lam = macro.lamina(prs, "Imce expectativas v/s Imacec no minero", 10)
    graficar(lam, df10, DOS[0], ANCHO_UNO, "10 IMCE vs Imacec",
             titulo="IMCE expectativas de inversión v/s Imacec no minero",
             unidad="(Var. trimestral a/a, %)", unidad_derecha="(Índice)", posicion_unidad="arriba",
             eje_derecho=["IMCE expectativas de inversión (der.)"], y_min=-20, y_max=25, y_paso=5,
             y2_min=10, y2_max=80, y2_paso=10, meses_entre_etiquetas=12, formato_fecha="yyyy",
             nota=f"Imacec no minero: {nota_mm.lower()} {nota_va} {nota_imce}")

    # --- guardar
    nombre = f"{hoy:%Y%m%d}_Chartpack_Expectativas"
    pptx = OUT / f"{nombre}.pptx"
    prs.save(pptx)
    with pd.ExcelWriter(OUT / f"{nombre}_datos.xlsx") as xl:
        for hoja, df in DATOS.items():
            df.rename_axis("Fecha").to_excel(xl, sheet_name=hoja[:31])
    print(pptx)
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(OUT), str(pptx)],
                       check=True, capture_output=True)
        print(OUT / f"{nombre}.pdf")


if __name__ == "__main__":
    main()
