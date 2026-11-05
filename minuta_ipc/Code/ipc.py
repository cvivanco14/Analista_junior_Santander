"""Datos y cálculos de la minuta IPC.

Fuentes:
  * INE (último dato): índices por producto, base anual 2023=100 (desde ene-2024) y serie
    referencial 2023 con la misma canasta. Se descargan en Inputs/INE/.
  * Inputs/matriz_productos.csv: clasificación de cada producto en los agregados volátiles /
    sin volátiles (coincide con la matriz oficial del Banco Central para IPC SV, bienes y
    servicios SV).
  * Inputs/historico_bcch_indices.csv: agregados empalmados del Banco Central hasta dic-2023
    (canasta 2018), usados solo para la historia de los gráficos.
  * Inputs/pesos_canasta_2018.csv, difusion_historica.csv: insumos fijos de la canasta 2018.
  * Inputs/ipc_base2018_alimentos.csv: IPC base 2018 del INE (división 1 y grupo 11.1) para
    actualizar la CBA entre mar-2022 y dic-2023.

CBA (metodología 2024 del MDSF): $66.896 a marzo de 2022, actualizada con un índice de alimentos
que pondera, con las ponderaciones del IPC, la división 1 (alimentos y bebidas no alcohólicas) y
el grupo 11.1 (restaurantes, cafés y similares); se encadena entre canastas en dic-2023.

Los agregados desde 2023 se calculan como promedio ponderado de los índices de producto
(Laspeyres base 2023): I_agregado = sum(w_i * I_i) / sum(w_i).
"""

from pathlib import Path

import pandas as pd
import requests

URL_BASE = ("https://www.ine.gob.cl/docs/default-source/%C3%ADndice-de-precios-al-consumidor/"
            "cuadros-estadisticos/base-anual-2023_100/")
ARCHIVOS = {
    "ipc-base-2023.xlsx": URL_BASE + "series-de-tiempo/ipc-xls.xlsx",
    "ipc-referencial-2023.xlsx": URL_BASE + "series-referenciales/ipc-referencial-base-2023-xls.xlsx",
}
CAMBIO_CANASTA = pd.Period("2023-12", "M")  # último mes con historia del Banco Central (canasta 2018)

# Agregados de la minuta: nombre en la matriz -> nombre para mostrar
AGREGADOS = {
    "IPC": "IPC",
    "IPC sin volátiles": "IPC sin volátiles",
    "Bienes sin volátiles": "Bienes sin volátiles",
    "Alimentos no volátiles": "Alimentos no volátiles",
    "Bienes sv. ex. alim.": "Bienes SV exc. alimentos",
    "Servicios sin volátiles": "Servicios sin volátiles",
    "Servicios_A&I_sv": "Servicios SV A&I",
    "Servicios_ Resto_sv": "Servicios SV exc. A&I",
    "IPC volátiles": "IPC volátiles",
    "Energía Volátiles": "Energía volátiles",
    "Electricidad_v": "Electricidad",
    "Combustible_v": "Combustibles",
    "Alim. Volát.": "Alimentos volátiles",
    "Alim. Resto_v": "Alimentos volátiles exc. FyV",
    "FyV (proxy)_v": "Frutas y verduras",
    "Resto volátiles": "Resto volátiles",
    "Bienes volát.": "Bienes volátiles",
    "Servicios volát.": "Servicios volátiles",
    "Servicios_A&I_v": "Servicios volátiles A&I",
    "Servicios_Resto_v": "Servicios volátiles exc. A&I",
}


def descargar(carpeta: Path) -> None:
    carpeta.mkdir(parents=True, exist_ok=True)
    s = requests.Session()
    s.headers["User-Agent"] = "Mozilla/5.0 (minuta-ipc)"
    for nombre, url in ARCHIVOS.items():
        r = s.get(url, timeout=120)
        r.raise_for_status()
        if not r.content.startswith(b"PK"):
            raise RuntimeError(f"La descarga de {nombre} no es un Excel válido.")
        (carpeta / nombre).write_bytes(r.content)
        print(f"  descargado {nombre} ({len(r.content) / 1e6:.1f} MB)")


def _leer_ine(ruta: Path) -> pd.DataFrame:
    d = pd.read_excel(ruta, header=None)
    fila = d.index[d[0].astype(str).str.strip() == "Año"][0]
    d = d.iloc[fila + 1:, :13]
    d.columns = ["anio", "mes", "div", "grp", "clase", "subclase", "prod", "glosa", "pond", "indice",
                 "var_mm", "var_acum", "var_aa"]
    d = d[pd.to_numeric(d["anio"], errors="coerce").notna()].copy()
    d["periodo"] = pd.PeriodIndex.from_fields(year=d.anio.astype(int), month=d.mes.astype(int), freq="M")
    for c in ["div", "grp", "clase", "subclase", "prod"]:
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0).astype(int)
    for c in ["indice", "var_mm", "var_acum", "var_aa"]:
        d[c] = pd.to_numeric(d[c], errors="coerce")
    d["glosa"] = d["glosa"].astype(str).str.strip()
    return d


def cargar(datos: Path, insumos: Path) -> dict:
    ine = pd.concat([_leer_ine(datos / "ipc-referencial-2023.xlsx"), _leer_ine(datos / "ipc-base-2023.xlsx")])
    ine = ine.drop_duplicates(["periodo", "div", "grp", "clase", "subclase", "prod", "glosa"], keep="last")

    general = ine[ine["div"] == 0].set_index("periodo")[["indice", "var_mm", "var_acum", "var_aa"]].sort_index()
    prod = ine[ine["prod"] > 0].copy()
    prod["codigo"] = (prod[["div", "grp", "clase", "subclase", "prod"]].astype(str).agg("".join, axis=1)
                      .astype(int))
    indices = prod.pivot_table(index="periodo", columns="codigo", values="indice").sort_index()

    mat = pd.read_csv(insumos / "matriz_productos.csv").set_index("codigo")
    faltan = set(mat.index) ^ set(indices.columns)
    if faltan:
        raise RuntimeError(f"Los productos del INE no calzan con la matriz: {sorted(faltan)[:10]}")
    indices = indices[mat.index]
    w = mat["ponderacion"]

    agregados = pd.DataFrame({
        a: (indices[mat.index[mat[a] == 1]] * w[mat[a] == 1]).sum(axis=1) / w[mat[a] == 1].sum()
        for a in AGREGADOS
    })
    pesos = pd.Series({a: w[mat[a] == 1].sum() for a in AGREGADOS})

    # índices de la división 1 y del grupo 11.1 (restaurantes), para la CBA
    nivel = ine[ine["clase"] == 0]
    d1 = nivel[(nivel["div"] == 1) & (nivel["grp"] == 0)].set_index("periodo")
    rest = nivel[(nivel["div"] == 11) & (nivel["grp"] == 1)].set_index("periodo")
    alim23 = ((d1["indice"] * d1["pond"].astype(float) + rest["indice"] * rest["pond"].astype(float))
              / (d1["pond"].astype(float) + rest["pond"].astype(float))).sort_index()
    b18 = pd.read_csv(insumos / "ipc_base2018_alimentos.csv")
    b18.index = pd.PeriodIndex(b18.pop("periodo"), freq="M")
    alim18 = ((b18.div1_alimentos * b18.pond_div1 + b18.grupo11_1_restaurantes * b18.pond_grupo11_1)
              / (b18.pond_div1 + b18.pond_grupo11_1))

    hist = pd.read_csv(insumos / "historico_bcch_indices.csv")
    hist.index = pd.PeriodIndex(hist.pop("periodo"), freq="M")
    pesos18 = pd.read_csv(insumos / "pesos_canasta_2018.csv", index_col=0)["ponderacion_2018"]

    return {
        "general_ine": general, "indices": indices, "matriz": mat, "pesos": pesos,
        "agregados": agregados, "hist": hist, "pesos18": pesos18,
        "dif_hist": pd.read_csv(insumos / "difusion_historica.csv"),
        "alim18": alim18, "alim23": alim23,
    }


# ------------------------------------------------------------------ variaciones
def variacion(d: dict, agregado: str, rezago: int) -> pd.Series:
    """Variación % (m/m si rezago=1, a/a si rezago=12): historia del Banco Central hasta
    dic-2023 y, desde ene-2024, la canasta 2023 del INE (serie referencial 2023 como base)."""
    h = d["hist"][agregado] if agregado in d["hist"] else None
    nuevo = d["agregados"][agregado]
    v_nuevo = (nuevo / nuevo.shift(rezago, freq="M") - 1) * 100
    v_nuevo = v_nuevo[v_nuevo.index > CAMBIO_CANASTA].dropna()
    if h is None:
        return v_nuevo
    v_hist = (h / h.shift(rezago, freq="M") - 1) * 100
    return pd.concat([v_hist[v_hist.index <= CAMBIO_CANASTA], v_nuevo]).dropna()


def incidencia(d: dict, agregado: str, rezago: int) -> pd.Series:
    """Incidencia (pp.) del agregado en la variación del IPC total."""
    def calc(idx, ipc, w):
        return w * (idx - idx.shift(rezago, freq="M")) / ipc.shift(rezago, freq="M")
    nuevo = calc(d["agregados"][agregado], d["agregados"]["IPC"], d["pesos"][agregado])
    nuevo = nuevo[nuevo.index > CAMBIO_CANASTA].dropna()
    if agregado not in d["hist"] or agregado not in d["pesos18"]:
        return nuevo
    h = calc(d["hist"][agregado], d["hist"]["IPC"], d["pesos18"][agregado])
    return pd.concat([h[h.index <= CAMBIO_CANASTA], nuevo]).dropna()


def productos(d: dict, p: pd.Period) -> pd.DataFrame:
    """Variaciones e incidencias por producto en el mes p."""
    I, mat, ipc = d["indices"], d["matriz"], d["agregados"]["IPC"]
    out = mat[["glosa", "ponderacion", "categoria", "subcategoria", "subsubcategoria"]].copy()
    out["var_mm"] = (I.loc[p] / I.loc[p - 1] - 1) * 100
    out["var_aa"] = (I.loc[p] / I.loc[p - 12] - 1) * 100 if p - 12 in I.index else float("nan")
    out["inc_mm"] = out["ponderacion"] * (I.loc[p] - I.loc[p - 1]) / ipc[p - 1]
    return out


def difusion(d: dict, agregado: str) -> pd.Series:
    """% de productos del agregado con variación mensual positiva (canasta 2023)."""
    cols = d["matriz"].index[d["matriz"][agregado] == 1]
    I = d["indices"][cols]
    sube = (I > I.shift(1)).sum(axis=1)
    return (sube / len(cols) * 100).iloc[1:]


CBA_MAR2022 = 66896  # valor mensual por persona de la CBA, metodología 2024 (MDSF)


def cba(d: dict) -> pd.Series:
    """Canasta Básica de Alimentos ($ por persona), réplica de la metodología 2024 del MDSF."""
    a18, a23 = d["alim18"], d["alim23"]
    indice = pd.concat([a18[a18.index <= CAMBIO_CANASTA],
                        a18[CAMBIO_CANASTA] * a23[a23.index > CAMBIO_CANASTA] / a23[CAMBIO_CANASTA]])
    return CBA_MAR2022 * indice / indice[pd.Period("2022-03", "M")]
