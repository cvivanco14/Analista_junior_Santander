"""Scraping diario de precios de supermercados para el nowcasting de alimentos.

Para cada producto IPC de `data/nowcast/canasta_supermercado.csv` busca en las tiendas online
los términos indicados y guarda una foto diaria de precios:
    data/nowcast/supermercados/AAAA-MM-DD.csv
    (fecha, tienda, codigo_ipc, sku, nombre, precio, precio_lista, unidad)

Tiendas: catálogos VTEX públicos (Jumbo, Santa Isabel, Unimarc). Los endpoints están en TIENDAS
y pueden cambiar sin aviso; si una tienda falla, se registra y se sigue con las demás.

Uso (correr una vez al día, idealmente a la misma hora):
    python scripts/06_scraper_supermercados.py
Requisitos de red: el entorno debe permitir los dominios de TIENDAS.
"""
import csv
import datetime as dt
import json
import os
import re
import time
import unicodedata
import urllib.parse
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CANASTA = os.path.join(ROOT, "data", "nowcast", "canasta_supermercado.csv")
DESTINO = os.path.join(ROOT, "data", "nowcast", "supermercados")
UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"

# API de búsqueda de catálogo VTEX: ?ft=<texto>&_from=<i>&_to=<j> (máximo 50 por página)
TIENDAS = {
    "jumbo": "https://www.jumbo.cl/api/catalog_system/pub/products/search/",
    "santa_isabel": "https://www.santaisabel.cl/api/catalog_system/pub/products/search/",
    "unimarc": "https://www.unimarc.cl/api/catalog_system/pub/products/search/",
}
MAX_ITEMS = 50


def normalizar(s):
    s = unicodedata.normalize("NFKD", s.lower())
    return "".join(c for c in s if not unicodedata.combining(c))


def buscar_vtex(url_base, termino):
    q = urllib.parse.urlencode({"ft": termino, "_from": 0, "_to": MAX_ITEMS - 1})
    req = urllib.request.Request(f"{url_base}?{q}", headers={"User-Agent": UA, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as r:
        datos = json.loads(r.read().decode("utf-8"))
    for p in datos:
        for item in p.get("items", []):
            oferta = (item.get("sellers") or [{}])[0].get("commertialOffer", {})
            if not oferta.get("Price") or not oferta.get("IsAvailable", True):
                continue
            yield {"sku": item.get("itemId"), "nombre": p.get("productName", ""),
                   "precio": oferta["Price"], "precio_lista": oferta.get("ListPrice"),
                   "unidad": item.get("measurementUnit", "")}


def main():
    hoy = dt.date.today().isoformat()
    os.makedirs(DESTINO, exist_ok=True)
    canasta = list(csv.DictReader(open(CANASTA, encoding="utf-8")))
    filas, errores = [], []
    for tienda, url in TIENDAS.items():
        try:  # prueba de conexión: si la tienda no responde se salta completa
            next(buscar_vtex(url, "leche"), None)
        except Exception as e:  # noqa: BLE001
            errores.append(f"{tienda}: sin conexión ({e})")
            continue
        for prod in canasta:
            filtro = normalizar(prod["debe_contener"])
            for termino in prod["termino_busqueda"].split("|"):
                try:
                    for it in buscar_vtex(url, termino):
                        if filtro and filtro not in normalizar(it["nombre"]):
                            continue
                        filas.append({"fecha": hoy, "tienda": tienda, "codigo_ipc": prod["codigo_ipc"], **it})
                except Exception as e:  # noqa: BLE001 - se registra y se sigue con la siguiente búsqueda
                    errores.append(f"{tienda} / {termino}: {e}")
                time.sleep(0.5)
    # un SKU puede aparecer en varias búsquedas del mismo producto
    vistos, unicas = set(), []
    for f in filas:
        k = (f["tienda"], f["codigo_ipc"], f["sku"])
        if k not in vistos:
            vistos.add(k)
            unicas.append(f)
    if unicas:
        with open(os.path.join(DESTINO, f"{hoy}.csv"), "w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=list(unicas[0]))
            w.writeheader()
            w.writerows(unicas)
    print(f"{len(unicas)} precios guardados; {len(errores)} búsquedas con error")
    for e in errores[:10]:
        print("  !", re.sub(r"\s+", " ", e)[:160])


if __name__ == "__main__":
    main()
