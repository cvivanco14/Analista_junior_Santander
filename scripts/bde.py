"""Cliente mínimo para la API REST de la Base de Datos Estadísticos (BDE) del Banco Central de Chile.
Credenciales vía variables de entorno BCCH_USER y BCCH_PASS."""
import os, json, urllib.request, urllib.parse

URL = "https://si3.bcentral.cl/SieteRestWS/SieteRestWS.ashx"

def _get(params):
    p = {"user": os.environ["BCCH_USER"], "pass": os.environ["BCCH_PASS"], **params}
    with urllib.request.urlopen(URL + "?" + urllib.parse.urlencode(p), timeout=60) as r:
        return json.loads(r.read().decode("latin-1"))

def get_series(code, first=None, last=None):
    params = {"function": "GetSeries", "timeseries": code}
    if first: params["firstdate"] = first
    if last: params["lastdate"] = last
    return _get(params)

def search_series(frequency="MONTHLY"):
    return _get({"function": "SearchSeries", "frequency": frequency})
