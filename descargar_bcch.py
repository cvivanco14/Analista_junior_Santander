"""Descarga las planillas de resultados de encuestas del Banco Central (EEE y EOF).

El sitio del Banco Central está detrás de Incapsula, que bloquea descargas con
curl/requests. Por eso se usa Chromium (Playwright): se abre la página de
resultados anteriores y cada planilla se descarga como lo haría un navegador.

Uso:
    python descargar_bcch.py eee            # descarga lo que falte en eee/raw/
    python descargar_bcch.py eof --desde 2010
"""
import argparse
import os
import re
import time
import urllib.parse
from pathlib import Path

from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent
CHROMIUM = os.environ.get("CHROMIUM_PATH", "/opt/pw-browsers/chromium")
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36")

ENCUESTAS = {
    # resultado012014.xls, resultado_012021.xlsx, resultado_12017.xls,
    # "Datos mensuales enero 2026.xlsx"
    "eee": ("https://www.bcentral.cl/encuestas-economicas/resultados-anteriores-eee",
            re.compile(r"(?i)^(resultado_?\d{5,6}|datos mensuales \w+ \d{4})\.xlsx?$")),
    # EOFResultado20100115.xls, EOFRESULTADOS_201801q1.xlsx,
    # EOF_RESULTADOS_Pre_RPM_201905.xlsx, "EOF RESULTADOS Post RPM Enero 2026.xlsx"
    "eof": ("https://www.bcentral.cl/encuestas-economicas/resultados-anteriores-eof",
            re.compile(r"(?i)^eof.*\.xls[xm]?$")),
}


def nombre_archivo(href):
    return urllib.parse.unquote_plus(urllib.parse.urlparse(href).path.split("/")[4])


def anio(nombre):
    m = re.search(r"(20\d\d)", nombre)
    return int(m.group(1)) if m else 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("encuesta", choices=ENCUESTAS)
    ap.add_argument("--desde", type=int, default=2010, help="primer año a descargar")
    args = ap.parse_args()
    url, patron = ENCUESTAS[args.encuesta]
    raw = BASE / args.encuesta / "raw"
    raw.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path=CHROMIUM if Path(CHROMIUM).exists() else None,
                                    args=["--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(accept_downloads=True, user_agent=UA)
        page = ctx.new_page()
        page.goto(url, wait_until="networkidle", timeout=90_000)

        hrefs = page.eval_on_selector_all("a[href*='/documents/']", "els => els.map(e => e.href)")
        planillas = {}
        for h in hrefs:
            n = nombre_archivo(h)
            if patron.match(n) and anio(n) >= args.desde:
                planillas[n] = h
        print(f"{len(planillas)} planillas en la página")

        for n, h in sorted(planillas.items()):
            destino = raw / n
            if destino.exists():
                continue
            try:
                with page.expect_download(timeout=60_000) as d:
                    page.evaluate("""(u) => { const a = document.createElement('a');
                        a.href = u; a.download = ''; document.body.appendChild(a); a.click(); a.remove(); }""", h)
                d.value.save_as(destino)
                if destino.stat().st_size < 10_000:  # página de bloqueo, no un Excel
                    destino.unlink()
                    print("bloqueado:", n)
                else:
                    print("ok:", n)
            except Exception as e:  # noqa: BLE001
                print("error:", n, e)
            time.sleep(1.5)
        browser.close()


if __name__ == "__main__":
    main()
