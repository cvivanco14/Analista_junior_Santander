"""Descarga las planillas mensuales de la Encuesta de Expectativas Económicas (EEE).

El sitio del Banco Central está detrás de Incapsula, que bloquea descargas con
curl/requests. Por eso se usa Chromium (Playwright): se abre la página de
resultados anteriores y cada planilla se descarga como lo haría un navegador.

Uso:
    python eee/descargar.py            # descarga lo que falte en eee/raw/
    python eee/descargar.py --desde 2014
"""
import argparse
import os
import re
import time
import urllib.parse
from pathlib import Path

from playwright.sync_api import sync_playwright

URL = "https://www.bcentral.cl/encuestas-economicas/resultados-anteriores-eee"
RAW = Path(__file__).resolve().parent / "raw"
CHROMIUM = os.environ.get("CHROMIUM_PATH", "/opt/pw-browsers/chromium")
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36")

# resultado012014.xls, resultado_012021.xlsx, resultado_12017.xls,
# "Datos mensuales enero 2026.xlsx"
PATRON = re.compile(r"(?i)^(resultado_?\d{5,6}|datos mensuales \w+ \d{4})\.xlsx?$")


def nombre_archivo(href):
    return urllib.parse.unquote_plus(urllib.parse.urlparse(href).path.split("/")[4])


def anio(nombre):
    return int(re.search(r"(\d{4})\.xlsx?$", nombre).group(1))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--desde", type=int, default=2014, help="primer año a descargar")
    args = ap.parse_args()
    RAW.mkdir(exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path=CHROMIUM if Path(CHROMIUM).exists() else None,
                                    args=["--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(accept_downloads=True, user_agent=UA)
        page = ctx.new_page()
        page.goto(URL, wait_until="networkidle", timeout=90_000)

        hrefs = page.eval_on_selector_all("a[href*='/documents/']", "els => els.map(e => e.href)")
        planillas = {}
        for h in hrefs:
            n = nombre_archivo(h)
            if PATRON.match(n) and anio(n) >= args.desde:
                planillas[n] = h
        print(f"{len(planillas)} planillas en la página")

        for n, h in sorted(planillas.items()):
            destino = RAW / n
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
