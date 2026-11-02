# IMCE: proceso mensual

`imce_principal.py` actualiza `IMCE.xlsx` desde la API del BCCh y genera el chartbook en PPTX
(gráficos nativos editables) y PDF. No define formato propio: usa `formato_graficos.py`
(en este repositorio, [`../formato/`](../formato/)), igual que el chartpack de Expectativas.

## Carpetas en SharePoint (sirven para cualquier usuario)

| Qué | Dónde |
|---|---|
| Este script | `...\Estudios - Documentos\General\Expectativas y Encuestas\IMCE\Code` |
| `formato_graficos.py` + `assets\` | `...\Estudios - Documentos\General\Administrativos\Manual de Estilo` |

- El script ubica el Manual de Estilo subiendo desde su propia carpeta; si no lo encuentra,
  lo busca en la carpeta sincronizada `Santander Office 365` del usuario que lo corre. No usa el
  nombre de usuario (`C:\Users\<usuario>`), así que funciona para cualquiera.
  `FORMATO_GRAFICOS_DIR` permite forzar otra carpeta.
No hay que editar nada:

- Credenciales desde variables de entorno (`BDE_TOKEN`/`BCCH_API_token`, o `BCCH_USER` y
  `BCCH_PASS`); nunca escritas en el código.
- Busca `IMCE.xlsx` primero en la carpeta `IMCE` del proyecto (`IMCE\`, `IMCE\Code\` y sus
  subcarpetas, por ejemplo `IMCE\Base\`) y luego en la ubicación anterior
  (`.../Bases - Chartpack/Bases/Local/IMCE/Base/IMCE.xlsx`). `IMCE_XLSX` permite forzar otra ruta.
- El PPTX y el PDF quedan en `tmp_chartbook_imce_santander/`, junto a `IMCE.xlsx`. El PDF lo
  exporta PowerPoint en Windows (con `pywin32`) o LibreOffice si está instalado.

```powershell
python imce/imce_principal.py
```
