# IMCE: proceso mensual

`imce_principal.py` actualiza `IMCE.xlsx` desde la API del BCCh y genera el chartbook PDF.
No hay que editar nada:

- Credenciales desde variables de entorno (`BDE_TOKEN`/`BCCH_API_token`, o `BCCH_USER` y
  `BCCH_PASS`); nunca escritas en el código.
- Busca `IMCE.xlsx` en las carpetas de OneDrive/SharePoint de Santander de cualquier usuario
  (`.../Bases - Chartpack/Bases/Local/IMCE/Base/IMCE.xlsx`); si no está, usa uno junto al script,
  y `IMCE_XLSX` permite forzar otra ruta.
- El PDF queda en `tmp_chartbook_imce_santander/`, junto a `IMCE.xlsx`.

```powershell
python imce/imce_principal.py
```
