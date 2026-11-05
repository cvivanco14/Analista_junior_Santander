# IPC: proceso mensual (subclases y analíticos)

Misma estructura que IMCE y masa salarial:

| Carpeta | Contenido |
|---|---|
| `Code/` | `ipc_subclases.py` (todo el proceso) y, para el paso X-13, `x13as.exe` o `x13as_ascii.exe`. |
| `Inputs/` | Archivos descargados del INE (`ipc-xls.xlsx`, `analiticos-xls.xlsx`, `ipc-xls-ine-subclases.xlsx`; se vuelven a bajar si tienen más de 6 horas) y, opcional, `Base_IPC_plantilla.xlsx`. |
| `Outputs/` | `Base_IPC.xlsx` (se actualiza en cada corrida), `respaldo/` (copia antes de modificar), `Chartbook_SubClases_4panel.pdf`, `Chartbook_Analiticos_4panel.pdf` y los PNG de cada chartbook. |

```bash
cd ipc/Code
python ipc_subclases.py
```

- Las rutas se arman desde la ubicación del script (`Path(__file__)`), así que ya no depende del
  usuario de Windows ni de la ruta de OneDrive.
- Si `Outputs/Base_IPC.xlsx` no existe, se copia desde `Inputs/Base_IPC_plantilla.xlsx` o, si falta,
  desde la ubicación anterior `Bases - Chartpack\Bases\Local\IPC\Base_IPC.xlsx` en OneDrive/SharePoint.
- El paso X-13 necesita Windows con Excel (`xlwings`) y el ejecutable X-13 en `Code/`.
