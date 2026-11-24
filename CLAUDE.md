# Instrucciones para Claude

- Trabajar y hacer push directamente en `main` (no crear ramas ni PRs salvo que se pida).
- Cada minuta/proyecto va en su propia carpeta con `Code/`, `Inputs/` y `Outputs/`, rutas
  relativas al script con `Path(__file__)` (nada de rutas por usuario de Windows/OneDrive).
- La Minuta IPC vive en `minuta_ipc/` (script principal: `minuta_ipc/Code/minuta_ipc.py`).
- La Minuta de Remuneraciones vive en `remuneraciones/` (un solo script: `remuneraciones/Code/minuta_remuneraciones.py`, que también descarga los datos).
- La Minuta IMCE vive en `imce/` (lanzador `imce/run_minuta_imce.py`; scripts `imce/Code/minuta_imce.py` y `datos_imce.py`).
- El Chartpack de Expectativas vive en `chartpack_expectativas/` (lanzador `chartpack_expectativas/run_chartpack_expectativas.py`; `Code/encuestas_bcch.py` descarga EEE/EOF; `Code/build_chartpack.py` arma el chartpack y actualiza `Chartpack_Expectativas.xlsx`).
- La Proyección del IPC vive en `proyeccion_ipc/` (script único: `proyeccion_ipc/Code/proyeccion_ipc.py`).
- El módulo común `formato_graficos.py` vive en `formato/` (versión vigente; subir `VERSION` al cambiarlo y entregarlo para copiarlo a `00_Administrativos\Manual de Estilo` en SharePoint).
- Al terminar, entregar también el código para descargar.
