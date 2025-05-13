# PATH: src/extract.py

import pandas as pd
from pathlib import Path


def load_tablas_bd(carpeta: Path) -> dict:
    """Carga T_USUARIOS, T_OBRAS, T_PROCESOS y T_TAREAS en DataFrames."""
    return {
        "usuarios": pd.read_excel(carpeta / "T_USUARIOS.xlsx"),
        "obras": pd.read_excel(carpeta / "T_OBRAS.xlsx"),
        "procesos": pd.read_excel(carpeta / "T_PROCESOS.xlsx"),
        "tareas": pd.read_excel(carpeta / "T_TAREAS.xlsx"),
    }


def load_historico(fichero: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Carga las hojas *Base Datos* y *asignaciones_tareas*."""
    base = pd.read_excel(fichero, sheet_name="Base Datos")
    asign = pd.read_excel(fichero, sheet_name="asignaciones_tareas")
    return base, asign


def load_maestro_modificaciones(ruta: Path) -> pd.DataFrame:
    """Lee T_OBRAS_SUBIR_MAESTRO_MODIFICACIONES.xlsx."""
    return pd.read_excel(ruta)
