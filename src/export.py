# PATH: src/export.py

import pandas as pd
from datetime import datetime
from pathlib import Path

def crear_output_dir(base_path: Path) -> Path:
    """Crea una carpeta en ./archivos/output/output_{timestamp}/ y la devuelve."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = base_path / "output" / f"output_{timestamp}"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir

def exportar_dataframes(output_dir: Path, archivos: dict[str, pd.DataFrame]) -> None:
    """Exporta cada DataFrame a un Excel dentro del directorio indicado."""
    for nombre, df in archivos.items():
        ruta_salida = output_dir / f"{nombre}.xlsx"
        df.to_excel(ruta_salida, index=False)
