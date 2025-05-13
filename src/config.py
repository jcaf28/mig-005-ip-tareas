# PATH: src/config.py

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
ARCHIVOS = BASE_DIR / "archivos"

# Dev / prod con un solo flag externo
ENV = "dev"                  # cambiar a "prod" en despliegue
