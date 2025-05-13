# PATH: migracion.py

"""Punto de entrada de la migración IP-Tareas."""

from src.config import ARCHIVOS
from src.extract import (
    load_tablas_bd,
    load_historico,
    load_maestro_modificaciones,
)
from src.transform import (
    normaliza_columnas,
    limpia_idusuario,
    desglosa_proyecto,
    tablas_auxiliares,
    analizar_cargado_a,
    preparar_anotaciones,
)
from src.export import crear_output_dir, exportar_dataframes


def main() -> None:
    # ------------- EXTRACCIÓN -------------
    tablas_bd = load_tablas_bd(ARCHIVOS / "TABLAS_BD")
    base, asign = load_historico(
        ARCHIVOS / "20250512_Base datos historico IP Julen.xlsx"
    )
    maestro = load_maestro_modificaciones(
        ARCHIVOS / "T_OBRAS_SUBIR_MAESTRO_MODIFICACIONES.xlsx"
    )

    # ------------- TRANSFORMACIONES BÁSICAS -------------
    base = (
        base.pipe(normaliza_columnas)
        .pipe(limpia_idusuario)
        .pipe(desglosa_proyecto)
    )

    # Auxiliares (usuarios/obras que faltan) -------------------------------
    usuarios_subir, obras_subir = tablas_auxiliares(
        base, tablas_bd["usuarios"], tablas_bd["obras"]
    )

    # Obras que aparecen en «CARGADO A» y no existen todavía --------------
    cargado_a_subir = analizar_cargado_a(
        base, tablas_bd["obras"], obras_subir
    )

    # ------------- GENERACIÓN T_ANOTACIONES_SUBIR -------------
    anotaciones_subir = preparar_anotaciones(
        base=base,
        asignaciones=asign,
        usuarios_bd=tablas_bd["usuarios"],
        maestro_obras=maestro,
        tareas_bd=tablas_bd["tareas"],
        primer_id=47000,
    )

    # ------------- EXPORT -------------
    output_dir = crear_output_dir(ARCHIVOS)
    exportar_dataframes(
        output_dir,
        {
            "AUX_USUARIOS_SUBIR_DEBE_CONTENER": usuarios_subir,
            "AUX_T_OBRAS_SUBIR_DEBE_CONTENER": obras_subir,
            "AUX_T_OBRAS_CARGADO_A_SUBIR_DEBE_CONTENER": cargado_a_subir,
            "T_ANOTACIONES_SUBIR": anotaciones_subir,
            "BASE_PROCESADA": base,
        },
    )


if __name__ == "__main__":
    main()
