from src.config import ARCHIVOS
from src.extract import (load_tablas_bd, load_historico,
                         load_maestro_modificaciones)
from src.transform import (normaliza_columnas, limpia_idusuario,
                           desglosa_proyecto, tablas_auxiliares,
                           analizar_cargado_a, preparar_anotaciones)
from src.export import crear_output_dir, exportar_dataframes

def main() -> None:
   tablas = load_tablas_bd(ARCHIVOS / "TABLAS_BD")
   base, asign = load_historico(ARCHIVOS / "20250512_Base datos historico IP Julen.xlsx")
   maestro = load_maestro_modificaciones(ARCHIVOS / "T_OBRAS_SUBIR_MAESTRO_MODIFICACIONES.xlsx")

   base = (base
           .pipe(normaliza_columnas)
           .pipe(limpia_idusuario)
           .pipe(desglosa_proyecto))

   usuarios_subir, obras_subir = tablas_auxiliares(base,
                                                   tablas["usuarios"],
                                                   tablas["obras"])
   cargado_a_subir = analizar_cargado_a(base,
                                        tablas["obras"],
                                        obras_subir)

   anotaciones_subir = preparar_anotaciones(base,
                                            tablas["usuarios"],
                                            maestro)

   output_dir = crear_output_dir(ARCHIVOS)
   exportar_dataframes(output_dir, {
       "AUX_USUARIOS_SUBIR_DEBE_CONTENER": usuarios_subir,
       "AUX_T_OBRAS_SUBIR_DEBE_CONTENER": obras_subir,
       "AUX_T_OBRAS_CARGADO_A_SUBIR_DEBE_CONTENER": cargado_a_subir,
       "T_ANOTACIONES_SUBIR": anotaciones_subir,
       "BASE_PROCESADA": base
   })

if __name__ == "__main__":
   main()
