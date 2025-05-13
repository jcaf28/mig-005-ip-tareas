# PATH: src/transform.py

import pandas as pd, re
from datetime import datetime

RENAME = {
    "FECHA": "fecha",
    "PERSONA\n(Nº de chapa)": "idusuario_raw",
    "PROYECTO": "proyecto",
    "ACTIVIDAD/TAREA": "actividad",
    "CODIGO PLANO": "nplano",
    "HORAS": "choras",
    "OBSERVACIONES": "obs",
}

def normaliza_columnas(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns=RENAME)

def limpia_idusuario(df: pd.DataFrame) -> pd.DataFrame:
    df["idusuario"] = (df["idusuario_raw"]
                       .astype(str)
                       .str.replace(r"\.0$", "", regex=True)
                       .str.strip())
    return df

def desglosa_proyecto(df: pd.DataFrame) -> pd.DataFrame:
    pat = re.compile(r"\((\d+)\)")
    df["proyecto_codigo"] = df["proyecto"].apply(
        lambda s: pat.search(str(s)).group(1)
        if pat.search(str(s)) else str(s).strip())
    df["proyecto_nombre"] = df["proyecto"].apply(
        lambda s: pat.sub("", str(s)).strip(" -"))
    return df

def analizar_cargado_a(base: pd.DataFrame,
                       obras_exist: pd.DataFrame,
                       obras_subir: pd.DataFrame) -> pd.DataFrame:
    """Detecta valores únicos de 'CARGADO A' no presentes ni en obras existentes ni en las que ya vamos a subir."""
    cargados = base["CARGADO A"].dropna().astype(str).str.strip().unique()
    obras_existentes = obras_exist["ClaveObra"].astype(str)
    obras_a_subir = obras_subir["ClaveObra"].astype(str)

    faltantes = [v for v in cargados
                 if v not in obras_existentes.values
                 and v not in obras_a_subir.values]

    df_faltantes = pd.DataFrame({"ClaveObra": faltantes,
                                 "NomObra": [""] * len(faltantes)})
    return df_faltantes


def tablas_auxiliares(base: pd.DataFrame,
                      usuarios: pd.DataFrame,
                      obras: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    faltan_u = base.loc[~base["idusuario"]
                        .isin(usuarios["IdUsuario"])].copy()
    usuarios_subir = (faltan_u[["idusuario"]]
                      .drop_duplicates()
                      .assign(NomUsuario="",
                              ClaveUsuario="",
                              PagaHE=""))
    faltan_o = base.loc[~base["proyecto_codigo"]
                        .isin(obras["ClaveObra"])].copy()
    obras_subir = (faltan_o[["proyecto_codigo", "proyecto_nombre"]]
                   .drop_duplicates()
                   .rename(columns={"proyecto_codigo": "ClaveObra",
                                    "proyecto_nombre": "NomObra"}))
    return usuarios_subir, obras_subir

def _clave_obra(proyecto: str) -> str:
   pat = re.compile(r"\((\d+)\)")
   if pat.search(proyecto):
      return pat.search(proyecto).group(1)
   return proyecto.strip()

def _aplicar_maestro(df: pd.DataFrame,
                     maestro: pd.DataFrame) -> pd.DataFrame:
   maestro = maestro.rename(columns=str.strip)
   borrar = maestro.loc[maestro["CambiarAObra"].astype(str).str.lower()
                        .eq("borrar"), "ClaveObra"].astype(str)
   df = df.loc[~df["proyecto_codigo"].isin(borrar)].copy()
   cambios = maestro.dropna(subset=["CambiarAObra"])
   mapa = dict(zip(cambios["ClaveObra"].astype(str),
                   cambios["CambiarAObra"].astype(str)))
   df["ClaveObra"] = df["proyecto_codigo"].replace(mapa)
   return df

def preparar_anotaciones(base: pd.DataFrame,
                         usuarios: pd.DataFrame,
                         maestro: pd.DataFrame) -> pd.DataFrame:
   base = _aplicar_maestro(base, maestro)

   base["ClaveObra"] = base["ClaveObra"].fillna(base["proyecto_codigo"])
   base["CHoras"] = pd.to_numeric(base["choras"], errors="coerce")
   ahora = datetime.now()

   anot = pd.DataFrame({
       "IdAnot": range(47000, 47000 + len(base)),
       "Idusuario": base["idusuario"],
       "FAnotacion": pd.to_datetime(base["fecha"]).dt.strftime("%d/%m/%Y"),
       "ClaveObra": base["ClaveObra"],
       "CodObra": "",
       "IdProceso": "",
       "CodTarea": "",
       "DescAnot": base["obs"].fillna(""),
       "CEuros": "",
       "CHoras": base["CHoras"],
       "NPlano": base["nplano"].fillna(""),
       "FCREA": ahora,
       "FMODIFI": ahora,
       "IdUsuarioC": base["idusuario"],
       "IdTipo": base["idusuario"].map(dict(zip(usuarios["IdUsuario"],
                                               usuarios["PagaHE"]))),
       "TasaHora": 80,
       "NumModOT": ""
   })

   return anot
