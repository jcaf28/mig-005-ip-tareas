import pandas as pd, re
from datetime import datetime

# ---------------------------------------------------------------------- #
#  -------------------------  NORMALIZACIÓN  --------------------------- #
# ---------------------------------------------------------------------- #
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
    df["idusuario"] = (
        df["idusuario_raw"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    )
    return df


def desglosa_proyecto(df: pd.DataFrame) -> pd.DataFrame:
    pat = re.compile(r"\((\d+)\)")
    df["proyecto_codigo"] = df["proyecto"].apply(
        lambda s: pat.search(str(s)).group(1) if pat.search(str(s)) else str(s).strip()
    )
    df["proyecto_nombre"] = df["proyecto"].apply(
        lambda s: pat.sub("", str(s)).strip(" -")
    )
    return df


# ---------------------------------------------------------------------- #
#  -------------------------  AUXILIARES  ------------------------------ #
# ---------------------------------------------------------------------- #
def analizar_cargado_a(
    base: pd.DataFrame, obras_exist: pd.DataFrame, obras_subir: pd.DataFrame
) -> pd.DataFrame:
    """Detecta valores únicos de 'CARGADO A' no presentes ni en obras existentes ni en las que ya vamos a subir."""
    cargados = base["CARGADO A"].dropna().astype(str).str.strip().unique()
    obras_existentes = obras_exist["ClaveObra"].astype(str)
    obras_a_subir = obras_subir["ClaveObra"].astype(str)

    faltantes = [
        v
        for v in cargados
        if v not in obras_existentes.values and v not in obras_a_subir.values
    ]

    return pd.DataFrame({"ClaveObra": faltantes, "NomObra": [""] * len(faltantes)})


def tablas_auxiliares(
    base: pd.DataFrame, usuarios: pd.DataFrame, obras: pd.DataFrame
) -> tuple[pd.DataFrame, pd.DataFrame]:
    # ----- Usuarios
    faltan_u = base.loc[~base["idusuario"].isin(usuarios["IdUsuario"])].copy()
    usuarios_subir = (
        faltan_u[["idusuario"]]
        .drop_duplicates()
        .assign(NomUsuario="", ClaveUsuario="", PagaHE="")
    )

    # ----- Obras
    faltan_o = base.loc[~base["proyecto_codigo"].isin(obras["ClaveObra"])].copy()
    obras_subir = (
        faltan_o[["proyecto_codigo", "proyecto_nombre"]]
        .drop_duplicates()
        .rename(columns={"proyecto_codigo": "ClaveObra", "proyecto_nombre": "NomObra"})
    )
    return usuarios_subir, obras_subir


# ---------------------------------------------------------------------- #
#  --------------------  PREPARACIÓN ANOTACIONES  ---------------------- #
# ---------------------------------------------------------------------- #
def _aplicar_maestro(base: pd.DataFrame, maestro: pd.DataFrame) -> pd.DataFrame:
    """Aplica las reglas del maestro de modificaciones de obra."""
    maestro = maestro.rename(columns=str.strip)

    # 1) Filas que hay que eliminar completamente
    borrar = maestro.loc[
        maestro["CambiarAObra"].astype(str).str.lower() == "borrar", "ClaveObra"
    ].astype(str)
    base = base.loc[~base["proyecto_codigo"].isin(borrar)].copy()

    # 2) Cambios de clave obra
    cambios = maestro.dropna(subset=["CambiarAObra"])
    mapa = dict(
        zip(
            cambios["ClaveObra"].astype(str),
            cambios["CambiarAObra"].astype(str),
        )
    )
    base["ClaveObra"] = base["proyecto_codigo"].replace(mapa)

    # Donde no haya coincidencia se mantiene la clave original
    base["ClaveObra"] = base["ClaveObra"].fillna(base["proyecto_codigo"])
    return base


def _mapear_cod_tarea(
    base: pd.DataFrame, asignaciones: pd.DataFrame, tareas_bd: pd.DataFrame
) -> pd.Series:
    """
    Devuelve una Serie con el código de tarea final según la hoja *asignaciones_tareas*.

    - Si `AsignarATarea` aparece directamente en BD → se usa.
    - Si es '*'      → se deja texto placeholder.
    - Si es '#ESPECIAL#' → se deja en blanco (pendiente).
    - Si no existe en T_TAREAS se deja en blanco para revisión.
    """
    asign = asignaciones.rename(columns=str.strip)
    asign["Tarea"] = asign["Tarea"].str.strip()
    asign["AsignarATarea"] = asign["AsignarATarea"].astype(str).str.strip()

    # diccionario tarea original -> asignación
    mapa = dict(zip(asign["Tarea"], asign["AsignarATarea"]))

    # Serie resultante
    cod_tarea = []

    for act in base["actividad"]:
        act = str(act).strip()
        if act in mapa:
            dest = mapa[act]

            if dest == "*":
                cod_tarea.append("PEND_ASIGNACION_*")  # placeholder
            elif dest == "#ESPECIAL#":
                cod_tarea.append("")  # se resolverá más adelante
            else:
                cod_tarea.append(dest)
        else:
            cod_tarea.append("")  # no encontrado, revisar

    # Verificamos que el código exista en T_TAREAS; si no, vaciamos
    existentes = set(tareas_bd["CodTarea"].astype(str))
    cod_tarea = pd.Series(cod_tarea).apply(
        lambda x: x if x in existentes or x.startswith("PEND_") else ""
    )
    return cod_tarea


def preparar_anotaciones(
    *,
    base: pd.DataFrame,
    asignaciones: pd.DataFrame,
    usuarios_bd: pd.DataFrame,
    maestro_obras: pd.DataFrame,
    tareas_bd: pd.DataFrame,
    primer_id: int = 47000,
) -> pd.DataFrame:
    """
    A partir de *base* genera el DataFrame que se exportará como
    **T_ANOTACIONES_SUBIR.xlsx** (solo campos completados hasta el momento).
    """
    # Aplicar reglas de obra
    base = _aplicar_maestro(base, maestro_obras).copy()

    # Conversión horas
    base["CHoras"] = pd.to_numeric(base["choras"], errors="coerce")

    # CodTarea provisional
    base["CodTarea"] = _mapear_cod_tarea(base, asignaciones, tareas_bd)

    # Fecha/hora actual
    ahora = datetime.now()

    # Mapeo IdTipo (PagaHE)
    paga_he = dict(zip(usuarios_bd["IdUsuario"], usuarios_bd["PagaHE"]))

    # Construcción de la tabla final
    anot = pd.DataFrame(
        {
            "IdAnot": range(primer_id, primer_id + len(base)),
            "Idusuario": base["idusuario"],
            "FAnotacion": pd.to_datetime(base["fecha"]).dt.strftime("%d/%m/%Y"),
            "ClaveObra": base["ClaveObra"],
            "CodObra": "",
            "IdProceso": "",  # pendiente de reglas
            "CodTarea": base["CodTarea"],
            "DescAnot": base["obs"].fillna(""),
            "CEuros": "",
            "CHoras": base["CHoras"],
            "NPlano": base["nplano"].fillna(""),
            "FCREA": ahora,
            "FMODIFI": ahora,
            "IdUsuarioC": base["idusuario"],
            "IdTipo": base["idusuario"].map(paga_he),
            "TasaHora": 80,
            "NumModOT": "",
        }
    )

    return anot
