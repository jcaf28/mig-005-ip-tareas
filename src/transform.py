# PATH: src/transform.py

import pandas as pd, re
from datetime import datetime

from src.utils.reglas_asterisco_tareas import asignar_tarea_asterisco

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
    
    # 0) Guardar los valores originales para depuración
    base["DBG_Proyecto_original"] = base["proyecto_codigo"].copy()
    if "CARGADO A" in base.columns:
        base["DBG_CargadoA_original"] = base["CARGADO A"].copy()

    # 1) Identificar obras a borrar
    borrar = maestro.loc[
        maestro["CambiarAObra"].astype(str).str.lower() == "borrar", "ClaveObra"
    ].astype(str).tolist()
    
    # 1.a) Eliminar filas donde proyecto_codigo está en la lista de borrar
    base = base.loc[~base["proyecto_codigo"].isin(borrar)].copy()
    
    # 1.b) Eliminar filas donde CARGADO A está en la lista de borrar
    if "CARGADO A" in base.columns:
        base = base.loc[~base["CARGADO A"].astype(str).isin(borrar)].copy()

    # 2) Cambios de clave obra
    cambios = maestro.dropna(subset=["CambiarAObra"])
    cambios = cambios[cambios["CambiarAObra"].astype(str).str.lower() != "borrar"]  # Excluir las obras a borrar
    mapa = dict(
        zip(
            cambios["ClaveObra"].astype(str),
            cambios["CambiarAObra"].astype(str),
        )
    )
    base["ClaveObra"] = base["proyecto_codigo"].replace(mapa)

    # Donde no haya coincidencia se mantiene la clave original
    base["ClaveObra"] = base["ClaveObra"].fillna(base["proyecto_codigo"])
    
    # 3) Aplicar las mismas reglas al campo "CARGADO A"
    if "CARGADO A" in base.columns:
        # Solo intentamos reemplazar si hay valores no nulos
        cargado_mask = base["CARGADO A"].notna()
        if cargado_mask.any():
            # Aplicamos el mapeo solo a las filas con valores no nulos
            base.loc[cargado_mask, "CARGADO A"] = (
                base.loc[cargado_mask, "CARGADO A"]
                .astype(str)
                .replace(mapa)
            )
    
    return base

# … import existentes …
from src.utils.reglas_asterisco_tareas import asignar_tarea_asterisco
# ↑ nueva importación


def _mapear_cod_tarea(
    base: pd.DataFrame,
    asignaciones: pd.DataFrame,
    tareas_bd: pd.DataFrame,
) -> pd.DataFrame:
    """
    Devuelve un DataFrame con:
        • CodTarea         → código final asignado
        • AsignarATarea    → valor bruto de la hoja *asignaciones_tareas*
    """
    asign = asignaciones.rename(columns=str.strip)
    asign["Tarea"] = asign["Tarea"].str.strip()
    asign["AsignarATarea"] = asign["AsignarATarea"].astype(str).str.strip()
    mapa_asign = dict(zip(asign["Tarea"], asign["AsignarATarea"]))

    tareas_validas = set(tareas_bd["CodTarea"].astype(str))

    cod_out, asign_out = [], []

    for _, row in base.iterrows():
        act_original  = str(row["actividad"]).strip()
        asignacion    = mapa_asign.get(act_original, "")      # puede venir vacío
        chapa         = str(row["idusuario"]).strip()
        categoria     = str(row.get("CATEGORIA", "")).strip().lower()

        # ---------------- resolución -----------------
        if asignacion == "*":
            codigo = (
                asignar_tarea_asterisco(
                    actividad=act_original,
                    chapa=chapa,
                    categoria=categoria,
                )
                or "PEND_ASIGNACION_*"
            )
        elif asignacion == "#ESPECIAL#":
            codigo = ""
        else:
            codigo = asignacion

        # --- verificar que exista en T_TAREAS --------------
        if (
            codigo
            and not codigo.startswith("PEND_")
            and codigo not in tareas_validas
        ):
            codigo = ""        # marcar para revisión

        cod_out.append(codigo)
        asign_out.append(asignacion)

    return pd.DataFrame({"CodTarea": cod_out, "AsignarATarea": asign_out})


# ------------------------------------------------------------------ #
# ----------------  CONSTRUCCIÓN DEL DATAFRAME FINAL  -------------- #
# ------------------------------------------------------------------ #
# ------------------------------------------------------------------ #
# ----------------  CONSTRUCCIÓN DEL DATAFRAME FINAL  -------------- #
# ------------------------------------------------------------------ #
def preparar_anotaciones(
    *,
    base: pd.DataFrame,
    asignaciones: pd.DataFrame,
    usuarios_bd: pd.DataFrame,
    maestro_obras: pd.DataFrame,
    tareas_bd: pd.DataFrame,
    primer_id: int = 47000,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Devuelve una tupla:
        • T_ANOTACIONES_SUBIR (DataFrame listo para Excel)
        • BASE_PROCESADA      (mismo contenido que *base* pero con la
          columna IdAnot añadida para trazabilidad)
    """
    # --- ajustes de obra (borrar / renombrar claves)
    base = _aplicar_maestro(base, maestro_obras).copy()

    # --- numérico a CHoras
    base["CHoras"] = pd.to_numeric(base["choras"], errors="coerce")

    # --- mapeo de tarea + columnas debug
    tarea_df = _mapear_cod_tarea(base, asignaciones, tareas_bd)
    base = pd.concat([base.reset_index(drop=True), tarea_df], axis=1)

    # ---------- IdAnot para ambas tablas ---------------
    ids = list(range(primer_id, primer_id + len(base)))
    base["IdAnot"] = ids          # ← trazabilidad

    ahora   = datetime.now()
    paga_he = dict(zip(usuarios_bd["IdUsuario"], usuarios_bd["PagaHE"]))

    anot = pd.DataFrame(
        {
            "IdAnot": ids,
            "Idusuario": base["idusuario"],
            "FAnotacion": pd.to_datetime(base["fecha"]).dt.strftime("%d/%m/%Y"),
            "ClaveObra": base["ClaveObra"],
            "CodObra": "",
            "IdProceso": "",                    # se rellenará más adelante
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
            # --- columnas de depuración ---
            "DBG_TareaOriginal": base["actividad"],
            "DBG_AsignarATarea": base["AsignarATarea"]
        }
    )

    return anot, base   # ← devolvemos ambos

def preparar_anotaciones_valid(
    anotaciones_subir: pd.DataFrame,
    base_trazada: pd.DataFrame,
    fecha_validacion: str = "15/05/2025",
    id_usuario_validacion: str = "18287"
) -> pd.DataFrame:
    """
    Genera la tabla T_ANOTACIONES_VALID_SUBIR a partir de T_ANOTACIONES_SUBIR.
    
    Args:
        anotaciones_subir: DataFrame con las anotaciones a insertar
        base_trazada: DataFrame con los datos originales (incluye IdAnot y CARGADO A)
        fecha_validacion: Fecha de validación (formato dd/mm/yyyy)
        id_usuario_validacion: ID del usuario que realiza la validación
        
    Returns:
        DataFrame con la estructura de T_ANOTACIONES_VALID
    """
    # Crear un diccionario para mapear IdAnot a CARGADO A
    id_to_cargado = dict(zip(base_trazada["IdAnot"], base_trazada["CARGADO A"]))
    
    # Crear el DataFrame de anotaciones validadas
    valid = pd.DataFrame({
        "IdAnot": anotaciones_subir["IdAnot"],
        "ClaveObra": [id_to_cargado.get(id_anot, "") for id_anot in anotaciones_subir["IdAnot"]],
        "IdTipoV": "H",
        "FValid": fecha_validacion,
        "VEuros": 0,
        "VHoras": anotaciones_subir["CHoras"],
        "FCREAV": fecha_validacion,
        "FMODIFIV": fecha_validacion,
        "IdUsuarioCV": id_usuario_validacion,
        "DctaHoras": "S"
    })
    
    # Asegurar que ClaveObra no tenga valores nulos (reemplazar por string vacío)
    valid["ClaveObra"] = valid["ClaveObra"].fillna("")
    
    return valid