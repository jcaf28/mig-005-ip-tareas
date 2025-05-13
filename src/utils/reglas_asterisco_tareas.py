# PATH: src/utils/reglas_asterisco_tareas.py

"""
Asignación de CodTarea cuando en la hoja *asignaciones_tareas* aparece un '*'.

La función `asignar_tarea_asterisco()` recibe:
    - actividad  -> str  (valor de la columna 'actividad')
    - chapa      -> str  (idusuario ya limpio, sin decimales)
    - categoria  -> str  ('gg' / 'procesos' / 'utillajes' / '')

Devuelve:
    - str con el código de tarea asignado  **o**  '' (cadena vacía) si la fila
      debe descartarse / no computa / queda pendiente.
"""

from __future__ import annotations

from typing import Callable

# --------------------------------------------------------------------- #
#  -------------------------  REGLAS BÁSICAS  ------------------------- #
# --------------------------------------------------------------------- #

# Para facilitar la lectura:
PROC = "procesos"
UTI = "utillajes"
GG = "gg"

# Conjuntos de chapas reutilizados en varias reglas
CHAPAS = {
    "UE81": {"10168", "11773", "12591"},
    "A81": {"12705", "14031", "14272"},
    "E81": {"13007", "13831", "13835", "12578"},
    #
    "UEVAR01": {"10168", "11773", "13877", "16276"},
    "A91": {"12705", "13144", "10705", "14031", "14272"},
    "E91": {"13007", "11382", "11782", "12578"},
    #
    "AGG01": {"10705", "12705", "13144", "14031", "14272"},
    "EGG01": {"11382", "11782", "13007", "13788"},
}

# --------------------------------------------------------------------- #
#  ---------------------------  REGLAS  -------------------------------- #
# --------------------------------------------------------------------- #

def _ips(chapa: str, categoria: str) -> str:
    if categoria == PROC:
        if chapa in CHAPAS["UE81"]:
            return "UE81"
        if chapa in CHAPAS["A81"]:
            return "A81"
        if chapa in CHAPAS["E81"]:
            return "E81"
    elif categoria == UTI:
        if chapa == "12591":
            return "UE81"
        if chapa in {"11296", "12320"}:
            return "UT80"
    # GG o cualquier otro → descartar
    return ""


def _ipa(chapa: str, categoria: str) -> str:
    if categoria == PROC:
        if chapa in CHAPAS["UEVAR01"]:
            return "UEVAR01"
        if chapa in CHAPAS["A91"]:
            return "A91"
        if chapa in CHAPAS["E91"]:
            return "E91"
    elif categoria == UTI:
        return "UTVAR01"
    return ""


def _i20(chapa: str, categoria: str) -> str:
    if categoria == PROC:
        return "A39"
    if categoria == UTI:
        return "UT31"
    return ""


def _i30(chapa: str, categoria: str) -> str:
    if categoria == PROC:
        return "UE30"
    # “Resto: descartar”
    return ""


def _u40(chapa: str, categoria: str) -> str:
    if categoria == PROC:
        return "UE64"
    if categoria == UTI:
        return "UE64" if chapa == "12591" else "UT64"
    return ""


def _u80(_, categoria: str) -> str:
    return "UE70" if categoria == PROC else ("UT30" if categoria == UTI else "")


def _u21(chapa: str, categoria: str) -> str:
    if categoria == PROC:
        return "A20" if chapa == "12705" else "UE20"
    if categoria == UTI:
        return "UE20" if chapa == "12591" else "UT16"
    return ""


def _u20(chapa: str, categoria: str) -> str:
    # Misma lógica que U21
    return _u21(chapa, categoria)


def _f20(chapa: str, categoria: str) -> str:
    if categoria == UTI:
        if chapa == "11296":
            return "UT20"
        if chapa == "11780":
            return "UT21"
    if categoria == PROC:
        if chapa == "12705":
            return "A39"
        if chapa == "16276":
            return "UA11"
    return ""


def _i10(_, categoria: str) -> str:
    return "UE30" if categoria == PROC else ("UT30" if categoria == UTI else "")


def _f40(_, categoria: str) -> str:
    return "A39" if categoria == PROC else ("UT20" if categoria == UTI else "")


def _f30(_, categoria: str) -> str:
    return "UE12" if categoria == PROC else ("UT22" if categoria == UTI else "")


def _u50(_, categoria: str) -> str:
    return "UE50" if categoria == PROC else ("UT50" if categoria == UTI else "")


def _f10(_, categoria: str) -> str:
    return "UE10" if categoria == PROC else ("UT20" if categoria == UTI else "")


def _formacion(chapa: str, _categoria: str) -> str:
    if chapa in CHAPAS["AGG01"]:
        return "AGG01"
    if chapa in CHAPAS["EGG01"]:
        return "EGG01"
    return ""


def _uh(chapa: str, categoria: str) -> str:
    if categoria == PROC:
        if chapa == "12591":
            return "UE10"
        if chapa == "12705":
            return "A32"
    if categoria == UTI:
        return "UE10"
    return ""


def _um(_, categoria: str) -> str:
    return "A39" if categoria == PROC else ("UT42" if categoria == UTI else "")


def _uv(chapa: str, categoria: str) -> str:
    if categoria == PROC:
        return "A39" if chapa == "12705" else "UE10"
    if categoria == UTI:
        return "UE10" if chapa == "12591" else "UT40"
    return ""


def _varios(_, categoria: str) -> str:
    return "UTVAR01" if categoria == UTI else ("UEVAR01" if categoria == PROC else "")


# Mapeo actividad  ➜  función que decide la tarea
_RULES: dict[str, Callable[[str, str], str]] = {
    "IPS-GESTION SUBCONTRATACION": _ips,
    "IPA-ADMINISTRACION": _ipa,
    "I20-FASE IMPLANTACION UTILLAJES BOGIE ACABADO": _i20,
    "I30-FASE IMPLANTACION UTILLAJES CAJAS ESTRUCTURA": _i30,
    "U40-FASE MODIFICACIONES PRODUCTO": _u40,
    "U80-FASE CIERRE PROYECTO: Analisis Coste/Mejoras/Incidencias/Utillajes fin de obra": _u80,
    "U21-FASE COORDINACION: Analisis de planos y creación del listado de herramientas": _u21,
    "U20-FASE COORDINACION: Definir proceso fabr. + Reuniones IP/Fabr. + Acta + Informe mejora + Listado utillajes": _u20,
    "F20-FASE FABRICACION UTILLAJES BOGIE ACABADO": _f20,
    "I10-FASE IMPLANTACION UTILLAJES BOGIE ESTRUCTURA": _i10,
    "F40-FASE FABRICACION UTILLAJES CAJAS ACABADO": _f40,
    "F30-FASE FABRICACION UTILLAJES CAJAS ESTRUCTURA": _f30,
    "U50-FASE PRESERIE 2 INICIO FABRICACION": _u50,
    "F10-FASE FABRICACION UTILLAJES BOGIE ESTRUCTURA": _f10,
    "FORMACIÓN": _formacion,
    "UH-GESTION DE HERRAMIENTAS": _uh,
    "UM-MANTENIMIENTO": _um,
    "UV-VERIFICACION DE UTILLAJES": _uv,
    "VARIOS": _varios,
}


# --------------------------------------------------------------------- #
#  ---------------------  FUNCIÓN PÚBLICA  ---------------------------- #
# --------------------------------------------------------------------- #
def asignar_tarea_asterisco(
    actividad: str, chapa: str, categoria: str | None
) -> str:
    """
    Determina el código de tarea para los registros marcados con '*'.

    Parámetros
    ----------
    actividad : str
        Valor de la columna 'actividad' (tal cual viene en el Excel).
    chapa : str
        Nº de chapa (IdUsuario ya limpio).
    categoria : str | None
        Valor de la columna 'CATEGORIA' (minúsculas).

    Returns
    -------
    str
        CodTarea final o '' si la fila debe descartarse / sin asignar.
    """
    categoria = (categoria or "").strip().lower()

    func = _RULES.get(actividad.strip())
    if func is None:
        # Actividad sin regla específica → pendiente de decidir
        return ""

    return func(chapa, categoria)
