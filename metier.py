"""Règles métier côté application : référentiels, capacités, taux, règle Zone 7."""
import math
import re

import pandas as pd

from backend import (
    CAMION_CODE,
    CAPACITE_POIDS_ESTAFETTE,
    CAPACITE_VOLUME_ESTAFETTE,
    get_capacite_poids_camion,
    get_capacite_volume_camion,
)

# =====================================================
# RÉFÉRENTIELS
# =====================================================
VEHICULES_DISPONIBLES = [
    "SLG-VEH11", "SLG-VEH08", "SLG-VEH22", "SLG-VEH23", "SLG-VEH20",
    "SLG-VEH19", "SLG-VEH204", "SLG-VEH178", "SLG-VEH198", "code-Camion",
]
VEHICULE_CAMION = "code-Camion"

CHAUFFEURS_DETAILS = {
    "09254": "DAMMAK Karim", "00000": "NAJI Mohamed", "00001": "MSEKNI Mahdi",
    "00002": "GHAOUI Mohamed Amine", "00003": "HARABI Salah", "00004": "ALLAYA Maher",
    "10334": "BOUJELBENE Mohamed", "15144": "GADDOUR Rami", "08278": "DAMMAK Wissem",
    "00005": "MAKNI Hcine", "00006": "BARI Aymen", "18339": "REKIK Ahmed",
    "07250": "BARKIA Mustapha", "13321": "BADRI Moez", "99999": "Chauffeur Camion",
}
OPTIONS_CHAUFFEURS = [f"{m} - {n}" for m, n in CHAUFFEURS_DETAILS.items()]
CHAUFFEUR_CAMION = "99999 - Chauffeur Camion"

TYPES_CAMION = ["5 tonnes", "10 tonnes"]
VILLES_EXCLUES_ANALYSE = {"TRIPOLI"}

# Ordre standard des colonnes d'un voyage (écran + exports)
COLS_VOYAGE = [
    "Zone", "Véhicule N°", "Poids total chargé", "Volume total chargé",
    "Client(s) inclus", "Représentant(s) inclus", "BL inclus",
    "Taux occupation poids (%)", "Taux occupation volume (%)",
    "Location_camion", "Location_proposee", "Code Véhicule",
]
COLS_LISTES = ["Client(s) inclus", "Représentant(s) inclus", "BL inclus"]
COLS_ATTRIBUTION = ["Véhicule attribué", "Chauffeur attribué", "Matricule chauffeur"]


# =====================================================
# OUTILS
# =====================================================
def liste(valeur, sep=r"[;,\n]"):
    """Découpe une cellule 'A;B' ou 'A, B' en liste propre."""
    if valeur is None:
        return []
    try:
        if pd.isna(valeur):
            return []
    except (TypeError, ValueError):
        pass
    texte = str(valeur).strip()
    if texte.lower() in ("", "nan", "none"):
        return []
    return [e.strip() for e in re.split(sep, texte) if e.strip()]


def bls_de(valeur):
    return liste(valeur, r";")


def num_zone(zone):
    m = re.search(r"\d+", str(zone))
    return int(m.group()) if m else 999


def trier_par_zone(df):
    if df is None or df.empty or "Zone" not in df.columns:
        return df
    cles = ["_zn"] + (["Véhicule N°"] if "Véhicule N°" in df.columns else [])
    return (df.assign(_zn=df["Zone"].map(num_zone))
              .sort_values(cles, na_position="last", kind="stable")
              .drop(columns="_zn"))


# =====================================================
# CAPACITÉS ET TAUX D'OCCUPATION
# =====================================================
def est_camion(row):
    return str(row.get("Code Véhicule", "")).strip() == CAMION_CODE


def type_camion(row):
    t = row.get("Type_Camion", None)
    return t if isinstance(t, str) and t in TYPES_CAMION else "5 tonnes"


def capacites(row):
    """(poids max kg, volume max m³) du véhicule de la ligne."""
    if est_camion(row):
        t = type_camion(row)
        return get_capacite_poids_camion(t), get_capacite_volume_camion(t)
    return CAPACITE_POIDS_ESTAFETTE, CAPACITE_VOLUME_ESTAFETTE


def ajouter_taux(df):
    """Recalcule les taux d'occupation poids / volume / max pour chaque voyage."""
    if df is None or df.empty:
        return df
    df = df.loc[:, ~df.columns.duplicated()].copy()
    cap = df.apply(capacites, axis=1, result_type="expand")
    poids = pd.to_numeric(df["Poids total chargé"], errors="coerce").fillna(0)
    volume = pd.to_numeric(df["Volume total chargé"], errors="coerce").fillna(0)
    df["Poids total chargé"] = poids
    df["Volume total chargé"] = volume
    df["Taux occupation poids (%)"] = poids / cap[0] * 100
    df["Taux occupation volume (%)"] = volume / cap[1] * 100
    df["Taux d'occupation (%)"] = df[["Taux occupation poids (%)", "Taux occupation volume (%)"]].max(axis=1)
    return df


def libelle_type(row):
    return f"Camion loué {type_camion(row)}" if est_camion(row) else "Estafette"


def nb_estafettes(df):
    """Estafettes réellement utilisées (Zone 7 : plusieurs voyages = une estafette)."""
    if df is None or df.empty:
        return 0
    d = df[(df["Code Véhicule"] != CAMION_CODE) & (df["BL inclus"].map(lambda x: len(bls_de(x)) > 0))]
    z7 = d[d["Zone"] == "Zone 7"]
    autres = d[d["Zone"] != "Zone 7"]
    n7 = z7["Véhicule N°"].map(lambda x: str(x).split("-")[0] if "-Voyage" in str(x) else str(x)).nunique()
    return n7 + len(autres)


def nb_camions(df):
    return 0 if df is None or df.empty else int((df["Code Véhicule"] == CAMION_CODE).sum())


# =====================================================
# BESOIN EN ESTAFETTES (règle spéciale SFAX / Zone 7)
# =====================================================
# Nombre maximum de voyages par estafette en Zone 7 (SFAX).
# Provisoire : cette valeur (1, 2 ou 3) sera fournie par un module dédié.
MAX_VOYAGES_ZONE7 = 2


def regle_zone7(row):
    """Zone 7 (Sfax Nord / Sfax Sud) ou ville SFAX : une estafette peut faire plusieurs voyages."""
    if str(row.get("Ville", "")).strip().upper() == "SFAX":
        return True
    return str(row.get("Zone", "")).strip().upper().startswith("ZONE 7")


def estafettes_necessaires(val, zone7):
    """Nombre d'estafettes pour un besoin donné (en Zone 7 : plusieurs voyages par estafette)."""
    val = pd.to_numeric(val, errors="coerce")
    if pd.isna(val):
        return 0
    return int(math.ceil(val / MAX_VOYAGES_ZONE7)) if zone7 else float(val)


def libelle_besoin(val, zone7, sep=" → "):
    if val is None or pd.isna(val):
        return ""
    try:
        val = float(val)
    except (TypeError, ValueError):
        return str(val)
    if zone7:
        if val <= MAX_VOYAGES_ZONE7:
            return f"{int(val)} voyage(s) par estafette"
        return f"{int(val)} voyages{sep}{estafettes_necessaires(val, True)} estafette(s)"
    return f"{val:.1f}".replace(".", ",")


def total_besoin(df):
    if df is None or "Besoin estafette réel" not in df.columns:
        return 0.0
    return float(sum(estafettes_necessaires(r["Besoin estafette réel"], regle_zone7(r)) for _, r in df.iterrows()))
