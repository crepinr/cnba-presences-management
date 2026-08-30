from pathlib import Path
import os

# ============================================================
# GROUPES / PROFILS D'ANALYSE
# ============================================================

GROUP_ELITE = "Elite"
GROUP_ELITE_JEUNES = "Elite Jeunes"

GROUP_ALIASES = {
    "elite": GROUP_ELITE,
    "e": GROUP_ELITE,
    "elite jeunes": GROUP_ELITE_JEUNES,
    "elite jeunes": GROUP_ELITE_JEUNES,
    "elite_jeunes": GROUP_ELITE_JEUNES,
    "elite-jeunes": GROUP_ELITE_JEUNES,
    "ej": GROUP_ELITE_JEUNES,
}

GROUP_CODES = {
    GROUP_ELITE: "E",
    GROUP_ELITE_JEUNES: "EJ",
}

# ============================================================
# COULEURS TERMINAL
# ============================================================

BLUE = "\033[34m"
GREEN = "\033[32m"
YELLOW = "\033[33m"
RED = "\033[31m"
RESET = "\033[0m"

# ============================================================
# GOOGLE DRIVE
# ============================================================

DATA_SHEET_KEY_E = "1pHFj2D63dVnghFwxeNurqcGgtAIbrVJMc9ZOAYBr0rY"
DATA_SHEET_KEY_EJ = "1ni3HTHVQNjnaAkT4FZGsd-YlBbKV1NmSpAZrF4Q-Jqw"

# ============================================================
# SAISON
# ============================================================

SEASON_START_YEAR = 2025
SEASON_END_YEAR = 2026

# ============================================================
# FEUILLES EXCEL
# ============================================================

INFO_SHEETS = {
    "Groupe",
    "Bilan",
    "Classement",
    "Communications",
}

MONTH_SHEETS = [
    "septembre",
    "octobre",
    "novembre",
    "décembre",
    "janvier",
    "février",
    "mars",
    "avril",
    "mai",
    "juin",
    "juillet",
]

MONTH_NUMBER = {
    "septembre": 9,
    "octobre": 10,
    "novembre": 11,
    "décembre": 12,
    "decembre": 12,
    "janvier": 1,
    "février": 2,
    "fevrier": 2,
    "mars": 3,
    "avril": 4,
    "mai": 5,
    "juin": 6,
    "juillet": 7,
}

# ============================================================
# STRUCTURE DES FEUILLES MENSUELLES
# ============================================================

DATE_ROW = 2
DAY_CODE_ROW = 3
FIRST_SWIMMER_ROW = 4

NAME_COLUMN = 1
FIRSTNAME_COLUMN = 2
BIRTHDATE_COLUMN = 3
FIRST_SESSION_COLUMN = 4

# ============================================================
# STATUTS DE PRESENCE
# ============================================================

STATUS_MAPPING = {
    "V": "present",
    "C": "present",
    "X": "absent",
    "M": "maladie",
    "O": "excuse",
    "E": "excuse",
    "": "hors_analyse",
    None: "hors_analyse",
}

STATUTS_COMPTABILISES = {
    "present",
    "absent",
    "maladie",
    "excuse",
}

# ============================================================
# COLONNES DE CALCUL A IGNORER
# ============================================================

SUMMARY_COLUMN_KEYWORDS = {
    "présences %",
    "presences %",
    "presence %",
    "présence %",
    "présences",
    "presences",
    "%",
}

# ============================================================
# TYPES DE SEANCES
# ============================================================

MORNING_KEYWORDS = {
    "am",
}

WEEKEND_CODES = {
    "s",
    "d",
}

# ============================================================
# EXCLUSIONS POUR ANALYSE GROUPE
# ============================================================

EXCLUDED_SWIMMERS_E = {
    "CONTENT_Elie_2006-01-21": "Arrivé en cours de saison",
    "CAVADINI_Caroline_1990-04-01": "Senior",
    "SEYMOUR_Felix_2005-07-10": "Senior",
    "TOPBAG_Sami_2005-11-04": "Senior",
    "VAN HENTENRIJK_Matthieu_2003-03-30": "Senior",
    "VANDEN BROECK_Jana_1989-04-09": "Senior",
    "VATA_Gjon_2002-06-14": "Senior",
}

EXCLUDED_SWIMMERS_EJ = {
}

# ============================================================
# VALIDATION / CLEAN UP
# ============================================================

STOP_ON_CRITICAL_ERRORS = False

EXPECTED_SESSION_TYPES = {
    "matin",
    "soir",
    "week-end",
    "inconnu",
}

MIN_TRAINING_DATE = "2026-08-24"
MAX_TRAINING_DATE = "2027-07-31"

CLEAN_UP_EXTRA_BINARY_COLUMNS = [
    "statut_inconnu",
    "seance_comptabilisee",
]

CLEAN_UP_BASE_REQUIRED_DATA_CLEAN_COLUMNS = [
    "nageur_id",
    "nom",
    "prenom",
    "date_naissance",
    "sheet_name",
    "mois",
    "date",
    "jour_code",
    "type_seance",
    "colonne_excel",
    "ligne_excel",
    "statut_original",
    "statut_clean",
    "statut_inconnu",
    "seance_comptabilisee",
    "annee",
    "mois_numero",
    "semaine",
    "age_seance",
    "seance_id",
    "exclude_from_group_analysis",
    "exclusion_reason",
]

CLEAN_UP_REQUIRED_NON_EMPTY_DATA_COLUMNS = [
    "nageur_id",
    "nom",
    "prenom",
    "date_naissance",
    "sheet_name",
    "mois",
    "date",
    "jour_code",
    "type_seance",
    "colonne_excel",
    "ligne_excel",
    "statut_clean",
    "seance_id",
]

CLEAN_UP_MIN_AGE_WARNING = 5
CLEAN_UP_MAX_AGE_WARNING = 45
CLEAN_UP_TERMINAL_MAX_EXAMPLES_PER_GROUP = 15

# ============================================================
# EXPORT EXCEL DE CONTROLE
# ============================================================

EXPORT_EXCEL_CONTROL_FILE = True

# ============================================================
# CODES SOURCE POUR L'ANALYSE DE PRESENCE
# ============================================================

PRESENCE_CODES = {
    "V",
    "C",
}

COMPETITION_CODES = {
    "C",
}

PLANNED_ABSENCE_CODES = {
    "O",
}

SICKNESS_CODES = {
    "M",
}

ABSENCE_CODES = {
    "X",
}

EMPTY_CODES = {
    "",
}

# ============================================================
# PROFILS
# ============================================================

GROUP_CONFIGS = {
    GROUP_ELITE: {
        "code": GROUP_CODES[GROUP_ELITE],
        "sheet_key": DATA_SHEET_KEY_E,
        "excluded_swimmers": EXCLUDED_SWIMMERS_E,
    },
    GROUP_ELITE_JEUNES: {
        "code": GROUP_CODES[GROUP_ELITE_JEUNES],
        "sheet_key": DATA_SHEET_KEY_EJ,
        "excluded_swimmers": EXCLUDED_SWIMMERS_EJ,
    },
}

# ============================================================
# CONTEXTE ACTIF ET CHEMINS DERIVES
# ============================================================

def normalize_group_name(group_name: str) -> str:
    """
    Normalise les entrées utilisateur :
    - Elite
    - E
    - Elite Jeunes
    - EJ
    """
    if group_name is None:
        return GROUP_ELITE

    key = str(group_name).strip().lower().replace("é", "e").replace("è", "e")
    key = key.replace("  ", " ")

    if key in GROUP_ALIASES:
        return GROUP_ALIASES[key]

    valid_groups = ", ".join(GROUP_CONFIGS.keys())

    raise ValueError(
        f"Groupe inconnu : {group_name}. Groupes possibles : {valid_groups}"
    )


def available_groups():
    return list(GROUP_CONFIGS.keys())


def set_active_group(group_name: str):
    """
    Définit le groupe actif et recalcule tous les chemins utilisés par les scripts.

    Cette fonction permet d'exécuter exactement la même logique pour :
    - Elite
    - Elite Jeunes
    """
    global ACTIVE_GROUP, ACTIVE_GROUP_CODE, ACTIVE_GROUP_LABEL
    global DATA_SHEET_KEY, EXCLUDED_SWIMMERS
    global INPUT_DIR, INPUT_FILE, OUTPUT_DIR, OUTPUT_FILE
    global PICKLE_DIR, DATA_CLEAN_PICKLE, NAGEURS_PICKLE, SEANCES_PICKLE
    global AUDIT_GLOBAL_PICKLE, AUDIT_DATES_PICKLE, AUDIT_STATUTS_PICKLE
    global CLEAN_UP_DIR, CLEAN_UP_REPORT_XLSX, CLEAN_UP_ISSUES_PICKLE
    global DATA_CLEAN_VALIDATED_PICKLE, NAGEURS_VALIDATED_PICKLE, SEANCES_VALIDATED_PICKLE
    global ANALYSIS_DIR, ANALYSIS_REPORT_XLSX
    global ANALYSIS_INDIVIDUAL_PICKLE, ANALYSIS_MONTHLY_PICKLE
    global ANALYSIS_GROUP_MONTHLY_PICKLE, ANALYSIS_GROUP_SUMMARY_PICKLE
    global ANALYSIS_SESSION_DETAIL_PICKLE

    ACTIVE_GROUP = normalize_group_name(group_name)
    profile = GROUP_CONFIGS[ACTIVE_GROUP]

    ACTIVE_GROUP_CODE = profile["code"]
    ACTIVE_GROUP_LABEL = ACTIVE_GROUP

    DATA_SHEET_KEY = profile["sheet_key"]
    EXCLUDED_SWIMMERS = profile["excluded_swimmers"]

    INPUT_DIR = Path("Input")
    INPUT_FILE = INPUT_DIR / f"Input_{ACTIVE_GROUP_CODE}.xlsx"

    OUTPUT_DIR = Path("Output") / ACTIVE_GROUP_CODE
    OUTPUT_FILE = OUTPUT_DIR / "dataset_clean.xlsx"

    PICKLE_DIR = OUTPUT_DIR / "pickle"

    DATA_CLEAN_PICKLE = PICKLE_DIR / "data_clean.pkl"
    NAGEURS_PICKLE = PICKLE_DIR / "nageurs.pkl"
    SEANCES_PICKLE = PICKLE_DIR / "seances.pkl"

    AUDIT_GLOBAL_PICKLE = PICKLE_DIR / "audit_global.pkl"
    AUDIT_DATES_PICKLE = PICKLE_DIR / "audit_dates.pkl"
    AUDIT_STATUTS_PICKLE = PICKLE_DIR / "audit_statuts.pkl"

    CLEAN_UP_DIR = OUTPUT_DIR / "clean_up"

    CLEAN_UP_REPORT_XLSX = CLEAN_UP_DIR / "clean_up_report.xlsx"
    CLEAN_UP_ISSUES_PICKLE = CLEAN_UP_DIR / "clean_up_issues.pkl"

    DATA_CLEAN_VALIDATED_PICKLE = CLEAN_UP_DIR / "data_clean_validated.pkl"
    NAGEURS_VALIDATED_PICKLE = CLEAN_UP_DIR / "nageurs_validated.pkl"
    SEANCES_VALIDATED_PICKLE = CLEAN_UP_DIR / "seances_validated.pkl"

    ANALYSIS_DIR = OUTPUT_DIR / "analysis"

    ANALYSIS_REPORT_XLSX = ANALYSIS_DIR / "presence_analysis.xlsx"

    ANALYSIS_INDIVIDUAL_PICKLE = ANALYSIS_DIR / "presence_individual.pkl"
    ANALYSIS_MONTHLY_PICKLE = ANALYSIS_DIR / "presence_monthly.pkl"
    ANALYSIS_GROUP_MONTHLY_PICKLE = ANALYSIS_DIR / "presence_group_monthly.pkl"
    ANALYSIS_GROUP_SUMMARY_PICKLE = ANALYSIS_DIR / "presence_group_summary.pkl"
    ANALYSIS_SESSION_DETAIL_PICKLE = ANALYSIS_DIR / "presence_session_detail.pkl"

    return ACTIVE_GROUP


# Groupe actif par défaut.
# Peut être surchargé par variable d'environnement :
# CNBA_GROUP="Elite Jeunes" python3.11 03_analyze_data.py
set_active_group(os.environ.get("CNBA_GROUP", GROUP_ELITE))
