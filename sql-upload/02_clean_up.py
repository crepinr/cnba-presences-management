import argparse
import sys
from pathlib import Path
import pandas as pd
import numpy as np

import config

try:
    import attendance_yaml_config
    attendance_yaml_config.bootstrap(config)
except Exception as yaml_config_error:
    # Si aucun YAML n'est configuré, on conserve le comportement historique.
    if "CNBA_CONFIG_YAML" in __import__("os").environ:
        raise

def get_expected_clean_statuses():
    """
    Récupère automatiquement les statuts valides depuis config.STATUS_MAPPING.

    Exemple :
    STATUS_MAPPING = {
        "V": "present",
        "X": "absent",
        "M": "maladie",
        "O": "excuse",
        "E": "excuse",
        "": "hors_analyse",
        None: "hors_analyse",
    }

    donnera :
    {
        "present",
        "absent",
        "maladie",
        "excuse",
        "hors_analyse",
    }
    """
    if not hasattr(config, "STATUS_MAPPING"):
        raise AttributeError(
            "config.STATUS_MAPPING est manquant. "
            "Impossible de déterminer les statuts valides."
        )

    statuses = {
        clean_status
        for clean_status in config.STATUS_MAPPING.values()
        if clean_status is not None and str(clean_status).strip() != ""
    }

    statuses = {
        str(status).strip()
        for status in statuses
    }

    if not statuses:
        raise ValueError(
            "config.STATUS_MAPPING ne contient aucun statut valide."
        )

    return statuses

def get_status_binary_columns():
    """
    Déduit automatiquement les colonnes binaires de statut depuis STATUS_MAPPING.

    Exemple :
    statuts valides = {"present", "absent", "maladie", "excuse", "hors_analyse"}

    colonnes attendues :
    ["present", "absent", "maladie", "excuse", "hors_analyse"]
    """
    return sorted(get_expected_clean_statuses())

# ============================================================
# PARAMETRES DEDUITS DE config.py
# ============================================================

EXPECTED_CLEAN_STATUSES = get_expected_clean_statuses()
STATUS_BINARY_COLUMNS = get_status_binary_columns()


def get_binary_columns():
    """
    Colonnes binaires attendues.

    Les colonnes de statut viennent de config.STATUS_MAPPING.
    Les colonnes techniques supplémentaires viennent de
    config.CLEAN_UP_EXTRA_BINARY_COLUMNS.
    """
    return STATUS_BINARY_COLUMNS + list(config.CLEAN_UP_EXTRA_BINARY_COLUMNS)


def get_required_data_clean_columns():
    """
    Colonnes obligatoires de data_clean.

    La base vient de config.CLEAN_UP_BASE_REQUIRED_DATA_CLEAN_COLUMNS.
    Les colonnes de statut sont ajoutées dynamiquement depuis
    config.STATUS_MAPPING.
    """
    return (
        list(config.CLEAN_UP_BASE_REQUIRED_DATA_CLEAN_COLUMNS)
        + STATUS_BINARY_COLUMNS
    )


def get_required_non_empty_data_columns():
    """
    Colonnes obligatoires non vides.
    """
    return list(config.CLEAN_UP_REQUIRED_NON_EMPTY_DATA_COLUMNS)


# ============================================================
# COULEURS TERMINAL
# ============================================================

class Colors:
    RESET = "\033[0m"
    BOLD = "\033[1m"

    RED = "\033[91m"
    YELLOW = "\033[93m"
    GREEN = "\033[92m"
    BLUE = "\033[94m"
    CYAN = "\033[96m"
    MAGENTA = "\033[95m"
    GREY = "\033[90m"


def color_text(text, color):
    return f"{color}{text}{Colors.RESET}"


def print_title(title):
    print()
    print(color_text("=" * 80, Colors.BLUE))
    print(color_text(title, Colors.BOLD + Colors.BLUE))
    print(color_text("=" * 80, Colors.BLUE))


def print_section(title):
    print()
    print(color_text("-" * 80, Colors.CYAN))
    print(color_text(title, Colors.BOLD + Colors.CYAN))
    print(color_text("-" * 80, Colors.CYAN))


def print_success(message):
    print(color_text(f"✅ {message}", Colors.GREEN))


def print_warning(message):
    print(color_text(f"⚠️  {message}", Colors.YELLOW))


def print_error(message):
    print(color_text(f"❌ {message}", Colors.RED))


def print_info(message):
    print(color_text(f"ℹ️  {message}", Colors.GREY))


# ============================================================
# OUTILS GENERAUX
# ============================================================

def ensure_directories():
    config.CLEAN_UP_DIR.mkdir(parents=True, exist_ok=True)


def read_pickle_or_fail(path: Path, label: str) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(
            f"Fichier Pickle manquant : {label} -> {path}"
        )

    return pd.read_pickle(path)


def clean_string_series(series: pd.Series) -> pd.Series:
    return (
        series
        .fillna("")
        .astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )


def format_value(value):
    if pd.isna(value):
        return ""

    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")

    return str(value)


def excel_column_name(column_number):
    """
    Convertit 1 -> A, 2 -> B, 27 -> AA.
    """
    if pd.isna(column_number):
        return ""

    try:
        column_number = int(column_number)
    except Exception:
        return ""

    result = ""

    while column_number > 0:
        column_number, remainder = divmod(column_number - 1, 26)
        result = chr(65 + remainder) + result

    return result


def make_excel_cell(row, column):
    if pd.isna(row) or pd.isna(column):
        return ""

    return f"{excel_column_name(column)}{int(row)}"


def get_row_source_info(row: pd.Series):
    ligne_excel = row.get("ligne_excel", "")
    colonne_excel = row.get("colonne_excel", "")
    cell = make_excel_cell(ligne_excel, colonne_excel)

    return {
        "sheet_name": row.get("sheet_name", ""),
        "mois": row.get("mois", ""),
        "date": row.get("date", ""),
        "nom": row.get("nom", ""),
        "prenom": row.get("prenom", ""),
        "nageur_id": row.get("nageur_id", ""),
        "ligne_excel": ligne_excel,
        "colonne_excel": colonne_excel,
        "cellule_excel": cell,
    }


def add_issue(
    issues: list,
    severity: str,
    error_type: str,
    dataframe: str,
    column: str,
    message: str,
    sheet_name="",
    mois="",
    date="",
    nom="",
    prenom="",
    nageur_id="",
    ligne_excel="",
    colonne_excel="",
    cellule_excel="",
    value="",
    expected="",
):
    issues.append({
        "severity": severity,
        "error_type": error_type,
        "dataframe": dataframe,
        "column": column,
        "message": message,
        "sheet_name": format_value(sheet_name),
        "mois": format_value(mois),
        "date": format_value(date),
        "nom": format_value(nom),
        "prenom": format_value(prenom),
        "nageur_id": format_value(nageur_id),
        "ligne_excel": format_value(ligne_excel),
        "colonne_excel": format_value(colonne_excel),
        "cellule_excel": format_value(cellule_excel),
        "value": format_value(value),
        "expected": format_value(expected),
    })


def add_issue_from_row(
    issues: list,
    severity: str,
    error_type: str,
    dataframe: str,
    column: str,
    message: str,
    row: pd.Series,
    value=None,
    expected="",
):
    source = get_row_source_info(row)

    if value is None and column in row.index:
        value = row[column]

    add_issue(
        issues=issues,
        severity=severity,
        error_type=error_type,
        dataframe=dataframe,
        column=column,
        message=message,
        sheet_name=source["sheet_name"],
        mois=source["mois"],
        date=source["date"],
        nom=source["nom"],
        prenom=source["prenom"],
        nageur_id=source["nageur_id"],
        ligne_excel=source["ligne_excel"],
        colonne_excel=source["colonne_excel"],
        cellule_excel=source["cellule_excel"],
        value=value,
        expected=expected,
    )


# ============================================================
# LECTURE DES PICKLE
# ============================================================

def load_imported_pickles():
    print_section("Lecture des fichiers Pickle")

    df_clean = read_pickle_or_fail(
        config.DATA_CLEAN_PICKLE,
        "data_clean"
    )
    print_success(f"data_clean lu : {len(df_clean)} lignes")

    df_nageurs = read_pickle_or_fail(
        config.NAGEURS_PICKLE,
        "nageurs"
    )
    print_success(f"nageurs lu : {len(df_nageurs)} lignes")

    df_seances = read_pickle_or_fail(
        config.SEANCES_PICKLE,
        "seances"
    )
    print_success(f"seances lu : {len(df_seances)} lignes")

    audit_global = read_pickle_or_fail(
        config.AUDIT_GLOBAL_PICKLE,
        "audit_global"
    )
    print_success(f"audit_global lu : {len(audit_global)} lignes")

    audit_dates = read_pickle_or_fail(
        config.AUDIT_DATES_PICKLE,
        "audit_dates"
    )
    print_success(f"audit_dates lu : {len(audit_dates)} lignes")

    audit_statuts = read_pickle_or_fail(
        config.AUDIT_STATUTS_PICKLE,
        "audit_statuts"
    )
    print_success(f"audit_statuts lu : {len(audit_statuts)} lignes")

    return (
        df_clean,
        df_nageurs,
        df_seances,
        audit_global,
        audit_dates,
        audit_statuts,
    )


# ============================================================
# NORMALISATION DES TYPES
# ============================================================

def normalize_data_clean(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    string_columns = [
        "nageur_id",
        "nom",
        "prenom",
        "sheet_name",
        "mois",
        "jour_code",
        "type_seance",
        "statut_original",
        "statut_clean",
        "seance_id",
        "exclusion_reason",
    ]

    for column in string_columns:
        if column in df.columns:
            df[column] = clean_string_series(df[column])

    for column in ["date_naissance", "date"]:
        if column in df.columns:
            df[column] = pd.to_datetime(df[column], errors="coerce")

    integer_columns = [
        "colonne_excel",
        "ligne_excel",
        "annee",
        "mois_numero",
        "semaine",
    ]

    for column in integer_columns:
        if column in df.columns:
            df[column] = pd.to_numeric(
                df[column],
                errors="coerce"
            ).astype("Int64")

    for column in get_binary_columns():
        if column in df.columns:
            df[column] = pd.to_numeric(
                df[column],
                errors="coerce"
            ).astype("Int64")

    if "age_seance" in df.columns:
        df["age_seance"] = pd.to_numeric(
            df["age_seance"],
            errors="coerce"
        )

    if "exclude_from_group_analysis" in df.columns:
        df["exclude_from_group_analysis"] = (
            df["exclude_from_group_analysis"]
            .fillna(False)
            .astype(bool)
        )

    return df


def normalize_nageurs(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    for column in ["nageur_id", "nom", "prenom"]:
        if column in df.columns:
            df[column] = clean_string_series(df[column])

    if "date_naissance" in df.columns:
        df["date_naissance"] = pd.to_datetime(
            df["date_naissance"],
            errors="coerce"
        )

    return df


def normalize_seances(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    for column in [
        "sheet_name",
        "mois",
        "jour_code",
        "type_seance",
        "seance_id",
    ]:
        if column in df.columns:
            df[column] = clean_string_series(df[column])

    if "date" in df.columns:
        df["date"] = pd.to_datetime(
            df["date"],
            errors="coerce"
        )

    if "colonne_excel" in df.columns:
        df["colonne_excel"] = pd.to_numeric(
            df["colonne_excel"],
            errors="coerce"
        ).astype("Int64")

    return df


def normalize_audit_dates(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    for column in ["sheet_name", "raw_date", "jour_code", "issue"]:
        if column in df.columns:
            df[column] = clean_string_series(df[column])

    if "parsed_date" in df.columns:
        df["parsed_date"] = pd.to_datetime(
            df["parsed_date"],
            errors="coerce"
        )

    for column in ["column", "expected_month", "date_month"]:
        if column in df.columns:
            df[column] = pd.to_numeric(
                df[column],
                errors="coerce"
            ).astype("Int64")

    if "date_mismatch" in df.columns:
        df["date_mismatch"] = df["date_mismatch"].fillna(False).astype(bool)

    return df


def normalize_audit_statuts(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    for column in [
        "sheet_name",
        "nom",
        "prenom",
        "value",
        "statut_original",
        "statut_clean",
    ]:
        if column in df.columns:
            df[column] = clean_string_series(df[column])

    for column in ["row", "column"]:
        if column in df.columns:
            df[column] = pd.to_numeric(
                df[column],
                errors="coerce"
            ).astype("Int64")

    return df


# ============================================================
# VALIDATIONS GENERALES
# ============================================================

def validate_required_columns(
    df: pd.DataFrame,
    dataframe_name: str,
    required_columns: list,
    issues: list,
):
    for column in required_columns:
        if column not in df.columns:
            add_issue(
                issues=issues,
                severity="CRITICAL",
                error_type="Colonne manquante",
                dataframe=dataframe_name,
                column=column,
                message="Colonne obligatoire absente du dataframe.",
                expected="Colonne présente",
            )


def validate_required_values(
    df: pd.DataFrame,
    dataframe_name: str,
    required_columns: list,
    issues: list,
):
    for column in required_columns:
        if column not in df.columns:
            continue

        if pd.api.types.is_string_dtype(df[column]) or df[column].dtype == "object":
            mask = df[column].isna() | (df[column].astype(str).str.strip() == "")
        else:
            mask = df[column].isna()

        bad_rows = df[mask]

        for _, row in bad_rows.iterrows():
            add_issue_from_row(
                issues=issues,
                severity="CRITICAL",
                error_type="Valeur obligatoire manquante",
                dataframe=dataframe_name,
                column=column,
                message="Une valeur obligatoire est manquante.",
                row=row,
                expected="Valeur non vide",
            )


def validate_duplicates(
    df: pd.DataFrame,
    dataframe_name: str,
    subset: list,
    issues: list,
    message: str,
):
    missing = [column for column in subset if column not in df.columns]

    if missing:
        return

    duplicated = df[df.duplicated(subset=subset, keep=False)]

    for _, row in duplicated.iterrows():
        add_issue_from_row(
            issues=issues,
            severity="CRITICAL",
            error_type="Doublon",
            dataframe=dataframe_name,
            column=", ".join(subset),
            message=message,
            row=row,
            expected="Combinaison unique",
        )


# ============================================================
# VALIDATION DATA CLEAN
# ============================================================

def validate_status_values(df: pd.DataFrame, issues: list):
    if "statut_clean" not in df.columns:
        return

    mask = ~df["statut_clean"].isin(EXPECTED_CLEAN_STATUSES)

    for _, row in df[mask].iterrows():
        add_issue_from_row(
            issues=issues,
            severity="CRITICAL",
            error_type="Statut inconnu",
            dataframe="data_clean",
            column="statut_clean",
            message="Statut nettoyé non reconnu.",
            row=row,
            expected=", ".join(sorted(EXPECTED_CLEAN_STATUSES)),
        )


def validate_binary_columns(df: pd.DataFrame, issues: list):
    for column in get_binary_columns():
        if column not in df.columns:
            continue

        mask = ~df[column].isin([0, 1])

        for _, row in df[mask].iterrows():
            add_issue_from_row(
                issues=issues,
                severity="CRITICAL",
                error_type="Valeur binaire invalide",
                dataframe="data_clean",
                column=column,
                message="Colonne binaire avec une valeur différente de 0 ou 1.",
                row=row,
                expected="0 ou 1",
            )

def validate_one_status_per_row(df: pd.DataFrame, issues: list):
    """
    Vérifie qu'une ligne appartient à une seule catégorie de statut.

    Les colonnes de statut sont déduites automatiquement de config.STATUS_MAPPING.
    """
    status_columns = [
        column
        for column in STATUS_BINARY_COLUMNS
        if column in df.columns
    ]

    if not status_columns:
        add_issue(
            issues=issues,
            severity="CRITICAL",
            error_type="Colonnes statut manquantes",
            dataframe="data_clean",
            column="statuts",
            message="Aucune colonne binaire de statut n'a été trouvée.",
            expected=", ".join(STATUS_BINARY_COLUMNS),
        )
        return

    missing_columns = [
        column
        for column in STATUS_BINARY_COLUMNS
        if column not in df.columns
    ]

    for column in missing_columns:
        add_issue(
            issues=issues,
            severity="CRITICAL",
            error_type="Colonne statut manquante",
            dataframe="data_clean",
            column=column,
            message="Colonne binaire de statut manquante.",
            expected="Colonne créée automatiquement depuis config.STATUS_MAPPING",
        )

    if missing_columns:
        return

    status_sum = df[STATUS_BINARY_COLUMNS].sum(axis=1)

    mask = status_sum != 1

    for _, row in df[mask].iterrows():
        add_issue_from_row(
            issues=issues,
            severity="CRITICAL",
            error_type="Incohérence statut",
            dataframe="data_clean",
            column="/".join(STATUS_BINARY_COLUMNS),
            message="Une ligne doit appartenir à une seule catégorie de statut.",
            row=row,
            value=status_sum.loc[row.name],
            expected="Somme des colonnes statut = 1",
        )
def validate_status_binary_consistency(df: pd.DataFrame, issues: list):
    """
    Vérifie que chaque colonne binaire correspond bien à statut_clean.

    Cette validation est dynamique :
    - les statuts viennent de config.STATUS_MAPPING
    - les colonnes binaires attendues sont les valeurs de STATUS_MAPPING
    """
    if "statut_clean" not in df.columns:
        return

    for status in STATUS_BINARY_COLUMNS:
        if status not in df.columns:
            add_issue(
                issues=issues,
                severity="CRITICAL",
                error_type="Colonne statut manquante",
                dataframe="data_clean",
                column=status,
                message="Colonne binaire de statut manquante.",
                expected="Colonne créée depuis config.STATUS_MAPPING",
            )
            continue

        expected_value = (
            df["statut_clean"] == status
        ).astype(int)

        mask = df[status] != expected_value

        for _, row in df[mask].iterrows():
            add_issue_from_row(
                issues=issues,
                severity="CRITICAL",
                error_type="Incohérence statut",
                dataframe="data_clean",
                column=status,
                message="La colonne binaire ne correspond pas au statut_clean.",
                row=row,
                value=row[status],
                expected=f"1 si statut_clean = {status}, sinon 0",
            )
def validate_seance_comptabilisee(df: pd.DataFrame, issues: list):
    if "seance_comptabilisee" not in df.columns:
        return

    if "statut_clean" not in df.columns:
        return

    expected = df["statut_clean"].isin(
        config.STATUTS_COMPTABILISES
    ).astype(int)

    mask = df["seance_comptabilisee"] != expected

    for _, row in df[mask].iterrows():
        add_issue_from_row(
            issues=issues,
            severity="CRITICAL",
            error_type="Incohérence séance comptabilisée",
            dataframe="data_clean",
            column="seance_comptabilisee",
            message="seance_comptabilisee ne correspond pas à config.STATUTS_COMPTABILISES.",
            row=row,
            value=row["seance_comptabilisee"],
            expected=expected.loc[row.name],
        )

def validate_config_status_consistency(issues: list):
    """
    Vérifie que la configuration des statuts est cohérente.

    Objectif :
    - STATUS_MAPPING doit exister
    - STATUTS_COMPTABILISES doit être inclus dans les statuts propres connus
    """
    if not hasattr(config, "STATUS_MAPPING"):
        add_issue(
            issues=issues,
            severity="CRITICAL",
            error_type="Configuration invalide",
            dataframe="config",
            column="STATUS_MAPPING",
            message="config.STATUS_MAPPING est manquant.",
            expected="Dictionnaire de mapping des statuts",
        )
        return

    if not hasattr(config, "STATUTS_COMPTABILISES"):
        add_issue(
            issues=issues,
            severity="CRITICAL",
            error_type="Configuration invalide",
            dataframe="config",
            column="STATUTS_COMPTABILISES",
            message="config.STATUTS_COMPTABILISES est manquant.",
            expected="Set de statuts comptabilisés",
        )
        return

    expected_statuses = get_expected_clean_statuses()

    unknown_comptabilises = set(config.STATUTS_COMPTABILISES) - expected_statuses

    for status in unknown_comptabilises:
        add_issue(
            issues=issues,
            severity="CRITICAL",
            error_type="Configuration incohérente",
            dataframe="config",
            column="STATUTS_COMPTABILISES",
            message="Un statut comptabilisé n'existe pas dans STATUS_MAPPING.",
            value=status,
            expected=f"Un des statuts connus : {', '.join(sorted(expected_statuses))}",
        )

def validate_session_types(df: pd.DataFrame, issues: list):
    if "type_seance" not in df.columns:
        return

    mask = ~df["type_seance"].isin(config.EXPECTED_SESSION_TYPES)

    for _, row in df[mask].iterrows():
        add_issue_from_row(
            issues=issues,
            severity="WARNING",
            error_type="Type de séance inconnu",
            dataframe="data_clean",
            column="type_seance",
            message="Type de séance non reconnu.",
            row=row,
            expected=", ".join(sorted(config.EXPECTED_SESSION_TYPES)),
        )


def validate_dates(df: pd.DataFrame, issues: list):
    if "date" not in df.columns:
        return

    min_date = pd.to_datetime(config.MIN_TRAINING_DATE)
    max_date = pd.to_datetime(config.MAX_TRAINING_DATE)

    mask_missing = df["date"].isna()

    for _, row in df[mask_missing].iterrows():
        add_issue_from_row(
            issues=issues,
            severity="CRITICAL",
            error_type="Date invalide",
            dataframe="data_clean",
            column="date",
            message="Date d'entraînement manquante ou invalide.",
            row=row,
            expected="Date valide",
        )

    mask_outside = (
        df["date"].notna()
        & (
            (df["date"] < min_date)
            | (df["date"] > max_date)
        )
    )

    for _, row in df[mask_outside].iterrows():
        add_issue_from_row(
            issues=issues,
            severity="CRITICAL",
            error_type="Date hors saison",
            dataframe="data_clean",
            column="date",
            message="Date d'entraînement hors de la saison attendue.",
            row=row,
            expected=f"{config.MIN_TRAINING_DATE} à {config.MAX_TRAINING_DATE}",
        )

    if "date_naissance" in df.columns:
        mask_birth_missing = df["date_naissance"].isna()

        for _, row in df[mask_birth_missing].iterrows():
            add_issue_from_row(
                issues=issues,
                severity="CRITICAL",
                error_type="Date de naissance invalide",
                dataframe="data_clean",
                column="date_naissance",
                message="Date de naissance manquante ou invalide.",
                row=row,
                expected="Date valide",
            )

    if "annee" in df.columns:
        mask_bad_year = (
            df["date"].notna()
            & df["annee"].notna()
            & (df["annee"] != df["date"].dt.year)
        )

        for _, row in df[mask_bad_year].iterrows():
            add_issue_from_row(
                issues=issues,
                severity="CRITICAL",
                error_type="Incohérence date",
                dataframe="data_clean",
                column="annee",
                message="La colonne annee ne correspond pas à la date.",
                row=row,
                value=row["annee"],
                expected=row["date"].year,
            )

    if "mois_numero" in df.columns:
        mask_bad_month = (
            df["date"].notna()
            & df["mois_numero"].notna()
            & (df["mois_numero"] != df["date"].dt.month)
        )

        for _, row in df[mask_bad_month].iterrows():
            add_issue_from_row(
                issues=issues,
                severity="CRITICAL",
                error_type="Incohérence date",
                dataframe="data_clean",
                column="mois_numero",
                message="La colonne mois_numero ne correspond pas à la date.",
                row=row,
                value=row["mois_numero"],
                expected=row["date"].month,
            )


def validate_age(df: pd.DataFrame, issues: list):
    if "age_seance" not in df.columns:
        return

    mask = (
        df["age_seance"].isna()
        | (df["age_seance"] < config.CLEAN_UP_MIN_AGE_WARNING)
        | (df["age_seance"] > config.CLEAN_UP_MAX_AGE_WARNING)
    )

    for _, row in df[mask].iterrows():
        add_issue_from_row(
            issues=issues,
            severity="WARNING",
            error_type="Âge inhabituel",
            dataframe="data_clean",
            column="age_seance",
            message="Âge calculé inhabituel. Vérifier la date de naissance.",
            row=row,
            value=row["age_seance"],
            expected=f"Âge entre {config.CLEAN_UP_MIN_AGE_WARNING} et {config.CLEAN_UP_MAX_AGE_WARNING} ans",
        )


def validate_source_location(df: pd.DataFrame, issues: list):
    columns = [
        "sheet_name",
        "ligne_excel",
        "colonne_excel",
    ]

    if any(column not in df.columns for column in columns):
        return

    mask = (
        df["sheet_name"].isna()
        | (df["sheet_name"].astype(str).str.strip() == "")
        | df["ligne_excel"].isna()
        | df["colonne_excel"].isna()
    )

    for _, row in df[mask].iterrows():
        add_issue_from_row(
            issues=issues,
            severity="CRITICAL",
            error_type="Localisation source manquante",
            dataframe="data_clean",
            column="source",
            message="Impossible de retrouver précisément la cellule dans Excel.",
            row=row,
            expected="sheet_name + ligne_excel + colonne_excel",
        )


def validate_swimmer_id_consistency(df: pd.DataFrame, issues: list):
    required = [
        "nageur_id",
        "nom",
        "prenom",
        "date_naissance",
    ]

    if any(column not in df.columns for column in required):
        return

    expected_id = (
        df["nom"].astype(str)
        + "_"
        + df["prenom"].astype(str)
        + "_"
        + df["date_naissance"].dt.strftime("%Y-%m-%d").fillna("")
    )

    mask = df["nageur_id"] != expected_id

    for _, row in df[mask].iterrows():
        add_issue_from_row(
            issues=issues,
            severity="WARNING",
            error_type="Identifiant nageur incohérent",
            dataframe="data_clean",
            column="nageur_id",
            message="nageur_id incohérent avec nom, prénom et date de naissance.",
            row=row,
            value=row["nageur_id"],
            expected=expected_id.loc[row.name],
        )


def validate_exclusions(df: pd.DataFrame, issues: list):
    """
    Vérifie que les nageurs listés dans config.EXCLUDED_SWIMMERS existent bien
    dans le dataset importé.
    """
    if "nageur_id" not in df.columns:
        return

    existing_ids = set(df["nageur_id"].unique())

    for swimmer_id, reason in config.EXCLUDED_SWIMMERS.items():
        if swimmer_id not in existing_ids:
            add_issue(
                issues=issues,
                severity="WARNING",
                error_type="Exclusion introuvable",
                dataframe="config",
                column="EXCLUDED_SWIMMERS",
                message="Un nageur exclu dans config.py n'existe pas dans data_clean.",
                nageur_id=swimmer_id,
                value=reason,
                expected="Identifiant présent dans data_clean",
            )


# ============================================================
# VALIDATION TABLES ANNEXES
# ============================================================

def validate_nageurs(df_nageurs: pd.DataFrame, issues: list):
    required_columns = [
        "nageur_id",
        "nom",
        "prenom",
        "date_naissance",
    ]

    validate_required_columns(
        df=df_nageurs,
        dataframe_name="nageurs",
        required_columns=required_columns,
        issues=issues,
    )

    validate_duplicates(
        df=df_nageurs,
        dataframe_name="nageurs",
        subset=["nageur_id"],
        issues=issues,
        message="Doublon de nageur_id dans la table nageurs.",
    )

    for column in required_columns:
        if column not in df_nageurs.columns:
            continue

        if pd.api.types.is_string_dtype(df_nageurs[column]) or df_nageurs[column].dtype == "object":
            mask = (
                df_nageurs[column].isna()
                | (df_nageurs[column].astype(str).str.strip() == "")
            )
        else:
            mask = df_nageurs[column].isna()

        for _, row in df_nageurs[mask].iterrows():
            add_issue(
                issues=issues,
                severity="CRITICAL",
                error_type="Valeur obligatoire manquante",
                dataframe="nageurs",
                column=column,
                message="Valeur obligatoire manquante dans la table nageurs.",
                nom=row.get("nom", ""),
                prenom=row.get("prenom", ""),
                nageur_id=row.get("nageur_id", ""),
                value=row.get(column, ""),
                expected="Valeur non vide",
            )


def validate_seances(df_seances: pd.DataFrame, issues: list):
    required_columns = [
        "seance_id",
        "sheet_name",
        "mois",
        "date",
        "jour_code",
        "type_seance",
        "colonne_excel",
    ]

    validate_required_columns(
        df=df_seances,
        dataframe_name="seances",
        required_columns=required_columns,
        issues=issues,
    )

    validate_duplicates(
        df=df_seances,
        dataframe_name="seances",
        subset=["seance_id"],
        issues=issues,
        message="Doublon de seance_id dans la table seances.",
    )

    validate_duplicates(
        df=df_seances,
        dataframe_name="seances",
        subset=["sheet_name", "date", "colonne_excel"],
        issues=issues,
        message="Doublon de séance source dans la table seances.",
    )


def validate_referential_integrity(
    df_clean: pd.DataFrame,
    df_nageurs: pd.DataFrame,
    df_seances: pd.DataFrame,
    issues: list,
):
    if "nageur_id" in df_clean.columns and "nageur_id" in df_nageurs.columns:
        known_swimmers = set(df_nageurs["nageur_id"].dropna())

        mask = ~df_clean["nageur_id"].isin(known_swimmers)

        for _, row in df_clean[mask].iterrows():
            add_issue_from_row(
                issues=issues,
                severity="WARNING",
                error_type="Nageur absent de la table Groupe",
                dataframe="data_clean",
                column="nageur_id",
                message="Nageur présent dans les feuilles mensuelles mais absent de la table nageurs.",
                row=row,
                expected="Nageur présent dans la feuille Groupe",
            )

    if "seance_id" in df_clean.columns and "seance_id" in df_seances.columns:
        known_sessions = set(df_seances["seance_id"].dropna())

        mask = ~df_clean["seance_id"].isin(known_sessions)

        for _, row in df_clean[mask].iterrows():
            add_issue_from_row(
                issues=issues,
                severity="CRITICAL",
                error_type="Séance absente de la table séances",
                dataframe="data_clean",
                column="seance_id",
                message="Séance présente dans data_clean mais absente de la table seances.",
                row=row,
                expected="seance_id présent dans seances.pkl",
            )


# ============================================================
# VALIDATION AUDIT SOURCE
# ============================================================

def validate_audit_dates(audit_dates: pd.DataFrame, issues: list):
    if audit_dates.empty:
        return

    if "issue" in audit_dates.columns:
        missing_sheets = audit_dates[
            audit_dates["issue"].astype(str).str.strip() != ""
        ]

        for _, row in missing_sheets.iterrows():
            add_issue(
                issues=issues,
                severity="CRITICAL",
                error_type="Feuille mensuelle absente",
                dataframe="audit_dates",
                column="sheet_name",
                message=row.get("issue", "Feuille mensuelle absente"),
                sheet_name=row.get("sheet_name", ""),
                mois=row.get("sheet_name", ""),
                expected="Feuille présente dans le fichier Excel source",
            )

    if "parsed_date" in audit_dates.columns:
        invalid_dates = audit_dates[
            audit_dates["parsed_date"].isna()
            & (
                audit_dates.get("issue", "")
                .astype(str)
                .str.strip()
                == ""
            )
        ]

        for _, row in invalid_dates.iterrows():
            column_number = row.get("column", "")
            cell = make_excel_cell(config.DATE_ROW, column_number)

            add_issue(
                issues=issues,
                severity="CRITICAL",
                error_type="Date d'en-tête invalide",
                dataframe="audit_dates",
                column="raw_date",
                message="Date d'entraînement impossible à interpréter dans l'en-tête.",
                sheet_name=row.get("sheet_name", ""),
                mois=row.get("sheet_name", ""),
                ligne_excel=config.DATE_ROW,
                colonne_excel=column_number,
                cellule_excel=cell,
                value=row.get("raw_date", ""),
                expected="Date valide : 01/09, 2026-07-01, ou vraie date Excel",
            )

    if "jour_code" in audit_dates.columns:
        missing_day_code = audit_dates[
            audit_dates["jour_code"].isna()
            | (audit_dates["jour_code"].astype(str).str.strip() == "")
        ]

        for _, row in missing_day_code.iterrows():
            column_number = row.get("column", "")
            cell = make_excel_cell(config.DAY_CODE_ROW, column_number)

            add_issue(
                issues=issues,
                severity="WARNING",
                error_type="Code jour manquant",
                dataframe="audit_dates",
                column="jour_code",
                message="Code jour/séance manquant dans l'en-tête.",
                sheet_name=row.get("sheet_name", ""),
                mois=row.get("sheet_name", ""),
                ligne_excel=config.DAY_CODE_ROW,
                colonne_excel=column_number,
                cellule_excel=cell,
                value=row.get("jour_code", ""),
                expected="Code jour : l, m, mAM, j, jAM, s, d...",
            )

    if "date_mismatch" in audit_dates.columns:
        mismatches = audit_dates[audit_dates["date_mismatch"] == True]

        for _, row in mismatches.iterrows():
            column_number = row.get("column", "")
            cell = make_excel_cell(config.DATE_ROW, column_number)

            add_issue(
                issues=issues,
                severity="CRITICAL",
                error_type="Date dans le mauvais mois",
                dataframe="audit_dates",
                column="raw_date",
                message="La date ne correspond pas au mois de la feuille.",
                sheet_name=row.get("sheet_name", ""),
                mois=row.get("sheet_name", ""),
                ligne_excel=config.DATE_ROW,
                colonne_excel=column_number,
                cellule_excel=cell,
                value=row.get("raw_date", ""),
                expected=f"Mois attendu : {row.get('expected_month', '')}",
            )


def validate_audit_statuts(audit_statuts: pd.DataFrame, issues: list):
    if audit_statuts.empty:
        return

    for _, row in audit_statuts.iterrows():
        column_number = row.get("column", "")
        row_number = row.get("row", "")
        cell = make_excel_cell(row_number, column_number)

        add_issue(
            issues=issues,
            severity="CRITICAL",
            error_type="Statut source inconnu",
            dataframe="audit_statuts",
            column="value",
            message="Statut présent dans Excel mais absent du mapping config.STATUS_MAPPING.",
            sheet_name=row.get("sheet_name", ""),
            mois=row.get("sheet_name", ""),
            nom=row.get("nom", ""),
            prenom=row.get("prenom", ""),
            ligne_excel=row_number,
            colonne_excel=column_number,
            cellule_excel=cell,
            value=row.get("value", ""),
            expected="Corriger Excel ou ajouter le code dans config.STATUS_MAPPING",
        )


# ============================================================
# SYNTHESES
# ============================================================

def build_issues_dataframe(issues: list) -> pd.DataFrame:
    columns = [
        "severity",
        "error_type",
        "dataframe",
        "column",
        "message",
        "sheet_name",
        "mois",
        "date",
        "nom",
        "prenom",
        "nageur_id",
        "ligne_excel",
        "colonne_excel",
        "cellule_excel",
        "value",
        "expected",
    ]

    if not issues:
        return pd.DataFrame(columns=columns)

    return pd.DataFrame(issues)[columns]


def build_summary(issues_df: pd.DataFrame) -> pd.DataFrame:
    if issues_df.empty:
        return pd.DataFrame([
            {
                "severity": "OK",
                "error_type": "Aucun problème",
                "count": 0,
            }
        ])

    return (
        issues_df
        .groupby(["severity", "error_type"])
        .size()
        .reset_index(name="count")
        .sort_values(
            by=["severity", "count"],
            ascending=[True, False]
        )
    )


def build_quality_summary(df_clean: pd.DataFrame) -> pd.DataFrame:
    rows = []

    rows.append({
        "metric": "Lignes data_clean",
        "value": len(df_clean),
    })

    if "nageur_id" in df_clean.columns:
        rows.append({
            "metric": "Nageurs uniques",
            "value": df_clean["nageur_id"].nunique(),
        })

    if "seance_id" in df_clean.columns:
        rows.append({
            "metric": "Séances uniques",
            "value": df_clean["seance_id"].nunique(),
        })

    if "statut_clean" in df_clean.columns:
        for statut, count in df_clean["statut_clean"].value_counts(dropna=False).items():
            rows.append({
                "metric": f"Statut - {statut}",
                "value": count,
            })

    if "exclude_from_group_analysis" in df_clean.columns:
        rows.append({
            "metric": "Lignes exclues analyse groupe",
            "value": int(df_clean["exclude_from_group_analysis"].sum()),
        })

        excluded_swimmers = (
            df_clean[df_clean["exclude_from_group_analysis"] == True]
            ["nageur_id"]
            .nunique()
        )

        rows.append({
            "metric": "Nageurs exclus analyse groupe",
            "value": excluded_swimmers,
        })

    return pd.DataFrame(rows)


# ============================================================
# AFFICHAGE TERMINAL GROUPE ET COLORE
# ============================================================

def severity_color(severity):
    if severity == "CRITICAL":
        return Colors.RED

    if severity == "WARNING":
        return Colors.YELLOW

    return Colors.GREY


def print_issues_to_terminal(issues_df: pd.DataFrame, max_examples_per_group=config.CLEAN_UP_TERMINAL_MAX_EXAMPLES_PER_GROUP):
    print_title("RAPPORT CLEAN UP")

    if issues_df.empty:
        print_success("Aucune erreur détectée.")
        return

    critical_count = int((issues_df["severity"] == "CRITICAL").sum())
    warning_count = int((issues_df["severity"] == "WARNING").sum())

    if critical_count > 0:
        print_error(f"Erreurs critiques : {critical_count}")
    else:
        print_success("Erreurs critiques : 0")

    if warning_count > 0:
        print_warning(f"Warnings : {warning_count}")
    else:
        print_success("Warnings : 0")

    summary = build_summary(issues_df)

    print_section("Résumé par type d'erreur")

    for _, row in summary.iterrows():
        severity = row["severity"]
        error_type = row["error_type"]
        count = row["count"]

        color = severity_color(severity)

        print(
            color_text(
                f"{severity:<10} | {count:>5} | {error_type}",
                color
            )
        )

    print_section("Détail des erreurs groupées")

    sort_order = {
        "CRITICAL": 0,
        "WARNING": 1,
        "INFO": 2,
    }

    issues_df = issues_df.copy()
    issues_df["_sort_order"] = issues_df["severity"].map(sort_order).fillna(99)

    grouped = (
        issues_df
        .sort_values(["_sort_order", "error_type", "sheet_name", "nom", "date"])
        .groupby(["severity", "error_type"], dropna=False)
    )

    for (severity, error_type), group in grouped:
        color = severity_color(severity)

        print()
        print(color_text(f"{severity} — {error_type} ({len(group)} erreur(s))", color))

        first_message = group["message"].iloc[0]
        print(color_text(f"Message : {first_message}", color))

        display_group = group.head(max_examples_per_group)

        for _, row in display_group.iterrows():
            sheet = row.get("sheet_name", "")
            mois = row.get("mois", "")
            date = row.get("date", "")
            nom = row.get("nom", "")
            prenom = row.get("prenom", "")
            cell = row.get("cellule_excel", "")
            value = row.get("value", "")
            expected = row.get("expected", "")
            column = row.get("column", "")

            nageur_txt = f"{nom} {prenom}".strip()
            if not nageur_txt:
                nageur_txt = row.get("nageur_id", "")

            location_parts = []

            if sheet:
                location_parts.append(f"feuille={sheet}")

            if cell:
                location_parts.append(f"cellule={cell}")

            if date:
                location_parts.append(f"date={date}")

            if nageur_txt:
                location_parts.append(f"nageur={nageur_txt}")

            location = " | ".join(location_parts)

            print(
                color_text(
                    f"  - {location}",
                    color
                )
            )

            print(
                color_text(
                    f"    colonne={column} | valeur='{value}' | attendu='{expected}'",
                    Colors.GREY
                )
            )

        remaining = len(group) - len(display_group)

        if remaining > 0:
            print(
                color_text(
                    f"  ... {remaining} autre(s) erreur(s) du même type. Voir le rapport Excel complet.",
                    Colors.GREY
                )
            )

    issues_df.drop(columns=["_sort_order"], inplace=True)


# ============================================================
# EXPORT
# ============================================================

def export_results(
    df_clean: pd.DataFrame,
    df_nageurs: pd.DataFrame,
    df_seances: pd.DataFrame,
    issues_df: pd.DataFrame,
):
    config.CLEAN_UP_DIR.mkdir(parents=True, exist_ok=True)

    df_clean.to_pickle(config.DATA_CLEAN_VALIDATED_PICKLE)
    df_nageurs.to_pickle(config.NAGEURS_VALIDATED_PICKLE)
    df_seances.to_pickle(config.SEANCES_VALIDATED_PICKLE)
    issues_df.to_pickle(config.CLEAN_UP_ISSUES_PICKLE)

    summary_df = build_summary(issues_df)
    quality_df = build_quality_summary(df_clean)

    critical_df = (
        issues_df[issues_df["severity"] == "CRITICAL"]
        if not issues_df.empty
        else pd.DataFrame()
    )

    warnings_df = (
        issues_df[issues_df["severity"] == "WARNING"]
        if not issues_df.empty
        else pd.DataFrame()
    )

    with pd.ExcelWriter(config.CLEAN_UP_REPORT_XLSX, engine="openpyxl") as writer:
        summary_df.to_excel(
            writer,
            sheet_name="summary",
            index=False
        )

        quality_df.to_excel(
            writer,
            sheet_name="quality_summary",
            index=False
        )

        issues_df.to_excel(
            writer,
            sheet_name="all_issues",
            index=False
        )

        critical_df.to_excel(
            writer,
            sheet_name="critical",
            index=False
        )

        warnings_df.to_excel(
            writer,
            sheet_name="warnings",
            index=False
        )

        df_clean.head(1000).to_excel(
            writer,
            sheet_name="sample_data_clean",
            index=False
        )

        df_nageurs.to_excel(
            writer,
            sheet_name="nageurs",
            index=False
        )

        df_seances.to_excel(
            writer,
            sheet_name="seances",
            index=False
        )

    print_section("Exports générés")
    print_success(f"Données validées : {config.DATA_CLEAN_VALIDATED_PICKLE}")
    print_success(f"Nageurs validés : {config.NAGEURS_VALIDATED_PICKLE}")
    print_success(f"Séances validées : {config.SEANCES_VALIDATED_PICKLE}")
    print_success(f"Erreurs Pickle : {config.CLEAN_UP_ISSUES_PICKLE}")
    print_success(f"Rapport Excel : {config.CLEAN_UP_REPORT_XLSX}")


# ============================================================
# PIPELINE CLEAN UP
# ============================================================

def run_clean_up():
    ensure_directories()

    print_title("02 CLEAN UP — CONTROLE DES DONNEES")

    (
        df_clean,
        df_nageurs,
        df_seances,
        audit_global,
        audit_dates,
        audit_statuts,
    ) = load_imported_pickles()

    issues = []

    validate_config_status_consistency(issues)
    
    print_section("Normalisation des formats")

    df_clean = normalize_data_clean(df_clean)
    df_nageurs = normalize_nageurs(df_nageurs)
    df_seances = normalize_seances(df_seances)
    audit_dates = normalize_audit_dates(audit_dates)
    audit_statuts = normalize_audit_statuts(audit_statuts)

    print_success("Formats normalisés")

    print_section("Validation structurelle")

    validate_required_columns(
        df=df_clean,
        dataframe_name="data_clean",
        required_columns=get_required_data_clean_columns(),
        issues=issues,
    )

    validate_required_values(
        df=df_clean,
        dataframe_name="data_clean",
        required_columns=get_required_non_empty_data_columns(),
        issues=issues,
    )

    validate_duplicates(
        df=df_clean,
        dataframe_name="data_clean",
        subset=["nageur_id", "seance_id"],
        issues=issues,
        message="Doublon nageur/séance dans data_clean.",
    )

    print_success("Validation structurelle terminée")

    print_section("Validation métier")

    validate_status_values(df_clean, issues)
    validate_binary_columns(df_clean, issues)
    validate_one_status_per_row(df_clean, issues)
    validate_status_binary_consistency(df_clean, issues)
    validate_seance_comptabilisee(df_clean, issues)
    validate_session_types(df_clean, issues)
    validate_dates(df_clean, issues)
    validate_age(df_clean, issues)
    validate_source_location(df_clean, issues)
    validate_swimmer_id_consistency(df_clean, issues)
    validate_exclusions(df_clean, issues)

    print_success("Validation métier terminée")

    print_section("Validation tables annexes")

    validate_nageurs(df_nageurs, issues)
    validate_seances(df_seances, issues)
    validate_referential_integrity(df_clean, df_nageurs, df_seances, issues)

    print_success("Validation tables annexes terminée")

    print_section("Validation source Excel")

    validate_audit_dates(audit_dates, issues)
    validate_audit_statuts(audit_statuts, issues)

    print_success("Validation source Excel terminée")

    issues_df = build_issues_dataframe(issues)

    print_issues_to_terminal(
        issues_df=issues_df,
        max_examples_per_group=config.CLEAN_UP_TERMINAL_MAX_EXAMPLES_PER_GROUP
    )

    export_results(
        df_clean=df_clean,
        df_nageurs=df_nageurs,
        df_seances=df_seances,
        issues_df=issues_df,
    )

    critical_count = (
        int((issues_df["severity"] == "CRITICAL").sum())
        if not issues_df.empty
        else 0
    )

    if critical_count > 0 and config.STOP_ON_CRITICAL_ERRORS:
        raise ValueError(
            f"Clean up terminé avec {critical_count} erreur(s) critique(s). "
            f"Voir le rapport : {config.CLEAN_UP_REPORT_XLSX}"
        )

    if critical_count > 0:
        print_error(
            f"Clean up terminé avec {critical_count} erreur(s) critique(s). "
            "Corriger le fichier source avant l'analyse finale."
        )
    else:
        print_success("Clean up terminé sans erreur critique.")

    return df_clean, df_nageurs, df_seances, issues_df


def main(group=None):
    if group is not None:
        config.set_active_group(group)

    run_clean_up()


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--group",
        default=None,
        choices=config.available_groups(),
        help="Groupe à traiter.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    main(group=args.group)
