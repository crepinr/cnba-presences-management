#!/usr/bin/env python3
"""
05_export_attendance_sql.py

Génère un fichier SQL annuel importable manuellement dans phpMyAdmin
à partir des pickles validés du projet Presence Analytics.

Chaîne prévue :
    00_download_data.py
    01_import_data.py
    02_clean_up.py
    05_export_attendance_sql.py

Principes :
- un fichier SQL annuel par groupe ;
- lecture de config.DATA_CLEAN_VALIDATED_PICKLE ;
- contrôle des membres et du groupe dans la DB locale ;
- cellule vide = aucun attendance_record ;
- une attendance_session est créée seulement si au moins un statut reconnu
  existe pour la date et le créneau ;
- réimport = remplacement transactionnel de toute la saison pour le groupe ;
- aucune connexion à la DB de production ;
- le SQL généré est compatible MySQL 5.6 / phpMyAdmin.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable, Sequence

import pandas as pd
import pymysql

import config

try:
    import attendance_yaml_config
    attendance_yaml_config.bootstrap(config)
except Exception as yaml_config_error:
    # Si aucun YAML n'est configuré, on conserve le comportement historique.
    if "CNBA_CONFIG_YAML" in __import__("os").environ:
        raise

DB_HOST = os.environ.get("CNBA_DB_HOST", "localhost")
DB_USER = os.environ.get("CNBA_DB_USER", "root")
DB_NAME = os.environ.get("CNBA_DB_NAME", "pcxa_cnba")
SQL_BATCH_SIZE = 500

SOURCE_STATUS_TO_DB = {
    "V": "PRESENT",
    "C": "COMPETITION",
    "X": "ABSENT",
    "O": "PLANNED_ABSENCE",
    "E": "PLANNED_ABSENCE",
    "M": "SICKNESS",
}
EMPTY_SOURCE_CODES = {"", "NAN", "NONE", "<NA>"}
REQUIRED_DB_STATUS_CODES = {
    "PRESENT",
    "COMPETITION",
    "ABSENT",
    "PLANNED_ABSENCE",
    "SICKNESS",
}
MEMBER_REQUIRED_COLUMNS = {"id", "nom", "prenom", "date_naissance"}
GROUP_TEXT_COLUMN_CANDIDATES = (
    "nom", "name", "groupe", "nom_groupe", "group_name", "label",
    "libelle", "libellé", "description", "code",
)
REQUIRED_PICKLE_COLUMNS = {
    "nom", "prenom", "date_naissance", "date", "jour_code",
    "type_seance", "statut_original",
}


@dataclass(frozen=True)
class MemberMatch:
    source_name: str
    source_firstname: str
    source_birthdate: date
    member_id: int
    member_group_id: int
    db_name: str
    db_firstname: str
    db_birthdate: date


@dataclass(frozen=True)
class AttendanceRow:
    session_date: date
    time_slot: str
    group_id: int
    member_id: int
    status_code: str
    source_group_name: str = ""
    source_group_label: str = ""


class Color:
    GREEN = "\033[32m"
    YELLOW = "\033[33m"
    RED = "\033[31m"
    CYAN = "\033[36m"
    RESET = "\033[0m"


def info(message: str) -> None:
    print(f"{Color.CYAN}ℹ {message}{Color.RESET}")


def success(message: str) -> None:
    print(f"{Color.GREEN}✓ {message}{Color.RESET}")


def warning(message: str) -> None:
    print(f"{Color.YELLOW}⚠ {message}{Color.RESET}")


def fail(message: str) -> None:
    raise RuntimeError(message)


def clean_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()


def strip_accents(value: str) -> str:
    return "".join(
        char for char in unicodedata.normalize("NFD", value)
        if unicodedata.category(char) != "Mn"
    )


def normalize_identity(value) -> str:
    return strip_accents(clean_text(value)).upper()


def normalize_status(value) -> str:
    return clean_text(value).upper()


def normalize_date(value, label: str) -> date:
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        fail(f"Date invalide pour {label}: {value!r}")
    return parsed.date()


def sql_quote(value: str) -> str:
    escaped = str(value).replace("\\", "\\\\").replace("'", "''")
    return f"'{escaped}'"


def sql_date(value: date) -> str:
    return sql_quote(value.isoformat())


def sql_identifier(value: str) -> str:
    """Quote un identifiant MySQL avec backticks."""
    return "`" + str(value).replace("`", "``") + "`"


def slugify(value: str) -> str:
    value = strip_accents(clean_text(value)).lower()
    return re.sub(r"[^a-z0-9]+", "_", value).strip("_") or "groupe"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Exporte les présences validées vers un SQL annuel phpMyAdmin."
    )
    parser.add_argument(
        "--group",
        action="append",
        default=None,
        help=(
            "Groupe à exporter. Peut être indiqué plusieurs fois. "
            "Compatible avec l'ancien usage mono-groupe."
        ),
    )
    parser.add_argument(
        "--groups",
        default=None,
        help=(
            "Liste de groupes séparés par une virgule ou un point-virgule. "
            "Exemple : 'Elite,Elite Jeunes'."
        ),
    )
    parser.add_argument("--group-id", type=int, default=None)
    parser.add_argument(
        "--group-id-map",
        default=None,
        help=(
            "Correspondances optionnelles groupe=id, séparées par ';'. "
            "Exemple : 'Elite=1;Elite Jeunes=2'."
        ),
    )
    parser.add_argument("--season-code", default=None)
    parser.add_argument("--output", type=Path, default=None)
    parser.add_argument("--db-host", default=DB_HOST)
    parser.add_argument("--db-user", default=DB_USER)
    parser.add_argument("--db-name", default=DB_NAME)
    parser.add_argument("--allow-unmatched", action="store_true")
    return parser.parse_args()


def split_group_list(value: str | None) -> list[str]:
    if not value:
        return []
    return [
        clean_text(item)
        for item in re.split(r"[,;]", value)
        if clean_text(item)
    ]


def resolve_requested_groups(args: argparse.Namespace) -> list[str]:
    groups: list[str] = []
    groups.extend(split_group_list(os.environ.get("CNBA_GROUPS")))
    groups.extend(split_group_list(args.groups))

    if args.group:
        for item in args.group:
            groups.extend(split_group_list(item))

    legacy_group = os.environ.get("CNBA_GROUP", getattr(config, "ACTIVE_GROUP", None))
    if not groups and legacy_group:
        groups.append(clean_text(legacy_group))

    deduped: list[str] = []
    seen = set()
    for group in groups:
        key = normalize_identity(group)
        if key not in seen:
            seen.add(key)
            deduped.append(group)

    return deduped


def parse_group_id_map(value: str | None) -> dict[str, int]:
    mapping: dict[str, int] = {}
    if not value:
        return mapping

    for item in value.split(";"):
        item = clean_text(item)
        if not item:
            continue
        if "=" not in item:
            fail(f"Entrée --group-id-map invalide : {item!r}")
        name, raw_id = item.split("=", 1)
        name = clean_text(name)
        try:
            mapping[normalize_identity(name)] = int(clean_text(raw_id))
        except ValueError:
            fail(f"ID groupe invalide dans --group-id-map : {item!r}")

    return mapping


def activate_group(group_name: str) -> None:
    if not group_name:
        fail('Aucun groupe défini. Utiliser --group "Elite" ou "Elite Jeunes".')
    if hasattr(config, "set_active_group"):
        config.set_active_group(group_name)
    else:
        warning("config.set_active_group() absent : chemins config actuels utilisés.")


def get_validated_pickle_path() -> Path:
    path = getattr(config, "DATA_CLEAN_VALIDATED_PICKLE", None)
    if path is None:
        fail("config.DATA_CLEAN_VALIDATED_PICKLE est absent.")
    path = Path(path)
    if not path.exists():
        fail(
            f"Pickle validé introuvable : {path}\n"
            "Exécuter d'abord 01_import_data.py puis 02_clean_up.py."
        )
    return path


def load_validated_data() -> pd.DataFrame:
    path = get_validated_pickle_path()
    info(f"Lecture du pickle validé : {path}")
    df = pd.read_pickle(path).copy()
    missing = sorted(REQUIRED_PICKLE_COLUMNS - set(df.columns))
    if missing:
        fail("Colonnes manquantes : " + ", ".join(missing))
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["date_naissance"] = pd.to_datetime(df["date_naissance"], errors="coerce")
    if df["date"].isna().any():
        fail("Le pickle contient des dates de séance invalides.")
    if df["date_naissance"].isna().any():
        fail("Le pickle contient des dates de naissance invalides.")
    return df


def validate_season_dates(df: pd.DataFrame) -> None:
    start = date(int(config.SEASON_START_YEAR), 8, 24)
    end = date(int(config.SEASON_END_YEAR), 7, 31)
    dates = df["date"].dt.date
    bad = df[(dates < start) | (dates > end)]
    if not bad.empty:
        fail(
            f"{len(bad)} ligne(s) hors saison {start} → {end}.\n"
            + bad[["nom", "prenom", "date"]].head(20).to_string(index=False)
        )


def start_mysql_server() -> None:
    """Démarre MySQL local. mysql.server gère lui-même le cas déjà démarré."""
    try:
        print("##### STARTING MYSQL SERVER #####")
        exit_code = os.system("mysql.server start")
        if exit_code != 0:
            fail(f"La commande `mysql.server start` a échoué (code {exit_code}).")
    except Exception as exc:
        fail(f"Erreur pendant le démarrage de MySQL : {exc}")


def stop_mysql_server() -> None:
    """Arrête MySQL local dans tous les cas à la fin du script."""
    try:
        print("##### STOPPING MYSQL SERVER #####")
        exit_code = os.system("mysql.server stop")
        if exit_code != 0:
            warning(f"La commande `mysql.server stop` a retourné le code {exit_code}.")
    except Exception as exc:
        warning(f"Erreur pendant l'arrêt de MySQL : {exc}")


def connect_db(args: argparse.Namespace):
    """Connexion volontairement simple à la DB locale."""
    info(f"Connexion DB locale : {args.db_user}@{args.db_host}/{args.db_name}")
    return pymysql.connect(
        host=args.db_host,
        user=args.db_user,
        database=args.db_name,
    )


def get_table_columns(connection, table_name: str) -> list[dict]:
    query = """
        SELECT COLUMN_NAME, COLUMN_TYPE, IS_NULLABLE, COLUMN_KEY
        FROM information_schema.COLUMNS
        WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = %s
        ORDER BY ORDINAL_POSITION
    """
    with connection.cursor(pymysql.cursors.DictCursor) as cursor:
        cursor.execute(query, (table_name,))
        return list(cursor.fetchall())


def fetch_dataframe(connection, query: str, params=None) -> pd.DataFrame:
    """Exécute une requête avec un curseur dictionnaire, sans SQLAlchemy."""
    with connection.cursor(pymysql.cursors.DictCursor) as cursor:
        cursor.execute(query, params or ())
        rows = cursor.fetchall()
    return pd.DataFrame(list(rows))


def require_non_empty_table(connection, table_name: str) -> int:
    with connection.cursor(pymysql.cursors.DictCursor) as cursor:
        cursor.execute(f"SELECT COUNT(*) AS row_count FROM `{table_name}`")
        row = cursor.fetchone()
    count = int(row["row_count"] if row else 0)
    if count == 0:
        fail(
            f"La table `{table_name}` est vide dans la DB locale. "
            "Export SQL annulé : aucune donnée ne sera générée."
        )
    return count


def load_members(connection) -> pd.DataFrame:
    columns = {row["COLUMN_NAME"] for row in get_table_columns(connection, "membres")}
    missing = MEMBER_REQUIRED_COLUMNS - columns
    if missing:
        fail("Colonnes manquantes dans membres : " + ", ".join(sorted(missing)))
    if "groupe" not in columns:
        fail("La table `membres` ne contient pas la colonne `groupe`.")
    selected = ["id", "nom", "prenom", "date_naissance", "groupe"]
    if "numero_membre" in columns:
        selected.append("numero_membre")
    require_non_empty_table(connection, "membres")
    query = "SELECT " + ", ".join(f"`{c}`" for c in selected) + " FROM `membres`"
    members = fetch_dataframe(connection, query)
    if members.empty:
        fail("La requête sur `membres` ne retourne aucune ligne. Export annulé.")
    members["date_naissance"] = pd.to_datetime(members["date_naissance"], errors="coerce")
    members = members[members["date_naissance"].notna()].copy()
    members["_nom_key"] = members["nom"].map(normalize_identity)
    members["_prenom_key"] = members["prenom"].map(normalize_identity)
    members["_birth_key"] = members["date_naissance"].dt.strftime("%Y-%m-%d")
    # Les doublons globaux de la table membres ne bloquent pas l'export.
    # Ils seront contrôlés uniquement s'ils correspondent à un nageur
    # réellement présent dans le fichier de présences à importer.
    success(f"{len(members)} membres chargés depuis la DB locale")
    return members


def resolve_group(connection, group_name: str, explicit_group_id: int | None) -> tuple[int, str]:
    columns_info = get_table_columns(connection, "groupes")
    columns = {row["COLUMN_NAME"] for row in columns_info}
    if "id" not in columns:
        fail("La table `groupes` ne contient pas de colonne `id`.")

    require_non_empty_table(connection, "groupes")

    # Cherche d'abord les noms connus, puis toute colonne texte exploitable.
    text_columns = [c for c in GROUP_TEXT_COLUMN_CANDIDATES if c in columns]
    if not text_columns:
        text_types = ("char", "varchar", "text", "tinytext", "mediumtext", "longtext")
        text_columns = [
            row["COLUMN_NAME"]
            for row in columns_info
            if str(row["COLUMN_TYPE"]).lower().split("(", 1)[0] in text_types
            and row["COLUMN_NAME"] != "id"
        ]

    selected = ["id"] + text_columns
    groups = fetch_dataframe(
        connection,
        "SELECT " + ", ".join(f"`{c}`" for c in selected) + " FROM `groupes`",
    )
    if groups.empty:
        fail("La table `groupes` ne retourne aucune ligne. Export annulé.")

    if explicit_group_id is not None:
        match = groups[groups["id"].astype(int) == explicit_group_id]
        if len(match) != 1:
            fail(f"--group-id {explicit_group_id} introuvable dans `groupes`.")
        row = match.iloc[0]
        label = next(
            (clean_text(row.get(c)) for c in text_columns if clean_text(row.get(c))),
            group_name,
        )
        return int(row["id"]), label

    if not text_columns:
        fail(
            "Aucune colonne texte exploitable dans `groupes`. "
            "Utiliser --group-id <id>."
        )

    target = normalize_identity(group_name)
    indexes = set()
    for idx, row in groups.iterrows():
        for column in text_columns:
            if normalize_identity(row.get(column)) == target:
                indexes.add(idx)
                break

    matches = groups.loc[sorted(indexes)]
    if len(matches) == 0:
        preview = groups[selected].head(50).to_string(index=False)
        fail(
            f"Groupe {group_name!r} non trouvé dans `groupes`. "
            "Utiliser --group-id si le libellé diffère.\n"
            f"Colonnes inspectées : {', '.join(text_columns)}\n"
            f"Aperçu :\n{preview}"
        )
    if len(matches) > 1:
        fail(
            f"Groupe {group_name!r} ambigu. Utiliser --group-id.\n"
            + matches[selected].to_string(index=False)
        )

    row = matches.iloc[0]
    label = next(
        (
            clean_text(row.get(c))
            for c in text_columns
            if normalize_identity(row.get(c)) == target
        ),
        group_name,
    )
    return int(row["id"]), label

def validate_local_attendance_schema(connection) -> None:
    required = {"attendance_seasons", "attendance_sessions", "attendance_records", "attendance_statuses"}
    with connection.cursor(pymysql.cursors.DictCursor) as cursor:
        cursor.execute("""
            SELECT TABLE_NAME FROM information_schema.TABLES
            WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME LIKE 'attendance_%'
        """)
        found = {row["TABLE_NAME"] for row in cursor.fetchall()}
    missing = required - found
    if missing:
        warning("Tables attendance absentes localement : " + ", ".join(sorted(missing)))
        return
    with connection.cursor(pymysql.cursors.DictCursor) as cursor:
        placeholders = ",".join(["%s"] * len(REQUIRED_DB_STATUS_CODES))
        cursor.execute(
            f"SELECT code FROM attendance_statuses WHERE code IN ({placeholders})",
            tuple(sorted(REQUIRED_DB_STATUS_CODES)),
        )
        found_statuses = {row["code"] for row in cursor.fetchall()}
    missing_statuses = REQUIRED_DB_STATUS_CODES - found_statuses
    if missing_statuses:
        fail("Statuts DB manquants : " + ", ".join(sorted(missing_statuses)))


def build_source_swimmers(df: pd.DataFrame) -> pd.DataFrame:
    swimmers = df[["nom", "prenom", "date_naissance"]].drop_duplicates().copy()
    swimmers["_nom_key"] = swimmers["nom"].map(normalize_identity)
    swimmers["_prenom_key"] = swimmers["prenom"].map(normalize_identity)
    swimmers["_birth_key"] = swimmers["date_naissance"].dt.strftime("%Y-%m-%d")
    duplicates = swimmers[swimmers.duplicated(["_nom_key", "_prenom_key", "_birth_key"], keep=False)]
    if not duplicates.empty:
        fail("Identités source dupliquées après normalisation.\n" + duplicates.to_string(index=False))
    return swimmers


def match_swimmers(source: pd.DataFrame, members: pd.DataFrame):
    """
    Associe uniquement les nageurs présents dans le fichier source.

    Règles :
    - 0 correspondance en DB  -> nageur non reconnu ;
    - 1 correspondance en DB  -> association valide ;
    - >1 correspondance en DB -> blocage uniquement pour ce nageur source.

    Les doublons présents ailleurs dans la table membres sont ignorés.
    """
    key_columns = ["_nom_key", "_prenom_key", "_birth_key"]

    source_keys = set(
        tuple(row)
        for row in source[key_columns].itertuples(index=False, name=None)
    )

    concerned_members = members[
        members[key_columns].apply(tuple, axis=1).isin(source_keys)
    ].copy()

    duplicate_members = concerned_members[
        concerned_members.duplicated(key_columns, keep=False)
    ].copy()

    if not duplicate_members.empty:
        duplicate_members = duplicate_members.sort_values(key_columns + ["id"])
        fail(
            "Plusieurs membres de la DB correspondent à un nageur présent "
            "dans le fichier de présences.\n"
            "Corriger ou désactiver le doublon avant l'export :\n"
            + duplicate_members[
                ["id", "nom", "prenom", "date_naissance"]
            ].to_string(index=False)
        )

    merged = source.merge(
        concerned_members,
        on=key_columns,
        how="left",
        suffixes=("_source", "_db"),
    )

    unmatched = merged[merged["id"].isna()].copy()
    matched = merged[merged["id"].notna()].copy()

    mapping = {}
    for _, row in matched.iterrows():
        key = (row["_nom_key"], row["_prenom_key"], row["_birth_key"])
        mapping[key] = MemberMatch(
            clean_text(row["nom_source"]),
            clean_text(row["prenom_source"]),
            pd.to_datetime(row["date_naissance_source"]).date(),
            int(row["id"]),
            int(row["groupe"]),
            clean_text(row["nom_db"]),
            clean_text(row["prenom_db"]),
            pd.to_datetime(row["date_naissance_db"]).date(),
        )

    return mapping, unmatched


def member_key(row: pd.Series) -> tuple[str, str, str]:
    birthdate = normalize_date(row["date_naissance"], f"{row.get('nom', '')} {row.get('prenom', '')}")
    return normalize_identity(row["nom"]), normalize_identity(row["prenom"]), birthdate.isoformat()


def derive_time_slot(row: pd.Series) -> str:
    session_type = normalize_identity(row.get("type_seance"))
    day_code = normalize_identity(row.get("jour_code"))
    return "AM" if session_type == "MATIN" or "AM" in day_code else "PM"


def build_member_labels(member_mapping: dict) -> dict[int, str]:
    """Construit un libellé lisible par membre DB pour les messages d'erreur."""
    labels: dict[int, str] = {}

    for match in member_mapping.values():
        labels[match.member_id] = (
            f"{match.db_name} {match.db_firstname} "
            f"({match.db_birthdate.isoformat()})"
        )

    return labels


def format_attendance_row_detail(item: AttendanceRow, member_labels: dict[int, str]) -> str:
    """Retourne une ligne de diagnostic lisible pour une présence exportée."""
    nageur = member_labels.get(item.member_id, f"membre_id={item.member_id}")
    source_group = clean_text(item.source_group_name) or "?"
    source_group_db = clean_text(item.source_group_label) or "?"

    return (
        f"source={source_group} | source_db={source_group_db} | "
        f"group_id_export={item.group_id} | date={item.session_date} | "
        f"créneau={item.time_slot} | membre_id={item.member_id} | "
        f"nageur={nageur} | statut={item.status_code}"
    )


def build_attendance_rows(
    df,
    member_mapping,
    allow_unmatched,
    source_group_name: str = "",
    source_group_label: str = "",
):
    rows = []
    status_counts = {code: 0 for code in sorted(REQUIRED_DB_STATUS_CODES)}
    empty_count = 0
    unknown = []
    unmatched_rows = []
    for index, row in df.iterrows():
        source_status = normalize_status(row["statut_original"])
        if source_status in EMPTY_SOURCE_CODES:
            empty_count += 1
            continue
        db_status = SOURCE_STATUS_TO_DB.get(source_status)
        if db_status is None:
            unknown.append(f"ligne {index}: {row.get('nom')} {row.get('prenom')} statut={source_status!r}")
            continue
        key = member_key(row)
        match = member_mapping.get(key)
        if match is None:
            unmatched_rows.append(f"ligne {index}: {row.get('nom')} {row.get('prenom')} {key[2]}")
            continue
        item = AttendanceRow(
            normalize_date(row["date"], f"séance ligne {index}"),
            derive_time_slot(row),
            match.member_group_id,
            match.member_id,
            db_status,
            source_group_name,
            source_group_label,
        )
        rows.append(item)
        status_counts[db_status] += 1
    if unknown:
        fail(f"{len(unknown)} statut(s) inconnu(s).\n" + "\n".join(unknown[:30]))
    if unmatched_rows and not allow_unmatched:
        fail(f"{len(unmatched_rows)} relevé(s) avec nageur non reconnu.\n" + "\n".join(unmatched_rows[:30]))
    if not rows:
        fail("Aucun relevé à exporter.")
    member_labels = build_member_labels(member_mapping)
    grouped = {}
    for item in rows:
        grouped.setdefault(
            (item.group_id, item.session_date, item.time_slot, item.member_id),
            [],
        ).append(item)

    duplicates = {key: values for key, values in grouped.items() if len(values) > 1}
    if duplicates:
        lines = []
        for _, values in list(duplicates.items())[:30]:
            for item in values:
                lines.append(format_attendance_row_detail(item, member_labels))
            lines.append("---")

        fail(
            "Doublons nageur/séance détectés dans un même groupe source.\n"
            + "\n".join(lines)
        )
    rows.sort(key=lambda x: (x.group_id, x.session_date, x.time_slot, x.member_id))
    return rows, status_counts, empty_count


def chunked(items: Sequence[AttendanceRow], size: int) -> Iterable[Sequence[AttendanceRow]]:
    for start in range(0, len(items), size):
        yield items[start:start + size]


def generate_sql(*, db_name, group_names, group_db_labels, season_code, rows, status_counts, source_pickles):
    sessions = sorted({(r.group_id, r.session_date, r.time_slot) for r in rows})
    involved_group_ids = sorted({r.group_id for r in rows})
    group_ids_sql = ", ".join(str(group_id) for group_id in involved_group_ids)

    group_labels_sql = ", ".join(
        f"{name} -> {label}" for name, label in zip(group_names, group_db_labels)
    )
    source_pickles_sql = ", ".join(str(path) for path in source_pickles)

    lines = [
        "-- =====================================================================",
        "-- IMPORT ANNUEL DES PRESENCES CNBA",
        f"-- Généré le          : {datetime.now():%Y-%m-%d %H:%M:%S}",
        f"-- Groupes source     : {', '.join(group_names)}",
        f"-- Groupes source DB  : {group_labels_sql}",
        f"-- Groupes réellement importés : {group_ids_sql}",
        f"-- Saison             : {season_code}",
        f"-- Pickles source     : {source_pickles_sql}",
        f"-- Séances            : {len(sessions)}",
        f"-- Relevés            : {len(rows)}",
        "-- Le groupe de chaque relevé provient de membres.groupe.",
        "-- =====================================================================",
        "",
        "SET NAMES utf8mb4;",
        f"USE {sql_identifier(db_name)};",
        "START TRANSACTION;",
        "",
        "CREATE TEMPORARY TABLE tmp_attendance_guard (",
        "    guard_key VARCHAR(100) NOT NULL PRIMARY KEY,",
        "    guard_id BIGINT NOT NULL",
        ") ENGINE=InnoDB;",
        "",
        "INSERT INTO tmp_attendance_guard (guard_key, guard_id) VALUES",
        f"    ('season', (SELECT id FROM attendance_seasons WHERE code = {sql_quote(season_code)} LIMIT 1)),",
    ]

    statuses = sorted(REQUIRED_DB_STATUS_CODES)
    for i, status in enumerate(statuses):
        comma = "," if i < len(statuses) - 1 else ";"
        lines.append(
            f"    ('status:{status}', (SELECT id FROM attendance_statuses "
            f"WHERE code = {sql_quote(status)} LIMIT 1)){comma}"
        )

    lines += [
        "",
        "SET @attendance_season_id := (SELECT guard_id FROM tmp_attendance_guard WHERE guard_key = 'season');",
        "",
        "CREATE TEMPORARY TABLE tmp_attendance_import (",
        "    group_id INT NOT NULL,",
        "    session_date DATE NOT NULL,",
        "    time_slot ENUM('AM', 'PM') NOT NULL,",
        "    membre_id INT NOT NULL,",
        "    status_code VARCHAR(30) NOT NULL,",
        "    PRIMARY KEY (group_id, session_date, time_slot, membre_id),",
        "    KEY idx_tmp_attendance_status (status_code)",
        ") ENGINE=InnoDB;",
        "",
    ]

    for batch_number, batch in enumerate(chunked(rows, SQL_BATCH_SIZE), 1):
        lines += [
            f"-- Lot {batch_number}",
            "INSERT INTO tmp_attendance_import (group_id, session_date, time_slot, membre_id, status_code) VALUES",
            ",\n".join(
                f"    ({r.group_id}, {sql_date(r.session_date)}, {sql_quote(r.time_slot)}, {r.member_id}, {sql_quote(r.status_code)})"
                for r in batch
            ) + ";",
            "",
        ]

    lines += [
        "-- Vérifie que chaque membre existe et appartient encore au groupe attendu.",
        "CREATE TEMPORARY TABLE tmp_attendance_member_guard (",
        "    validation_result TINYINT NOT NULL",
        ") ENGINE=InnoDB;",
        "",
        "INSERT INTO tmp_attendance_member_guard (validation_result)",
        "SELECT IF(",
        "    COUNT(*) = COUNT(m.id),",
        "    1,",
        "    NULL",
        ")",
        "FROM tmp_attendance_import t",
        "LEFT JOIN membres m",
        "    ON m.id = t.membre_id",
        "   AND m.groupe = t.group_id;",
        "",
        "-- Vérifie que tous les groupes utilisés existent en production.",
        "CREATE TEMPORARY TABLE tmp_attendance_group_guard (",
        "    validation_result TINYINT NOT NULL",
        ") ENGINE=InnoDB;",
        "",
        "INSERT INTO tmp_attendance_group_guard (validation_result)",
        "SELECT IF(",
        "    COUNT(DISTINCT t.group_id) = COUNT(DISTINCT g.id),",
        "    1,",
        "    NULL",
        ")",
        "FROM tmp_attendance_import t",
        "LEFT JOIN groupes g ON g.id = t.group_id;",
        "",
        "-- Remplace la saison pour tous les groupes réellement présents dans ce fichier.",
        "DELETE FROM attendance_sessions",
        "WHERE season_id = @attendance_season_id",
        f"  AND group_id IN ({group_ids_sql});",
        "",
        "INSERT INTO attendance_sessions (season_id, group_id, session_date, time_slot, created_at, updated_at)",
        "SELECT DISTINCT @attendance_season_id, group_id, session_date, time_slot, NOW(), NOW()",
        "FROM tmp_attendance_import",
        "ORDER BY group_id, session_date, time_slot;",
        "",
        "INSERT INTO attendance_records (session_id, membre_id, status_id, comment, created_at, updated_at)",
        "SELECT s.id, t.membre_id, st.id, NULL, NOW(), NOW()",
        "FROM tmp_attendance_import t",
        "INNER JOIN attendance_sessions s",
        "    ON s.season_id = @attendance_season_id",
        "   AND s.group_id = t.group_id",
        "   AND s.session_date = t.session_date",
        "   AND s.time_slot = t.time_slot",
        "INNER JOIN attendance_statuses st ON st.code = t.status_code",
        "INNER JOIN membres m",
        "    ON m.id = t.membre_id",
        "   AND m.groupe = t.group_id;",
        "",
        "SELECT g.id AS group_id, COUNT(*) AS imported_sessions",
        "FROM attendance_sessions s",
        "INNER JOIN groupes g ON g.id = s.group_id",
        "WHERE s.season_id = @attendance_season_id",
        f"  AND s.group_id IN ({group_ids_sql})",
        "GROUP BY g.id",
        "ORDER BY g.id;",
        "",
        "SELECT s.group_id, COUNT(*) AS imported_records",
        "FROM attendance_records ar",
        "INNER JOIN attendance_sessions s ON s.id = ar.session_id",
        "WHERE s.season_id = @attendance_season_id",
        f"  AND s.group_id IN ({group_ids_sql})",
        "GROUP BY s.group_id",
        "ORDER BY s.group_id;",
        "",
        "SELECT s.group_id, st.code, COUNT(*) AS records_count",
        "FROM attendance_records ar",
        "INNER JOIN attendance_sessions s ON s.id = ar.session_id",
        "INNER JOIN attendance_statuses st ON st.id = ar.status_id",
        "WHERE s.season_id = @attendance_season_id",
        f"  AND s.group_id IN ({group_ids_sql})",
        "GROUP BY s.group_id, st.code",
        "ORDER BY s.group_id, st.sort_order;",
        "",
        "COMMIT;",
        "",
        "DROP TEMPORARY TABLE IF EXISTS tmp_attendance_import;",
        "DROP TEMPORARY TABLE IF EXISTS tmp_attendance_member_guard;",
        "DROP TEMPORARY TABLE IF EXISTS tmp_attendance_group_guard;",
        "DROP TEMPORARY TABLE IF EXISTS tmp_attendance_guard;",
        "",
        f"-- Valeur attendue totale : imported_sessions = {len(sessions)}",
        f"-- Valeur attendue totale : imported_records = {len(rows)}",
    ]

    for status_code, count in sorted(status_counts.items()):
        lines.append(f"-- {status_code:<20} = {count}")
    lines.append("")
    return "\n".join(lines)

def write_unmatched_report(unmatched: pd.DataFrame, output_sql_path: Path):
    if unmatched.empty:
        return None
    path = output_sql_path.with_name(output_sql_path.stem + "_unmatched_members.csv")
    columns = [c for c in ["nom_source", "prenom_source", "date_naissance_source"] if c in unmatched.columns]
    unmatched[columns].to_csv(path, index=False, encoding="utf-8-sig")
    return path


def main() -> int:
    args = parse_args()
    connection = None

    try:
        # Démarrage systématique demandé : mysql.server gère le cas déjà actif.
        start_mysql_server()

        groups = resolve_requested_groups(args)
        if not groups:
            fail("Au moins un groupe doit être indiqué avec --group, --groups ou CNBA_GROUPS.")

        if args.group_id is not None and len(groups) > 1:
            fail("--group-id ne peut être utilisé qu'en mono-groupe. Utiliser --group-id-map pour plusieurs groupes.")

        group_id_map = parse_group_id_map(args.group_id_map)
        season_code = args.season_code or f"{config.SEASON_START_YEAR}-{config.SEASON_END_YEAR}"

        connection = connect_db(args)
        members = load_members(connection)
        validate_local_attendance_schema(connection)

        all_rows: list[AttendanceRow] = []
        all_source_pickles: list[Path] = []
        all_group_db_labels: list[str] = []
        total_source_swimmers = 0
        total_matched_swimmers = 0
        total_empty_count = 0
        total_status_counts = {code: 0 for code in sorted(REQUIRED_DB_STATUS_CODES)}
        all_member_labels: dict[int, str] = {}

        output_dir = Path(config.OUTPUT_DIR) / "sql"
        output_name_groups = "_".join(slugify(group) for group in groups)
        output_path = args.output or output_dir / f"attendance_{output_name_groups}_{season_code}.sql"
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        for group_name in groups:
            print()
            info(f"Traitement export SQL du groupe : {group_name}")
            activate_group(group_name)

            df = load_validated_data()
            validate_season_dates(df)
            source_swimmers = build_source_swimmers(df)

            explicit_group_id = (
                args.group_id
                if len(groups) == 1 and args.group_id is not None
                else group_id_map.get(normalize_identity(group_name))
            )
            group_id, group_db_label = resolve_group(connection, group_name, explicit_group_id)
            success(f"Groupe DB reconnu : {group_db_label} (id={group_id})")
            all_group_db_labels.append(f"{group_db_label} (id={group_id})")

            # Le matching ne tient volontairement pas compte du groupe actuel :
            # nom + prénom + date de naissance uniquement.
            member_mapping, unmatched = match_swimmers(source_swimmers, members)
            all_member_labels.update(build_member_labels(member_mapping))
            success(f"{len(member_mapping)}/{len(source_swimmers)} nageurs reconnus")

            unmatched_report = write_unmatched_report(unmatched, output_path.with_name(
                f"{output_path.stem}_{slugify(group_name)}{output_path.suffix}"
            ))
            if not unmatched.empty:
                message = f"{len(unmatched)} nageur(s) non reconnu(s) pour {group_name}."
                if unmatched_report:
                    message += f" Rapport : {unmatched_report}"
                if not args.allow_unmatched:
                    fail(message)
                warning(message)

            rows, status_counts, empty_count = build_attendance_rows(
                df,
                member_mapping,
                args.allow_unmatched,
                source_group_name=group_name,
                source_group_label=f"{group_db_label} (id={group_id})",
            )

            all_rows.extend(rows)
            all_source_pickles.append(get_validated_pickle_path())
            total_source_swimmers += len(source_swimmers)
            total_matched_swimmers += len(member_mapping)
            total_empty_count += empty_count
            for code, count in status_counts.items():
                total_status_counts[code] = total_status_counts.get(code, 0) + count

        if not all_rows:
            fail("Aucun relevé à exporter pour les groupes demandés.")

        grouped = {}
        for item in all_rows:
            grouped.setdefault(
                (item.group_id, item.session_date, item.time_slot, item.member_id),
                []
            ).append(item)

        duplicates = {key: values for key, values in grouped.items() if len(values) > 1}
        if duplicates:
            lines = []
            for key, values in list(duplicates.items())[:30]:
                group_id, session_date, time_slot, member_id = key
                nageur = all_member_labels.get(member_id, f"membre_id={member_id}")

                lines.append(
                    f"Doublon: group_id_export={group_id} | date={session_date} | "
                    f"créneau={time_slot} | membre_id={member_id} | nageur={nageur}"
                )

                for item in values:
                    lines.append("  - " + format_attendance_row_detail(item, all_member_labels))

            fail(
                "Doublons inter-groupes nageur/séance détectés dans l'export agrégé.\n"
                "Chaque doublon liste le nageur, le groupe source Google Sheet "
                "et le groupe DB réellement attribué à la ligne.\n"
                + "\n".join(lines)
            )

        all_rows.sort(key=lambda x: (x.group_id, x.session_date, x.time_slot, x.member_id))
        sessions_count = len({(r.group_id, r.session_date, r.time_slot) for r in all_rows})
        involved_groups = sorted({r.group_id for r in all_rows})

        sql = generate_sql(
            db_name=args.db_name,
            group_names=groups,
            group_db_labels=all_group_db_labels,
            season_code=season_code,
            rows=all_rows,
            status_counts=total_status_counts,
            source_pickles=all_source_pickles,
        )
        output_path.write_text(sql, encoding="utf-8")

        success(f"Fichier SQL unique généré : {output_path}")
        print(f"Saison                  : {season_code}")
        print(f"Groupes demandés        : {', '.join(groups)}")
        print(f"Groupes attribués DB    : {', '.join(map(str, involved_groups))}")
        print(f"Nageurs source cumulés  : {total_source_swimmers}")
        print(f"Nageurs reconnus cumulés: {total_matched_swimmers}")
        print(f"Séances exportées       : {sessions_count}")
        print(f"Relevés exportés        : {len(all_rows)}")
        print(f"Cellules vides ignorées : {total_empty_count}")
        for code, count in sorted(total_status_counts.items()):
            print(f"{code:<23}: {count}")

        return 0

    finally:
        if connection is not None:
            try:
                connection.close()
            except Exception as exc:
                warning(f"Erreur pendant la fermeture de la connexion DB : {exc}")

        # Arrêt systématique, y compris après une erreur de validation/export.
        stop_mysql_server()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"{Color.RED}✗ {exc}{Color.RESET}", file=sys.stderr)
        sys.exit(1)
