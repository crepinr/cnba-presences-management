#!/usr/bin/env python3
"""
attendance_yaml_config.py

Adaptateur YAML pour le pipeline Presence Analytics.

Objectif :
- garder les scripts existants basés sur config.py ;
- permettre de déclarer les groupes dans attendance_config.yaml ;
- utiliser uniquement les groupes dont active: true dans l'orchestrateur ;
- surcharger dynamiquement les chemins et la clé Google Sheet du groupe actif.

Le YAML est optionnel pour les anciens usages : si CNBA_CONFIG_YAML n'est pas
défini et si attendance_config.yaml est absent, le config.py existant reste
la source de vérité.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path
from typing import Any

try:
    import yaml
except ImportError:  # pragma: no cover
    yaml = None


DEFAULT_CONFIG_FILE = "attendance_config.yaml"


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()


def slugify(value: str) -> str:
    import unicodedata

    value = clean_text(value).lower()
    value = "".join(
        char for char in unicodedata.normalize("NFD", value)
        if unicodedata.category(char) != "Mn"
    )
    return re.sub(r"[^a-z0-9]+", "_", value).strip("_") or "groupe"


def extract_google_sheet_key(value: str) -> str:
    value = clean_text(value)
    if not value:
        return ""
    match = re.search(r"/spreadsheets/d/([^/]+)", value)
    if match:
        return match.group(1)
    return value


def get_yaml_path(path: str | None = None) -> Path | None:
    raw = path or os.environ.get("CNBA_CONFIG_YAML") or DEFAULT_CONFIG_FILE
    candidate = Path(raw)
    if candidate.exists():
        return candidate
    if path or os.environ.get("CNBA_CONFIG_YAML"):
        raise FileNotFoundError(f"Fichier YAML introuvable : {candidate}")
    return None


def load_yaml(path: str | None = None) -> dict:
    yaml_path = get_yaml_path(path)
    if yaml_path is None:
        return {}
    if yaml is None:
        raise RuntimeError(
            "Le module PyYAML est requis pour lire attendance_config.yaml. "
            "Installer avec : pip install pyyaml"
        )
    with yaml_path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}
    if not isinstance(data, dict):
        raise ValueError("Le fichier YAML doit contenir un objet racine.")
    return data


def iter_group_entries(data: dict):
    groups = data.get("groups", [])
    if isinstance(groups, dict):
        for key, group in groups.items():
            if not isinstance(group, dict):
                continue
            item = dict(group)
            item.setdefault("key", key)
            yield item
    elif isinstance(groups, list):
        for group in groups:
            if isinstance(group, dict):
                yield dict(group)


def group_label(group: dict) -> str:
    return clean_text(group.get("label") or group.get("name") or group.get("key"))


def is_active(group: dict) -> bool:
    return bool(group.get("active", False))


def find_group(data: dict, requested_group: str) -> dict:
    requested = clean_text(requested_group)
    requested_key = slugify(requested)
    for group in iter_group_entries(data):
        candidates = {
            clean_text(group.get("key")),
            group_label(group),
            clean_text(group.get("name")),
        }
        candidate_slugs = {slugify(item) for item in candidates if item}
        if requested in candidates or requested_key in candidate_slugs:
            return group
    raise ValueError(f"Groupe {requested_group!r} introuvable dans le YAML.")


def active_group_labels(data: dict) -> list[str]:
    return [group_label(group) for group in iter_group_entries(data) if is_active(group)]


def group_id_map(data: dict, labels: list[str] | None = None) -> dict[str, int]:
    allowed = {slugify(label) for label in labels} if labels else None
    mapping = {}
    for group in iter_group_entries(data):
        label = group_label(group)
        if not label:
            continue
        if allowed is not None and slugify(label) not in allowed:
            continue
        raw_id = group.get("db_group_id")
        if raw_id is None or clean_text(raw_id) == "":
            continue
        mapping[label] = int(raw_id)
    return mapping


def _path_from_template(template: str, *, label: str, key: str, slug: str, base_dir: Path) -> Path:
    return Path(str(template).format(
        label=label,
        key=key,
        slug=slug,
        base_dir=str(base_dir),
    ))


def apply_group_to_config(config_module, group: dict, data: dict) -> None:
    label = group_label(group)
    key = clean_text(group.get("key")) or slugify(label)
    slug = slugify(key or label)

    google = group.get("google_sheet") or {}
    if isinstance(google, dict):
        sheet_key = clean_text(google.get("key")) or extract_google_sheet_key(google.get("url"))
    else:
        sheet_key = extract_google_sheet_key(str(google))

    if not sheet_key:
        raise ValueError(f"Clé Google Sheet absente pour le groupe {label!r}.")

    paths = data.get("paths", {}) or {}
    base_dir = Path(paths.get("base_dir", "."))
    input_dir = Path(paths.get("input_dir", getattr(config_module, "INPUT_DIR", "Input")))
    output_dir = Path(paths.get("output_dir", getattr(config_module, "OUTPUT_DIR", "output")))
    pickle_dir_template = paths.get("pickle_dir_template", "output/pickle/{slug}")
    clean_up_dir_template = paths.get("clean_up_dir_template", "output/clean_up/{slug}")

    input_file = group.get("input_file")
    if input_file:
        input_file = Path(str(input_file))
    else:
        input_file = input_dir / f"Input_{slug}.xlsx"

    pickle_dir = _path_from_template(
        group.get("pickle_dir") or pickle_dir_template,
        label=label,
        key=key,
        slug=slug,
        base_dir=base_dir,
    )
    clean_up_dir = _path_from_template(
        group.get("clean_up_dir") or clean_up_dir_template,
        label=label,
        key=key,
        slug=slug,
        base_dir=base_dir,
    )

    config_module.ACTIVE_GROUP = label
    config_module.ACTIVE_GROUP_LABEL = label
    config_module.ACTIVE_GROUP_KEY = key
    config_module.DATA_SHEET_KEY = sheet_key
    config_module.INPUT_DIR = input_dir
    config_module.INPUT_FILE = Path(input_file)
    config_module.OUTPUT_DIR = output_dir
    config_module.OUTPUT_FILE = output_dir / f"controle_{slug}.xlsx"
    config_module.PICKLE_DIR = pickle_dir
    config_module.CLEAN_UP_DIR = clean_up_dir

    config_module.DATA_CLEAN_PICKLE = pickle_dir / "data_clean.pkl"
    config_module.NAGEURS_PICKLE = pickle_dir / "nageurs.pkl"
    config_module.SEANCES_PICKLE = pickle_dir / "seances.pkl"
    config_module.AUDIT_GLOBAL_PICKLE = pickle_dir / "audit_global.pkl"
    config_module.AUDIT_DATES_PICKLE = pickle_dir / "audit_dates.pkl"
    config_module.AUDIT_STATUTS_PICKLE = pickle_dir / "audit_statuts.pkl"

    config_module.DATA_CLEAN_VALIDATED_PICKLE = clean_up_dir / "data_clean_validated.pkl"
    config_module.NAGEURS_VALIDATED_PICKLE = clean_up_dir / "nageurs_validated.pkl"
    config_module.SEANCES_VALIDATED_PICKLE = clean_up_dir / "seances_validated.pkl"
    config_module.CLEAN_UP_ISSUES_PICKLE = clean_up_dir / "clean_up_issues.pkl"
    config_module.CLEAN_UP_REPORT_XLSX = clean_up_dir / f"clean_up_report_{slug}.xlsx"


def bootstrap(config_module, yaml_path: str | None = None) -> bool:
    data = load_yaml(yaml_path)
    if not data:
        return False

    original_available_groups = getattr(config_module, "available_groups", None)
    original_set_active_group = getattr(config_module, "set_active_group", None)

    def available_groups() -> list[str]:
        labels = [group_label(group) for group in iter_group_entries(data) if group_label(group)]
        if labels:
            return labels
        if callable(original_available_groups):
            return list(original_available_groups())
        return []

    def set_active_group(group_name: str):
        group = find_group(data, group_name)
        if callable(original_set_active_group):
            try:
                original_set_active_group(group_label(group))
            except Exception:
                # Le YAML peut déclarer de nouveaux groupes inconnus de l'ancien config.py.
                pass
        apply_group_to_config(config_module, group, data)

    config_module.available_groups = available_groups
    config_module.set_active_group = set_active_group

    season = data.get("season", {}) or {}
    if "start_year" in season:
        config_module.SEASON_START_YEAR = int(season["start_year"])
    if "end_year" in season:
        config_module.SEASON_END_YEAR = int(season["end_year"])

    active_labels = active_group_labels(data)
    if active_labels:
        config_module.ACTIVE_GROUPS = active_labels

    return True


def main() -> int:
    parser = argparse.ArgumentParser(description="Utilitaires YAML du pipeline CNBA.")
    parser.add_argument("command", choices=["active-labels", "group-id-map", "all-labels"])
    parser.add_argument("--config", default=None)
    parser.add_argument("--groups", default=None)
    args = parser.parse_args()

    data = load_yaml(args.config)
    if args.command == "active-labels":
        for label in active_group_labels(data):
            print(label)
        return 0

    if args.command == "all-labels":
        for group in iter_group_entries(data):
            label = group_label(group)
            if label:
                print(label)
        return 0

    labels = [clean_text(item) for item in (args.groups or "").split(",") if clean_text(item)]
    mapping = group_id_map(data, labels or None)
    print(";".join(f"{label}={db_id}" for label, db_id in mapping.items()))
    return 0


if __name__ == "__main__":
    sys.exit(main())
