import argparse
import locale
from pathlib import Path

import gspread
from gspread.utils import ExportFormat
from oauth2client.service_account import ServiceAccountCredentials

import config

try:
    import attendance_yaml_config
    attendance_yaml_config.bootstrap(config)
except Exception as yaml_config_error:
    # Si aucun YAML n'est configuré, on conserve le comportement historique.
    if "CNBA_CONFIG_YAML" in __import__("os").environ:
        raise

try:
    locale.setlocale(locale.LC_TIME, "fr_FR.UTF-8")
except locale.Error:
    pass


# ============================================================
# GOOGLE DRIVE
# ============================================================

def init_client_drive():
    """
    Initialise le client Google Drive à partir de client_key.json.
    """
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]

    file_name = "client_key.json"
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        file_name,
        scope,
    )

    client = gspread.authorize(creds)

    return client


def download_data(group=None):
    """
    Télécharge le Google Sheet du groupe actif vers le fichier Input dédié.

    Exemples :
    - Elite        -> Input/Input_E.xlsx
    - Elite Jeunes -> Input/Input_EJ.xlsx
    """
    if group is not None:
        config.set_active_group(group)

    config.INPUT_DIR.mkdir(parents=True, exist_ok=True)

    client = init_client_drive()

    print(f"Groupe actif : {config.ACTIVE_GROUP_LABEL}")
    print(f"Google Sheet : {config.DATA_SHEET_KEY}")
    print(f"Export vers : {config.INPUT_FILE}")

    spreadsheet = client.export(
        config.DATA_SHEET_KEY,
        format=ExportFormat.EXCEL,
    )

    with open(config.INPUT_FILE, "wb") as file:
        file.write(spreadsheet)

    print(f"Téléchargement terminé : {config.INPUT_FILE}")

    return config.INPUT_FILE


# Compatibilité ancienne utilisation éventuelle.
def download_E():
    return download_data("Elite")


def download_EJ():
    return download_data("Elite Jeunes")


# ============================================================
# CLI
# ============================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description="Télécharge les données de présence depuis Google Drive."
    )

    parser.add_argument(
        "--group",
        default=None,
        choices=config.available_groups(),
        help="Groupe à télécharger.",
    )

    return parser.parse_args()


def main(group=None):
    download_data(group=group)


if __name__ == "__main__":
    args = parse_args()
    main(group=args.group)
