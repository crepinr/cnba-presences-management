#!/usr/bin/env python3
"""
06_selenium_upload_sql.py

Importe dans phpMyAdmin le fichier SQL généré par 05_export_attendance_sql.py.

Le chemin du SQL n'est plus codé en dur :
    python 06_selenium_upload_sql.py --file output/sql/attendance_2025-2026.sql

Les paramètres sensibles viennent d'un fichier .env chargé au démarrage,
puis restent surchargeables par variables d'environnement déjà présentes.

Variables attendues :
    PMA_URL
    PMA_SERVER
    PMA_USERNAME
    PMA_PASSWORD

Variables optionnelles :
    PMA_DATABASE_INDEX   index du lien DB dans l'arbre phpMyAdmin, défaut 2
    PMA_HEADLESS         1/0, défaut 1
"""

from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


IMPLICIT_WAIT = 10


def clean_text(value: str | None) -> str:
    return (value or "").strip()


def load_env_file(path: Path, *, override: bool = False) -> None:
    """
    Charge un fichier .env simple sans dépendance externe.

    Format accepté :
        CLE=valeur
        CLE="valeur avec espaces"
        CLE='valeur avec espaces'

    Les lignes vides et les commentaires # sont ignorés.
    Par défaut, une variable déjà présente dans l'environnement n'est pas écrasée.
    """
    path = path.expanduser().resolve()

    if not path.exists():
        if path.name == ".env":
            return
        raise FileNotFoundError(f"Fichier .env introuvable : {path}")

    if not path.is_file():
        raise ValueError(f"Le chemin .env n'est pas un fichier : {path}")

    for line_number, raw_line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        line = raw_line.strip()

        if not line or line.startswith("#"):
            continue

        if line.startswith("export "):
            line = line[len("export "):].strip()

        if "=" not in line:
            raise ValueError(f"Ligne .env invalide {path}:{line_number} -> {raw_line!r}")

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip()

        if not key:
            raise ValueError(f"Clé .env vide {path}:{line_number}")

        if (
            (value.startswith('"') and value.endswith('"'))
            or (value.startswith("'") and value.endswith("'"))
        ):
            value = value[1:-1]

        if override or key not in os.environ:
            os.environ[key] = value


def env_bool(name: str, default: bool) -> bool:
    value = clean_text(os.environ.get(name))
    if not value:
        return default
    return value.lower() not in {"0", "false", "no", "non", "off"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Upload d'un fichier SQL dans phpMyAdmin via Selenium."
    )
    parser.add_argument(
        "--file",
        required=True,
        type=Path,
        help="Fichier SQL à importer."
    )
    parser.add_argument(
        "--env-file",
        type=Path,
        default=Path(".env"),
        help="Fichier .env à charger. Défaut: .env"
    )
    parser.add_argument(
        "--override-env",
        action="store_true",
        help="Le fichier .env écrase les variables déjà présentes dans l'environnement."
    )
    parser.add_argument(
        "--headless",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="Active/désactive Chrome headless. Défaut: PMA_HEADLESS dans .env, sinon activé."
    )
    return parser.parse_args()


def validate_sql_file(path: Path) -> Path:
    path = path.expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Fichier SQL introuvable : {path}")
    if not path.is_file():
        raise ValueError(f"Le chemin n'est pas un fichier : {path}")
    if path.suffix.lower() != ".sql":
        raise ValueError(f"Le fichier doit avoir l'extension .sql : {path}")
    if path.stat().st_size == 0:
        raise ValueError(f"Le fichier SQL est vide : {path}")
    return path


def required_env(name: str) -> str:
    value = clean_text(os.environ.get(name))
    if not value:
        raise RuntimeError(f"Variable d'environnement obligatoire manquante : {name}")
    return value


def init_driver(headless: bool):
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    return webdriver.Chrome(options=chrome_options)


def run_upload(sql_file: Path, *, headless: bool) -> None:
    pma_url = os.environ.get("PMA_URL", "https://h2-phpmyadmin.infomaniak.com/index.php")
    pma_server = required_env("PMA_SERVER")
    pma_username = required_env("PMA_USERNAME")
    pma_password = required_env("PMA_PASSWORD")
    database_index = int(os.environ.get("PMA_DATABASE_INDEX", "2"))

    driver = None
    try:
        driver = init_driver(headless=headless)
        print("Chrome Driver Init() successful")

        driver.maximize_window()
        driver.implicitly_wait(IMPLICIT_WAIT)

        print(f"Ouverture phpMyAdmin : {pma_url}")
        driver.get(pma_url)
        time.sleep(2)

        driver.find_element(By.ID, "serverNameInput").send_keys(pma_server)
        driver.find_element(By.ID, "input_username").send_keys(pma_username)
        driver.find_element(By.ID, "input_password").send_keys(pma_password)
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "input_go"))
        ).click()

        driver.implicitly_wait(IMPLICIT_WAIT)

        try:
            db_xpath = f'//*[@id="pma_navigation_tree_content"]/ul/li[{database_index}]/a'
            WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, db_xpath))
            ).click()
        except Exception as exc:
            print(f"Failed to click on database index {database_index}: {exc}")

        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="topmenu"]/li[5]/a'))
        ).click()
        time.sleep(5)

        print(f"Import du fichier SQL : {sql_file}")
        file_input = driver.find_element(By.ID, "input_import_file")
        file_input.send_keys(str(sql_file))

        button_import = driver.find_element(By.XPATH, '//*[@id="buttonGo"]')
        ActionChains(driver).scroll_to_element(button_import).perform()
        ActionChains(driver).scroll_by_amount(0, 20).perform()
        time.sleep(2)
        button_import.click()

        time.sleep(25)
        print("Upload SQL terminé.")

    finally:
        if driver is not None:
            driver.quit()


def main() -> int:
    args = parse_args()
    load_env_file(args.env_file, override=args.override_env)
    sql_file = validate_sql_file(args.file)
    headless = args.headless if args.headless is not None else env_bool("PMA_HEADLESS", True)
    run_upload(sql_file, headless=headless)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"✗ {exc}", file=sys.stderr)
        sys.exit(1)
