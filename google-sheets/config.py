#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Apr  6 23:35:26 2022

@author: remicrepin
"""

import os
from pathlib import Path


def _load_env_file():
    """Charge google-sheets/.env sans ajouter de dépendance externe."""
    env_file = Path(__file__).resolve().with_name(".env")
    if not env_file.is_file():
        return

    for raw_line in env_file.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip("\"'")
        if key:
            os.environ.setdefault(key, value)


_load_env_file()

class config():
    def __init__(self):
        self.folder_id = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "")
        self.EMAIL_ADDRESS = os.environ.get("CNBA_EMAIL_ADDRESS", "")
        self.PASSWORD = os.environ.get("CNBA_EMAIL_PASSWORD", "")

    
