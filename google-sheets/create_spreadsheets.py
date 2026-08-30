#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Apr 14 23:16:23 2022

@author: remicrepin
"""

import pandas as pd
import string
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import model
import time
import config

# 1: Masters Licence | 2: Elite | 3: Avenir 1
# 4: Elite Jeunes | 5: Elite Jeunes 2 | 6: Avenir 2
# 7: Avenir Compétition 1 | 8: Avenir Compétition 2 | 10: Cadre
# 17: Aucun | 18: Invité | 36: Start To Swim Jeunes
# 37: Perf. Jeunes | 41: Perf. Ados | 43: Masters Perf. 1
# 44: Masters Perf. 2 | 45: Triathlon | 99: Réservé
# 104: Masters Comp. | 105: Start to Swim 1 | 106: Couloir parents
# 108: Elite Licence | 109: Elite Senior | 110: Accoutumance 09 2025
# 111: Accoutumance 02 2026 | 118: Start to Swim 2 | 119: Start to Swim 3

UPDATE = False
TEST = False
ADD_CONDITIONAL_FORMAT = True
list_groups = [36,37,41,43,44,45,105,110,118,119]

if __name__ == '__main__':
    conf = config.config()
    model.create_spreadsheets(
        conf,
        update=UPDATE,
        test=TEST,
        list_groups=list_groups,
        add_conditional_format=ADD_CONDITIONAL_FORMAT
    )
