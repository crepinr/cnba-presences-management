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

UPDATE = False
TEST = False
ADD_CONDITIONAL_FORMAT = True
list_groups = [2]   

if __name__ == '__main__':
    conf = config.config()
    model.create_spreadsheets(
        conf,
        update=UPDATE,
        test=TEST,
        list_groups=list_groups,
        add_conditional_format=ADD_CONDITIONAL_FORMAT
    )
