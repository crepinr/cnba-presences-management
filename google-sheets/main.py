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
import pprint
import datetime
import urllib
import model
import calendar
import locale
import itertools
import time

locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')

def date_global(month):
    if month > 7:
        year = int(datetime.datetime.now().strftime("%Y"))
    else:
        year = int(datetime.datetime.now().strftime("%Y")) + 1

    num_days = calendar.monthrange(year, month)[1]

    dates = [datetime.date(year, month, day).strftime("%d/%m") for day in range(1, num_days+1)]
    days_names = [datetime.date(year, month, day).strftime("%A") for day in range(1, num_days+1)]
    days_names = [day_name[0].lower() for day_name in days_names]
    list_collumns, indexes = model.get_list_collumns_letters(len(dates))
   
    return [list_collumns, indexes, dates, days_names]

def date_elite(month):
    if month > 7:
        year = int(datetime.datetime.now().strftime("%Y"))
    else:
        year = int(datetime.datetime.now().strftime("%Y")) + 1

    num_days = calendar.monthrange(year, month)[1]

    dates = [datetime.date(year, month, day).strftime("%d/%m") for day in range(1, num_days+1)]
    days_names = [datetime.date(year, month, day).strftime("%A") for day in range(1, num_days+1)]

    days_names_out = days_names
    dates_out = dates
    i=0
    
    while i < len(days_names):
        day = days_names[i]
        if day == 'Mardi':
            days_names.insert(i, 'mAM')
            dates_out.insert(i, dates_out[i])
            i+=2            
        elif day == 'Jeudi':
            days_names.insert(i, 'jAM')
            dates_out.insert(i, dates_out[i])
            i+=2 
        else:
            i+=1

    days_names_out = []
    for day_name in days_names:
        if day_name[-2:] == 'AM':
            days_names_out.append(day_name) 
        else :
            days_names_out.append(day_name[0].lower()) 

    list_collumns, indexes = model.get_list_collumns_letters(len(dates))
   
    return [list_collumns, indexes, dates_out, days_names_out]

def date_ej2(month):
    if month > 7:
        year = int(datetime.datetime.now().strftime("%Y"))
    else:
        year = int(datetime.datetime.now().strftime("%Y")) + 1

    num_days = calendar.monthrange(year, month)[1]

    dates = [datetime.date(year, month, day).strftime("%d/%m") for day in range(1, num_days+1)]
    days_names = [datetime.date(year, month, day).strftime("%A") for day in range(1, num_days+1)]

    days_names_out = days_names
    dates_out = dates
    i=0
    
    while i < len(days_names):
        day = days_names[i]
        if day == 'Samedi':
            days_names.insert(i, 'sAM')
            dates_out.insert(i, dates_out[i])
            i+=2      
        else:
            i+=1

    days_names_out = []
    for day_name in days_names:
        if day_name[-2:] == 'AM':
            days_names_out.append(day_name) 
        else :
            days_names_out.append(day_name[0].lower()) 

    list_collumns, indexes = model.get_list_collumns_letters(len(dates))
   
    return [list_collumns, indexes, dates_out, days_names_out]


list_months = [9,10,11,12,1,2,3,4,5,6]
sleep = 15
UPDATE = False

if __name__ == "__main__":
  # DF MEMBRES AND GROUPES
  df_membres, df_groupes = model.read_db(update=UPDATE)
  client = model.init_client_drive()

  for index,groupe in df_groupes.iterrows():
    if groupe["id"] not in (10,17,18):
        if groupe["id"] in ([45]): #only one groupe for test, remove for prod
            #DATA MEMBERS
            if groupe["id"] != 2:
                membres_groupe = df_membres.query(f'groupe == {groupe["id"]}')[['nom','prenom','date_naissance']]
                membres_groupe = membres_groupe.sort_values('nom')
            else:
                membres_elite = df_membres.query(f'groupe == {groupe["id"]}')[['nom','prenom','date_naissance']]
                membres_elite = membres_elite.sort_values('nom')
                print(f'ELITE =={membres_elite}')
                membres_senior = df_membres.query(f'groupe == {109}')[['nom','prenom','date_naissance']]
                membres_senior = membres_senior.sort_values('nom')
                print(f'ELITE SENIOR =={membres_senior}')
                membres_groupe = pd.concat([membres_elite, membres_senior])


            membres_groupe['date_naissance'] = membres_groupe['date_naissance'].astype(str)
            print(membres_groupe)
            #print(f'{groupe["id"]} - {groupe["nom_groupe"]}')
            
            #CREATE SHEET
            nom_sheet = f'Presences {groupe["nom_groupe"]}'
            file = model.create_spreadsheet(client, nom_sheet)
            
            #RENAME SHEET 'Groupe' AND CONTENT
            print('----- SHEET GROUPE')
            worksheet = file.get_worksheet(0)
            worksheet.update_title("Groupe")
            worksheet.update([membres_groupe.columns.values.tolist()] + membres_groupe.values.tolist())
            worksheet.update_acell('D2',file.id)
            worksheet.add_protected_range('A1:D40', ['pythonaccount@pythonsheet-346808.iam.gserviceaccount.com','remicrepin25@gmail.com','technique.cnba@gmail.com'])
            print('---------- SUCCESS SHEET RENAMED AND FILLED IN')
            
            #ADD SHEETS FOR EACH MONTH AND CONTENT
            print('----- SHEETS FOR EACH MONTH')
            for month in list_months: #REMOVE RANGE FOR PROD
                month_name = datetime.date(2025, month, 1).strftime("%B")
                print(f'------ ADDING SHEET FOR {month_name}')
                worksheet = file.add_worksheet(title=month_name, rows=100, cols=50)
                
                #TITLE 
                print(f'---------- ADDING TITLE')
                """
                worksheet.update('A1',month_name)
                worksheet.format("A1", { 'backgroundColor': {
                                                            'red':236/255,
                                                            'green':155/255,
                                                            'blue':155/255},
                                        'textFormat': {'bold': True}})
                """
                requests = [
                    {
                        "updateCells": {
                            "range": {
                                "sheetId": worksheet.id,
                                "startRowIndex": 0,
                                "endRowIndex": 1,
                                "startColumnIndex": 0,
                                "endColumnIndex": 1
                            },
                            "rows": [
                                {
                                    "values": [
                                        {
                                            "userEnteredValue": {"stringValue": month_name},
                                            "userEnteredFormat": {
                                                "backgroundColor": {
                                                    "red": 236/255,
                                                    "green": 155/255,
                                                    "blue": 155/255
                                                },
                                                "textFormat": {"bold": True}
                                            }
                                        }
                                    ]
                                }
                            ],
                            "fields": "userEnteredValue,userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.bold"
                        }
                    }
                ]
                worksheet.spreadsheet.batch_update({"requests": requests})
                
                #IMPORT GROUPE VALUES
                print(f'---------- ADDING GROUPE NAMES')
                worksheet.update('A4','=IMPORTRANGE(Groupe!D2;"Groupe!A2:C30")', value_input_option='USER_ENTERED')

                #DATES
                print(f'---------- ADDING DATES')
                worksheet.update('C2','Dates')
                worksheet.format("C2", { 'backgroundColor': {
                                                            'red':252/255,
                                                            'green':202/255,
                                                            'blue':159/255}})
                #GET DATES depending on GROUP
                if groupe["id"] not in (2,5):
                    results = date_global(month)
                elif groupe["id"] == 2: #Elite
                    results = date_elite(month)
                elif groupe["id"] == 5: #EJ2
                    results = date_global(month)
                
                list_collumns = results[0]
                indexes = results[1]
                dates = results[2]
                days_names = results[3]

                range_cols = f'{list_collumns[0]}:{list_collumns[-3]}'
                range_dates = f'{list_collumns[0]}2:{list_collumns[-3]}2'
                range_days = f'{list_collumns[0]}3:{list_collumns[-3]}3'

                #print(dates)
                #print(days_names)
                #print(range_dates)
                
                worksheet.update(range_dates,[dates])

                worksheet.format(range_dates, { 'backgroundColor': {
                                                            'red':252/255,
                                                            'green':202/255,
                                                            'blue':159/255},
                                                'textRotation': {
                                                            'angle' : 90}})
                

                #JOURS
                print(f'---------- ADDING DAYS')
                worksheet.update('C3','Jours')
                worksheet.format("C3", { 'backgroundColor': {
                                                            'red':156/255,
                                                            'green':198/255,
                                                            'blue':230/255}})
                worksheet.update(range_days,[days_names])
                worksheet.format(range_days, { 'backgroundColor': {
                                                        'red':156/255,
                                                        'green':198/255,
                                                        'blue':230/255}})

                #ADD DATA COUNT
                worksheet.add_cols(2)
                worksheet.update(f'{list_collumns[-2]}3:{list_collumns[-1]}3',[['Présences %','Nbre Entr.']])
                
                #PERCENTAGE
                print(f'---------- ADDING PERCENTAGE')
                formulas_list_percent = []
                formulas_list_presence = []
                for i in range(len(membres_groupe.values.tolist())):
                    range_pourcentage = f'D{i+4}:{list_collumns[-3]}{i+4}'
                    formulas_list_percent.append([f'=NB.SI({range_pourcentage};"v")/(${list_collumns[-2]}$1-NB.VIDE({range_pourcentage})-NB.SI({range_pourcentage};"b")-NB.SI({range_pourcentage};"o")-NB.SI({range_pourcentage};"m"))'])
                    formulas_list_presence.append([f'=NB.SI({range_pourcentage};"v")'])

                range_formulas = f'{list_collumns[-2]}4:{list_collumns[-2]}{str(3+len(membres_groupe.values.tolist()))}'
                worksheet.update(range_formulas,
                                 formulas_list_percent, 
                                 value_input_option='USER_ENTERED')
                
                worksheet.format(range_formulas, { 'numberFormat': {
                                                        'type': 'PERCENT'}})
                
                #NUMBER EMPTY CELLS
                worksheet.update(f'{list_collumns[-2]}1',
                                 f'=NB.VIDE(D1:{list_collumns[-3]}1)', 
                                 value_input_option='USER_ENTERED')
                
                #NUMBER TRAINING
                print(f'---------- ADDING NB TRAININGS')
                range_presence = f'{list_collumns[-1]}4:{list_collumns[-1]}{str(3+len(membres_groupe.values.tolist()))}'
                worksheet.update(range_presence,
                                 formulas_list_presence, 
                                 value_input_option='USER_ENTERED')
                
                #NUMBER OF SWIMMERS
                print(f'---------- ADDING NB SWIMMERS')
                range_swimmers_formulas = f'{list_collumns[0]}{str(5+len(membres_groupe.values.tolist()))}:{list_collumns[-3]}{str(5+len(membres_groupe.values.tolist()))}'
                formulas_list_swimmers = []
                for i,day in enumerate(dates):
                    range_swimmer = f'{list_collumns[i]}4:{list_collumns[i]}{str(3+len(membres_groupe.values.tolist()))}'
                    formulas_list_swimmers.append(f'=NB.SI({range_swimmer};"v")')

                worksheet.update(range_swimmers_formulas,
                                 [formulas_list_swimmers], 
                                 value_input_option='USER_ENTERED')
                
                #AUTO RESIZE COLUMNS WIDTH
                worksheet.columns_auto_resize(indexes[0],indexes[1])
                #PROTECTED RANGE
                worksheet.add_protected_range('A1:C40', ['remicrepin25@gmail.com','technique.cnba@gmail.com'])
                print(f'##### SLEEPING {sleep} SECONDS BETWEEN SPREADSHEETS #####')
                time.sleep(sleep)
                #CONDITIONAL FORMATTING
                """
                range_formatting = f'D4:{list_collumns[-2]}{str(3+len(membres_groupe.values.tolist()))}'
                worksheet.batch_update([{
                    'range':range_formatting,
                    'response_value_render_option':
                        {"booleanRule": 
                            {"condition": 
                                {"type": "CUSTOM_FORMULA",
                                "values": [
                                    {
                                        "userEnteredValue": (
                                            "='x'"
                                        )
                                    }
                                ]},
                            },
                            "format": {
                                "textFormat": {"foregroundColor": {"red": 0.8}}
                            }
                        }
                    }])
                """
                

  """
  

  # Pass: title
  model.create_spreadsheet(client, "presences_test_1")
  
  
                range_formatting = f'D4:{list_collumns[-2]}{str(3+len(membres_groupe.values.tolist()))}'
                worksheet.batch_update([{
                    'range':range_formatting,
                    'response_value_render_option': { 
                        "booleanRule": {
                            "condition": {
                                "type": "CUSTOM_FORMULA",
                                "values": [
                                    {
                                        "userEnteredValue": (
                                            "='x'"
                                        )
                                    }
                                ],
                            },
                            
                        }
                    } 
                }])
  """
