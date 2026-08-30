# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import smtplib
import config
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import base64
import urllib.parse
import string
import itertools
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import pymysql
import datetime
import urllib
import calendar
import locale
import time

locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')

########################
### GLOBAL VARIABLES ###
########################

BLUE = "\033[34m"
GREEN = "\033[32m"
YELLOW = "\033[33m"
RED = "\033[31m"
RESET = "\033[0m"
LIST_MONTHS_TO_ADAPT = ['janvier','février','mars', 'avril', 'mai', 'juin', 'juillet', 'septembre', 'octobre', 'novembre', 'décembre']

######################
######## DB #########
######################

def update_db():
    updated = False
    try:
        print("##### UPDATING DB #####")
        bashCommand = "bash /Users/remicrepin/Library/CloudStorage/OneDrive-Personnel/Python/ListContact/update_db.sh"
        os.system(bashCommand)
        updated = True
    except Exception as e:
        print(f"Error while updating databse : {e}")
        
    return updated

def read_db(update=False):
    """
    UPDATE DB
    """   

    if update:
        updated = update_db()
    else:
        updated = True
    
    """
    EXTRACT MEMBRES AND GROUPES
    """   
    if updated:
        print("##### STARTING SERVER #####")
        bashCommand = "mysql.server start"
        os.system(bashCommand)

        # Connect to the database
        connection = pymysql.connect(
            host='localhost',
            user='root',
            database='pcxa_cnba',
        )

        table_membres = "membres" 
        table_groupes = "groupes"

        # Select all from tables membres and groupes
        query_membres = f"SELECT * FROM {table_membres};" # WHERE paye > 0;"
        query_groupes = f"SELECT * FROM {table_groupes};"

        # Execute the query
        df_membres = pd.read_sql(query_membres,connection)
        df_groupes = pd.read_sql(query_groupes,connection)

        # Close the database connection
        connection.close()

        print("##### STOPING SERVER #####")
        bashCommand = "mysql.server stop"
        os.system(bashCommand)

    else:
        print('ERROR - UPDATE_DB FAILED')
        df_membres = False
        df_groupes = False
    
    return df_membres, df_groupes

######################
###### GSPREAD #######
######################

def init_client_drive():
    ###---------------###
    ###-----INIT------###
    ###---------------###
    #return gspread client 
    ### API CONNECTION ###
    #pythonsheet@cellular-motif-346609.iam.gserviceaccount.com
    #Authorize the API

    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    file_name = os.environ.get(
        "GOOGLE_APPLICATION_CREDENTIALS",
        str(os.path.join(os.path.dirname(__file__), "client_key.json")),
    )
    if not os.path.isabs(file_name):
        file_name = os.path.join(os.path.dirname(__file__), file_name)
    creds = ServiceAccountCredentials.from_json_keyfile_name(file_name,scope)
    client = gspread.authorize(creds) 

    return client

def create_spreadsheet(client, title, conf):
  """
  Creates the Sheet the user has access to.
  """
  try:
    sh = client.create(title, folder_id = conf.folder_id)
    print(f"{GREEN}SUCCESS SPREADSHEET ---- {title} ---- CREATED{RESET}")
    return sh
  except Exception as error:
    print(f"An error occurred: {error}")
    return error
  
def open_sheet_by_key(client, key):
   output_spreadsheet = client.open_by_key(key)
   return output_spreadsheet

def column_letter_to_index(column_letter):
    """Convert a column letter (e.g., 'A', 'Z', 'AA', 'AD') to a zero-based column index."""
    index = 0
    for char in column_letter:
        index = index * 26 + (ord(char.upper()) - ord('A')) + 1
    return index - 1  # Convert to zero-based index

def _presence_conditional_format_rule(worksheet, end_column_index, swimmer_count):
    return {
        "ranges": [
            {
                "sheetId": worksheet.id,
                "startRowIndex": 3,
                "endRowIndex": 3 + swimmer_count,
                "startColumnIndex": 3,
                "endColumnIndex": end_column_index
            }
        ],
        "booleanRule": {
            "condition": {
                "type": "TEXT_EQ",
                "values": [{"userEnteredValue": "V"}]
            },
            "format": {
                "backgroundColor": {
                    "red": 0.0,
                    "green": 1.0,
                    "blue": 0.0
                }
            }
        }
    }

def _is_presence_conditional_format(rule):
    ranges = rule.get("ranges", [])
    condition = rule.get("booleanRule", {}).get("condition", {})
    values = condition.get("values", [])
    return (
        len(ranges) == 1
        and ranges[0].get("startColumnIndex") == 3
        and ranges[0].get("startRowIndex") == 3
        and condition.get("type") == "TEXT_EQ"
        and values == [{"userEnteredValue": "V"}]
    )

def set_presence_conditional_format(worksheet, end_column_index, swimmer_count, existing_rules=None):
    """Add or update the V/green rule for one attendance input grid."""
    if swimmer_count <= 0:
        return False

    rule = _presence_conditional_format_rule(worksheet, end_column_index, swimmer_count)
    existing_rules = existing_rules or []

    for index, existing_rule in enumerate(existing_rules):
        if _is_presence_conditional_format(existing_rule):
            if existing_rule == rule:
                return False
            request = {
                "updateConditionalFormatRule": {
                    "sheetId": worksheet.id,
                    "index": index,
                    "rule": rule
                }
            }
            break
    else:
        request = {
            "addConditionalFormatRule": {
                "rule": rule,
                "index": 0
            }
        }

    worksheet.spreadsheet.batch_update({"requests": [request]})
    return True

######################
####### LOGIC ########
######################

def get_list_collumns_letters(n):
  list_collumns = list(
    itertools.chain(
        string.ascii_uppercase, 
        (''.join(pair) for pair in itertools.product(string.ascii_uppercase, repeat=2))
    ))
  return list_collumns[3:n+5], [3,n+5]

def date_global(month):
    if month > 7:
        year = int(datetime.datetime.now().strftime("%Y"))
    else:
        year = int(datetime.datetime.now().strftime("%Y")) + 1

    num_days = calendar.monthrange(year, month)[1]

    dates = [datetime.date(year, month, day).strftime("%d/%m") for day in range(1, num_days+1)]
    days_names = [datetime.date(year, month, day).strftime("%A") for day in range(1, num_days+1)]
    days_names = [day_name[0].lower() for day_name in days_names]

    days_names_out = days_names
    i=0

    days_names_out = []
    for day_name in days_names:
        if day_name == 'Mardi':
            days_names_out.append('ma')
        elif day_name == 'Mercredi':
            days_names_out.append('me')    
        else :
            days_names_out.append(day_name[0].lower()) 

    list_collumns, indexes = get_list_collumns_letters(len(dates))
   
    return [list_collumns, indexes, dates, days_names_out]

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
            days_names.insert(i, 'maAM')
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
        elif day_name == 'Mardi':
            days_names_out.append('ma')
        elif day_name == 'Mercredi':
            days_names_out.append('me')    
        else :
            days_names_out.append(day_name[0].lower()) 

    list_collumns, indexes = get_list_collumns_letters(len(dates))
   
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
        elif day_name == 'Mardi':
            days_names_out.append('ma')
        elif day_name == 'Mercredi':
            days_names_out.append('me')
        else :
            days_names_out.append(day_name[0].lower()) 

    list_collumns, indexes = get_list_collumns_letters(len(dates))
   
    return [list_collumns, indexes, dates_out, days_names_out]

def file2groupe(file_name):
    group_name = file_name[10:]
    return group_name

def list_np_drive(worksheet):        
        noms = worksheet.col_values(1)[1:]
        prenoms = worksheet.col_values(2)[1:]

        list_nomsprenoms_drive = [f'{x}{prenoms[i]}' for i,x in enumerate(noms)]     

        return list_nomsprenoms_drive

def list_np_db(sheet, df_membres, df_groupes):
    nom_groupe = sheet["name"]
    if nom_groupe == 'Elite':
        membres_groupe_db = df_membres.query(f'groupe in (2,109)')[['nom','prenom','date_naissance']].sort_values('nom')
    else:
        groupe_id = df_groupes.loc[df_groupes['nom_groupe'] == nom_groupe, 'id'].iloc[0]
        membres_groupe_db = df_membres.query(f'groupe == {groupe_id}')[['nom','prenom','date_naissance']].sort_values('nom')

    noms_db = membres_groupe_db['nom'].tolist()
    prenoms_db = membres_groupe_db['prenom'].tolist()
    list_nomsprenoms_db = [f'{x}{prenoms_db[i]}' for i,x in enumerate(noms_db)]
    print(list_nomsprenoms_db)

    return list_nomsprenoms_db, membres_groupe_db

def list_np_db_aucun(df_membres):
    
    membres_groupe_db = df_membres.query(f'groupe == 17')[['nom','prenom','date_naissance']].sort_values('nom')
    
    noms_db = membres_groupe_db['nom'].tolist()
    prenoms_db = membres_groupe_db['prenom'].tolist()
    list_nomsprenoms_db = [f'{x}{prenoms_db[i]}' for i,x in enumerate(noms_db)]
    print(f'###LIST_NP_AUCUN : {list_nomsprenoms_db}')

    return list_nomsprenoms_db

def list_sheets(list_files):
    list_sheets = []
    for file in list_files:
        #CONVERT FILE NAME TO GROUP NAME
        group_name= file2groupe(file['name'])
        file_id = file['id']
        list_sheets.append({'name':group_name, 'key':file_id})

    return list_sheets
    
#######################
####### FEATURES ######
#######################

def create_presence_spreadsheet(membres_groupe, client, groupe, conf, test=False, add_conditional_format=True):
    if not test:
        list_months = [9,10,11,12,1,2,3,4,5,6]
    else:
        list_months = [9,10]
    sleep = 20

    #CREATE SHEET
    nom_sheet = f'Presences {groupe["nom_groupe"]}'
    file = create_spreadsheet(client, nom_sheet, conf)
    
    #RENAME SHEET 'Groupe' AND CONTENT
    print(f'{BLUE}----- SHEET GROUPE{RESET}')
    worksheet = file.get_worksheet(0)
    worksheet.update_title("Groupe")
    worksheet.update([membres_groupe.columns.values.tolist()] + membres_groupe.values.tolist())
    worksheet.update_acell('D2',file.id)
    worksheet.add_protected_range('A1:D40', ['pythonaccount@pythonsheet-346808.iam.gserviceaccount.com','remicrepin25@gmail.com','technique.cnba@gmail.com'])
    print('---------- SUCCESS SHEET RENAMED AND FILLED IN')

    #ADD SHEETS FOR EACH MONTH AND CONTENT
    print(f'{BLUE}----- SHEETS FOR EACH MONTH{RESET}')
    for month in list_months: #REMOVE RANGE FOR PROD
        month_name = datetime.date(2025, month, 1).strftime("%B")
        print(f'------ ADDING SHEET FOR {BLUE}{month_name}{RESET}')
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
        worksheet.update([['=IMPORTRANGE(Groupe!D2;"Groupe!A2:C30")']],'A4', value_input_option='USER_ENTERED')

        #DATES
        print(f'---------- ADDING DATES')
        worksheet.update_acell('C2','Dates')
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
        worksheet.update_acell('C3','Jours')
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

        if add_conditional_format:
            set_presence_conditional_format(
                worksheet,
                column_letter_to_index(list_collumns[-3]) + 1,
                len(membres_groupe)
            )
        
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
        worksheet.update([[f'=NB.VIDE(D1:{list_collumns[-3]}1)']], 
                         f'{list_collumns[-2]}1',
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
        worksheet.add_protected_range('A1:C40', ['technique.cnba@gmail.com'])
        
        print(f'{YELLOW}##### SLEEPING {sleep} SECONDS BETWEEN WORKSHEET #####{RESET}')
        time.sleep(sleep)

def create_spreadsheets(conf, update=False, test=False, list_groups=[], add_conditional_format=True):
    
    # DF MEMBRES AND GROUPES
    df_membres, df_groupes = read_db(update)
    client = init_client_drive()

    if test : 
        list_of_groups_excluded = [1,3,4,5,6,7,8,10,17,18,36,37,41,43,44,45,99,104,106,108,109,110,111,118,119,120,121,122]
    else :
        list_of_groups_excluded = [] #[1,5,10,17,18,99,104,106,108,109,120,121,122] + list_groups

    for index,groupe in df_groupes.iterrows():
        if groupe["id"] not in list_of_groups_excluded and groupe["id"] in list_groups:
            #DATA MEMBERS
            if groupe["id"] != 2:
                membres_groupe = df_membres.query(f'groupe == {groupe["id"]}')[['nom','prenom','date_naissance']]
                membres_groupe = membres_groupe.sort_values('nom')
            else:
                membres_elite = df_membres.query(f'groupe == {groupe["id"]}')[['nom','prenom','date_naissance']]
                membres_elite = membres_elite.sort_values('nom')
                #print(f'ELITE =={membres_elite}')
                membres_senior = df_membres.query(f'groupe == {109}')[['nom','prenom','date_naissance']]
                membres_senior = membres_senior.sort_values('nom')
                #print(f'ELITE SENIOR =={membres_senior}')
                membres_groupe = pd.concat([membres_elite, membres_senior])

            membres_groupe['date_naissance'] = membres_groupe['date_naissance'].astype(str)
            #print(membres_groupe)
            
            create_presence_spreadsheet(
                membres_groupe,
                client,
                groupe,
                conf,
                test=test,
                add_conditional_format=add_conditional_format
            )

    return None

def remove_members(list_nomsprenoms_drive,list_nomsprenoms_db, worksheet, spreadsheet):

            #compare and get deleted names
            list_deleted_members = []
            for i,np in enumerate(list_nomsprenoms_drive):
                if np not in list_nomsprenoms_db:
                    list_deleted_members.append({'np':np, 'index':i})

            print(f"MEMBERS to delete : {list_deleted_members}")
            list_index_to_delete = [i['index'] for i in list_deleted_members]
            
            print(f"INDEX to delete : {list_index_to_delete}")
            #remove 0 if in
            if 0 in list_index_to_delete:
                #special case 0 tbd 
                list_index_to_delete = list_index_to_delete[1:]
            list_index_to_delete.reverse()

            print(f"INDEX to delete : {list_index_to_delete}")
            
            if len(list_index_to_delete)>=1:
                print(f"--------- DELETING -- 'GROUPE' ---------")
                requests = [
                    {
                        "deleteDimension": {
                            "range": {
                                "sheetId": worksheet.id, #remove from 'GROUPE'
                                "dimension": "ROWS",
                                "startIndex": row + 1,
                                "endIndex": row + 1 + 1
                            }
                        }
                    } for row in list_index_to_delete]
                # Perform the batch update
                print(requests)
                worksheet.spreadsheet.batch_update({"requests": requests})
                
                
                #Delete in month 
                list_ws = spreadsheet.worksheets()
                list_ws = [sheet for sheet in list_ws if sheet.title in LIST_MONTHS_TO_ADAPT]
                for i, ws in enumerate(list_ws):
                    ws_name = ws.title
                    print(ws_name)
                    if ws_name != 'Groupe':
                        worksheet = ws
                        print(f"--------- DELETING -- {ws_name} ---------")
                        request_month = ([
                            {
                                "deleteDimension": {
                                    "range": {
                                        "sheetId": ws.id,
                                        "dimension": "ROWS",
                                        "startIndex": row + 3,
                                        "endIndex": row + 1 + 3
                                    }
                                }
                            } for row in list_index_to_delete
                        ])
                        print(request_month)
                        # Perform the batch update
                        worksheet.spreadsheet.batch_update({"requests": request_month})
            else:
                print(f'{YELLOW}---- NOTHING TO DELETE ----{RESET}')

def remove_members_aucun(list_nomsprenoms_drive,list_nomsprenoms_aucun, worksheet, spreadsheet):
    #compare and get deleted names
    list_deleted_members = []
    for i,np in enumerate(list_nomsprenoms_drive):
        if np in list_nomsprenoms_aucun:
            list_deleted_members.append({'np':np, 'index':i})

    print(f"MEMBERS to delete : {list_deleted_members}")
    list_index_to_delete = [i['index'] for i in list_deleted_members]
    
    print(f"INDEX to delete : {list_index_to_delete}")
    #remove 0 if in
    if 0 in list_index_to_delete:
        #special case 0 tbd 
        list_index_to_delete = list_index_to_delete[1:]
    list_index_to_delete.reverse()

    print(f"INDEX to delete : {list_index_to_delete}")
    
    if len(list_index_to_delete)>=1:
        print(f"--------- DELETING -- 'GROUPE' ---------")
        requests = [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": worksheet.id, #remove from 'GROUPE'
                        "dimension": "ROWS",
                        "startIndex": row + 1,
                        "endIndex": row + 1 + 1
                    }
                }
            } for row in list_index_to_delete]
        # Perform the batch update
        print(requests)
        worksheet.spreadsheet.batch_update({"requests": requests})
        
        
        #Delete in month 
        list_ws = spreadsheet.worksheets()
        list_ws = [sheet for sheet in list_ws if sheet.title in LIST_MONTHS_TO_ADAPT]

        for i, ws in enumerate(list_ws):
            ws_name = ws.title
            print(ws_name)
            if ws_name != 'Groupe':
                worksheet = ws
                print(f"--------- DELETING -- {ws_name} ---------")
                request_month = ([
                    {
                        "deleteDimension": {
                            "range": {
                                "sheetId": ws.id,
                                "dimension": "ROWS",
                                "startIndex": row + 3,
                                "endIndex": row + 1 + 3
                            }
                        }
                    } for row in list_index_to_delete
                ])
                print(request_month)
                # Perform the batch update
                worksheet.spreadsheet.batch_update({"requests": request_month})
    else:
        print(f'{YELLOW}---- NOTHING TO DELETE ----{RESET}')

def get_column_index_presences(ws):
    data = ws.get_all_values()
    
    df = pd.DataFrame(data)
    target = 'Présences %'
    a = df.where(df==target).dropna(how='all').dropna(axis=1)
    column_index = a.columns.tolist()[0]
    column_index = int(column_index)

    return column_index

def add_presence_conditional_formatting(spreadsheet, swimmer_count):
    """Apply the attendance rule to existing month sheets without duplicating it."""
    metadata = spreadsheet.fetch_sheet_metadata(
        params={"fields": "sheets(properties(sheetId),conditionalFormats)"}
    )
    metadata_by_id = {
        sheet["properties"]["sheetId"]: sheet.get("conditionalFormats", [])
        for sheet in metadata.get("sheets", [])
    }

    updated = False
    month_worksheets = [
        worksheet
        for worksheet in spreadsheet.worksheets()
        if worksheet.title in LIST_MONTHS_TO_ADAPT
    ]
    for worksheet in month_worksheets:
        end_column_index = get_column_index_presences(worksheet)
        if set_presence_conditional_format(
            worksheet,
            end_column_index,
            swimmer_count,
            existing_rules=metadata_by_id.get(worksheet.id, [])
        ):
            updated = True

    if not updated:
        print(f'{YELLOW}---- NO MISSING CONDITIONAL FORMAT ----{RESET}')

    return updated

def add_members(list_nomsprenoms_drive,list_nomsprenoms_db, worksheet, spreadsheet, membres_groupe_db):
    #ADD MEMBERS
    #compare and get deleted names
    list_added_members = []
    for i,np in enumerate(list_nomsprenoms_db):
        if np not in list_nomsprenoms_drive:
            list_added_members.append({'np':np, 'index':i, 
                                        'nom':membres_groupe_db[i]['nom'], 
                                        'prenom':membres_groupe_db[i]['prenom'], 
                                        'date_naissance': membres_groupe_db[i]['date_naissance'].strftime('%Y-%m-%d')})

    list_index_to_add = [i['index'] for i in list_added_members]

    print(f"MEMBERS to add : {list_added_members}")
    print(f"INDEX to add : {list_index_to_add}")
    #remove 0 if in
    
    if 0 in list_index_to_add:
        #special case 0 tbd 
        list_index_to_add = list_index_to_add[1:]
        list_added_members = list_added_members[1:]
        print('######## ROW 0 TO BE ADDED - MANUAL ACTION REQUIRED ######')
    
    #BEGIN AT LOWEST LINES TO AVOID CHANGING LINES NUMBER DURING EXECUTION
    list_index_to_add.reverse()
    list_added_members.reverse() 

    print(f"INDEX to add : {list_index_to_add}")
    print(list_added_members)
    if len(list_added_members)>=1:
        print(f"--------- ADDING -- 'GROUPE' ---------")
        requests = []
        for row in list_added_members:
            requests.append(
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": worksheet.id, #add from 'GROUPE'
                            "dimension": "ROWS",
                            "startIndex": row['index'] + 1,
                            "endIndex": row['index'] + 1 + 1
                        },
                        "inheritFromBefore": "false"
                    }
                })
            requests.append({
                    "updateCells": {
                        "range": {
                            "sheetId": worksheet.id,
                            "startRowIndex": row['index'] + 1,
                            "endRowIndex": row['index'] + 1 + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 3
                        },
                        "rows": [
                            {
                                "values": [
                                    {"userEnteredValue": {"stringValue": row['nom']}},
                                    {"userEnteredValue": {"stringValue": row['prenom']}},
                                    {"userEnteredValue": {"stringValue": row['date_naissance']}}
                                ]
                            }
                        ],
                        "fields": "userEnteredValue"
                    }
                } )
        # Perform the batch update
    
        print(requests)

        worksheet.spreadsheet.batch_update({"requests": requests})
        
        
        #ADD IN MONTHS
        list_ws = spreadsheet.worksheets()
        list_ws = [sheet for sheet in list_ws if sheet.title in LIST_MONTHS_TO_ADAPT]
        for i, ws in enumerate(list_ws):
            ws_name = ws.title
            print(ws_name)
            if ws_name != 'Groupe':
                
                print(f"--------- ADDING -- {ws_name} ---------")

                col_index_presences = get_column_index_presences(ws)
                request_month = []
                for row in list_index_to_add:
                    request_month.append(
                        {
                            "insertDimension": {
                                "range": {
                                    "sheetId": ws.id,
                                    "dimension": "ROWS",
                                    "startIndex": row + 3,
                                    "endIndex": row + 1 + 3
                                }
                            }
                        })
                    request_month.append(
                        {
                            "copyPaste": {
                                "source": {
                                    "sheetId": ws.id,
                                    "startRowIndex": row + 2,
                                    "endRowIndex": row + 3,
                                    "startColumnIndex": col_index_presences,
                                    "endColumnIndex": col_index_presences + 2
                                },
                                "destination": {
                                    "sheetId": ws.id,
                                    "startRowIndex": row + 3,
                                    "endRowIndex": row + 1 + 3,
                                    "startColumnIndex": col_index_presences,
                                    "endColumnIndex": col_index_presences + 2
                                },
                                "pasteType": "PASTE_FORMULA"
                            }
                        })
                print(request_month)
                # Perform the batch update
                ws.spreadsheet.batch_update({"requests": request_month})
    else:
        print(f'{YELLOW}---- NOTHING TO ADD ----{RESET}')

def missing_months(spreadsheet,list_months):
    worksheet_list = spreadsheet.worksheets()
    worksheets_list_names = [ws.title for ws in worksheet_list]
    #print(f"MONTHS PRESENT DRIVE = {worksheets_list_names}")
    month_names = [datetime.date(2025, month, 1).strftime("%B") for month in list_months]
    months_to_add = [] 
    for month in month_names:
        if month not in worksheets_list_names:
            months_to_add.append(month)
    print(f"MONTHS TO ADD = {months_to_add}")

    dict_months = {}
    for i,month_number in enumerate(list_months):
        dict_months[month_names[i]] = month_number

    months_to_add_numbers = [dict_months[month_name] for month_name in months_to_add]
    print(f"MONTHS_TO_ADD_NUMBERS = {months_to_add_numbers}")

    return months_to_add_numbers

def add_months(groupe, spreadsheet, membres_groupe_db, list_months=[9,10,11,12,1,2,3,4,5,6], add_conditional_format=True):
    sleep = 20
    file = spreadsheet

    list_missing_month = missing_months(spreadsheet,list_months)

    if len(list_missing_month) >= 1:
        #ADD SHEETS FOR EACH MONTH AND CONTENT
        print('----- SHEETS FOR EACH MONTH')
        for month in list_missing_month: #REMOVE RANGE FOR PROD
            month_name = datetime.date(2025, month, 1).strftime("%B")
            print(f'{BLUE}------ ADDING SHEET FOR {month_name}{RESET}')
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
            worksheet.update([['=IMPORTRANGE(Groupe!D2;"Groupe!A2:C30")']],'A4', value_input_option='USER_ENTERED')

            #DATES
            print(f'---------- ADDING DATES')
            worksheet.update_acell('C2','Dates')
            worksheet.format("C2", { 'backgroundColor': {
                                                        'red':252/255,
                                                        'green':202/255,
                                                        'blue':159/255}})
            #GET DATES depending on GROUP
            if groupe not in (2,5):
                results = date_global(month)
            elif groupe == 2: #Elite
                results = date_elite(month)
            elif groupe == 5: #EJ2
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
            worksheet.update_acell('C3','Jours')
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

            if add_conditional_format:
                set_presence_conditional_format(
                    worksheet,
                    column_letter_to_index(list_collumns[-3]) + 1,
                    len(membres_groupe_db)
                )
            
            #PERCENTAGE
            print(f'---------- ADDING PERCENTAGE')
            formulas_list_percent = []
            formulas_list_presence = []
            for i in range(len(membres_groupe_db)):
                range_pourcentage = f'D{i+4}:{list_collumns[-3]}{i+4}'
                formulas_list_percent.append([f'=NB.SI({range_pourcentage};"v")/(${list_collumns[-2]}$1-NB.VIDE({range_pourcentage})-NB.SI({range_pourcentage};"b")-NB.SI({range_pourcentage};"o")-NB.SI({range_pourcentage};"m"))'])
                formulas_list_presence.append([f'=NB.SI({range_pourcentage};"v")'])

            range_formulas = f'{list_collumns[-2]}4:{list_collumns[-2]}{str(3+len(membres_groupe_db))}'
            worksheet.update(range_formulas,
                                formulas_list_percent, 
                                value_input_option='USER_ENTERED')
            
            worksheet.format(range_formulas, { 'numberFormat': {
                                                    'type': 'PERCENT'}})
            
            #NUMBER EMPTY CELLS
            worksheet.update([[f'=NB.VIDE(D1:{list_collumns[-3]}1)']], 
                                f'{list_collumns[-2]}1',
                                value_input_option='USER_ENTERED')
            
            #NUMBER TRAINING
            print(f'---------- ADDING NB TRAININGS')
            range_presence = f'{list_collumns[-1]}4:{list_collumns[-1]}{str(3+len(membres_groupe_db))}'
            worksheet.update(range_presence,
                                formulas_list_presence, 
                                value_input_option='USER_ENTERED')
            
            #NUMBER OF SWIMMERS
            print(f'---------- ADDING NB SWIMMERS')
            range_swimmers_formulas = f'{list_collumns[0]}{str(5+len(membres_groupe_db))}:{list_collumns[-3]}{str(5+len(membres_groupe_db))}'
            formulas_list_swimmers = []
            for i,day in enumerate(dates):
                range_swimmer = f'{list_collumns[i]}4:{list_collumns[i]}{str(3+len(membres_groupe_db))}'
                formulas_list_swimmers.append(f'=NB.SI({range_swimmer};"v")')

            worksheet.update(range_swimmers_formulas,
                                [formulas_list_swimmers], 
                                value_input_option='USER_ENTERED')
            
            #AUTO RESIZE COLUMNS WIDTH
            worksheet.columns_auto_resize(indexes[0],indexes[1])
             #PROTECTED RANGE
            worksheet.add_protected_range('A1:C40', ['technique.cnba@gmail.com'])

            print(f'{YELLOW}##### SLEEPING {sleep} SECONDS BETWEEN WORKSHEET #####{RESET}')
            time.sleep(sleep)
    else:
        print(f'{YELLOW}---- NO MISSING MONTH ----{RESET}')

def add_protection(spreadsheet, list_months = [9,10,11,12,1,2,3,4,5,6]):  
    missing_protection_groupe = True
    missing_protection_months = False
    ### PROTECT 'GROUPE' WORKSHEET
    groupe_ws = spreadsheet.worksheet('Groupe')
    if spreadsheet.list_protected_ranges(groupe_ws.id) == []:
        print(f'{BLUE}---- ADDING PROTECTION {RESET}')
        print(f'------ ADDING PROTECTION FOR GROUPE')
        groupe_ws.add_protected_range('A1:D40', ['pythonaccount@pythonsheet-346808.iam.gserviceaccount.com','remicrepin25@gmail.com','technique.cnba@gmail.com'])
    else :
        missing_protection_groupe = False
    
    ### PROTECT MONTHS WORKSHEET
    
    worksheet_list = spreadsheet.worksheets()
    worksheets_list_names = [ws.title for ws in worksheet_list]

    for month in list_months: #REMOVE RANGE FOR PROD
        month_name = datetime.date(2025, month, 1).strftime("%B")
        if month_name in worksheets_list_names:
            ws = spreadsheet.worksheet(month_name)
            if spreadsheet.list_protected_ranges(ws.id) == []:
                if not missing_protection_groupe and not missing_protection_months:
                    print(f'{BLUE}---- ADDING PROTECTION {RESET}')
                print(f'------ ADDING PROTECTION FOR {month_name}')
                ws.add_protected_range('A1:C40', ['remicrepin25@gmail.com','technique.cnba@gmail.com'])
                missing_protection_months = True
    
    if not missing_protection_groupe and not missing_protection_months:
        print(f'{YELLOW}---- NO MISSING PROTECTION ----{RESET}')
    
    return None
