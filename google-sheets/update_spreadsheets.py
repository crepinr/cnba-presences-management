import model
import config 
import pandas as pd
import time

############################
### CONFIGURATION 
############################

### COLORS FOR THE SHELL 

BLUE = "\033[34m"
GREEN = "\033[32m"
YELLOW = "\033[33m"
RED = "\033[31m"
RESET = "\033[0m"

### PARAMETERS TO CHECK 
# conf.folder_id -> Google Folder in which the script is applied
# SHEETS_TO_MODIFY -> Optional parameter, name of the group not the sheet like 'Elite',
#                      if empty it is by default applied to all sheets in the folder

### SCRIPT PARAMETERS 

UPDATE_DB = False
REMOVE_MEMBERS = True
REMOVE_MEMBERS_AUCUN = True
ADD_MEMBERS = True
ADD_MONTHS = False
ADD_PROTECTED_RANGE = True
ADD_CONDITIONAL_FORMAT = True

SHEETS_TO_MODIFY = ['Elite Jeunes']

############################
### MAIN 
############################

if __name__ == "__main__":
    
    conf = config.config()     
    
    if UPDATE_DB :
        model.update_db()
    else :
        pass
    
    #gspread client to connect to drive
    client = model.init_client_drive() 

    #Get members and groups data from local DB
    df_membres, df_groupes = model.read_db()
    list_nomsprenoms_aucun = model.list_np_db_aucun(df_membres)

    # LIST OF FILES IN FOLDER DRIVE
    #Return : list of dict
    list_files = client.list_spreadsheet_files(folder_id = conf.folder_id) 
    
    # LIST OF DICT GROUPS & FILE ID
    list_sheets = model.list_sheets(list_files) 
    print(list_sheets)

    if SHEETS_TO_MODIFY != []:
        list_sheets_to_treat = [sheet for sheet in list_sheets if sheet['name'] in SHEETS_TO_MODIFY]
    else:
        list_sheets_to_treat = list_sheets
    print(list_sheets_to_treat)

    for sheet in list_sheets_to_treat:
        print(f"{GREEN}{sheet['name']}{RESET}")
        if sheet['name'] != '':
            if sheet['name'] != '':
                #OPEN WORKSHEET AND EXTRACT LIST MEMBRE DRIVE
                spreadsheet = client.open_by_key(sheet['key']) 
                worksheet = spreadsheet.worksheet('Groupe')
                list_nomsprenoms_drive = model.list_np_drive(worksheet)   

                #EXTRACT LIST MEMBRE DB
                list_nomsprenoms_db, membres_groupe_db = model.list_np_db(sheet, df_membres, df_groupes)
                membres_groupe_db = membres_groupe_db.to_dict('records')

                if REMOVE_MEMBERS:
                    # Enabled in the beginning of the year for updating members not starting the year
                    # During the year disabled, such that when a swimmer changes group, past attendences in previous group are kept
                    model.remove_members(list_nomsprenoms_drive,list_nomsprenoms_db, worksheet, spreadsheet)

                if REMOVE_MEMBERS_AUCUN:
                    # Enabled in the beginning of the year for updating members not starting the year
                    # During the year disabled, such that when a swimmer changes group, past attendences in previous group are kept
                    model.remove_members_aucun(list_nomsprenoms_drive,list_nomsprenoms_aucun, worksheet, spreadsheet)
                
                if ADD_MEMBERS: 
                    # Add new members to the group (new member, member changed group)               
                    model.add_members(list_nomsprenoms_drive,list_nomsprenoms_db, worksheet, spreadsheet, membres_groupe_db)

                if ADD_MONTHS:
                    # Add missing months if eventually not all would be present
                    nom_groupe = sheet['name']
                    groupe_id = df_groupes.loc[df_groupes['nom_groupe'] == nom_groupe, 'id'].iloc[0]
                    model.add_months(
                        groupe_id,
                        spreadsheet,
                        membres_groupe_db,
                        add_conditional_format=ADD_CONDITIONAL_FORMAT
                    )

                if ADD_PROTECTED_RANGE:
                    model.add_protection(spreadsheet)

                if ADD_CONDITIONAL_FORMAT:
                    model.add_presence_conditional_formatting(spreadsheet, len(membres_groupe_db))

                print(f'##### SLEEPING 10 SECONDS BETWEEN SPREADSHEETS #####')
                time.sleep(10)
        else :
            print("ERROR : EMPTY SHEET NAME")
            
          



        

    

   
