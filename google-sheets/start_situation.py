import model
import config 
import pandas as pd

if __name__ == "__main__":
    #INIT
    conf = config.config()
    client = model.init_client_drive()
    
    # LIST OF FILES FOLDER
    list_files = client.list_spreadsheet_files(folder_id = conf.folder_id) #list of dict

    #DELETE Existing file
    for file in list_files:
        #if file['name'] == 'Presences Elite':
        print(f'------ DELETING SPREADSHEET {file["name"]} ------')
        client.del_spreadsheet(file['id'])

    print(f'------ CREATING SPREADSHEET ------')
    model.create_spreadsheets(conf,test=True)

    spreadsheet = client.open('Presences Elite', folder_id=conf.folder_id)
    
    print('____ DELETING ROWS ____')
    worksheet = spreadsheet.worksheet('Groupe')
    worksheet.delete_rows(3)
    worksheet.delete_rows(7)

    print('____ ADDING SEQUENCE ____')
    worksheet_m = spreadsheet.worksheet('septembre')
    worksheet_m.update('D4:D18', [[i] for i in range(1,16)])
