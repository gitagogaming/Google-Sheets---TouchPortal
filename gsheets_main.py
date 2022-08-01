
import time
import json
import gspread
from gspread import utils, worksheet
import gsheets_auth
## https://github.com/burnash/gspread/issues



## Make a manual update action   # user can utilize load config for this ??
## Format a Cell Action not working - colors are changing when you dont add a value
## Issue with being able to get created_time and last_updated from .spreadsheet 

######### POSSIBLE TO DO #######

## Find & Return a cell based on string function - return true/false ?
## Dont forget about option to use API Key instead



### CUSTOM STATES AFTER WE FIGURE OUT HOW TO GET
# make states for - #Row Count, Column Count#, Frozen Column Count, Frozen Row Count


### Potential Setup by User Needed
"""
## Unsure if this step is needed yet... 
Go to https://console.developers.google.com/
- Create a new project; this can be done in multiple ways but here's one:
- In the upper left 'Select a project' drop down, open it up and use 'New Project'
- Give it a name example: TP GSheets  and create it (this may take a moment)
-- In the top-middle search bar, search for 'Google Sheets API', navigate to that and press the 'Enable' button, then go back to your project
-- Navigate through 'Credentials' in the left-side menu, click 'Create Credentials' at the top and choose 'OAuth client ID'

"""


class G_Sheets:
    spreadsheet_id = ""
    worksheet_name = ""
    the_cells = []
    worksheet = None
    auto_update = False
    


#--------------------#
#--- AUTH TRIGGER ---#
#--------------------#
gc = gsheets_auth.G_Sheets_Auth.do_auth()



class G_Sheets_Config:
#-------------------------#
#--- LOAD SHEET CONFIG ---#
#-------------------------#
### This loads the config file set by the user with Cell names, Sheet ID, Name ETC.
# We use these details to open the sheet and get the values of the cells we want to monitor
    def load_sheets_config(location):
        """
        - Loads the Config File from the location specified
        - Sets G_Sheets class variables for spreadsheet_id, name and cells to watch.
        """
        # E:\Other Assets\Tactical_Banditry\SheetsIO_TB\configs\TB_PANEL.json
        chat_config = open(location, "r")  ## Open and Read
        the_chatconfig_dict = json.load(chat_config)
        chat_config.close()

        G_Sheets.the_cells = the_chatconfig_dict['cells']
        G_Sheets.spreadsheet_id = the_chatconfig_dict['spreadsheetId']
        G_Sheets.worksheet_name = the_chatconfig_dict['worksheetName']
        return None


    # Open the Sheet - Automatically done when we load the config
    def open_sheet(spreadsheet_id, worksheet_name):
        """
        Open the Workbook by ID and desired Sheet Name 
        - Currently this is automatic based on config being loaded
        - In future we will allow choices to be set for any sheet
        """
        try:
          #  gc.worksheets()
            sheet = gc.open_by_key(spreadsheet_id)
            worksheet = sheet.worksheet(worksheet_name)
            worksheet.add_rows
            TPClient.stateUpdate(stateId="gitago.gsheets.state.sheet.title", stateValue=worksheet.title)
            TPClient.stateUpdate(stateId="gitago.gsheets.state.sheet.row_count", stateValue=str(worksheet.row_count))
            TPClient.stateUpdate(stateId="gitago.gsheets.state.sheet.col_count", stateValue=str(worksheet.col_count))
          #  print(worksheet.creationTime)
            print(worksheet.url)
           # print(worksheet.lastUpdateTime)
         #   print(worksheet.updated)

            return worksheet
        except:
            print("Error opening sheet")
            return None, None


    def load_worksheet(spreadsheet_id, worksheet_name):
        """ Loading Worksheet from spreadsheet_id
        - Seperated this as it may be needed more than once rather than loading the config every time as well"""
        
        ## Loading the Sheet Object to G_Sheets.worksheet
        G_Sheets.worksheet = G_Sheets_Config.open_sheet(spreadsheet_id, worksheet_name)



### NEED TO FIX IT BREAKING LOOP WHEN A CELL IS EMPTY AND WE WANT TO CAPTURE IT...
    def create_states_from_config(state_category=None):
        """
        Creates and Updates States in TouchPortal based on the config file
        - Loop thru the config, pull the value of the cell and update the state in TouchPortal
        """
        choice_list = []
        all_cells = G_Sheets.worksheet.get_values()
        try:
            for thing in G_Sheets.the_cells:
                if thing.get('fileExtension'):
                    print("trigger image stuff?")
                
                cell_num = thing.get('cell')
                cell_name = thing.get('name')
                coords = utils.a1_range_to_grid_range(cell_num)
                 #coords = utils.a1_to_rowcol(cell_num) this doesnt work??
                 
                 ## Appending to Choice List for Smart Actions
                choice_list.append(cell_name)
    
                ##########################################################
                ### HOW TO OVERCOME THIS ERROR WHEN LAST ONE IS EMPTY ? ##
                ##########################################################
                """ Set all the TP States needed based on whats found here"""
                # if Cell is EMPTY then it returns an error and we make that blank in the except
                cells_value = all_cells[coords['startRowIndex']][coords['startColumnIndex']]
                if state_category:
                    TPClient.createState(stateId=f"gitago.gsheets.state.{cell_name}", description=f"GS | {cell_name}", value=str(cells_value), parentGroup=state_category)
                else:
                    TPClient.stateUpdate(stateId=f"gitago.gsheets.state.{cell_name}", stateValue=str(cells_value))
        except IndexError:
            """ This Cell Must Be Empty"""
            print(cell_name, "is empty")
            TPClient.choiceUpdate(choiceId=f"gitago.gsheets.act.swap_cell.fromcell.smart", values=choice_list)
            TPClient.choiceUpdate(choiceId=f"gitago.gsheets.act.swap_cell.tocell.smart", values=choice_list)
            



# Delete Row(s) from the Sheet
def delete_rows(start_index, end_index):
    """ Delete a Range of Rows, or a Single Row
    - Need to figure out how to get new columns/rows
    """
    try:
        G_Sheets.worksheet.delete_rows(int(start_index), int(end_index))
    except:
        print("Error Deleting Rows")


# Add Row(s) to the Sheet
def add_rows(start_index):
    """ Add a Range of Rows, or a Single Row
    - Need to figure out how to get new columns/rows
    """
    try:
        G_Sheets.worksheet.add_rows(int(start_index))
    except:
        print("Error Adding Rows")


# Hide Row(s) on the Sheet
def hide_rows(start_index, end_index):
    """ Hide a Range of Rows, or a Single Row
    """
    try:
        G_Sheets.worksheet.hide_rows(int(start_index), int(end_index))
    except:
        print("Error Hiding Rows")


# Un-Hide Row(s) on the Sheet
def unhide_rows(start_index, end_index):
    """ Hide a Range of Rows, or a Single Row
    """
    try:
        G_Sheets.worksheet.unhide_rows(int(start_index), int(end_index))
    except gspread.exceptions.APIError as err:
        print("Error Hiding Rows", err)


# Delete Col(s) from the Sheet
def delete_cols(start_index, end_index):
    """ Delete a Range of Columns, or a Single Column
    - Need to figure out how to get new columns/rows
    """
    count = check_for_a1_notation(start_index, end_index)
    ## Now that we checked, lets do the thing
    if count >= 0:
        try:
            G_Sheets.worksheet.delete_columns(int(start_index), end_index)
        except:
            print("Error Deleting Columns - Numeric")
    elif count == -2 or count == -1:
        ## This allows user to use A1 notation to delete a column instead of numbers
        try:
            start_index = utils.column_letter_to_index(start_index)
            if end_index != "":
                end_index = utils.column_letter_to_index(end_index)
            G_Sheets.worksheet.delete_columns(start_index, end_index)
        except:
            print("Error deleting Columns - A1 Notation")


# Add Col(s) to the Sheet
def add_cols(start_index):
    """ Add a Range of Columns, or a Single Column
    - Need to figure out how to get new columns/rows
    """
    try:
        G_Sheets.worksheet.add_cols(int(start_index))
    except:
        print("Error Adding Columns")


# Hide Col(s) on the Sheet
def hide_cols(start_index, end_index):
    """ Hide a Range of Columns, or a Single Column
    """
    count = check_for_a1_notation(start_index, end_index)
    ## Now that we checked, lets do the thing
    if count >= 0:
        try:
            G_Sheets.worksheet.hide_columns(int(start_index), end_index)
        except gspread.exceptions.APIError as err:
            print("Error Hiding Columns - Numeric", err)
    
    elif count == -2 or count == -1:
        ## This allows user to use A1 notation to delete a column instead of numbers
        try:
            start_index = utils.column_letter_to_index(start_index)
            if end_index != "":
                end_index = utils.column_letter_to_index(end_index)
            G_Sheets.worksheet.hide_columns(start_index, end_index)
        except gspread.exceptions.APIError as err:
            print("Error Hiding Columns - A1 Notation", err)


# Un-Hide Col(s) on the Sheet
def unhide_cols(start_index, end_index):
    """ Hide a Range of Columns, or a Single Column
    """
    count = check_for_a1_notation(start_index, end_index)
    if count >= 0:
        try:
            G_Sheets.worksheet.unhide_columns(int(start_index), end_index)
        except gspread.exceptions.APIError as err:
            print("Error unhiding Columns - Numeric", err)
            
    elif count == -2 or count == -1:
        ## This allows user to use A1 notation to delete a column instead of numbers
        try:
            start_index = utils.column_letter_to_index(start_index)
            if end_index != "":
                end_index = utils.column_letter_to_index(end_index)
            G_Sheets.worksheet.unhide_columns(start_index, end_index)
        except gspread.exceptions.APIError as err:
            print("Error unhiding Columns - A1 Notation", err)



# Update a Cell with value
def update_gsheet_cell(cell, value, input=worksheet.ValueInputOption.raw):
    G_Sheets.worksheet.update(cell, value, value_input_option=input)  #unsure if raw vs userinput matters here
   #G_Sheets.worksheet.update(values=worksheet.ValueInputOption.user_entered)


# Copy Single Cell to Another
def copy_cell_to_cell(from_cell=None, to_cell=None):
    """ Copy a One Cell to Another Cell"""
    batch_get = G_Sheets.worksheet.batch_get([from_cell, to_cell])
    G_Sheets.worksheet.batch_update([{
            'range': to_cell,
            'values': batch_get[0],
        }, {
            'range': from_cell,
            'values': batch_get[1],
        }])
    
        # val1 = G_Sheets.worksheet.acell(from_cell).value
        # val2 = G_Sheets.worksheet.acell(to_cell).value
        # G_Sheets.worksheet.update(to_cell, val1)
        # G_Sheets.worksheet.update(from_cell, val2)


# Copy Range A to Range B
def copy_rangecell_to_rangecell(from_cells=None, to_cells=None):
    """ Copy a Range of Cells to another Range of cells
    """
    batch_get = G_Sheets.worksheet.batch_get([from_cells, to_cells])
    G_Sheets.worksheet.batch_update([{
            'range': to_cells,
            'values': batch_get[0],
        }, {
            'range': from_cells,
            'values': batch_get[1],
        }])




## Format a Cell - Needs work
def format_a_cell(cell, color, size):
    """ Format a Google Sheet cell
    - Still need to build an action
    - issue with text color changing when not changing the color...
    
    - Investigate the Library for this
    - need to be able to issue Nones to avoid issueing a command
    """
    G_Sheets.worksheet.format(cell, {"textFormat": {
                                                    "fontSize": size,
                                                    "bold": None}})
    # hiding some notes                          
    def athing():
        pass
        # batch format example   formats = [
        # batch format example       {
        # batch format example           "range": "A1:C1",
        # batch format example           "format": {
        # batch format example               "textFormat": {
        # batch format example                   "bold": True,
        # batch format example               },
        # batch format example           },
        # batch format example       },
        # batch format example       {
        # batch format example           "range": "A2:C2",
        # batch format example           "format": {
        # batch format example               "textFormat": {
        # batch format example                   "fontSize": 16,
        # batch format example               },
        # batch format example           },
        # batch format example       },
        # batch format example   ]
        # batch format example   G_Sheets.worksheet.batch_format(formats)
        ##     G_Sheets.worksheet.format(cell, {
        ##     "backgroundColor": {
        ##     "red": 1,
        ##     "green": 1,
        ##     "blue": 1
        ##     },
        ##     "horizontalAlignment": "CENTER",
        ##     "textFormat": {
        ##      "foregroundColor": {
        ##        "red": 0,
        ##        "green": 0,
        ##        "blue": 0
        ##      },
        ##   # "fontFamily": str,
        ##    "fontSize": size,
        ##   # "bold": bool,
        ##   # "italic": bool,
        ##   # "strikethrough": bool,
        ##   # "underline": bool,
        ##   # "link": { object ('Link') }
        ##     
        ##     }
        ##     })

# The Sheet Update Loop
def update_loop():
    """ Updates Every 1 second until it finds a change, then we update all the values"""
    before = []
    count = 0
    TPClient.stateUpdate(stateId="gitago.gsheets.state.auto_update", stateValue="RUNNING")
    
    
    """ whats all in this "object" ... we need this to get row count on auto update..."""
    # Saving the Current Sheet Object to G_Sheets.worksheet
    G_Sheets.worksheet = G_Sheets_Config.open_sheet(G_Sheets.spreadsheet_id, G_Sheets.worksheet_name)
    
    ## Comparing the Values 
    while G_Sheets.auto_update:
        new_values = G_Sheets.worksheet.get_all_values()
        if before != new_values:
            before = new_values
            G_Sheets_Config.create_states_from_config()
                    ### Think need to find a way to duplicate the worksheet so we arent constantly making calls ?
            TPClient.stateUpdate(stateId="gitago.gsheets.state.sheet.title", stateValue=G_Sheets.worksheet.title)
            TPClient.stateUpdate(stateId="gitago.gsheets.state.sheet.row_count", stateValue=str(G_Sheets.worksheet.row_count))
            TPClient.stateUpdate(stateId="gitago.gsheets.state.sheet.col_count", stateValue=str(G_Sheets.worksheet.col_count))
        
        time.sleep(2)
        count+=1
        print(count)
        

    TPClient.stateUpdate(stateId="gitago.gsheets.state.auto_update", stateValue="STOPPED")
        


def check_for_a1_notation(start_index, end_index):
    """ Returns Count """
    count = 0
    check = start_index.isdecimal()
    if check: 
        count +=1
    else: 
        count -=1
    
    ## Is end_index blank ? if so lets make it none
    if end_index == "":
        end_index = None
    else:
        check = end_index.isdecimal()
        if check:
            end_index = int(end_index)
            count +=1
        else: 
            count -=1
    return count



## unfinished / Unused.. wait for a person whom needs to know how to implement
def update_range_gsheet_cell(cell_range=None, value=None):
    """ Update a Range of Cells with a particular set of Values
    - Unsure how this may be used with TP"""
    G_Sheets.worksheet.batch_update([{
                'range': 'D1',
                'values': [['D2']],
            }, {
                'range': 'D2',
                'values': [['D1']],
            }])
    pass



import TouchPortalAPI  # Import the api

Debug = True
TPClient = TouchPortalAPI.Client('gitago.gsheets')

@TPClient.on('info')
def onStart(data):
    print('[CONNECTED]', data)
    TPClient.stateUpdate(stateId="gitago.gsheets.state.auto_update", stateValue="STOPPED")



# Action handlers, called when user activates one of this plugin's actions in Touch Portal.
@TPClient.on('action')
def onActions(data):
    print(data)
    
    # Load the Config and Create the States from it
    if data["actionId"] == "gitago.gsheets.act.sync_sheets":
        """ 
        - Data0 = Config File Location
        - Data1 = Name for the States Category 
        """
        # Load the Config file.
        G_Sheets_Config.load_sheets_config(data['data'][0]['value'])
        
        # Load the worksheet using the Sheet ID and Sheet Name
        G_Sheets_Config.load_worksheet(G_Sheets.spreadsheet_id, G_Sheets.worksheet_name)
        
        # Create the States for the Cells in the config file
        G_Sheets_Config.create_states_from_config(str(data['data'][1]['value']))
    
    # Update a Cell
    if data["actionId"] == "gitago.gsheets.act.update_cell":
        try:
            update_gsheet_cell(data['data'][0]['value'], data['data'][1]['value'])
        except:
            print("[ERROR] Error Updating Cell")
    
    # Update a Cell (Smart)
    if data["actionId"] == "gitago.gsheets.act.update_cell.smart":
        for x in G_Sheets.the_cells:
            if x['name'] == data['data'][0]['value']:
                the_cell = x['cell']
        try:
            update_gsheet_cell(the_cell, data['data'][1]['value'])
        except:
            print("error in swap_cell.smart")
    
    # Swap Range to Range
    if data["actionId"] == "gitago.gsheets.act.swap_cell_range":
        try:
            copy_rangecell_to_rangecell(data['data'][0]['value'], data['data'][1]['value'])
        except:
            print("[ERROR] Error Swapping Cell Range")
    
    # Swap CellA and CellB
    if data["actionId"] == "gitago.gsheets.act.swap_cell":
        try:
            copy_cell_to_cell(data['data'][0]['value'], data['data'][1]['value'])
        except:
            print("[ERROR] Error Swapping Cell")
    
    # Swap CellA and CellB (Smart)
    if data["actionId"] == "gitago.gsheets.act.swap_cell.smart":
        for x in G_Sheets.the_cells:
            if x['name'] == data['data'][0]['value']:
                the_from = x['cell']
            if x['name'] == data['data'][1]['value']:
                the_to = x['cell']
        try:
            copy_cell_to_cell(the_from, the_to)
        except:
            print("error in swap_cell.smart")
    
    # Toggle Auto Update
    if data["actionId"] == "gitago.gsheets.act.auto_update":
        if data['data'][0]['value']== "ON":
            G_Sheets.auto_update = True
            update_loop()
        else:
            G_Sheets.auto_update = False

    # Format a Cell
    if data["actionId"] == "gitago.gsheets.act.format_cell":
        format_a_cell(cell = data['data'][0]['value'], 
                      color = data['data'][1]['value'], 
                      size = data['data'][2]['value'])
        

    
    # Delete Columns / Rows
    if data["actionId"] == "gitago.gsheets.act.delete_columns_rows":
        if data['data'][0]['value'] == "Rows":
            delete_rows(data['data'][1]['value'], data['data'][2]['value'])
        if data['data'][0]['value'] == "Columns":
            delete_cols(data['data'][1]['value'], data['data'][2]['value'])
    
    # Add Rows / Cols
    if data["actionId"] == "gitago.gsheets.act.add_rows_columns":
        if data['data'][0]['value'] == "Rows":
            add_rows(data['data'][1]['value'])
        if data['data'][0]['value'] == "Columns":
            add_cols(data['data'][1]['value'])
    
    # Hide Columns / Rows
    if data["actionId"] == "gitago.gsheets.act.hide_columns_rows":
        if data['data'][0]['value'] == "Rows":
            hide_rows(data['data'][1]['value'], data['data'][2]['value'])
        if data['data'][0]['value'] == "Columns":
            hide_cols(data['data'][1]['value'], data['data'][2]['value'])
            
    if data["actionId"] == "gitago.gsheets.act.unhide_columns_rows":
        if data['data'][0]['value'] == "Rows":
            unhide_rows(data['data'][1]['value'], data['data'][2]['value'])
        if data['data'][0]['value'] == "Columns":
            unhide_cols(data['data'][1]['value'], data['data'][2]['value'])
        
        
        

@ TPClient.on('settings')
def onSettings(data):
    print('received data from settings!')
    print(data['values'])
    global Debug
    Debug = (data['values'][0]["Debug"])


@ TPClient.on('closePlugin')
def onShutdown(data):
    print('Received shutdown message!')
    TPClient.disconnect()



TPClient.connect()








""" 
WHEN WE RUN INTO ISSUE WITH TOO MANY UPDATES PER MINUTE 
UNTIL THEN WE WILL USE CLIENT AUTH AND MONITOR HOW OFTEN PEOPLE USE IT?
"""















### example using gsheets module and using API key to pull a sheet.. but havent done it yet

# import logging
# 
# from gsheets import Sheets
# import gsheets
# 
# 
# #WORKSHEET_IDS = [0, 1747240182, 894548711]
# SHEET_ID = '1bQ6wtP-OArPFJ5xZYfPqxOmPosRI9ilVzE65NfcWd0c'
# WORKSHEET_NAMES = ['Tactical_Bandits']
# 
# 
# logging.basicConfig(format='[%(levelname)s@%(name)s] %(message)s',
#                     level=logging.INFO)

## https://github.com/xflr6/gsheets/blob/24e99f67e5f6bc49facc85f5db305caa63ccbdef/tests/test_api.py
# user_API_key = 'AIzaSyDDrMWL_gs14nH1Q1lSYhDzSoyYn1bgxKA'
# sheets = Sheets.from_developer_key(user_API_key)
# gc = sheets[SHEET_ID]
# sheet_values = gc[0]
# 
# old = []
# while True:
#     gc = sheets[SHEET_ID]
#     sheet_values = gc[0]
#     new = sheet_values.values()
#     if old != new:
#         print("Its not the same")
#         old = new
#     time.sleep(1)
#     print("Not it")
# print(sheet_values['A1'])
#print(sheets[SHEET_ID])

###########################################################
###########################################################
# url = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}'
# s = sheets.get(url).first_sheet
# print(s.at(1,1))
#print(s.first_sheet.values())
#print(s.find(WORKSHEET_NAMES[1]))

#print(sheets.find(WORKSHEET_NAMES[1]))

#print(s[WORKSHEET_IDS[1]].at(row=1, col=1))

#s.sheets[1].to_csv('Spam.csv', encoding='utf-8', dialect='excel')

#csv_name = lambda infos: '%(title)s - %(sheet)s.csv' % infos
#s.to_csv(make_filename=csv_name)

#print(s.find(WORKSHEET_NAMES[1]).to_frame(index_col='spam'))














#sheet_id, sheet_name, cells = load_sheets_config('E:\Other Assets\Tactical_Banditry\SheetsIO_TB\configs\TB_PANEL.json')
#worksheet = open_sheet(sheet_id, sheet_name)
#worksheet.update('B1', 'Bingo!')

###  excel ?   import openpyxl
###  excel ?   
###  excel ?   
###  excel ?   
###  excel ?   class Excel:
###  excel ?       sheet_names = []
###  excel ?       
###  excel ?   
###  excel ?   
###  excel ?   
###  excel ?   ## make a class for excel
###  excel ?   class Excel_Funcs:
###  excel ?       ## make an init function
###  excel ?       wb = openpyxl.load_workbook(r'C:\Users\dbcoo\Downloads\example.xlsx')
###  excel ?       
###  excel ?       
###  excel ?       def get_all_sheet_names():
###  excel ?           """ Get all the Sheet Names in the Workbook """
###  excel ?           Excel.sheet_names=Excel_Funcs.wb.get_sheet_names()
###  excel ?   
###  excel ?   
###  excel ?       def get_sheet_titles():
###  excel ?           for x in Excel_Funcs.wb.sheetnames:
###  excel ?               print(x)
###  excel ?   
###  excel ?       def get_sheet_by_name(sheet_name):
###  excel ?           """ Get the Sheet by Name """
###  excel ?           Pull = Excel_Funcs.wb[sheet_name]
###  excel ?           for i in range(1, 8, 2):
###  excel ?               print(i, Pull.cell(row=i, column=2).value)
###  excel ?           return Pull
###  excel ?           
###  excel ?           
###  excel ?       def pull_cell_value(sheet_name, cell_name):
###  excel ?           """ Pull the Value from a Cell 
###  excel ?           - Specify Sheet Name & Cell Name 
###  excel ?           - Returns the Value inside of Cell"""
###  excel ?           Pull = Excel_Funcs.wb[sheet_name][cell_name].value
###  excel ?           return Pull
###  excel ?       
###  excel ?       
###  excel ?   Excel_Funcs.get_all_sheet_names()
###  excel ?   
###  excel ?   Excel_Funcs.get_sheet_titles()
###  excel ?   
###  excel ?   Response = Excel_Funcs.get_sheet_by_name('Sheet1')
###  excel ?   
###  excel ?   #print(Response.title)
###  excel ?   
###  excel ?   print(Excel_Funcs.pull_cell_value('Sheet1', 'A1'))




