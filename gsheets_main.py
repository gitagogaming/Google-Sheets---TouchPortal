## https://github.com/burnash/gspread/issues
import gspread
import json
from gspread import utils

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


class Google_Sheets:
    sheet_ID = ""
    sheet_name = ""
    the_cells = []
    worksheet = None
    auto_update = False


def do_auth():
    """ Authenticate with Google """
    ## using service account which means user doesnt have to auth, long as the sheet is public
  # gc = gspread.service_account(filename = r'touchportal-sheets-3b2c8684cab1.json')
  # return gc
    try:
        gc = gspread.oauth(
            credentials_filename='client_secrets.json')
        return gc
    except:
        print("Error Authenticating")
        return None


#--------------------#
#--- AUTH TRIGGER ---#
#--------------------#
gc = do_auth()




#########################################################
#########################################################

### This loads the config file set by the user with Cell names, Sheet ID, Name ETC.
# We use these details to open the sheet and get the values of the cells we want to monitor
#sheet_ID, sheet_name, the_cells = load_sheets_config()

def load_sheets_config(location):
    """
    - Loads the Config File from the location specified
    - Sets Google_Sheets class variables for sheet_ID, name and cells to watch.
    """
    # E:\Other Assets\Tactical_Banditry\SheetsIO_TB\configs\TB_PANEL.json
    chat_config = open(location, "r")  ## Open and Read
    the_chatconfig_dict = json.load(chat_config)
    chat_config.close()
    
    Google_Sheets.the_cells = the_chatconfig_dict['cells']
    Google_Sheets.sheet_ID = the_chatconfig_dict['spreadsheetId']
    Google_Sheets.sheet_name = the_chatconfig_dict['worksheetName']
    
    return None



def open_sheet(sheet_ID, sheet_name):
    """
    Open the Workbook by ID and desired Sheet Name 
    - Currently this is automatic based on config being loaded
    - In future we will allow choices to be set for any sheet
    """
    try:
        sheet = gc.open_by_key(sheet_ID)
        worksheet = sheet.worksheet(sheet_name)
        return worksheet
    except:
        print("Error opening sheet")
        return None, None



def load_worksheet(sheet_id, sheet_name):
    """ Setting 'worksheet' in Google_Sheets class to this sheetID 
    - Seperated this as it may be needed more than once rather than loading the config every time as well"""
    Google_Sheets.worksheet = open_sheet(sheet_id, sheet_name)


### NEED TO FIX IT BREAKING LOOP WHEN A CELL IS EMPTY AND WE WANT TO CAPTURE IT...
def create_states_from_config(state_category=None):
    """
    Creates and Updates States in TouchPortal based on the config file
    - Loop thru the config, pull the value of the cell and update the state in TouchPortal
    """
    choice_list = []
    all_cells = Google_Sheets.worksheet.get_values()
    
    try:
        for thing in Google_Sheets.the_cells:
            if thing.get('fileExtension'):
                print("trigger image stuff?")
            
            cell_num = thing.get('cell')
            cell_name = thing.get('name')
            coords = utils.a1_range_to_grid_range(cell_num)
             #coords = utils.a1_to_rowcol(cell_num) this doesnt work due to... well i dont know..
             
             ## appending name of cell to list, we can then reference the Google_Sheets.the_cells to determine the A1 notation
            choice_list.append(cell_name)

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
        
        





## Future Features:
# Change BG and FG Color / Size / Alignment etc 



def update_gsheet_cell(cell, value):
   # Google_Sheets.worksheet = open_sheet(Google_Sheets.sheet_ID, Google_Sheets.sheet_name)
    Google_Sheets.worksheet.update(cell, value)
   # format_a_cell()
    #update_range_gsheet_cell()
    copy_cell_to_cell('A3', 'A20')



def copy_cell_to_cell(from_cell=None, to_cell=None):
    """ Copy a One Cell to Another Cell"""
    
    val1 = Google_Sheets.worksheet.acell(from_cell).value
    val2 = Google_Sheets.worksheet.acell(to_cell).value
    
    Google_Sheets.worksheet.update(to_cell, val1)
    Google_Sheets.worksheet.update(from_cell, val2)




import time
def update_loop():
    """ Updates Every 1 second until it finds a change, then we update all the values"""
    before = []
    count = 0
    TPClient.stateUpdate(stateId="gitago.gsheets.state.auto_update", stateValue="RUNNING")
    Google_Sheets.worksheet = open_sheet(Google_Sheets.sheet_ID, Google_Sheets.sheet_name)
    
    ## Comparing the Values 
    while Google_Sheets.auto_update:
        new_values = Google_Sheets.worksheet.get_all_values()
        if before != new_values:
            before = new_values
            create_states_from_config()
        
        time.sleep(1)
        count+=1
        print(count)
    
    TPClient.stateUpdate(stateId="gitago.gsheets.state.auto_update", stateValue="STOPPED")
        


# worksheet.batch_get(['A1:B2', 'F12'])
# https://docs.gspread.org/en/v5.4.0/api/models/worksheet.html#valuerange
# worksheet.batch_update"""



## unfinished
def update_range_gsheet_cell(cell_range=None, value=None):
    """ Update a Range of Cells with a particular set of Values
    - Unsure how this may be used with TP"""
    Google_Sheets.worksheet.update('A1:B2', [[1, 2], [3, 4]])
    pass



## unfinished
def format_a_cell():
    """ Format a Google Sheet cell
    - Still need to build an action
    """
    Google_Sheets.worksheet.format("B2", {
    "backgroundColor": {
      "red": 0.0,
      "green": 0.0,
      "blue": 0.0
    },
    "horizontalAlignment": "CENTER",
    "textFormat": {
      "foregroundColor": {
        "red": 1.0,
        "green": 1.0,
        "blue": 1.0
      },
      "fontSize": 12,
      "bold": True
    }
    })
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
        load_sheets_config(data['data'][0]['value'])
        
        # Load the worksheet using the Sheet ID and Sheet Name
        load_worksheet(Google_Sheets.sheet_ID, Google_Sheets.sheet_name)
        
        # Create the States for the Cells in the config file
        create_states_from_config(str(data['data'][1]['value']))
    
    # Update a Cell
    if data["actionId"] == "gitago.gsheets.act.update_cell":
        try:
            update_gsheet_cell(data['data'][0]['value'], data['data'][1]['value'])
        except:
            print("[ERROR] Error Updating Cell")
    
    # Update a Cell (Smart)
    if data["actionId"] == "gitago.gsheets.act.update_cell.smart":
        for x in Google_Sheets.the_cells:
            if x['name'] == data['data'][0]['value']:
                the_cell = x['cell']
        try:
            update_gsheet_cell(the_cell, data['data'][1]['value'])
        except:
            print("error in swap_cell.smart")
    
    # Swap CellA and CellB
    if data["actionId"] == "gitago.gsheets.act.swap_cell":
        try:
            copy_cell_to_cell(data['data'][0]['value'], data['data'][1]['value'])
        except:
            print("[ERROR] Error Swapping Cell")
    
    # Swap CellA and CellB (Smart)
    if data["actionId"] == "gitago.gsheets.act.swap_cell.smart":
        for x in Google_Sheets.the_cells:
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
            Google_Sheets.auto_update = True
            update_loop()
        else:
            Google_Sheets.auto_update = False



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




