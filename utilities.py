#MORE INFO: https://developers.google.com/resources/api-libraries/documentation/sheets/v4/python/latest/
#MORE INFO: https://developers.google.com/sheets/quickstart/python

import sys, os
import datetime
import json
from google_methods import update_spreadsheet, spreadsheet_info, create_sheet_service, get_range, get_file_id

"""
Reference values (row.column):

Invoice Number: 2.5
Date: 8.5

Company Name: 8.3
Company Address: 9.3
Company City: 10.3
Company Postal Code: 11.3

Description: 15.3
Description Date: 16.3
Description Cost: 38.3

Unit Quantity: 38.2
Unit Cost: 38.4

"""




# write to json
def json_dump(input, file_name):
    with open(file_name, 'w') as json_file:
        json.dump(input, json_file)

def json_load(file_name):
    with open(file_name, 'r') as json_file:
        return(json.load(json_file))

#Backup procedure - creates bkup directory and backups all files with date
def backup(filepath):                                                             
    def get_root(filepath):
        index = filepath.rfind('/')
        root = filepath[0:index + 1]
        return root
    def get_filename(filepath):
        index = filepath.rfind('/')
        dot = filepath.rfind('.')
        filename = filepath[index + 1:dot]
        return filename
    def get_extension(filepath):
        dot = filepath.rfind('.')
        extension = filepath[dot:]
        return extension
    root = get_root(filepath)
    filename = get_filename(filepath)
    extension = get_extension(filepath)
    new_root = root + 'bkup/'
    def get_date():
        time = datetime.datetime.today()
        year = str(time.year)
        def add_zero(string):
            if len(string) == 1:
                return '0' + string
            elif len(string) == 2:
                return string
            else:
                raise ValueError
        month = add_zero(str(time.month))
        day = add_zero(str(time.day))
        hour = add_zero(str(time.hour))
        minute = add_zero(str(time.minute))
        return year + month + day + hour + minute
    new_filename = filename + '_' + get_date()
    new_filepath = new_root + new_filename + extension
    def write_file():
        with open(filepath, 'r') as source:
            with open (new_filepath, 'w') as target:
                target.write(source.read())
    try:
        write_file()
    except FileNotFoundError:
        os.mkdir(new_root)
        write_file()

def backup_dicts():
    backup('json/company_dict.json')
    backup('json/project_dict.json')
    backup('json/rate_dict.json')
    backup('json/invoice_number.json')

#Backup all json dicts to dropbox
#ARGS: None
#RETURNS: None
def dropbox_backup():
    company_dict = json_load('json/company_dict.json')
    project_dict = json_load('json/project_dict.json')
    rate_dict = json_load('json/rate_dict.json')
    invoice_number = json_load('json/invoice_number.json')

    json_dump(company_dict, '/Users/jimmy/Dropbox/Accounting_Backup/company_dict.json')
    json_dump(project_dict, '/Users/jimmy/Dropbox/Accounting_Backup/project_dict.json')
    json_dump(rate_dict, '/Users/jimmy/Dropbox/Accounting_Backup/rate_dict.json')
    json_dump(invoice_number, '/Users/jimmy/Dropbox/Accounting_Backup/invoice_number.json')

#Restores all json dicts from dropbox
#ARGS: None
#RETURNS: None
def dropbox_restore():
    if input('Type "y" to confirm you want to restore all dicts from Dropbox backup.\n') not in ('y'):
        sys.exit(0)
    company_dict = json_load('/Users/jimmy/Dropbox/Accounting_Backup/company_dict.json')
    project_dict = json_load('/Users/jimmy/Dropbox/Accounting_Backup/project_dict.json')
    rate_dict = json_load('/Users/jimmy/Dropbox/Accounting_Backup/rate_dict.json')
    invoice_number = json_load('/Users/jimmy/Dropbox/Accounting_Backup/invoice_number.json')

    json_dump(company_dict, 'json/company_dict.json')
    json_dump(project_dict, 'json/project_dict.json')
    json_dump(rate_dict, 'json/rate_dict.json')
    json_dump(invoice_number, 'json/invoice_number.json')

    print('\nAll dicts restored from Dropbox Backup')

def print_sheet_info(spreadsheet_id):
    output = ''
    for sheet in spreadsheet_info(spreadsheet_id)['sheets']:
        output += 'Sheet Title: ' + sheet['properties']['title']
        output += ' - Sheet ID: ' + str(sheet['properties']['sheetId'])
    print (output)

# changes spreadsheet title using  google_methods update_spreadsheet method
# ARGS: Spreadsheet ID (str) - New title (str)
# RETURNS: Void - changes spreadsheet title
def title_change(spreadsheet_id, title):
    model_request_dict = {
        'updateSpreadsheetProperties' :
            {
            'properties' : {'title' : 'JM_TEST'},
            'fields' : 'title'
            }
        }
    new_request_dict = model_request_dict
    new_request_dict['updateSpreadsheetProperties']['properties']['title'] = title
    update_spreadsheet(spreadsheet_id, [new_request_dict])

# changes spreadsheet cell values  using  google_methods update_spreadsheet method 
# Example format taken from: https://developers.google.com/sheets/guides/batchupdate
# Detailed reference at: https://developers.google.com/sheets/reference/rest/v4/spreadsheets/request#UpdateCellsRequest
# ARGS: Spreadsheet ID (str) - Change List (lst) - OPTIONAL: Sheet_number (default 1875455808 - STRANGE VALUE... DEBUG?)
#       Change List Content: Dictionaries of changes: [ { 'row' : (int), 'col' : (int), 'value_list' : (2-tuple, see below) }, {ibid...} ] 
#                            Value_list is a 2-tuple with the following content options:
#                            1. ('numberValue', int)
#                            2. ('stringValue', str)
#                            3. ('boolValue', bool)
#                            4. ('formulaValue', str)
#                            5. ('errorValue', { object(ErrorValue) }) #NOT USED

                          
# RETURNS: Void - changes sheet values
# DEBUG: - Can only change one value at a time - extended with value_changes function

def value_change(spreadsheet_id, change_list, sheet_id = 1875455808):
    model_request_dict = {
        'updateCells' :
            {
            'start' : {'sheetId' : 0, 'rowIndex': 0, 'columnIndex' : 0},
            'rows' : [ { 'values' :[ { 'userEnteredValue' : {} } ] } ], 
            'fields' : 'userEnteredValue'
            }
        }
    new_request_list = []

    for change in change_list:
        new_request_dict = model_request_dict
        new_request_dict['updateCells']['start']['sheetId'] = sheet_id
        new_request_dict['updateCells']['start']['rowIndex'] = change['row'] - 1 #Compensate for zero-index
        new_request_dict['updateCells']['start']['columnIndex'] = change['col'] - 1 #Compensate for zero-index
        new_request_dict['updateCells']['rows'][0]['values'][0]['userEnteredValue'][change['value_list'][0]] = change['value_list'][1]
        new_request_list.append(new_request_dict)

    update_spreadsheet(spreadsheet_id, new_request_list)

def value_changes(spreadsheet_id, change_list, sheet_id = 1875455808):
    for change in change_list:
        value_change(spreadsheet_id, [change], sheet_id)

#adds a row to spreadsheet - could be redone with sheet service instance
def insert_row(spreadsheet_id, row_number = 0, sheet_id = 1691011218):
    model_request_dict = {
        'insertDimension' :
            {
            'range' : {'sheetId' : sheet_id, 'dimension': 'ROWS', 'startIndex' : row_number, 'endIndex' : row_number + 1},
            'inheritFromBefore' : True 
            }
        }

    new_request_dict = model_request_dict
#    new_request_dict['updateCells']['start']['sheetId'] = sheet_id
#    new_request_dict['updateCells']['start']['rowIndex'] = change['row'] - 1 #Compensate for zero-index
#    new_request_dict['updateCells']['start']['columnIndex'] = change['col'] - 1 #Compensate for zero-index
#    new_request_dict['updateCells']['rows'][0]['values'][0]['userEnteredValue'][change['value_list'][0]] = change['value_list'][1]
#    new_request_list.append(new_request_dict)

    update_spreadsheet(spreadsheet_id, [new_request_dict])

#Get invoice description from invoice number
#ARGS: invoice number
#RETURNS: invoice description
def get_description(invoice_number):
    last_invoice_row = str(get_final_row())
    invoice_titles = get_range('A71:A' + last_invoice_row)['values']
    
    invoice_id = get_file_id('Trio_JM_Facture_1004')
    last_description = get_range('C15', spreadsheet_id = invoice_id, sheet_title = 'Sheet1')['values'][0][0]
    return(last_description)


#Returns the month in French from a given int
def get_month(mois):
    liste_mois = ['Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre']
    return liste_mois[mois - 1]

#Returns a date string with 0s before day + month if necessary. Defaults to current date.
#ARGS: Datetime object
#RETURNS: String - Format = "2016-09-22".
def get_date_string(date = datetime.date.today()):
    year = str(date.year)
    month = str(date.month)
    if int(month) < 10:
        month = '0' + month
    day = str(date.day)
    if int(day) < 10:
        day = '0' + day
    date_string = year + '-' + month + '-' + day
    return date_string

#Strips all values not in (0,1,2,3,4,5,6,7,8,9,",") from string
#ARGS: string
#RETURNS: string

def strip_dollar_string(dollar_string):
    output = ''
    valid_input = ('1','2','3','4','5','6','7','8','9','0',',')
    for letter in dollar_string:
        if letter in valid_input:
            output += letter
    return output

#Returns a float from dollar string
#ARGS: string
#RETURNS: float
def dollar_to_float(raw_dollar_string):
    dollar_string = strip_dollar_string(raw_dollar_string)
    comma = dollar_string.index(',')
    whole = int(dollar_string[0:comma])
    fraction = int(dollar_string[comma + 1:comma + 3])
    number = whole + (fraction * 0.01)
    #print(number)
    #print(type(number))
    return number



#Find line to insert - TODO: Make more general
#ARGS: NONE
#RETURNS: integer with correct row to insert new row before
def get_final_row():
    range_object = get_range('A1:A')
    return range_object['values'].index(['Total:']) - 1 #for zero-index purposes


def check_invoice_number(to_print = False):
    #general accounting ID - NOTE: Test version - Change for production
    general_accounting = '1GhovgewjuCoqJRdN-PrJvXv9Tq9INeKqgBH_s__xi2M'

    #current invoice number from json dict
    invoice_number = int(json_load('json/invoice_number.json'))

    last_invoice_row = str(get_final_row())

    #last invoice number from general accounting
    worksheet_invoice_number = int(get_range('A' + last_invoice_row)['values'][0][0][-4:])

    if to_print:
        print('google docs last invoice: ' + str(worksheet_invoice_number) + '\n')
        print('json dict last invoice: ' + str(invoice_number) + '\n')

    return invoice_number == worksheet_invoice_number + 1
