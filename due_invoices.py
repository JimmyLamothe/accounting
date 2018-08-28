#!/usr/bin/env python3

import sys
from itertools import combinations
from oauth2client.service_account import ServiceAccountCredentials
from google_methods import update_spreadsheet, get_range, set_range, get_file_id
from utilities import value_changes, get_final_row, json_load, json_dump, dollar_to_float

"""
Lists unpaid invoices by company.

Required Arguments: None.
Optional Arguments: "add_wip" to add an item to work in progress.
                    "list_wip" to see and-or remove work in progress 
"""
#company_dict = json_load('json/company_dict.json')
#for company in company_dict:
    #print (company_dict[company]['invoice_name'])

work_in_progress = json_load('json/work_in_progress.json')

try:
    wip_input = sys.argv[1]
    if wip_input == 'add_wip':
        new_wip = input('Type description of project to add to work in progress:\n')
        if new_wip:
            work_in_progress.append(new_wip)
        print('Current work in progress:\n')
        for index, item in enumerate(work_in_progress):
            print(str(index + 1) + '. ' + item)
        json_dump(work_in_progress, 'json/work_in_progress.json')
        sys.exit(0)
    elif wip_input == 'list_wip':
        print('Current work in progress:\n')
        for index, item in enumerate(work_in_progress):
            print(str(index + 1) + '. ' + item)
        index = input('Do you want to invoice an item now?' + '\n' +
                  'If so, please type the corresponding number to delete from list' +
                  ' and run new_invoice.py\n' +
                  'Use argument add_wip to add work in progress\n')
        if index:
            try:
                del work_in_progress[int(index) - 1]
                json_dump(work_in_progress, 'json/work_in_progress.json')
            except IndexError:
                print('Incorrect input. Rerun program to try again.')
        json_dump(work_in_progress, 'json/work_in_progress.json')
        sys.exit(0)
except IndexError:
    pass

print(type(work_in_progress))

#general accounting ID - NOTE: Test version - Change for production
general_accounting = '1GhovgewjuCoqJRdN-PrJvXv9Tq9INeKqgBH_s__xi2M'

last_invoice_row = str(get_final_row())

paid_bools = get_range('G4:G' + last_invoice_row)['values'] #List of true or false strings

bool_dict = { 'TRUE' : True, 'FALSE' : False, 'VRAI' : True, 'FAUX' : False } #Convert strings to booleans

value_dict = {}

#value_dict = json_load('json/test_check.json') #Uncomment to use test dictionary.


for index, bool in enumerate(paid_bools): #Get unpaid invoice info from Google Sheets
    if not bool_dict[bool[0]]:
        row = str(index + 4)
        invoice_name = get_range('A' + row)['values'][0][0] #invoice name
        underline_index = invoice_name.index('_')
        company_name = invoice_name[:underline_index] #company name
        invoice_date = get_range('B' + row)['values'][0][0] #invoice date
        amount_raw = get_range('F' + row)['values'][0][0]
        amount = dollar_to_float(amount_raw) #dollar amount of invoice
        invoice_id = get_file_id(invoice_name) 
        invoice_description = get_range('C15', spreadsheet_id = invoice_id, sheet_title = 'Sheet1')['values'][0][0] #invoice description
        if company_name not in value_dict:  
            value_dict[company_name] = [(invoice_name, invoice_date, amount, invoice_description)]
        else:
            value_dict[company_name].append((invoice_name, invoice_date, amount, invoice_description))




def parse_invoice(list, running_total = 0):
    #RETURNS: Running total
    #PRINTS: Parsed invoice
    parsed = 'Nom: '
    parsed += list[0] #invoice name
    parsed += '\nDate: '
    parsed += list[1] #invoice date
    parsed += '\nMontant: '
    parsed += str(list[2]) + '$'#invoice amount
    running_total += list[2]
    parsed += '\nDescription: '
    parsed += list[3] #invoice description
    print(parsed, '\n')    
    return(running_total)

total_amount_due = 0

for company in value_dict:
    running_total = 0
    print('\n' + 'Factures non-payées pour ' + company + ':\n')
    invoice_list = [item for item in value_dict[company]]
    for invoice in invoice_list:
        running_total = parse_invoice(invoice, running_total)
    print('Total amount due: ' + str(running_total) + '$\n')
    total_amount_due += running_total
    
print('Total amount due from  all companies: ' + str(int(total_amount_due)) + '$\n')

if work_in_progress:
    print('The following work in progress has not been invoiced yet.')
    for index, item in enumerate(work_in_progress):
        print(str(index + 1) + '. ' + item)
    index = input('Do you want to invoice an item now?' + '\n' +
                  'If so, please type the corresponding number to delete from list' +
                  ' and run new_invoice.py\n' +
                  'Use argument add_wip to add work in progress\n')
    if index:
        try:
            del work_in_progress[int(index) - 1]
            json_dump(work_in_progress, 'json/work_in_progress.json')
        except IndexError:
            print('Incorrect input. Rerun program to try again.')
    








#json_dump(value_dict, 'json/test_check.json') #Uncomment to create test dictionary.

