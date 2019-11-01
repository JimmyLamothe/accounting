#!/usr/bin/env python3 

import sys
from itertools import combinations
from google_methods import update_spreadsheet, get_range, set_range, get_file_id
from utilities import value_changes, get_final_row, json_load, json_dump, dollar_to_float

"""
Used when check received. Checks in general accounting Google Sheets document to see which
invoices check was for and marks them as paid. 

Required Arguments: None.
Optional Arguments: Use parameter 'cancel' to restore values modified by previous execution.

All other necessary information is input at runtime.
"""

cancel = False

try:
    if sys.argv[1] == 'cancel':
        last_change_list = json_load('json/last_change_list.json')
        for row in last_change_list:
            set_range('G' + row, [['=FALSE']])
        sys.exit(0)
except IndexError:
    pass

company_dict = json_load('json/company_dict.json')
for company in company_dict:
    print (company_dict[company]['invoice_name'])

#general accounting ID - NOTE: Test version - Change for production
general_accounting = '1GhovgewjuCoqJRdN-PrJvXv9Tq9INeKqgBH_s__xi2M'

last_invoice_row = str(get_final_row())

paid_bools = get_range('G4:G' + last_invoice_row)['values']

bool_dict = { 'TRUE' : True, 'FALSE' : False, 'VRAI' : True, 'FAUX' : False }

value_dict = {}

#Prints description of invoice in specific row.
def print_description(row):
    invoice_name = get_range('A' + row)['values'][0][0]
    print(invoice_name)
    invoice_id = get_file_id(invoice_name)
    if invoice_id is None:
        print("No description available - CHECK MANUALLY! POSSIBLE BUG!")
    last_description = get_range('C15', spreadsheet_id = invoice_id, sheet_title = 'Sheet1')['values'][0][0]
    print(last_description)

for index, bool in enumerate(paid_bools):
    if not bool_dict[bool[0]]:
        row = str(index + 4)
        #company = get_range('A' + row)['values'][0][0] #TESTING PROJECT INSTEAD OF COMPANY
        #underline_index = company.index('_')
        #company_name = company[:underline_index]
        project = get_range('I' + row)['values'][0][0]
        amount_raw = get_range('F' + row)['values'][0][0]
        amount = dollar_to_float(amount_raw)
        """
        if company_name not in value_dict:
            value_dict[company_name] = [(amount, row)]
        else:
            value_dict[company_name].append((amount, row))
        """
        if project not in value_dict:
            value_dict[project] = [(amount, row)]
        else:
            value_dict[project].append((amount, row))
        #print(row, company_name, amount)
        print(row, project, amount)
        #print_description(row)

print (value_dict)
# json_dump(value_dict, 'json/test_check.json') #Uncomment to create test dictionary.

check_string = input('Veuillez entrer le montant du chèque.\n')

check_float = float(check_string)

# value_dict = json_load('json/test_check.json') #Uncomment to use test dictionary.


def get_combinations(value_dictionary):
    possibilities = []
    #for company in value_dict:
    for project in value_dict:
        #print(company)
        print(project)
        #length = len(value_dict[company])
        length = len(value_dict[project])
        enumerate_length = ''
        for i in range (length):
            enumerate_length += str(i) #Create string in format '01234...' to use as index.
        for i in range (length):
            combinator = combinations(enumerate_length, length - i)
            for combination in combinator:
                total = 0
                for index in combination:
                    #total += value_dict[company][int(index)][0]
                    total += value_dict[project][int(index)][0]
                if round(total, 2) == check_float:
                    print(round(total,2))
                    print(combination)
                    combination_list = []
                    for index in combination:
                        #combination_list.append(value_dict[company][int(index)][1]) #append row
                        combination_list.append(value_dict[project][int(index)][1]) #append row
                    possibilities.append(combination_list)
    return possibilities


possibilities = get_combinations(value_dict)
print("possibilities: " + str(possibilities) + '\n')

last_change_list = []

def pay_invoice(row):
    set_range('G' + row, [['=TRUE']])
    print('Paid invoice: ' + get_range('A' + row)['values'][0][0] + '\n')
    print('Value: ' + get_range('F' + row)['values'][0][0] + '\n')

def manual_payment(row):    
    print_description(row)
    confirm = input("Press Y to pay this invoice, anything else to exit")
    if str.upper(confirm) == "Y":
        pay_invoice(row)
        
if len(possibilities) == 0:
    print('No possible combinations with these values. Please check input value or pay manually.\n')
    payable_row = input("To mark an invoice as paid, please enter the corresponding row. ENTER to exit")
    if payable_row:
        manual_payment(payable_row)

elif len(possibilities) == 1:
    for row in possibilities[0]:
        last_change_list.append(row)
        pay_invoice(row)
else:
    print('Different combinations are possible. Please change values manually in Google Sheets or select rows to pay.\n')
    print('The following row combinations are valid:\n')
    for possibility in enumerate(possibilities, start = 1):
        print("OPTION " + str(possibility[0]))
        for row in possibility[1]:
            print("Row :" + str(row))
            print_description(row)
            print('\n')
    selection = input("Please select an option or RETURN to exit.\n")
    if selection:
        rows = possibilities[int(selection) - 1]
        for row in rows:
            pay_invoice(row)
            

json_dump(last_change_list, 'json/last_change_list.json')

input("Don't forget to deposit check if necessary. Press RETURN to confirm.")
