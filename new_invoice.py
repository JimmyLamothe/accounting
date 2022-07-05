#!/usr/bin/env python3

import sys
import datetime
from google_methods import copy_file, download_pdf, update_spreadsheet, get_range, set_range
from utilities import title_change, value_changes, json_load, json_dump, get_month, insert_row, get_date_string
from utilities import dollar_to_float, dropbox_backup, get_final_row, check_invoice_number
from utilities import backup_dicts
import invoice_info

"""
Creates new invoice in Google Docs and adds it to accounting spreadsheet.
Outputs invoice as PDF in Documents/Factures..

Required Arguments: None.
Optional Arguments: Use paramater 'test' to keep dicts unchanged in Dropbox while testing program.
                    Run restore_dicts.py to restore them from Dropbox after testing.
                    Use parameter 'local' to not modify Google Sheets documents or create new invoices.

All  necessary information for invoicing is input at runtime.
"""

local = False  #Change to False to complete entire process including Google Docs changes, to True to stop before. Also bypasses Dropbox backup.

dropbox = True #Change to False by adding 'test' as an argument to keep dicts intact in Dropbox.

for index in [1,2]:
    try:
        if sys.argv[index].lower() in ['test']:
            dropbox = False
            print('\n' + 'Running in test mode, dicts won\'t be saved to Dropbox.' + '\n')
        elif sys.argv[index].lower() in ['local']:
            local = True
            print('\n' + 'Running in local mode, no changes will be made to Google Sheets documents' + '\n')
    except (IndexError):
        pass

#check current invoice number matches Google Docs general accounting document
if not check_invoice_number():
    print('current json invoice_number doesn\'t match Google Docs document. Fix before continuing.')
    sys.exit(1)

#model invoice ID
model_invoice = '1REZ_1wDyAhh7W6EFqPR65XVeeeB0XSEs0s7DNpiwe0g'

#general accounting ID - NOTE: Test version - Change for production
general_accounting = '1GhovgewjuCoqJRdN-PrJvXv9Tq9INeKqgBH_s__xi2M'

#get info for new invoice

#get company name
#ARGS: None
#RETURNS: Company object
def get_company():
    order_list = {} #temporarily associate company name with printed number to access dictionary by number.
    count = 1
    company_dict = json_load('json/company_dict.json')
    for company in sorted(company_dict, key=lambda company:company_dict[company]['last_invoice'], reverse = True): #Sort by latest invoice number
        print(str(count) + '. ' + company) #DEBUG TRY -company_dict[company]['company_name'])
        order_list[str(count)] = company 
        count += 1
    print('\r')
    company_input = input('Choisissez une compagnie de la liste ou tapez 0 pour créer une nouvelle compagnie.\n')

    if company_input is "0": # Create new company entry in company_dict.json and return the object
        name = input('Entrez le nom de la compagnie.\n')
        invoice_name = input('Entrez une abbréviation pour le nom de facture ou laissez blanc pour mettre le nom de compagnie. \n')
        if not invoice_name:
            invoice_name = name
        address = input('Entrez le numéro et la rue.\n')
        city = input('Entrez la ville et la province.\n')
        post_code = input('Entrez le code postal.\n')
        return invoice_info.Company(name, address, city, post_code, invoice_name)
    else:
        try: # Create company object from existing company info in company_dict.json and return it
            int(company_input)
            return invoice_info.Company(order_list[company_input]) 
        except (ValueError, KeyError):
            print('Entrée invalide, recommencez SVP.\n')
            return get_company() #Recursively calls the function again until proper input is received


active_company  = get_company()
#print(active_company.dict['projects'])

#get project name
#ARGS: None
#RETURNS: Project object
def get_project():
    print('\r')
    order_list = {} #temporarily associate project name with printed number
    count = 1
    project_dict = json_load('json/project_dict.json')
    for project in sorted(project_dict, key = lambda project:project_dict[project]['last_invoice'], reverse = True): #Sort by latest invoice number
        if project_dict[project]['company_name'] == active_company.dict['company_name']:
            print(str(count) + '. ' + project)
            order_list[str(count)] = project
            count += 1
    print('\r')
    project_input = input('Choisissez un projet de la liste ou tapez 0 pour créer un nouveau projet.\n')

    if project_input is "0": # Create new project entry in project_dict.json and return the object
        name = input('Entrez le nom du projet\n')
        if name in project_dict.keys():
            print('Name already exists, please pick another')
            return get_project()
        code = input('Entrez un code de facturation pour la comptabilité générale. Ex: AVRP - CHUCK\n')
        invoice = input('Entrez le nom de facturation ou laissez blanc pour mettre le nom de compagnie. \n')
        return active_company.add_project(name, invoice, code)
    else: 
        try: # Create existing project object from project_dict.json and return it
            int(project_input)
            return invoice_info.Project(active_company.dict['company_name'], order_list[project_input], invoice_code = '')
        except (ValueError, KeyError):
            print('Entrée invalide, recommencez SVP.\n')
            return get_project() #Recursively call function until proper input is received

active_project = get_project()



#get rate for invoice
#ARGS: None
#RETURNS: Rate object
def get_rate(project):
    print('\r')
    order_list = {} #temporarily associate rate name with printed number
    count = 1
    rate_dict = json_load('json/rate_dict.json')
    for rate in sorted(rate_dict, key = lambda rate:rate_dict[rate]['last_invoice'], reverse = True): #Sort by latest invoice number
        if rate_dict[rate]['project_name'] == active_project.dict['project_name']:
            print(str(count) + '. ' + rate_dict[rate]['taux'])
            order_list[str(count)] = rate_dict[rate]['taux'][:-1]
            count += 1
    print('\r')
    #print(order_list)
    rate_input = input('Choisissez un taux de la liste ou tapez 0 pour créer un nouveau taux.\n')

    if rate_input is "0": # Create new rate entry in rate_dict.json and return the object
        dollar_rate = input('Entrez le taux de facturation en dollars.\n')
        float_bool = False
        try:
            int(dollar_rate)
        except (ValueError):
            try:
                dollar_rate = dollar_rate.replace(',','.')
                float(dollar_rate)
                float_bool = True
                dollar_rate = str(float(dollar_rate))
            except (ValueError):
                print('Entrée invalide. Veuillez recommencer.')
                return get_rate(project)
        new_rate = active_project.add_rate(dollar_rate, float_bool)
        print('Voici la description de ce taux:\n')
        print(new_rate.rate_sentence(1))
        def get_answer(): #Temporary function to allow recursion if invalid answer.
            answer = input('Tapez Y si vous êtes satisfaits, ou N pour recommencer.\n') 
            if answer.lower() not in ('y', 'n'):
                print('Entrée invalide. Veuillez recommencer.')
                return get_answer()
            else:
                return answer 
        if get_answer().lower() == 'y':
            return new_rate
        else:
            rate_dict = json_load('json/rate_dict.json')
            rate_key_name = active_project.dict['project_name'] + '_' + dollar_rate
            rate_dict.pop(rate_key_name, None) #Remove unwanted new rate from rate_dict.
            json_dump(rate_dict, 'json/rate_dict.json') #Updates the rate JSON file.
            project.remove_last_rate()
            return get_rate(project)

    else: 
        try: # Create existing rate object from project_dict.json and return it
            try:
                int(rate_input)
                #print('int_' + order_list[rate_input])
                return invoice_info.Rate(active_project.dict['project_name'], order_list[rate_input], False) 
            except(ValueError):
                #print('check_value')
                float(rate_input)
                #print('float_' + order_list[rate_input])
                return invoice_info.Rate(active_project.dict['project_name'], order_list[rate_input], True) 
        except (ValueError, KeyError):
            #print('check_key')
            print('Entrée invalide, recommencez SVP.\n')
            return get_rate(project)

active_rate = get_rate(active_project) #called with active project to allow reanswering.

#get invoice description
#ARGS: None
#RETURNS: String - Description incuding episode and series name.
def get_description():
    print('\r')
    print("Voici la description pour la dernière facture facturée avec ce taux:")
    print(active_rate.get_last_description())
    print('\r')
    description = input('Entrez la description de la facture, incluant le nom de l\'épisode et de la série.\n')
    if description:
        return description
    else:
        print('Entrée invalide, recommencez SVP.\n')
        return get_description()

active_description = get_description()


#get month for invoice description - defaults to current month.
#ARGS: None
#RETURNS: String - Date work was executed with month and year specified.
def get_description_date():
    print('\r')
    description_date = input('Entrez la description du mois où le travail a été effectué, ou laissez vide pour le mois courant.\n')
    if description_date:
        return description_date
    else:
        description_date = get_month(datetime.date.today().month) + ' ' + str(datetime.date.today().year)
        return description_date
    
active_description_date = get_description_date()



invoice_number = int(json_load('json/invoice_number.json'))
title = active_company.dict['invoice_name'] + '_JM_Facture_' + str(invoice_number)
date_string = get_date_string() #Defaults to current day - add datetime object as parameter to change
invo_name = active_project.dict['invoice_name']
invo_code = active_project.dict['invoice_code']
comp_addr = active_company.dict['address']
comp_city = active_company.dict['city']
comp_post_code = active_company.dict['post_code']
unit_quan = float(input('Combien de ' + active_rate.dict['type'] + '(s) voulez-vous facturer?'))
desc_cost = active_rate.rate_sentence(unit_quan)
unit_cost = active_rate.dict['unit_cost']
prof_type = active_rate.dict['profession']
if prof_type == 'TRAD':
    translation_reminder = True

"""
print(type(invoice_number))
print(invoice_number)
print(type(title))
print(title)
print(type(date_string))
print(date_string)
print(type(invo_name))
print(invo_name)
print(type(comp_addr))
print(comp_addr)
print(type(comp_city))
print(comp_city)
print(type(comp_post_code))
print(comp_post_code)
print(type(desc_cost))
print(desc_cost)
print(type(unit_quan))
print(unit_quan)
print(type(unit_cost))
print(unit_cost)
"""

#Exit before networking processes if working on program
if local:
    sys.exit(0)

#create new invoice from model invoice

new_id = copy_file(model_invoice)['id']

#update new invoice with new title

title_change(new_id, title)



#set date parameters - TODO
#current_month = datetime.date.today().month

#update new invoice info - TODO
"""
Reference values (row.column):

Invoice Number: 2.5 - int
Date: 8.5 - str(date_object)

Company Name: 8.3 - str
Company Address: 9.3 - str
Company City: 10.3 - str
Company Postal Code: 11.3 - str

Description: 15.3 - str
Description Date: 16.3 - str
Description Cost: 38.3 - str

Unit Quantity: 38.2 - int
Unit Cost: 38.4 - int(str[0:4])

Change List Content: Dictionaries of changes: [ { 'row' : (int), 'col' : (int), 'value_list' : (2-tuple, see below) }, {ibid...} ] 
                     Value_list is a 2-tuple with the following content options:
                     1. ('numberValue', int)
                     2. ('stringValue', str)
                     3. ('boolValue', bool)
                     4. ('formulaValue', str)
                     5. ('errorValue', { object(ErrorValue) }) #NOT USED

"""
#Update invoice values
change_dict_list = []
change_dict_list.append( { 'row' : (2), 'col' : (5), 'value_list' : ('numberValue', invoice_number) } )
change_dict_list.append( { 'row' : (38), 'col' : (2), 'value_list' : ('numberValue', unit_quan) } )
change_dict_list.append( { 'row' : (38), 'col' : (4), 'value_list' : ('numberValue', unit_cost) } )
change_dict_list.append( { 'row' : (8), 'col' : (5), 'value_list' : ('stringValue', date_string) } )
change_dict_list.append( { 'row' : (8), 'col' : (3), 'value_list' : ('stringValue', invo_name) } )
if prof_type == 'TRAD':
    change_dict_list.append( { 'row' : (3), 'col' : (3), 'value_list' : ('stringValue', 'Traducteur') } )
change_dict_list.append( { 'row' : (9), 'col' : (3), 'value_list' : ('stringValue', comp_addr) } )
change_dict_list.append( { 'row' : (10), 'col' : (3), 'value_list' : ('stringValue', comp_city) } )
change_dict_list.append( { 'row' : (11), 'col' : (3), 'value_list' : ('stringValue', comp_post_code) } )
change_dict_list.append( { 'row' : (15), 'col' : (3), 'value_list' : ('stringValue', active_description) } )
change_dict_list.append( { 'row' : (16), 'col' : (3), 'value_list' : ('stringValue', active_description_date) } )
change_dict_list.append( { 'row' : (38), 'col' : (3), 'value_list' : ('stringValue', active_rate.rate_sentence(unit_quan)) } )

value_changes(new_id, change_dict_list)

#Update general accounting spreadsheet with new invoice info

new_row = get_final_row()

insert_row(general_accounting, new_row)

new_range = 'A' + str(new_row + 1) + ':I' + str(new_row + 1)

base_salary = '=IMPORTRANGE("' + new_id + '";"E41")'
tax_federal = '=IMPORTRANGE("' + new_id + '";"E42")'
tax_provincial = '=IMPORTRANGE("' + new_id + '";"E43")'
total_salary = '=IMPORTRANGE("' + new_id + '";"E44")'

new_values = [[title, date_string, base_salary, tax_federal, tax_provincial, total_salary, '=FAUX', prof_type, invo_code]]

set_range(new_range, new_values)

#download PDF

try:
    download_pdf(new_id, '/Users/jimmy/Documents/Factures/' + title)
except Exception:
    print('Failed to download invoice')
try:
    download_pdf(general_accounting, '/Users/jimmy/Documents/Bilans/' + str(invoice_number))
except Exception:
    print('Failed to download financial data')


#increment invoice number ONLY AT END WHEN EVERYTHING HAS WORKED
active_company.update_invoice_number(invoice_number)
active_project.update_invoice_number(invoice_number)
active_rate.update_invoice_number(invoice_number)
json_dump(str(invoice_number + 1), 'json/invoice_number.json')

#Backup all dicts to Dropbox if not in test mode
if dropbox:
    dropbox_backup()
    print('\nAll dicts backed up to Dropbox.')
else:
    print('\nTest mode enabled, dicts not saved to Dropbox.')

#Backup all dicts to directory with date and time - Keep all versions for now
backup_dicts()
