import sys
from utilities import json_load, json_dump
from google_methods import get_file_id, get_range

debug = False

#GENERAL NOTE: All company, project and rate objects are re-created from the .json everytime they are needed. 
#              This ensures that editing can be done directly in the .json if necessary.
#              Company info is stored in company_dict.json.
#              Project info is stored in project_dict.json.
#              Rate info is stored in rate_dict.json.
#              Current invoice number is stored in invoice_number.json

class Company:
    def __init__(self, company_name, address = '', city = '', post_code = '', invoice_name = ''):
        company_dict = json_load('json/company_dict.json') #Loads the latest general JSON file.
        if company_name in company_dict:
            self.dict = company_dict[company_name]
        else:
            self.dict = {}
            self.dict['projects'] = [] # List of all projects. Added to when creating a new Project object.
            self.dict['company_name'] = company_name
            self.dict['address'] = address
            self.dict['city'] = city
            self.dict['post_code'] = post_code
            self.dict['invoice_name'] = invoice_name
            self.dict['last_invoice'] = -1 #Latest invoice number. Starts at -1 on creation (never invoiced with this company).
            company_dict[company_name] = self.dict
            json_dump(company_dict, 'json/company_dict.json') #Updates the company JSON file.

    #adds a new project to a company
    def add_project(self, project_name, invoice_name, invoice_code):
        new_project = Project(self.dict['company_name'], project_name, invoice_code, invoice_name)
        company_dict = json_load('json/company_dict.json')#Loads the latest company JSON file.
        company_dict[self.dict['company_name']]['projects'].append(project_name)
        json_dump(company_dict, 'json/company_dict.json') #Updates the company JSON file.
        return new_project
    
    def update_invoice_number(self, invoice_number):
        company_dict = json_load('json/company_dict.json')#Loads the latest company JSON file.
        company_dict[self.dict['company_name']]['last_invoice'] = invoice_number
        json_dump(company_dict, 'json/company_dict.json') #Updates the company JSON file.
        

class Project:
    def __init__(self, company_name, project_name, invoice_code, invoice_name = ''):
        company_dict = json_load('json/company_dict.json')
        project_dict = json_load('json/project_dict.json')
        if project_name in company_dict[company_name]['projects']:
            self.dict = project_dict[project_name]
        else:
            self.dict = {}
            self.dict['company_name'] = company_name
            self.dict['project_name'] = project_name
            self.dict['invoice_code'] = invoice_code
            if invoice_name:
                self.dict['invoice_name'] = invoice_name
            else:
                self.dict['invoice_name'] = company_name
            self.dict['last_invoice'] = -1 #Latest invoice number. Starts at -1 on creation (never invoiced with this project).
            self.dict['rates'] = [] # List of all rates. Added to when creating a new Rate object.
            project_dict[project_name] = self.dict
            print(self.dict)
            json_dump(project_dict, 'json/project_dict.json') #Updates the project JSON file.
            
    #adds a new rate to a project
    def add_rate(self,  dollar_rate, float_bool):
        new_rate = Rate(self.dict['project_name'],dollar_rate, float_bool)
        project_dict = json_load('json/project_dict.json')#Loads the latest project JSON file.
        project_dict[self.dict['project_name']]['rates'].append(dollar_rate)
        json_dump(project_dict, 'json/project_dict.json') #Updates the project JSON file.
        return new_rate
    
    #removes last rate added
    def remove_last_rate(self):
        project_dict = json_load('json/project_dict.json')#Loads the latest project JSON file.
        project_dict[self.dict['project_name']]['rates'].pop()
        json_dump(project_dict, 'json/project_dict.json') #Updates the project JSON file.

    def update_invoice_number(self, invoice_number):
        project_dict = json_load('json/project_dict.json')#Loads the latest project JSON file.
        project_dict[self.dict['project_name']]['last_invoice'] = invoice_number
        json_dump(project_dict, 'json/project_dict.json') #Updates the project JSON file.

class Rate:
    def __init__(self, project_name, dollar_rate, float_bool):
        project_dict = json_load('json/project_dict.json')
        rate_dict = json_load('json/rate_dict.json')
        if dollar_rate in project_dict[project_name]['rates']:
            self.dict = rate_dict[project_name + '_' + dollar_rate]
        else:
            self.dict = {}
            self.dict['rate_name'] = dollar_rate #identical to "taux".
            if float_bool:
                self.dict['unit_cost'] = float(dollar_rate)
            else:
                self.dict['unit_cost'] = int(dollar_rate)
            self.dict['project_name'] = project_name
            def get_profession():
                def get_input():
                    number = input('Tapez 1 pour du montage ou 2 pour de la traduction.\n')
                    if number not in ('1','2'):
                        print('Entrez invalide, veuillez recommencer')
                        return get_input()
                    else:
                        return number
                answer = get_input()
                if answer == '1':
                    return 'MONT'
                else:
                    return 'TRAD'
            self.dict['profession'] = get_profession()
            self.dict['last_invoice'] = -1 #Latest invoice number. Starts at -1 on creation (never invoiced with this project).
            self.dict['début'] = 'Total: '
            self.dict['type'] = input('Entrez le type d\'unité. Ex: Forfait, journée, mot...\n').lower()
            self.dict['milieu'] = ' à '
            self.dict['taux'] = str(self.dict['unit_cost']) + '$'
            self.dict['fin'] = ' ' + input('Tapez la fin désirée. Ex: par animatique de 27 minutes.\n')
            rate_dict[project_name + '_' + dollar_rate] = self.dict
            json_dump(rate_dict, 'json/rate_dict.json') #Updates the rate JSON file.

    #Returns rate in a complete sentence.
    def rate_sentence(self, units):
        if int(units) == units:
            units = int(units)
        sentence = ''
        sentence += self.dict['début']
        sentence += str(units) + ' '
        sentence += self.dict['type']
        if units > 1:
            sentence += 's'
        sentence += self.dict['milieu']
        sentence += self.dict['taux']
        sentence += self.dict['fin']
        return sentence

    def update_invoice_number(self, invoice_number):
        rate_dict = json_load('json/rate_dict.json')#Loads the latest rate JSON file.
        project_name = self.dict['project_name']
        dollar_rate = self.dict['rate_name']
        rate_dict[project_name + '_' + dollar_rate]['last_invoice'] = invoice_number
        json_dump(rate_dict, 'json/rate_dict.json') #Updates the project JSON file.

    #Returns name of latest invoice with this rate
    def get_last_invoice_name(self):
        company_dict = json_load('json/company_dict.json')
        project_dict = json_load('json/project_dict.json')
        project_name = self.dict['project_name']
        company_name = project_dict[project_name]['company_name']
        invoice_name = company_dict[company_name]['invoice_name']
        last_invoice_name = invoice_name + '_JM_Facture_' + str(self.dict['last_invoice'])
        return(last_invoice_name)

    #Returns description for latest invoice with this rate
    def get_last_description(self):
        invoice_name = self.get_last_invoice_name()
        invoice_id = get_file_id(invoice_name)
        if invoice_id is None:
            return("No description available")
        last_description = get_range('C15', spreadsheet_id = invoice_id, sheet_title = 'Sheet1')['values'][0][0]
        return(last_description)
                                                               
        
"""
# FOLLOWING PROBABLY NOT  NECESSARY - forgot what this was for - commented out for now

company_dict = json_load('json/company_dict.json')

description_dict = {}

unit_dict = {}

invoice_dict = {
    'invoice_number' : 1000,
    'company_dict' : company_dict,
    'description_dict' : description_dict,
    'unit_dict' : unit_dict
}
"""
