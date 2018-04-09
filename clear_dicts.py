from utilities import json_dump

company_dict = {}
project_dict = {}
rate_dict = {}
invoice_number = "1000" #dict is actually a string - DOESN'T MAKE SENSE

json_dump(company_dict, 'json/company_dict.json')
json_dump(project_dict, 'json/project_dict.json')
json_dump(rate_dict, 'json/rate_dict.json')
json_dump(invoice_number, 'json_invoice_dict.json')

answer = input('Delete dropbox dicts? Type y to delete')
if answer == 'y':
    json_dump(company_dict, '/Users/jimmy/Dropbox/Accounting_Backup/company_dict.json')
    json_dump(project_dict, '/Users/jimmy/Dropbox/Accounting_Backup/project_dict.json')
    json_dump(rate_dict, '/Users/jimmy/Dropbox/Accounting_Backup/rate_dict.json')
    json_dump(invoice_number, '/Users/jimmy/Dropbox/Accounting_Backup/invoice_dict.json')
    print('Local and Dropbox dicts deleted')
else:
    print('Local dicts deleted')
