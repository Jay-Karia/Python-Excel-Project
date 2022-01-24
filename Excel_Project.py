from openpyxl import load_workbook
import json

# Globals
input_file_name = "Template.xlsx"
output_file_name = "Final.xlsx"

def ReadJSONData():
    with open("CE_Analytics.json") as json_file:
        data = json.load(json_file)
        # Getting the total account numbers
        bank_accounts = data['response']['bank_accounts']
        total_bank_accounts = len(bank_accounts)
        # Update the excel file with json bank_accounts number
        workbook = load_workbook(input_file_name)
        worksheet = workbook['Deposits']
        worksheet['F2'] = total_bank_accounts
        workbook.save(output_file_name)
        # Getting the last 4 digits of account number
        for i in range(0, total_bank_accounts):
            account_number = data['response']['bank_accounts'][i]['account_number']
            last_4 = account_number[len(account_number)-1]+account_number[len(account_number)-2]+account_number[len(account_number)-3]+account_number[len(account_number)-4]

ReadJSONData()
