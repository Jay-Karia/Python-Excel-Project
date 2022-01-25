from openpyxl import load_workbook
import json

# Globals
input_file_name = "Template.xlsx"
output_file_name = "Final.xlsx"

account_numbers_list = []
last_4_digits = []

workbook = load_workbook(input_file_name)
worksheet = workbook['Deposits']

def ReadJSONData():
    global account_numbers_list
    global last_4_digits

    with open("CE_Analytics.json") as json_file:
        data = json.load(json_file)
        # Getting the total account numbers
        bank_accounts = data['response']['bank_accounts']
        # Update the excel file with json bank_accounts number and replacing with F2 block in excel
        total_bank_accounts = len(bank_accounts)
        worksheet['F2'] = total_bank_accounts
        # Getting the last 4 digits of account number
        for i in range(0, total_bank_accounts):
            account_number = data['response']['bank_accounts'][i]['account_number']
            last_4_digits.append(account_number[len(account_number)-4]+account_number[len(account_number)-3]+account_number[len(account_number)-2]+account_number[len(account_number)-1])
            account_numbers_list.append(last_4_digits[i])
            WriteAccountNo(account_numbers_list)

        # Getting the info under block E22
        for i in range(0, total_bank_accounts):
            daily_balance = data['response']['bank_accounts'][i]['daily_balances']
            beginning_date = list(daily_balance.keys())[0]
            beginning_balance = list(daily_balance.values())[0]
            end_date = list(daily_balance.keys())[len(daily_balance)-1]
            end_balance = list(daily_balance.values())[len(daily_balance)-1]
        
        


def WriteAccountNo(account_numbers):
    if len(account_numbers) == 1:
        worksheet['G5'] = account_numbers[0]
    elif len(account_numbers) == 2:
        worksheet['G28'] = account_numbers[1]
    elif len(account_numbers) == 3:
        worksheet['G51'] = account_numbers[2]
    elif len(account_numbers) == 4:
        worksheet['G74'] = account_numbers[3]

ReadJSONData()
workbook.save(output_file_name)
