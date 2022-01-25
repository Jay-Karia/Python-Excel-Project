from openpyxl import load_workbook
import json

# Globals
input_file_name = "Template.xlsx"
output_file_name = "Final.xlsx"

account_numbers_list = []
last_4_digits = []

beginning_date = []
beginning_balance = []
end_date = []
end_balance = []

workbook = load_workbook(input_file_name)
worksheet = workbook['Deposits']

def ReadJSONData():
    global account_numbers_list
    global last_4_digits
    global beginning_date
    global beginning_balance
    global end_date
    global end_balance

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
            beginning_date.insert(i, list(daily_balance.keys())[0])
            beginning_balance.insert(i, list(daily_balance.values())[0])
            end_date.insert(i, list(daily_balance.keys())[len(daily_balance)-1])
            end_balance.insert(i, list(daily_balance.values())[len(daily_balance)-1])
            WriteBalanceAndDate(beginning_balance, end_balance)

        # Getting other data under block E22
        estimated_revenue_list = []
        for j in range(0, total_bank_accounts):
            sum = 0
            estimated_revenue = data['response']['bank_accounts'][j]['estimated_revenue_by_month']
            temp_values = list(estimated_revenue.values())
            for i in range(0, len(temp_values)):
                temp_values[i] = float(temp_values[i])
                sum+=temp_values[i]
            estimated_revenue_list.append(sum)
            WriteEstimatedRevenue(estimated_revenue_list)
            
        #deposits_month_sum = 0
        #for j in range(0, total_bank_accounts):
        #    deposits_month = data['response']['bank_accounts'][i]['deposits_sum_by_month']
        #    temp_deposits = list(deposits_month.values())
        #    temp_deposits[j] = float(temp_deposits[j])
        #    deposits_month_sum+=temp_deposits[j]
        #deposits_month_sum -= sum

# Other Write Methods
def WriteAccountNo(account_numbers):
    if len(account_numbers) == 1:
        worksheet['G5'] = account_numbers[0]
    elif len(account_numbers) == 2:
        worksheet['G28'] = account_numbers[1]
    elif len(account_numbers) == 3:
        worksheet['G51'] = account_numbers[2]
    elif len(account_numbers) == 4:
        worksheet['G74'] = account_numbers[3]
    else:
        print("Bank Accounts cannot be more than 4")

def WriteBalanceAndDate(b_balance, e_balance):
    if len(b_balance) == 1 or len(e_balance) == 1:
        worksheet['H22'] = b_balance[0]
        worksheet['L22'] = e_balance[0]

    elif len(b_balance) == 2 or e_balance == 2:

        worksheet['H45'] = b_balance[1]
        worksheet['L45'] = e_balance[1]

    elif len(b_balance) == 3 or e_balance == 3:

        worksheet['H68'] = b_balance[2]
        worksheet['L68'] = e_balance[2]

    elif len(b_balance) == 4 or e_balance == 4:

        worksheet['H45'] = b_balance[3]
        worksheet['L45'] = e_balance[3]

def WriteEstimatedRevenue(estimated_revenue):
    if len(estimated_revenue) == 1:
        worksheet['G23'] = estimated_revenue[0]
    elif len(estimated_revenue) == 2:
        worksheet['G46'] = estimated_revenue[1]
    elif len(estimated_revenue) == 3:
        worksheet['G69'] = estimated_revenue[2] 

ReadJSONData()
workbook.save(output_file_name)
