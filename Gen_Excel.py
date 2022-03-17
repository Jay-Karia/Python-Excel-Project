from tarfile import LENGTH_NAME
from requests.auth import HTTPBasicAuth
from openpyxl import load_workbook
from datetime import datetime
from tkinter import *
import requests
import json
import csv
import tkinter as tk
import time

# Globals

account_numbers_list = []
last_4_digits = []

beginning_date = []
beginning_balance = []
end_date = []
end_balance = []

input_file_name = "Template.xlsx"
output_file_name = "Report"
input_json_file = "APIResponse.json"

def ReadAndWrite(bookpk):
    global input_file_name
    global input_json_file
    global output_file_name
    

    workbook = load_workbook(input_file_name)
    worksheet = workbook['Deposits']


    def ReadJSONData():
        global account_numbers_list
        global last_4_digits
        global beginning_date
        global beginning_balance
        global end_date
        global end_balance
        global input_file_name
        global input_json_file
        global output_file_name

        if not input_json_file.endswith('.json'):
            input_json_file = f"{input_json_file}.json"
        with open(input_json_file) as json_file:
            data = json.load(json_file)
            
            # Getting the total bank account numbers
        bank_accounts = data['response']['bank_accounts']
        
        # Update the excel file with json bank_accounts number and replacing with F2 block in excel
        total_bank_accounts = len(bank_accounts)
        if total_bank_accounts > 4:
            total_bank_accounts = 4
        worksheet['F2'] = total_bank_accounts
        
        # Getting the last 4 digits of account number
        for i in range(0, total_bank_accounts):
            account_number = data['response']['bank_accounts'][i]['account_number']
            if len(account_number) > 4:
                last_4_digits.append(account_number[len(account_number)-4]+account_number[len(
                    account_number)-3]+account_number[len(account_number)-2]+account_number[len(account_number)-1])
                account_numbers_list.append(last_4_digits[i])
            elif len(account_number) <= 4:
                account_numbers_list.append(account_number)
            WriteAccountNo(account_numbers_list)

        # Getting the info under block E22
        for i in range(0, total_bank_accounts):
            daily_balance = data['response']['bank_accounts'][i]['daily_balances']
            beginning_date.insert(i, list(daily_balance.keys())[0])
            beginning_balance.insert(i, list(daily_balance.values())[0])
            end_date.insert(i, list(daily_balance.keys())
                            [len(daily_balance)-1])
            end_balance.insert(i, list(daily_balance.values())[
                            len(daily_balance)-1])
            WriteBalanceAndDate(beginning_balance,
                                end_balance, beginning_date, end_date)

        # Getting other data under block E22
        estimated_revenue_list = []
        for j in range(0, total_bank_accounts):
            sum = 0
            estimated_revenue = data['response']['bank_accounts'][j]['estimated_revenue_by_month']
            temp_values = list(estimated_revenue.values())
            for i in range(0, len(temp_values)):
                temp_values[i] = float(temp_values[i])
                sum += temp_values[i]
            estimated_revenue_list.append(sum)
            WriteEstimatedRevenue(estimated_revenue_list)

        # Begin Date
        for i in range(0, total_bank_accounts):
            begin_dates = []
            deposit_sums = []
            temp_deposits = data['response']['bank_accounts'][i]['estimated_revenue_by_month']
            temp_dates = data['response']['bank_accounts'][i]['estimated_revenue_by_month']

            deposit_sums = list(temp_deposits.values())
            begin_dates = list(temp_dates.keys())

            deposit_sums.reverse()
            begin_dates.reverse()

            WriteRemainingData(deposit_sums, begin_dates, i)

        # Deposits Box (G:N)
        len_amounts = []
        for i in range(0, total_bank_accounts):
            temp_dates_and_amount = {}
            temp_raw = data['response']['bank_accounts'][i]['non_estimated_revenue_txns_list']
            for j in range(0, len(temp_raw)):
                temp_amounts = temp_raw[j]['amount']
                temp_dates = temp_raw[j]['txn_date']
                temp_dates_and_amount[temp_dates] = temp_amounts

            temp_dates_and_amount = list(sorted(temp_dates_and_amount.items(), key = lambda x:datetime.strptime(x[0], '%m/%d/%Y'), reverse=False))

            temp_amount_list = []
            temp_dates_list = []

            for g in range(0, len(temp_dates_and_amount)):
                temp_amount_list.append(temp_dates_and_amount[g][1])
                temp_dates_list.append(temp_dates_and_amount[g][0])

            len_amounts.append(len(temp_amount_list))
            WriteInExcel_non_estimated_revenue_txns_list(temp_amount_list,temp_dates_list, i, len_amounts)

    # Other Write Methods
    def WriteInExcel_non_estimated_revenue_txns_list(sorted_amounts, dates,total_bank_accounts, total_amounts):
        months = []
        index = []
        for i in range(0, len(sorted_amounts)):
            months.append(dates[i][0]+dates[i][1])
            index.append(12-int(months[i]))
            print(f"{sorted_amounts[i]}\t\t{months[i]}\t{index[i]}")

    def WriteAccountNo(account_numbers):
        if len(account_numbers) == 1:
            worksheet['G5'] = account_numbers[0]
        elif len(account_numbers) == 2:
            worksheet['G28'] = account_numbers[1]
        elif len(account_numbers) == 3:
            worksheet['G51'] = account_numbers[2]
        elif len(account_numbers) == 4:
            worksheet['G74'] = account_numbers[3]

    def WriteBalanceAndDate(b_balance, e_balance, b_date, e_date):
        for i in range(len(b_balance)):
            b_balance[i] = float(b_balance[i])
            e_balance[i] = float(e_balance[i])
        if len(b_balance) == 1 or len(e_balance) == 1:
            worksheet['I22'] = b_balance[0]
            worksheet['O22'] = e_balance[0]
            worksheet['F22'] = b_date[0]
            worksheet['L22'] = e_date[0]

        elif len(b_balance) == 2 or e_balance == 2:

            worksheet['I45'] = b_balance[1]
            worksheet['O45'] = e_balance[1]
            worksheet['F45'] = b_date[1]
            worksheet['L45'] = e_date[1]

        elif len(b_balance) == 3 or e_balance == 3:

            worksheet['I68'] = b_balance[2]
            worksheet['O68'] = e_balance[2]
            worksheet['F68'] = b_date[2]
            worksheet['L68'] = e_date[2]

        elif len(b_balance) == 4 or e_balance == 4:

            worksheet['I91'] = b_balance[3]
            worksheet['O91'] = e_balance[3]
            worksheet['F91'] = b_date[3]
            worksheet['L91'] = e_date[3]

    def WriteEstimatedRevenue(estimated_revenue):
        if len(estimated_revenue) == 1:
            worksheet['I23'] = estimated_revenue[0]
        elif len(estimated_revenue) == 2:
            worksheet['I46'] = estimated_revenue[1]
        elif len(estimated_revenue) == 3:
            worksheet['I69'] = estimated_revenue[2]
        elif len(estimated_revenue) == 4:
            worksheet['I92'] = estimated_revenue[3]

    def WriteRemainingData(deposit_sum, begin_date, total_bank_accounts):
        start_block = [8, 31, 54, 77]
        try:
            if len(deposit_sum)>12:
                del deposit_sum[12:]
                del begin_date[12:]
            for i in range(0, len(begin_date)):
                deposit_sum[i] = float(deposit_sum[i])
                if total_bank_accounts == 0:
                    worksheet[f'E{start_block[0]+i}'] = begin_date[i]
                    worksheet[f'F{start_block[0]+i}'] = deposit_sum[i]
                elif total_bank_accounts == 1:
                    worksheet[f'E{start_block[1]+i}'] = begin_date[i]
                    worksheet[f'F{start_block[1]+i}'] = deposit_sum[i]
                elif total_bank_accounts == 2:
                    worksheet[f'E{start_block[2]+i}'] = begin_date[i]
                    worksheet[f'F{start_block[2]+i}'] = deposit_sum[i]
                elif total_bank_accounts == 3:
                    worksheet[f'E{start_block[3]+i}'] = begin_date[i]
                    worksheet[f'F{start_block[3]+i}'] = deposit_sum[i]
        except:
            pass

    # try:
    ReadJSONData()
    # output_file_name = "Report  for PK " + bookpk  + "_created_" + time.strftime("%Y%b%d-%H%M%S") + ".xlsx"
    output_file_name = "Final.xlsx"
    workbook.save(output_file_name)
    print(f"Excel File: \"{output_file_name}\", Successfully created!")    

# Graphical User Interface

def Call_Request_API(bookpk):
    try:
        # Calling API
        with open(r'C:\\Ocrolus_Input\\Credential.txt', newline='') as csvfile:
            reader = csv.reader(csvfile, delimiter=' ', quotechar='|')
            for row in reader:
                if row[0][0:4] == 'key:':
                    uname = row[0][4:]
                if row[0][0:7] == 'secret:':
                    apikey = row[0][7:]

        url_summary = 'https://api.ocrolus.com/v1/book/summary'
        headers = {'content-type': 'application/json'}
        params_summary = {
                    'pk': bookpk,
                    'extra_fields': 'estimated_revenue_txns_list, non_estimated_revenue_txns_list'}
        ra = requests.get(url_summary, params=params_summary, auth=HTTPBasicAuth(uname, apikey), headers=headers)

        summary = ra.content.decode("utf-8") 
        with open('APIResponse.json', 'w') as file:
            file.write(str(summary))

    except Exception as e: 
        print(str(e))
        exit()

    ReadAndWrite(bookpk)


def GUI():

    HEIGHT = 500
    WIDTH = 800

    root = tk.Tk()

    canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH)
    canvas.pack()

    root.title("Generate Excel Report")

    frame = tk.Frame(root, bg='#80c1ff', bd=5)
    frame.place(relx=0.5, rely=0.1, relwidth=0.75, relheight=0.1, anchor='n')

    entry = tk.Entry(frame, font=40)
    entry.place(relwidth=0.65, relheight=1)

    button = tk.Button(frame, text="Submit Book PK", font=40, command=lambda: Call_Request_API(entry.get()))
    button.place(relx=0.7, relheight=1, relwidth=0.3)

    lower_frame = tk.Frame(root, bg='#80c1ff', bd=10)
    lower_frame.place(relx=0.5, rely=0.25, relwidth=0.75, relheight=0.6, anchor='n')

    label = tk.Label(lower_frame)
    label.place(relwidth=1, relheight=1)

    root.mainloop()

# GUI()
ReadAndWrite("")
