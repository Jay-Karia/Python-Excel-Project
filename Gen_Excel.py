from requests.auth import HTTPBasicAuth
from openpyxl import load_workbook
from datetime import datetime
from tkinter import *
import requests
import json
import csv
import tkinter as tk
import time

excel_report_location = ""

strInput_Excel_filename = "Template.xlsx"
strOutput_Excel_filename = "Report.xlsx"
strInput_JSON_file = "APIResponse.json"
strBookPK = ""
All_JSON_Data = "test"

Bank_account_numbers_list = []
Bank_account_number_last_4_digits = []
worksheetDeposit = ""

beginning_date = []
beginning_balance = []
end_date = []
end_balance = []

def Report_Location(location):
    global excel_report_location
    excel_report_location = location

def UpdateDataOnExcel(bookpk):
    global worksheetDeposit

    def WriteBankAccountNo(account_numbers):
        if len(account_numbers) == 1:
            worksheetDeposit['G5'] = account_numbers[0]
        elif len(account_numbers) == 2:
            worksheetDeposit['G28'] = account_numbers[1]
        elif len(account_numbers) == 3:
            worksheetDeposit['G51'] = account_numbers[2]
        elif len(account_numbers) == 4:
            worksheetDeposit['G74'] = account_numbers[3]

    def WriteBalanceAndDate(b_balance, e_balance, b_date, e_date):
        for i in range(len(b_balance)):
            b_balance[i] = float(b_balance[i])
            e_balance[i] = float(e_balance[i])
        if len(b_balance) == 1 or len(e_balance) == 1:
            worksheetDeposit['I22'] = b_balance[0]
            worksheetDeposit['O22'] = e_balance[0]
            worksheetDeposit['F22'] = b_date[0]
            worksheetDeposit['L22'] = e_date[0]

        elif len(b_balance) == 2 or e_balance == 2:

            worksheetDeposit['I45'] = b_balance[1]
            worksheetDeposit['O45'] = e_balance[1]
            worksheetDeposit['F45'] = b_date[1]
            worksheetDeposit['L45'] = e_date[1]

        elif len(b_balance) == 3 or e_balance == 3:

            worksheetDeposit['I68'] = b_balance[2]
            worksheetDeposit['O68'] = e_balance[2]
            worksheetDeposit['F68'] = b_date[2]
            worksheetDeposit['L68'] = e_date[2]

        elif len(b_balance) == 4 or e_balance == 4:

            worksheetDeposit['I91'] = b_balance[3]
            worksheetDeposit['O91'] = e_balance[3]
            worksheetDeposit['F91'] = b_date[3]
            worksheetDeposit['L91'] = e_date[3]

    def WriteEstimatedRevenue(estimated_revenue):
        if len(estimated_revenue) == 1:
            worksheetDeposit['I23'] = estimated_revenue[0]
        elif len(estimated_revenue) == 2:
            worksheetDeposit['I46'] = estimated_revenue[1]
        elif len(estimated_revenue) == 3:
            worksheetDeposit['I69'] = estimated_revenue[2]
        elif len(estimated_revenue) == 4:
            worksheetDeposit['I92'] = estimated_revenue[3]

    def WriteDepositDateAndAmount(deposit_sum, begin_date, total_bank_accounts):
        start_block = [8, 31, 54, 77]
        try:
            if len(deposit_sum)>12:
                del deposit_sum[12:]
                del begin_date[12:]
            for i in range(0, len(begin_date)):
                deposit_sum[i] = float(deposit_sum[i])
                if total_bank_accounts == 0:
                    worksheetDeposit[f'E{start_block[0]+i}'] = begin_date[i]
                    worksheetDeposit[f'R{start_block[0]+i}'] = deposit_sum[i]
                elif total_bank_accounts == 1:
                    worksheetDeposit[f'E{start_block[1]+i}'] = begin_date[i]
                    worksheetDeposit[f'R{start_block[1]+i}'] = deposit_sum[i]
                elif total_bank_accounts == 2:
                    worksheetDeposit[f'E{start_block[2]+i}'] = begin_date[i]
                    worksheetDeposit[f'R{start_block[2]+i}'] = deposit_sum[i]
                elif total_bank_accounts == 3:
                    worksheetDeposit[f'E{start_block[3]+i}'] = begin_date[i]
                    worksheetDeposit[f'R{start_block[3]+i}'] = deposit_sum[i]
        except Exception as e:
            print(str(e))
            pass

    def Write_non_estimated_revenue_txns_list_and_dates():
        Rows_List = [8, 31, 54, 77]
        Cols_List = ['G','H','I','J','K','L','M']
        try:
            for i in range(0, len(List_Of_Dict_Dates_amounts)):
                for j in range(0,12):
                    dateKey = worksheetDeposit[f'E{Rows_List[i] + j}'].value

                    #print(List_Of_Dict_Dates_amounts[i].get(dateKey))
                    ListOfTxnAmounts = List_Of_Dict_Dates_amounts[i].get(dateKey)
                    #print(ListOfTxnAmounts)
                    if (ListOfTxnAmounts is not None) and (',' in ListOfTxnAmounts):
                        # List of txn amount
                        ListOfTxnAmounts = ListOfTxnAmounts.split(',')
                        for colCounter in range(0,7):
                            if(colCounter < len(ListOfTxnAmounts)):
                                worksheetDeposit[f'{Cols_List[colCounter]}{Rows_List[i] + j}'] = float(ListOfTxnAmounts[colCounter])                       
                        
                    elif (ListOfTxnAmounts is not None):
                        # Single txn amount
                        ListOfTxnAmounts = float(ListOfTxnAmounts)
                        worksheetDeposit[f'{Cols_List[0]}{Rows_List[i] + j}'] = ListOfTxnAmounts
                     
                #print("\n")
        except Exception as e:
            pass

    workbook = load_workbook(strInput_Excel_filename)
    worksheetDeposit = workbook['Deposits']
    
    worksheetDeposit['F2'] = total_bank_accounts
    
    # Getting the last 4 digits of account number
    for each_bank_acct in range(0, total_bank_accounts):
        account_number = All_JSON_Data['response']['bank_accounts'][each_bank_acct]['account_number']
        if len(account_number) > 4:
            Bank_account_number_last_4_digits.append(account_number[len(account_number)-4]+account_number[len(
                account_number)-3]+account_number[len(account_number)-2]+account_number[len(account_number)-1])
            Bank_account_numbers_list.append(Bank_account_number_last_4_digits[each_bank_acct])
        elif len(account_number) <= 4:
            Bank_account_numbers_list.append(account_number)
        WriteBankAccountNo(Bank_account_numbers_list)  

    # Write Begin/End Balance/Date
    for each_bank_acct in range(0, total_bank_accounts):
        daily_balance = All_JSON_Data['response']['bank_accounts'][each_bank_acct]['daily_balances']
        beginning_date.insert(each_bank_acct, list(daily_balance.keys())[0])
        beginning_balance.insert(each_bank_acct, list(daily_balance.values())[0])
        end_date.insert(each_bank_acct, list(daily_balance.keys())[len(daily_balance)-1])
        end_balance.insert(each_bank_acct, list(daily_balance.values())[len(daily_balance)-1])
        WriteBalanceAndDate(beginning_balance, end_balance, beginning_date, end_date)

    # Write Summary Estimated Revenue Transactions
    estimated_revenue_list = []
    for j in range(0, total_bank_accounts):
        sum = 0
        estimated_revenue = All_JSON_Data['response']['bank_accounts'][j]['estimated_revenue_by_month']
        temp_values = list(estimated_revenue.values())
        for each_bank_acct in range(0, len(temp_values)):
            temp_values[each_bank_acct] = float(temp_values[each_bank_acct])
            sum += temp_values[each_bank_acct]
        estimated_revenue_list.append(sum)
        WriteEstimatedRevenue(estimated_revenue_list)

    # Write deposit date(MM/YYYY) and amount for the estimated_revenue_by_month
    excel_dates = []
    for each_bank_acct in range(0, total_bank_accounts):
        begin_dates = []
        deposit_sums = []
        temp_deposits = All_JSON_Data['response']['bank_accounts'][each_bank_acct]['estimated_revenue_by_month']
        temp_dates = All_JSON_Data['response']['bank_accounts'][each_bank_acct]['estimated_revenue_by_month']

        deposit_sums = list(temp_deposits.values())
        begin_dates = list(temp_dates.keys())

        excel_dates.append(begin_dates)

        deposit_sums.reverse()
        begin_dates.reverse()

        WriteDepositDateAndAmount(deposit_sums, begin_dates, each_bank_acct)

    List_Of_Dict_Dates_amounts = []
    List_of_unique_txn_dates = [[],[],[],[]]
    # non_estimated_revenue_txns_list - Deposits Box (G:N)
    for each_bank_acct in range(0, total_bank_accounts):
        list_non_estimated_revenue_txns_list = All_JSON_Data['response']['bank_accounts'][each_bank_acct]['non_estimated_revenue_txns_list']
        for j in range(0, len(list_non_estimated_revenue_txns_list)):
            #txn_amounts = list_non_estimated_revenue_txns_list[j]['amount']
            txn_date = list_non_estimated_revenue_txns_list[j]['txn_date']

            txn_date = txn_date.replace(f'{txn_date[2]+txn_date[3]+txn_date[4]}', '',1)
            
            List_of_unique_txn_dates[each_bank_acct].append(txn_date)
        List_of_unique_txn_dates[each_bank_acct] = list(set(List_of_unique_txn_dates[each_bank_acct]))   # Now we have a list of unique txn dates
    
    Dict_of_unique_txn_dates = {}
    #Converting the dates List to a Dictionary Using dict.fromkeys()
    for eachList in range (0,len(List_of_unique_txn_dates)):
        Dict_of_unique_txn_dates = dict.fromkeys(List_of_unique_txn_dates[eachList],"")
        List_Of_Dict_Dates_amounts.append(Dict_of_unique_txn_dates)
    
    #Now fill the List of dict with date and txn amount
    for each_bank_acct in range(0, total_bank_accounts):
        list_non_estimated_revenue_txns_list = All_JSON_Data['response']['bank_accounts'][each_bank_acct]['non_estimated_revenue_txns_list']
        for j in range(0, len(list_non_estimated_revenue_txns_list)):
            txn_date = list_non_estimated_revenue_txns_list[j]['txn_date']
            txn_date = txn_date.replace(f'{txn_date[2]+txn_date[3]+txn_date[4]}', '',1)
            txn_amount = list_non_estimated_revenue_txns_list[j]['amount']

            if ((List_Of_Dict_Dates_amounts[each_bank_acct][txn_date]) == ""):
                List_Of_Dict_Dates_amounts[each_bank_acct][txn_date] = txn_amount
            else:
                List_Of_Dict_Dates_amounts[each_bank_acct][txn_date] = List_Of_Dict_Dates_amounts[each_bank_acct][txn_date] + "," + txn_amount
            #print(List_Of_Dict_Dates_amounts[each_bank_acct][txn_date])
            
        print()
                #print(List_Of_Dict_Dates_amounts[each_bank_acct][j])

    #Write the list of non_estimated_revenue_txns_list. for eah month
    Write_non_estimated_revenue_txns_list_and_dates()

    #Write the OUTPUT Excel file
    file_location = ""
    if not excel_report_location == "":
        file_location = excel_report_location + "\\"
    Output_Excel_filename = file_location + "Report  for PK " + bookpk  + "_created_" + time.strftime("%Y%b%d-%H%M%S") + ".xlsx"
    workbook.save(Output_Excel_filename)
    print(f"Excel File: \"{Output_Excel_filename}\", Successfully created!")    

    print("<<<function end UpdateDataOnExcel")

def ReadAllJSONData(bookpk):

    global All_JSON_Data
    global total_bank_accounts
    global Bank_account_numbers_list

    try:
        if not strInput_JSON_file.endswith('.json'):
            input_json_file = f"{Input_JSON_file}.json"
        with open(strInput_JSON_file) as json_file:
            All_JSON_Data = json.load(json_file)

        # Getting the total bank account numbers
        bank_accounts = All_JSON_Data['response']['bank_accounts']
        # Update the excel file with json bank_accounts number and replacing with F2 block in excel
        total_bank_accounts = len(bank_accounts)
        if total_bank_accounts > 4:
            total_bank_accounts = 4
    
        print("<<<function end ReadAllJSONData")
        UpdateDataOnExcel(bookpk)

    except Exception as e: 
        print(str(e))

def Call_Request_API(bookpk, location):
    Report_Location(location)
    try:
        # Calling API
        with open(r'Credential.txt', newline='') as csvfile:
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

    ReadAllJSONData(bookpk)

def GUI():

    HEIGHT = 500
    WIDTH = 800

    root = tk.Tk()

    canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH)
    canvas.pack()

    root.title("Generate Excel Report")

    frame = tk.Frame(root, bg='#80c1ff', bd=5)

    frame_2 = tk.Frame(root, bg='#80c1ff', bd=5)

    entry = tk.Entry(frame, font=40)

    entry_2 = tk.Entry(frame_2, font=40)

    label_2 = tk.Label(frame_2, text="Report location", font=40)
    button = tk.Button(frame, text="Submit Book PK", font=40, command=lambda: Call_Request_API(entry.get(), entry_2.get()))


    frame_2.place(relx=0.5, rely=0.1, relwidth=0.75, relheight=0.1, anchor='n')
    entry_2.place(relwidth=0.65, relheight=1)
    label_2.place(relx=0.7, relheight=1, relwidth=0.3)

    frame.place(relx=0.5, rely=0.3, relwidth=0.75, relheight=0.1, anchor='n')
    entry.place(relwidth=0.65, relheight=1)
    button.place(relx=0.7, relheight=1, relwidth=0.3)


    root.mainloop()


GUI()
