from ctypes import alignment
from openpyxl import load_workbook
from tkinter import *
import json

# Globals

account_numbers_list = []
last_4_digits = []

beginning_date = []
beginning_balance = []
end_date = []
end_balance = []

# Graphical User Interface

root = Tk()
root.title("Excel Generation")
root.geometry("700x550")
root.iconbitmap("excel.ico")
root.resizable(width=False, height=False)

input_file_name = StringVar()
output_file_name = StringVar()
input_json_file = StringVar()

header = Label(root, text="Excel Generation", bg="yellow",
               fg="black", font="Consolas 20 bold")
header.pack(fill="x")

input_excel = Label(root, text="Input Excel File Path (*.xlsx): ", font="Consolas 10 bold", anchor="w", bg="black", fg="white", width=53, justify=CENTER)
input_excel.pack(pady=20)
input_excel_text_field = Entry(root, bg="light yellow", textvariable=input_file_name, width=100, justify=CENTER)
input_excel_text_field.pack()

output_excel = Label(root, text="Output Excel File Path (*.xlsx): ", font="Consolas 10 bold", anchor="w", bg="black", fg="white", width=53, justify=CENTER)
output_excel.pack(pady=30)
output_excel_text_field = Entry(root, bg="light yellow", textvariable=output_file_name, width=100, justify=CENTER)
output_excel_text_field.pack()

input_json = Label(root, text="Input JSON File Path (*.json): ", font="Consolas 10 bold", anchor="w", bg="black", fg="white", width=53, justify=CENTER)
input_json.pack(pady=30)
input_json_text_field = Entry(root, bg="light yellow", textvariable=input_json_file, width=100, justify=CENTER)
input_json_text_field.pack()

input_json_file = input_json_file.get()
input_file_name = input_file_name.get()
output_file_name = output_file_name.get()

# gen_btn = Button(root, text="Generate", bg="blue", fg="white", font="Consolas 10 bold", anchor="w", width=13, command=call_read)
# gen_btn.pack(pady=40)

root.mainloop()

workbook = load_workbook(input_file_name)
worksheet = workbook['Deposits']


def ReadJSONData(input_json_file):
    global account_numbers_list
    global last_4_digits
    global beginning_date
    global beginning_balance
    global end_date
    global end_balance

    if not input_json_file.endswith('.json'):
        input_json_file = f"{input_json_file}.json"
    with open(input_json_file) as json_file:
        data = json.load(json_file)
        # Getting the total account numbers
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

        deposits_list = []
        final_deposits_list = []
        for j in range(0, total_bank_accounts):
            final_deposits = 0
            d_sum = 0
            deposits_month = data['response']['bank_accounts'][j]['deposits_sum_by_month']
            temp_deposits = list(deposits_month.values())
            for i in range(0, len(temp_values)):
                temp_deposits[i] = float(temp_deposits[i])
                d_sum += temp_deposits[i]
            deposits_list.append(d_sum)
            for i in range(0, len(estimated_revenue_list)):
                final_deposits = d_sum - estimated_revenue_list[i]
            final_deposits_list.append(final_deposits)
            WriteDeposits(final_deposits_list)

            # Begin Date
            for i in range(0, total_bank_accounts):
                begin_dates = []
                deposit_sums = []
                total_periods = len(
                    data['response']['bank_accounts'][i]['periods'])
                for j in range(0, total_periods):
                    temp_begin_date = data['response']['bank_accounts'][i]['periods'][j]['begin_date']
                    temp_deposits_sum = data['response']['bank_accounts'][i]['periods'][j]['deposit_sum']
                    begin_dates.append(temp_begin_date)
                    deposit_sums.append(temp_deposits_sum)
                WriteReamainingData(deposit_sums, begin_dates, i)

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


def WriteBalanceAndDate(b_balance, e_balance, b_date, e_date):
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


def WriteDeposits(deposits):
    if len(deposits) == 1:
        worksheet['I24'] = deposits[0]
    elif len(deposits) == 2:
        worksheet['I47'] = deposits[1]
    elif len(deposits) == 3:
        worksheet['I70'] = deposits[2]
    elif len(deposits) == 4:
        worksheet['I93'] = deposits[3]


def WriteReamainingData(deposit_sum, begin_date, total_bank_accounts):
    start_block = [8, 31, 54, 77]
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


# try:
ReadJSONData(input_json_file)
workbook.save(output_file_name)
print(f"Excel File: \"{output_file_name}\", Successfully created!")
# except:
#     print("Could not generate Excel File!")
