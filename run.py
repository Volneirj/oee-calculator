import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import openpyx1

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
    ]

CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
SHEET = GSPREAD_CLIENT.open('oee_calculator')

def ask_yes_no_question(question, data, function):
    response = input(question + " (yes/no): ").lower()
    if response == 'yes':
        return data
    elif response == 'no':
        print("The data collection will be restarted")
        function()
    else:
        print("Invalid response. Please enter 'yes' or 'no'.")
        return ask_yes_no_question(question, data, function)

def get_valid_date_input(prompt):
    while True:
        try:
            date_str = input(prompt)
            date_obj = datetime.strptime(date_str, '%d/%m/%Y')
            formatted_date = date_obj.strftime('%d/%m/%Y')
            return formatted_date
        except ValueError:
            print("Invalid date format. Please use dd/mm/yyyy.")

def get_valid_string_input(prompt):
    while True:
        value = input(prompt)
        if value.isalpha():
            return value.upper()
        else:
            print("Invalid input. Please enter letters only.")

def get_valid_integer_input(prompt):
    while True:
        value = input(prompt)
        if value.isdigit():
            return int(value)
        else:
            print("Invalid input. Please enter a valid integer.")


def get_daily_data():
    """
    Get daily report data verifying 
    in every input if it is correct
    """
    print("Please enter the daily report data.")
    daily_data=[]
    date = get_valid_date_input("Enter the date of your data input (dd/mm/yyyy): ")
    daily_data.append(date)
    supervisor = get_valid_string_input("Please name of supervisor: ")
    daily_data.append(supervisor)
    shift_length = get_valid_integer_input("Please add shift lenght (minutes): ")
    daily_data.append(shift_length)
    short_breaks = get_valid_integer_input("Please add short breaks (minutes): ")
    daily_data.append(short_breaks)
    meal_breaks = get_valid_integer_input("Please add meal breaks (minutes): ")
    daily_data.append(meal_breaks)
    downtime = get_valid_integer_input("Please add machine down time (minutes): ")
    daily_data.append(downtime)
    ideal_run = get_valid_integer_input("Please ideal run rate (part per minutes): ")
    daily_data.append(ideal_run)
    total_pieces = get_valid_integer_input("Please add the total pieces processed: ")
    daily_data.append(total_pieces)
    rejected_pieces = get_valid_integer_input("Please add rejected pieces: ")
    daily_data.append(rejected_pieces)
    headers = ['Date', 'Supervisor', 'Shift Length', 'Short Breaks', 'Meal Break', 'Down Time', 'Ideal Run Rate', 'Total Pieces', 'Reject Pieces']
    data_dict = dict(zip(headers, daily_data))
    print(data_dict)
    ask_yes_no_question("Are the information given correct?", daily_data, get_daily_data)

    return daily_data

def update_worksheet(data,worksheet):
    """
    Update worksheet, add new row with the report data provided
    """
    print(f"Updating {worksheet} worksheet...\n")
    sheet = SHEET.worksheet(worksheet)
    sheet.append_row(data)
    print(f"Report {worksheet} updated sucessfully \n")

def calculate_variables(data):
    """
    Calculate the variables planned production time, 
    operating time and good pieces based on daily data.
    """
    variables = []
        
    date = data[0]
    variables.append(date)
    
    planned_pro_time = int(data[3]) - int(data[4]) - int(data[5])
    variables.append(planned_pro_time)
    
    operating_time = planned_pro_time - int(data[6])
    variables.append(operating_time)
    
    good_pieces = int(data[7]) - int(data[8])
    variables.append(good_pieces)
    
    return variables


def main():
    """
    Run all program functions
    """
    day_data = get_daily_data()
    update_worksheet(day_data, "report")
    variables = calculate_variables()
    update_worksheet(variables, "variables")
    
main()
