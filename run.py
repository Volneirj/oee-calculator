import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
    ]

CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
SHEET = GSPREAD_CLIENT.open('oee_calculator')

def get_daily_overall_data():
    """
    Get daily overall data
    """
    print("Please enter the daily overall data.")
    overall_data=[]
    date = get_valid_date_input("Enter the date of your data input (dd/mm/yyyy): ")
    overall_data.append(date)
    supervisor = get_valid_string_input("Please name of supervisor: ")
    overall_data.append(supervisor)
    shift_length = get_valid_integer_input("Please add shift lenght (minutes): ")
    overall_data.append(shift_length)
    short_breaks = get_valid_integer_input("Please add short breaks (minutes): ")
    overall_data.append(short_breaks)
    meal_breaks = get_valid_integer_input("Please add meal breaks (minutes): ")
    overall_data.append(meal_breaks)
    downtime = get_valid_integer_input("Please add machine down time (minutes): ")
    overall_data.append(downtime)
    ideal_run = get_valid_integer_input("Please ideal run rate (part per minutes): ")
    overall_data.append(ideal_run)
    total_pieces = get_valid_integer_input("Please add the total pieces processed: ")
    overall_data.append(total_pieces)
    rejected_pieces = get_valid_integer_input("Please add rejected pieces: ")
    overall_data.append(rejected_pieces)
    headers = ['Date', 'Supervisor', 'Shift Length', 'Short Breaks', 'Meal Break', 'Down Time', 'Ideal Run Rate', 'Total Pieces', 'Reject Pieces']
    data_dict = dict(zip(headers, overall_data))
    print(data_dict)
    ask_yes_no_question("Are the information given correct?", overall_data, get_daily_overall_data)


def ask_yes_no_question(question, data, function):
    response = input(question + " (yes/no): ").lower()
    if response == 'yes':
        return data
    elif response == 'no':
        print("The data collection will be restarted")
        function()
    else:
        print("Invalid response. Please enter 'yes' or 'no'.")
        return ask_yes_no_question(question)

def get_valid_date_input(prompt):
    while True:
        try:
            date_str = input(prompt)
            date_obj = datetime.strptime(date_str, "%d/%m/%Y")
            return date_obj.strftime("%d/%m/%Y")
        except ValueError:
            print("Invalid date format. Please use dd/mm/yyyy.")

def get_valid_string_input(prompt):
    while True:
        value = input(prompt)
        if value.isalpha():
            return value
        else:
            print("Invalid input. Please enter letters only.")

def get_valid_integer_input(prompt):
    while True:
        value = input(prompt)
        if value.isdigit():
            return int(value)
        else:
            print("Invalid input. Please enter a valid integer.")

get_daily_overall_data()