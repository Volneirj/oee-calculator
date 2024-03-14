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


def ask_yes_no_question(question, data, function):
    response = input(question + " (yes/no): ").lower()
    if response == 'yes':
        if data is None:
            print('Thank you for using OEE Calculator,'
                  'this softawe has been developed by Volnei Resena Junior.\n'
                  'This code can be found at'
                  ' https://github.com/Volneirj/oee-calculator\n')
            print("Exiting the program.")
            raise SystemExit
        else:
            return data
    elif response == 'no':
        print("\nThe data collection will be restarted\n")
        function()
    else:
        print("\nInvalid response. Please enter 'yes' or 'no'.")
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
        if value.replace(" ", "").isalpha() and len(value) >= 3:
            return value.upper()
        else:
            print('Invalid input.'
                  'The name must contain 3 or more letters.\n'
                  'Please enter letters only.')


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
    daily_data = []
    date = get_valid_date_input("Date (dd/mm/yyyy): ")
    daily_data.append(date)
    supervisor = get_valid_string_input("Supervisor Name: ")
    daily_data.append(supervisor)
    shift_length = get_valid_integer_input("Shift length (minutes): ")
    daily_data.append(shift_length)
    short_breaks = get_valid_integer_input("Short breaks (minutes): ")
    daily_data.append(short_breaks)
    meal_breaks = get_valid_integer_input("Meal breaks (minutes): ")
    daily_data.append(meal_breaks)
    downtime = get_valid_integer_input("Machine down time (minutes): ")
    daily_data.append(downtime)
    ideal_run = get_valid_integer_input("Ideal run rate (parts per minute): ")
    daily_data.append(ideal_run)
    total_pieces = get_valid_integer_input("Total processed pieces: ")
    daily_data.append(total_pieces)
    rejected_pieces = get_valid_integer_input("Total rejected pieces: ")
    daily_data.append(rejected_pieces)
    headers = ['Date', 'Supervisor', 'Shift Length',
               'Short Breaks', 'Meal Break', 'Down Time',
               'Ideal Run Rate', 'Total Pieces', 'Reject Pieces']
    data_dict = dict(zip(headers, daily_data))
    print(data_dict)
    ask_yes_no_question(
        "Are the information given correct?", daily_data, get_daily_data
    )
    return daily_data


def update_worksheet(data, worksheet):
    """
    Update worksheet, add new row with the report data provided
    """
    print(f"Updating {worksheet} worksheet...\n")
    sheet = SHEET.worksheet(worksheet)
    sheet.append_row(data)
    print(f"Worksheet {worksheet} updated sucessfully \n")


def calculate_variables(data):
    """
    Calculate the variables planned production time,
    operating time and good pieces based on daily data.
    """
    variables = []

    date = data[0]
    variables.append(date)

    planned_prod_time = int(data[2]) - int(data[3]) - int(data[4])
    variables.append(planned_prod_time)

    operating_time = planned_prod_time - int(data[5])
    variables.append(operating_time)

    good_pieces = int(data[7]) - int(data[8])
    variables.append(good_pieces)

    return variables


def calculate_oee(variables, data):
    """
    Calculate availability, performance,
    quality and Overal OEE based on calulated
    variables and the daily data supplied.

    variables = variables result list
    data = report list
    """
    oee_factor = []

    date = data[0]
    oee_factor.append(date)

    availability = (int(variables[2])/int(variables[1]))
    oee_factor.append(round(availability, 2))

    performance = ((int(data[7])/int(variables[2]))/data[6])
    oee_factor.append(round(performance, 2))

    quality = (int(variables[3])/int(data[7]))
    oee_factor.append(round(quality, 2))

    overall_oee = availability * performance * quality
    oee_factor.append(round(overall_oee, 2))

    return oee_factor


def main():
    """
    Run all program functions
    """
    day_data = get_daily_data()

    update_worksheet(day_data, "report")

    day_variables = calculate_variables(day_data)

    update_worksheet(day_variables, "variables")

    day_oee = calculate_oee(day_variables, day_data)

    update_worksheet(day_oee, "oee_factor")

    print(f'The production today {day_data[0]} with the supervision of'
          f'{day_data[1]} reached the availability of:'
          f'{day_oee[1]*100:.2f}%, performance: {day_oee[2]*100:.2f}%,'
          f'and quality: {day_oee[3]*100:.2f}%.'
          f'In general, the Overall OEE (Overall Equipment Effectiveness)'
          f' reached: {day_oee[4]*100:.2f}%.')'


main()
