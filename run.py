import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import os

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
    ]

CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
SHEET = GSPREAD_CLIENT.open('oee_calculator')


class YesNoQuestion:
    def __init__(self, question, data, function):
        """
        Parameters:
        - question (str): The question to ask the user.
        - data (any): Associated data to return if the
        user responds with 'yes'.
        - function (function): The function to call
        if the user responds with 'no'.
        """
        self.question = question
        self.data = data
        self.function = function

    def ask_question(self):
        """
        Ask the user a yes/no question and handle the response.
        """
        response = input(self.question + " (yes/no): ").lower()
        if response == 'yes':
            if self.data is None:
                print('No data found')
            else:
                return self.data
        elif response == 'no':
            print("\nThe data collection will be restarted\n")
            self.function()
        else:
            print("\nInvalid response. Please enter 'yes' or 'no'.")
            return self.ask_question()


def clear_screen():
    """
    Clear the screen when called (win/mac)
    """
    if os.name == 'nt':
        _ = os.system('cls')
    else:
        _ = os.system('clear')


def press_any_key_to_return():
    """
    Added to user decide when the want clear
    the screen
    """
    input("\nPress Any key to return main menu....")
    clear_screen()


def get_date_input_wfilter(prompt):
    """
    Check if the data is a date,
    Using the filter_data compare if the date already exists
    """
    error_message = (f'Data for this date already exists.\n'
                     f'Please choose a different date or'
                     f' delete it on worksheet.')
    while True:
        try:
            date_str = input(prompt)
            if date_str.lower() == 'q':
                return 'q'
            date_obj = datetime.strptime(date_str, '%d/%m/%Y')
            formatted_date = date_obj.strftime('%d/%m/%Y')
            report = get_data_worksheet('report')
            filtered_data = filter_data(report, formatted_date)
            if filtered_data and filtered_data[0][0] == formatted_date:
                raise ValueError(error_message)
            return formatted_date
        except ValueError as e:
            if str(e) != error_message:
                print("Invalid date format. Please use dd/mm/yyyy.")
            else:
                print(e)


def get_valid_date_input(prompt):
    """
    Check if the data input is a date
    Also show the user the data format expected
    """
    while True:
        try:
            date_str = input(prompt)
            if date_str.lower() == 'q':
                return 'q'
            date_obj = datetime.strptime(date_str, '%d/%m/%Y')
            formatted_date = date_obj.strftime('%d/%m/%Y')
            return formatted_date
        except ValueError:
            print("Invalid date format. Please use dd/mm/yyyy.")


def get_valid_string_input(prompt):
    """
    Check if the input is only string, has not only spaces, len >= 3, and
    return it uppercase, show what could be wront to the user
    """
    while True:
        value = input(prompt)
        if value.lower() == 'q':
            return 'q'
        if value.replace(" ", "").isalpha() and len(value) >= 3:
            return value.upper()
        else:
            print('Invalid input.'
                  'The name must contain 3 or more letters.\n'
                  'Please enter letters only.')


def get_valid_integer_input(prompt):
    """
    Check if input is an integer
    """
    while True:
        value = input(prompt)
        if value == 'q':
            return 'q'
        else:
            if value.isdigit():
                return int(value)
            else:
                print("Invalid input. Please enter a valid integer.")


def code_break(variable):
    """
    Break the code when user input "q"
    """
    if variable == 'q':
        raise KeyboardInterrupt


def get_daily_data():
    """
    Get daily report data verifying
    in every input if it is correct
    """
    try:
        while True:
            print('\nPlease enter the daily report data.\n'
                  'To return Main Menu type "q" \n')
            daily_data = []
            date = get_date_input_wfilter("Date (dd/mm/yyyy): ")
            code_break(date)
            daily_data.append(date)

            supervisor = get_valid_string_input("Supervisor Name: ")
            code_break(supervisor)
            daily_data.append(supervisor)

            shift_length = get_valid_integer_input("Shift length (minutes): ")
            code_break(shift_length)
            daily_data.append(shift_length)

            short_breaks = get_valid_integer_input("Short breaks (minutes): ")
            code_break(short_breaks)
            daily_data.append(short_breaks)

            meal_breaks = get_valid_integer_input("Meal breaks (minutes): ")
            code_break(meal_breaks)
            daily_data.append(meal_breaks)

            downtime = get_valid_integer_input("Machine down time (minutes): ")
            code_break(downtime)
            daily_data.append(downtime)

            ideal_run = get_valid_integer_input('Ideal run rate'
                                                '(parts per minute): ')
            code_break(ideal_run)
            daily_data.append(ideal_run)

            total_pieces = get_valid_integer_input("Total processed pieces: ")
            code_break(total_pieces)
            daily_data.append(total_pieces)

            rejected_pieces = get_valid_integer_input('Total '
                                                      'rejected pieces: ')
            code_break(rejected_pieces)
            daily_data.append(rejected_pieces)

            headers = ['Date', 'Supervisor', 'Shift Length',
                       'Short Breaks', 'Meal Break', 'Down Time',
                       'Ideal Run Rate', 'Total Pieces', 'Reject Pieces']
            data_dict = dict(zip(headers, daily_data))
            print(data_dict)
            is_the_data_correct = YesNoQuestion('Are the '
                                                'information given correct?',
                                                daily_data, get_daily_data)
            is_the_data_correct.ask_question()

            clear_screen()

            return daily_data

    except KeyboardInterrupt:
        press_any_key_to_return()
        return None


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

    variables refer to variables result list
    data refer to report list
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


def oee_results(report, oee_factor):
    """
    Print daily results based on daily reports and oee calculated
    """
    print(f'\nThe production from {report[0]} with the supervision of '
          f'{report[1]} reached:\n\nAvailability:'
          f'{oee_factor[1]*100:.2f}%\nPerformance: {oee_factor[2]*100:.2f}%\n'
          f'Quality: {oee_factor[3]*100:.2f}%\n\n'
          f'Overall OEE (Overall Equipment Effectiveness)'
          f': {oee_factor[4]*100:.2f}%.\n')


def display_menu():
    """
    Print main menu with the available Options
    """
    print("\nMain Menu\n")
    print("1. Add new report")
    print("2. Show all reports")
    print("3. Load OEE by date")
    print("4. Exit\n")


def add_new_report():
    """
    Collect the input data, import it to googlesheets
    calculate the oee factor based on data and variables.
    """
    clear_screen()

    day_report = get_daily_data()

    if day_report is not None:
        update_worksheet(day_report, "report")

        day_variables = calculate_variables(day_report)

        update_worksheet(day_variables, "variables")

        day_oee = calculate_oee(day_variables, day_report)

        update_worksheet(day_oee, "oee_factor")

        clear_screen()

        oee_results(day_report, day_oee)

        press_any_key_to_return()

    else:
        return None


def print_report(worksheet, header, filter_date=None):
    """
    Using the API, connect to the Google Sheet and
    retrieve all data from the specified worksheet, then print it.
    Optionally, filter the data by date before printing.
    """
    clear_screen()
    print(header)

    report = SHEET.worksheet(worksheet)
    data = report.get_all_values()

    if filter_date:
        filtered_data = [row for row in data if row[0] == filter_date]
        if not filtered_data:
            print("No data available.")
            return None
        else:
            data = filtered_data

    column_widths = [max(len(str(cell)) for cell in column)
                     for column in zip(*data)]
    print("|".join(cell.ljust(width) for cell,
                   width in zip(data[0], column_widths)))
    print("-" * (sum(column_widths) + 8))

    for row in data[1:]:
        print("|".join(cell.ljust(width) for cell,
              width in zip(row, column_widths)))

    press_any_key_to_return()


def get_data_worksheet(worksheet):
    """
    Fetches data from the specified worksheet in Google Sheets.
    """
    data = SHEET.worksheet(worksheet)

    return data.get_all_values()


def filter_data(data, selected_date):
    """
    Filters the data based on the selected date.
    """
    filtered_data = [row for row in data if row[0] == selected_date]

    return filtered_data


def validate_data(data):
    """
    Converts numeric strings in the data list to integers or floats.
    """
    converted_data = []
    for row in data:
        converted_row = []
        for cell in row:
            try:
                converted_cell = int(cell)
            except ValueError:
                try:
                    converted_cell = float(cell.replace(',', '.'))
                except ValueError:
                    converted_cell = cell
            converted_row.append(converted_cell)
        converted_data.append(converted_row)

    return converted_data


def oee_by_date():
    """
    Retrieves data for the selected date and prints OEE.
    """
    clear_screen()
    print('\nOEE calculation by date\n'
          'To Return to the Main Menu type "q"\n')

    selected_date = get_valid_date_input('Choose date (dd/mm/yyyy) '
                                         'to print the report: ')
    if selected_date != "q":
        report_data = get_data_worksheet('report')
        oee_factor_data = get_data_worksheet('oee_factor')

        filtered_report = filter_data(report_data, selected_date)
        filtered_oee_factor = filter_data(oee_factor_data, selected_date)

        report_int = validate_data(filtered_report)
        oee_int = validate_data(filtered_oee_factor)

        if report_int and oee_int:
            oee_results(report_int[0], oee_int[0])
        else:
            print(f'\nNo data Available for {selected_date}')

    press_any_key_to_return()


def main():
    """
    Display the menu where the user can select between
    the options available
    """
    while True:
        display_menu()
        choice = input("Please select an option from 1 to 4: ")

        if choice == '1':
            add_new_report()
        elif choice == '2':
            print_report('report', "\nShowing all reports\n")
        elif choice == '3':
            oee_by_date()
        elif choice == '4':
            clear_screen()
            print('\nThank you for using OEE Calculator\n'
                  'This software has been developed by Volnei Resena Junior.\n'
                  'This code can be found at'
                  ' https://github.com/Volneirj/oee-calculator\n')
            print("Exiting the program.")
            break
        else:
            print("Invalid choice. Please enter a number from 1 to 4.")


main()
