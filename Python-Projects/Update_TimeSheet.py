# Script to change the dates in a time sheet
from openpyxl import load_workbook
import sys

# Locally access the file // can be put directly into the folder
file_path = "/Users/KVohra/Desktop/time.xlsx"
# Loads the Workbook and assigns it the object "wb"
wb = load_workbook(filename=file_path)
# Sets the current worksheet to "sheet"
sheet = wb.active
# List that holds the number of months
months = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
# List that holds the total amount of days in a month
month_days = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
# List of all of the cells that need to be modified
wb_days = ['A18', 'A20', 'A22', 'A24', 'A26', 'A28', 'A36', 'A38', 'A40', 'A42', 'A44', 'A46']
# List that holds all of the start dates
start_days =["08/05", "08/19", "09/02", "09/16", "09/30", "10/14"]
# List that holds all of the end days
end_days = ["08/17", "08/31", "09/14", "09/28", "10/12", "10/26"]
# List of days
days_of_week = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]


# Function that will take in a start and an end
def change_start_end(start, end):
    # Counts how many days left to modify
    date_counter = 0
    day_counter = 0
    # 4 Lines modify the start and the end dates
    new_start = start + "/2019"
    new_end = end + "/2019"
    sheet['B6'].value = new_start
    sheet['D6'].value = new_end

    # Loop that iterates over the string
    for i in range(len(start)):
        # Splits on the "/"
        split_date = start.split("/")
        parsed_month = int(split_date[0])
        parsed_day = int(split_date[1])
        # gets the total number of days in a given month
        total_days_in_month = month_days[parsed_month]

    print("\n The following dates are added: \n")

    # Inserts the modified dates into the cells
    while date_counter < 12:
        # Checks to make sure days dont pass their max
        if parsed_day > total_days_in_month:
            parsed_day = 1
            parsed_month = parsed_month + 1
        if day_counter == 6:
            day_counter = 0
            parsed_day = parsed_day + 1
        # Changes the value of the cell
        sheet[wb_days[date_counter]].value = str(parsed_month)+"/"+str(parsed_day)+"/2019"
        print(days_of_week[day_counter] + " : " + sheet[wb_days[date_counter]].value)
        parsed_day = parsed_day + 1
        # Condition for the while loop
        date_counter = date_counter + 1
        day_counter = day_counter + 1
    # Saves the file
    wb.save(file_path)
# End of function


# Function that returns all of a the pay periods
def list_dates():
    print("\n" + "Upcoming Pay Periods: ")
    for i in range(len(start_days)):
        print("\n " + start_days[i] + " - " + end_days[i] + "\n")


# Takes in the starting date of the pay period
def user_input():
    flag1 = True
    while flag1:
        param1 = input("Please enter the start date (mm/dd): ")
        # Checks the list to make sure that the date that was selected is valid
        for slots in start_days:
            if param1 == slots:
                end_index = start_days.index(param1)
                flag1 = False
                param2 = end_days[end_index]
                break
        if flag1:
            print("The Start date you entered is invalid please enter a valid date")
            continue

    change_start_end(param1, param2)


# Allows user to choose a command
def selection():
    loop = True

    while loop:
        select = input("Time sheet update script \n Commands: \n U -- Update \n"
                       " L -- List all start and end dates \n E -- Exit \n")

        if select == "U" or select == "u":
            user_input()
            break
        elif select == "L" or select == "l":
            list_dates()
            continue
        elif select == "E" or select == "e":
            sys.exit()
        else:
            print("Please enter a valid command: ")


# Start the program
if __name__ == '__main__':
    selection()
