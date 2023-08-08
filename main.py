import calendar
import datetime
import xlsxwriter


# Requesting for month and year
month = int(input("Podaj miesiąc: "))
year = int(input("Podaj rok: "))
num_days = calendar.monthrange(year, month)[1]
dates = []

# Making list of days in month
for day in range(1,num_days+1):
    date = datetime.datetime(year,month,day)
    dates.append(date)

# Ask for quantity of people in excel
people_quantity = int(input("Podaj ilość osób do dodania w excelu: "))
people = []

# Requesting for people and adding them to list
for i in range(people_quantity):
    person = input(f"Podaj dane {i+1}. osoby: ")
    people.append(person)

my_date = datetime.datetime.now()
my_date = my_date.strftime("%d %b %y")

# Creating excel file
month_name = dates[0].strftime("%B")
workbook_name = f"{month_name}.xlsx"
workbook = xlsxwriter.Workbook(workbook_name)
worksheet = workbook.add_worksheet()

# Must have data for loops
rows_num = num_days + 2
columns_num = 4 + people_quantity * 3
start_date = datetime.datetime(1899, 12, 30)

# Formats
weekday_format = workbook.add_format()
weekday_format.set_num_format("ddd")
weekday_format.set_left(5)
weekday_format.set_right(5)

week_num_format = workbook.add_format()
week_num_format.set_left(5)
week_num_format.set_right(5)

day_month_year_format = workbook.add_format()
day_month_year_format.set_num_format("d mmm yy")

# Loops to fill excel
for column in range(columns_num):
    for row in range(rows_num):

        # Week of year
        if column == 0:
            # Name
            if row == 1:
                worksheet.write(row, column, "Tydz. roku")
            # Values
            elif row > 1 and row < rows_num - 1:
                worksheet.write(row, column, f"=WEEKNUM(B{row+1})", week_num_format)
            elif row == rows_num - 1:
                date = dates[row - 2]
                days = (date - start_date).days
                down_cell_format = workbook.add_format()
                down_cell_format.set_left(5)
                down_cell_format.set_right(5)
                down_cell_format.set_bottom(5)
                worksheet.write(row, column, f"=WEEKNUM(B{row+1})", down_cell_format)

        # Date
        elif column == 1:
            # Name
            if row == 1:
                worksheet.write(row, column, "Data")
            # Values
            elif row > 1 and row < rows_num - 1:
                date = dates[row-2]
                days = (date - start_date).days
                worksheet.write(row, column, days, day_month_year_format)
            elif row == rows_num - 1:
                date = dates[row - 2]
                days = (date - start_date).days
                cell_format = workbook.add_format()
                cell_format.set_num_format("d mmm yy")
                cell_format.set_bottom(5)
                worksheet.write(row, column, days, cell_format)

        # Day of week
        elif column == 2:
            # Name
            if row == 1:
                worksheet.write(row, column, "Dzień")
            # Values
            elif row > 1 and row < rows_num - 1:
                date = dates[row - 2]
                days = (date - start_date).days
                worksheet.write(row, column, days, weekday_format)
            elif row == rows_num - 1:
                date = dates[row - 2]
                days = (date - start_date).days
                down_cell_format = workbook.add_format()
                down_cell_format.set_num_format("ddd")
                down_cell_format.set_left(5)
                down_cell_format.set_right(5)
                down_cell_format.set_bottom(5)
                worksheet.write(row, column, days, down_cell_format)

        # Empty column
        elif row == 3:
            cell_format = workbook.add_format()
            


workbook.close()
