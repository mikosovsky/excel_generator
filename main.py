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
rows_num = 4 + people_quantity * 3
columns_num = num_days + 2
start_date = datetime.datetime(1899, 12, 30)

for row in range(rows_num):
    for column in range(columns_num):
        if row == 0:
            if column == 1:
                worksheet.write(column, row, "Tydz. roku")
            elif column > 1:
                worksheet.write(column, row, f"=WEEKNUM(B{column+1})")
        elif row == 1:
            if column == 1:
                worksheet.write(column,row, "Data")
            elif column > 1:
                date = dates[column-2]
                days = (date - start_date).days
                cell_format = workbook.add_format()
                cell_format.set_num_format("d mmm yy")
                worksheet.write(column, row, days, cell_format)
        elif row == 2:
            if column == 1:
                worksheet.write(column, row, "Dzień")
            elif column > 1:
                cell_format = workbook.add_format()
                cell_format.set_num_format("ddd")
                date = dates[column - 2]
                days = (date - start_date).days
                worksheet.write(column, row, days, cell_format)
        elif row == 3:
            cell_format = workbook.add_format()



workbook.close()
