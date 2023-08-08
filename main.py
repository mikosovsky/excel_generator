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

rows_num = 4 + people_quantity * 3
columns_num = num_days + 1
for row in range(rows_num):
    for column in range(columns_num):
        if row == 0:
            if column == 0:
                worksheet.write(column, row, "Tydz. roku")
            else:
                week_num = dates[column - 1].strftime("%V")
                worksheet.write(column, row, week_num)


workbook.close()
