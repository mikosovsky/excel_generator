import calendar
import datetime
import xlsxwriter

start_date = datetime.datetime(1899, 12, 30)

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

                worksheet.write(column, row, f"=WEEKNUM(B{column+1})")
        elif row == 1:
            if column == 0:
                worksheet.write(column,row, "Data")
            else:
                date = dates[column-1]
                days = (date - start_date).days
                cell_format = workbook.add_format()
                cell_format.set_num_format("d mmm yy")
                worksheet.write(column, row, days, cell_format)
        elif row == 2:
            if column == 0:
                worksheet.write(column, row, "Dzień")
            else:
                cell_format = workbook.add_format()
                cell_format.set_num_format("ddd")
                day = dates[column-1].strftime("%a")
                worksheet.write(column, row, day,cell_format)


workbook.close()
