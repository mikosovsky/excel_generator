import calendar
import datetime
import xlsxwriter
import math
from azureml.opendatasets import PublicHolidays

def column_to_char(value):
    return_value = ""
    if value <= 25:
        return_value =  chr(65 + value)
    else:
        return_value = chr(65 + int(math.floor(value / 26)) - 1) + chr(65 + value % 26)
    return return_value

# Requesting for month and year
month = int(input("Podaj miesiąc: "))
year = int(input("Podaj rok: "))
num_days = calendar.monthrange(year, month)[1]
dates = []
holiday_rows = []

# Ask for quantity of people in excel
people_quantity = int(input("Podaj ilość osób do dodania w excelu: "))
people = []

# Requesting for people and adding them to list
for i in range(people_quantity):
    person = input(f"Podaj dane {i+1}. osoby: ")
    people.append(person)

# Making list of days in month
for day in range(1,num_days+1):
    date = datetime.datetime(year,month,day)
    dates.append(date)

# Download holidays from Microsoft server
end_date = dates[len(dates) - 1]
start_date = dates[0]
hol = PublicHolidays(country_or_region = 'PL', start_date=start_date, end_date=end_date)
hol_df = hol.to_pandas_dataframe()
holiday_list = hol_df["date"].tolist()

# Making list of rows with holiday in month
for day in range(1,num_days+1):
    date = datetime.datetime(year,month,day)
    if date.weekday() == 5 or date.weekday() == 6 or date in holiday_list:
        holiday_rows.append(day + 1)

# Creating excel file
month_name = dates[0].strftime("%B")
workbook_name = f"{month_name}.xlsx"
workbook = xlsxwriter.Workbook(workbook_name)
worksheet = workbook.add_worksheet()

# Must-have data for loops
rows_num = num_days + 2
columns_num = 4 + people_quantity * 3
start_excel_date = datetime.datetime(1899, 12, 30)

# Formats
weekday_format = workbook.add_format()
weekday_format.set_num_format("ddd")
weekday_format.set_left(2)
weekday_format.set_right(2)

week_num_format = workbook.add_format()
week_num_format.set_left(2)
week_num_format.set_right(2)

day_month_year_format = workbook.add_format()
day_month_year_format.set_num_format("d mmm yy")

empty_format = workbook.add_format()
empty_format.set_bg_color("gray")

name_format = workbook.add_format()
name_format.set_left(5)
name_format.set_right(5)
name_format.set_bottom(5)
name_format.set_top(5)

person_format = workbook.add_format()
person_format.set_left(5)
person_format.set_right(5)
person_format.set_bottom(5)
person_format.set_top(5)
person_format.set_bold(True)
person_format.set_bg_color("yellow")
person_format.set_align("center")

person_info_format = workbook.add_format()
person_info_format.set_left(5)
person_info_format.set_right(5)
person_info_format.set_bottom(5)
person_info_format.set_top(5)
person_info_format.set_align("center")

hour_format = workbook.add_format()
hour_format.set_left(2)
hour_format.set_right(2)
hour_format.set_num_format("h:mm")

end_hour_format = workbook.add_format()
end_hour_format.set_left(2)
end_hour_format.set_right(2)
end_hour_format.set_bottom(2)
end_hour_format.set_num_format("h:mm")

sum_time_format = workbook.add_format()
sum_time_format.set_right(2)
sum_time_format.set_num_format("h:mm")

holiday_hour_format = workbook.add_format()
holiday_hour_format.set_left(2)
holiday_hour_format.set_right(2)
holiday_hour_format.set_bg_color("red")
holiday_hour_format.set_align("center")

holiday_end_hour_format = workbook.add_format()
holiday_end_hour_format.set_left(2)
holiday_end_hour_format.set_right(2)
holiday_end_hour_format.set_bottom(2)
holiday_end_hour_format.set_bg_color("red")
holiday_end_hour_format.set_align("center")

# Loops to fill excel
for column in range(columns_num):
    for row in range(rows_num):

        # Week of year
        if column == 0:
            # Empty cell
            if row == 0:
                worksheet.write(row,column,"",empty_format)
            # Name
            elif row == 1:
                worksheet.write(row, column, "Tydz. roku", name_format)
            # Values
            elif 1 < row < rows_num - 1:
                date = dates[row - 2]
                days = (date - start_excel_date).days
                down_cell_format = workbook.add_format()
                if row in holiday_rows:
                    holiday_week_num_format = workbook.add_format()
                    holiday_week_num_format.set_left(2)
                    holiday_week_num_format.set_right(2)
                    holiday_week_num_format.set_bg_color("red")
                    worksheet.write(row, column, f"=WEEKNUM(B{row + 1})", holiday_week_num_format)
                else:
                    worksheet.write(row, column, f"=WEEKNUM(B{row + 1})", week_num_format)
            elif row == rows_num - 1:
                date = dates[row - 2]
                days = (date - start_excel_date).days
                down_cell_format = workbook.add_format()
                down_cell_format.set_left(2)
                down_cell_format.set_right(2)
                down_cell_format.set_bottom(2)
                if row in holiday_rows:
                    down_cell_format.set_bg_color("red")
                    worksheet.write(row, column, f"=WEEKNUM(B{row + 1})", down_cell_format)
                else:
                    worksheet.write(row, column, f"=WEEKNUM(B{row + 1})", down_cell_format)

        # Date
        elif column == 1:
            # Empty cell
            if row == 0:
                worksheet.write(row, column, "", empty_format)
            # Name
            elif row == 1:
                worksheet.write(row, column, "Data", name_format)
            # Values
            elif 1 < row < rows_num - 1:
                date = dates[row-2]
                days = (date - start_excel_date).days
                if row in holiday_rows:
                    holiday_day_month_year_format = workbook.add_format()
                    holiday_day_month_year_format.set_left(2)
                    holiday_day_month_year_format.set_right(2)
                    holiday_day_month_year_format.set_bg_color("red")
                    holiday_day_month_year_format.set_num_format("d mmm yy")
                    worksheet.write(row, column, days, holiday_day_month_year_format)
                else:
                    worksheet.write(row, column, days, day_month_year_format)
            elif row == rows_num - 1:
                date = dates[row - 2]
                days = (date - start_excel_date).days
                cell_format = workbook.add_format()
                cell_format.set_num_format("d mmm yy")
                cell_format.set_bottom(2)
                if row in holiday_rows:
                    cell_format.set_bg_color("red")
                    worksheet.write(row, column, days, cell_format)
                else:
                    worksheet.write(row, column, days, cell_format)

        # Day of week
        elif column == 2:
            # Empty cell
            if row == 0:
                worksheet.write(row, column, "", empty_format)
            # Name
            elif row == 1:
                worksheet.write(row, column, "Dzień", name_format)
            # Values
            elif 1 < row < rows_num - 1:
                date = dates[row - 2]
                days = (date - start_excel_date).days
                if row in holiday_rows:
                    holiday_weekday_format = workbook.add_format()
                    holiday_weekday_format.set_left(2)
                    holiday_weekday_format.set_right(2)
                    holiday_weekday_format.set_num_format("ddd")
                    holiday_weekday_format.set_bg_color("red")
                    worksheet.write(row,column,days,holiday_weekday_format)
                else:
                    worksheet.write(row, column, days, weekday_format)
            elif row == rows_num - 1:
                date = dates[row - 2]
                days = (date - start_excel_date).days
                down_cell_format = workbook.add_format()
                down_cell_format.set_num_format("ddd")
                down_cell_format.set_left(2)
                down_cell_format.set_right(2)
                down_cell_format.set_bottom(2)
                if row in holiday_rows:
                    down_cell_format.set_bg_color("red")
                    worksheet.write(row, column, days, down_cell_format)
                else:
                    worksheet.write(row, column, days, down_cell_format)

        # Empty column
        elif column == 3:
            worksheet.write(row, column, "", empty_format)

        # 1st column of person
        elif column > 3 and column % 3 == 1:
            if row == 0:
                worksheet.merge_range(row, column, row, column+2, people[int(column / 3 - 1)], person_format)
            elif row == 1:
                worksheet.write(row, column, "Od", person_info_format)
            elif row < rows_num - 1:
                if row in holiday_rows:
                    worksheet.write(row, column, "", holiday_hour_format)
                else:
                    worksheet.write(row, column, "", hour_format)
            else:
                if row in holiday_rows:
                    worksheet.write(row, column, "", holiday_end_hour_format)
                else:
                    worksheet.write(row, column, "", end_hour_format)

        # 2nd column of person
        elif column > 3 and column % 3 == 2:
            if row == 1:
                worksheet.write(row, column, "Do", person_info_format)
            elif 1 < row < rows_num - 1:
                if row in holiday_rows:
                    worksheet.write(row, column, "", holiday_hour_format)
                else:
                    worksheet.write(row, column, "", hour_format)
            elif row == rows_num - 1:
                if row in holiday_rows:
                    worksheet.write(row, column, "", holiday_end_hour_format)
                else:
                    worksheet.write(row, column, "", end_hour_format)

        # 3rd column of person
        elif column > 3 and column % 3 == 0:
            if row == 1:
                worksheet.write(row, column, "Suma", person_info_format)
            elif 1 < row < rows_num - 1:
                if row in holiday_rows:
                    worksheet.write(row, column, "", holiday_hour_format)
                else:
                    first_column = column_to_char(column - 2)
                    second_column = column_to_char(column - 1)
                    worksheet.write(row, column, f"={second_column}{row+1} - {first_column}{row+1}", sum_time_format)
            elif row == rows_num - 1:
                if row in holiday_rows:
                    worksheet.write(row, column, "", holiday_end_hour_format)
                else:
                    first_column = column_to_char(column - 2)
                    second_column = column_to_char(column - 1)
                    down_cell_format = workbook.add_format()
                    down_cell_format.set_num_format("h:mm")
                    down_cell_format.set_right(2)
                    down_cell_format.set_bottom(2)
                    worksheet.write(row, column, f"={second_column}{row + 1} - {first_column}{row + 1}", down_cell_format)

workbook.close()
