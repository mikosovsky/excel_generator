import calendar
import datetime
import xlsxwriter
import math
from azureml.opendatasets import PublicHolidays


class ExcelGenerator:

    indigo = "4c5d93"
    light_gray = "7f7f7f"
    light_blue = "79b1e2"

    def __init__(self):
        self.people = None
        self.dates = []
        self.holiday_rows = []

    def change_month(self, month: int, year: int):
        num_days = calendar.monthrange(year, month)[1]

        # Making list of days in month
        for day in range(1, num_days + 1):
            date = datetime.datetime(year, month, day)
            self.dates.append(date)
        self.__make_holiday_rows_list()

    def change_people_list(self, people):
        self.people = people

    def __make_holiday_rows_list(self):
        # Download holidays from Microsoft server
        num_end_date = len(self.dates) - 1
        end_date = self.dates[num_end_date]
        start_date = self.dates[0]
        hol = PublicHolidays(country_or_region='PL', start_date=start_date, end_date=end_date)
        hol_df = hol.to_pandas_dataframe()
        holiday_list = hol_df["date"].tolist()

        # Making list of rows with holiday in month
        for date in self.dates:
            if date.weekday() == 5 or date.weekday() == 6 or date in holiday_list:
                day = date.day
                self.holiday_rows.append(day + 1)

    @staticmethod
    def __weekday_format_init(weekday_format):
        weekday_format.set_num_format("ddd")
        weekday_format.set_left(2)
        weekday_format.set_right(2)
        return weekday_format

    @staticmethod
    def __week_num_format_init(week_num_format):
        week_num_format.set_left(2)
        week_num_format.set_right(2)
        return week_num_format

    @staticmethod
    def __day_month_year_format_init(day_month_year_format):
        day_month_year_format.set_num_format("d mmm yy")
        return day_month_year_format

    def __empty_format_init(self, empty_format):
        empty_format.set_bg_color(self.indigo)
        return empty_format

    @staticmethod
    def __name_format_init(name_format):
        name_format.set_left(5)
        name_format.set_right(5)
        name_format.set_bottom(5)
        name_format.set_top(5)
        return name_format

    def __person_format_init(self, person_format):
        person_format.set_left(5)
        person_format.set_right(5)
        person_format.set_bottom(5)
        person_format.set_top(5)
        person_format.set_bold(True)
        person_format.set_bg_color(self.light_blue)
        person_format.set_align("center")
        return person_format

    @staticmethod
    def __person_info_format_init(person_info_format):
        person_info_format.set_left(5)
        person_info_format.set_right(5)
        person_info_format.set_bottom(5)
        person_info_format.set_top(5)
        person_info_format.set_align("center")
        return person_info_format

    @staticmethod
    def __hour_format_init(hour_format):
        hour_format.set_left(2)
        hour_format.set_right(2)
        hour_format.set_num_format("h:mm")
        return hour_format

    @staticmethod
    def __end_hour_format_init(end_hour_format):
        end_hour_format.set_left(2)
        end_hour_format.set_right(2)
        end_hour_format.set_bottom(2)
        end_hour_format.set_num_format("h:mm")
        return end_hour_format

    @staticmethod
    def __sum_time_format_init(sum_time_format):
        sum_time_format.set_right(2)
        sum_time_format.set_num_format("h:mm")
        return sum_time_format

    def __holiday_hour_format_init(self, holiday_hour_format):
        holiday_hour_format.set_left(2)
        holiday_hour_format.set_right(2)
        holiday_hour_format.set_bg_color(self.light_gray)
        holiday_hour_format.set_align("center")
        return holiday_hour_format

    def __holiday_end_hour_format_init(self, holiday_end_hour_format):
        holiday_end_hour_format.set_left(2)
        holiday_end_hour_format.set_right(2)
        holiday_end_hour_format.set_bottom(2)
        holiday_end_hour_format.set_bg_color(self.light_gray)
        holiday_end_hour_format.set_align("center")
        return holiday_end_hour_format

    @staticmethod
    def __summary_format_init(summary_format):
        summary_format.set_left(2)
        summary_format.set_bottom(2)
        summary_format.set_align("right")
        return summary_format

    @staticmethod
    def __sum_value_format_init(sum_value_format):
        sum_value_format.set_bottom(2)
        sum_value_format.set_right(2)
        sum_value_format.set_num_format("h:mm")
        return sum_value_format

    def __holiday_week_num_format_init(self, holiday_week_num_format):
        holiday_week_num_format.set_left(2)
        holiday_week_num_format.set_right(2)
        holiday_week_num_format.set_bg_color(self.light_gray)
        return holiday_week_num_format

    @staticmethod
    def column_to_char(value):
        return_value = ""
        if value <= 25:
            return_value = chr(65 + value)
        else:
            return_value = chr(65 + int(math.floor(value / 26)) - 1) + chr(65 + value % 26)
        return return_value

    def generate_excel(self, path):
        # Creating excel file
        month_name = self.dates[0].strftime("%B")
        workbook_name = f"{month_name}.xlsx"
        workbook = xlsxwriter.Workbook(path + "/" + workbook_name)
        worksheet = workbook.add_worksheet()

        # Must-have data for loops
        num_days = len(self.dates)
        people_quantity = len(self.people)
        rows_num = num_days + 2
        columns_num = 4 + people_quantity * 3
        start_excel_date = datetime.datetime(1899, 12, 30)

        # Formats
        weekday_format = workbook.add_format()
        weekday_format = self.__weekday_format_init(weekday_format)

        week_num_format = workbook.add_format()
        week_num_format = self.__week_num_format_init(week_num_format)

        day_month_year_format = workbook.add_format()
        day_month_year_format = self.__day_month_year_format_init(day_month_year_format)

        empty_format = workbook.add_format()
        empty_format = self.__empty_format_init(empty_format)

        name_format = workbook.add_format()
        name_format = self.__name_format_init(name_format)

        person_format = workbook.add_format()
        person_format = self.__person_format_init(person_format)

        person_info_format = workbook.add_format()
        person_info_format = self.__person_info_format_init(person_info_format)

        hour_format = workbook.add_format()
        hour_format = self.__hour_format_init(hour_format)

        end_hour_format = workbook.add_format()
        end_hour_format = self.__end_hour_format_init(end_hour_format)

        sum_time_format = workbook.add_format()
        sum_time_format = self.__sum_time_format_init(sum_time_format)

        holiday_hour_format = workbook.add_format()
        holiday_hour_format = self.__holiday_hour_format_init(holiday_hour_format)

        holiday_end_hour_format = workbook.add_format()
        holiday_end_hour_format = self.__holiday_end_hour_format_init(holiday_end_hour_format)

        summary_format = workbook.add_format()
        summary_format = self.__summary_format_init(summary_format)

        sum_value_format = workbook.add_format()
        sum_value_format = self.__sum_value_format_init(sum_value_format)

        holiday_week_num_format = workbook.add_format()
        holiday_week_num_format = self.__holiday_week_num_format_init(holiday_week_num_format)

        # Loops to fill excel
        for column in range(columns_num):
            for row in range(rows_num):
                date = self.dates[row - 2]
                # Week of year
                if column == 0:
                    # Empty cell
                    if row == 0:
                        worksheet.write(row, column, "", empty_format)
                    # Name
                    elif row == 1:
                        worksheet.write(row, column, "Tydz. roku", name_format)
                    # Values
                    elif 1 < row < rows_num - 1:

                        if row in self.holiday_rows:
                            worksheet.write(row, column, f"=WEEKNUM(B{row + 1})", holiday_week_num_format)
                        else:
                            worksheet.write(row, column, f"=WEEKNUM(B{row + 1})", week_num_format)
                    elif row == rows_num - 1:
                        down_cell_format = workbook.add_format()
                        down_cell_format.set_left(2)
                        down_cell_format.set_right(2)
                        down_cell_format.set_bottom(2)
                        if row in self.holiday_rows:
                            down_cell_format.set_bg_color(self.light_gray)
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
                        days = (date - start_excel_date).days
                        if row in self.holiday_rows:
                            holiday_day_month_year_format = workbook.add_format()
                            holiday_day_month_year_format.set_left(2)
                            holiday_day_month_year_format.set_right(2)
                            holiday_day_month_year_format.set_bg_color(self.light_gray)
                            holiday_day_month_year_format.set_num_format("d mmm yy")
                            worksheet.write(row, column, days, holiday_day_month_year_format)
                        else:
                            worksheet.write(row, column, days, day_month_year_format)
                    elif row == rows_num - 1:
                        days = (date - start_excel_date).days
                        cell_format = workbook.add_format()
                        cell_format.set_num_format("d mmm yy")
                        cell_format.set_bottom(2)
                        if row in self.holiday_rows:
                            cell_format.set_bg_color(self.light_gray)
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
                        days = (date - start_excel_date).days
                        if row in self.holiday_rows:
                            holiday_weekday_format = workbook.add_format()
                            holiday_weekday_format.set_left(2)
                            holiday_weekday_format.set_right(2)
                            holiday_weekday_format.set_num_format("ddd")
                            holiday_weekday_format.set_bg_color(self.light_gray)
                            worksheet.write(row, column, days, holiday_weekday_format)
                        else:
                            worksheet.write(row, column, days, weekday_format)
                    elif row == rows_num - 1:
                        days = (date - start_excel_date).days
                        down_cell_format = workbook.add_format()
                        down_cell_format.set_num_format("ddd")
                        down_cell_format.set_left(2)
                        down_cell_format.set_right(2)
                        down_cell_format.set_bottom(2)
                        if row in self.holiday_rows:
                            down_cell_format.set_bg_color(self.light_gray)
                            worksheet.write(row, column, days, down_cell_format)
                        else:
                            worksheet.write(row, column, days, down_cell_format)

                # Empty column
                elif column == 3:
                    worksheet.write(row, column, "", empty_format)

                # 1st column of person
                elif column > 3 and column % 3 == 1:
                    if row == 0:
                        worksheet.merge_range(row, column, row, column + 2, self.people[int(column / 3 - 1)], person_format)
                    elif row == 1:
                        worksheet.write(row, column, "Od", person_info_format)
                    elif row < rows_num - 1:
                        if row in self.holiday_rows:
                            worksheet.write(row, column, "", holiday_hour_format)
                        else:
                            worksheet.write(row, column, "", hour_format)
                    else:
                        if row in self.holiday_rows:
                            worksheet.write(row, column, "", holiday_end_hour_format)
                        else:
                            worksheet.write(row, column, "", end_hour_format)

                # 2nd column of person
                elif column > 3 and column % 3 == 2:
                    if row == 1:
                        worksheet.write(row, column, "Do", person_info_format)
                    elif 1 < row < rows_num - 1:
                        if row in self.holiday_rows:
                            worksheet.write(row, column, "", holiday_hour_format)
                        else:
                            worksheet.write(row, column, "", hour_format)
                    elif row == rows_num - 1:
                        if row in self.holiday_rows:
                            worksheet.write(row, column, "", holiday_end_hour_format)
                        else:
                            worksheet.write(row, column, "", end_hour_format)

                        worksheet.write(rows_num, column, "SUMA:", summary_format)

                # 3rd column of person
                elif column > 3 and column % 3 == 0:
                    if row == 1:
                        worksheet.write(row, column, "Suma", person_info_format)
                    elif 1 < row < rows_num - 1:
                        if row in self.holiday_rows:
                            worksheet.write(row, column, "", holiday_hour_format)
                        else:
                            first_column = self.column_to_char(column - 2)
                            second_column = self.column_to_char(column - 1)
                            worksheet.write(row, column, f"={second_column}{row + 1} - {first_column}{row + 1}",
                                            sum_time_format)
                    elif row == rows_num - 1:
                        if row in self.holiday_rows:
                            worksheet.write(row, column, "", holiday_end_hour_format)
                        else:
                            first_column = self.column_to_char(column - 2)
                            second_column = self.column_to_char(column - 1)
                            down_cell_format = workbook.add_format()
                            down_cell_format.set_num_format("h:mm")
                            down_cell_format.set_right(2)
                            down_cell_format.set_bottom(2)
                            worksheet.write(row, column, f"={second_column}{row + 1} - {first_column}{row + 1}",
                                            down_cell_format)
                        sum_column = self.column_to_char(column)
                        worksheet.write(rows_num, column, f"=SUM({sum_column}3:{sum_column}{row + 1})",
                                        sum_value_format)
        workbook.close()