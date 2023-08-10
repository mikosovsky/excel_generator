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

    def change_people_list(self, people: [str]):
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
    def __weekday_format_init(weekday_format: xlsxwriter.Workbook.Format) -> xlsxwriter.Workbook.Format:
        weekday_format.set_num_format("ddd")
        weekday_format.set_left(2)
        weekday_format.set_right(2)
        return weekday_format

    @staticmethod
    def __week_num_format_init(week_num_format: xlsxwriter.Workbook.Format) -> xlsxwriter.Workbook.Format:
        week_num_format.set_left(2)
        week_num_format.set_right(2)
        return week_num_format

    @staticmethod
    def __day_month_year_format_init(day_month_year_format: xlsxwriter.Workbook.Format) -> xlsxwriter.Workbook.Format:
        day_month_year_format.set_num_format("d mmm yy")
        return day_month_year_format

    def __empty_format_init(self, empty_format: xlsxwriter.Workbook.Format) -> xlsxwriter.Workbook.Format:
        empty_format.set_bg_color(self.indigo)
        return empty_format

    @staticmethod
    def __name_format_init(name_format: xlsxwriter.Workbook.Format) -> xlsxwriter.Workbook.Format:
        name_format.set_left(5)
        name_format.set_right(5)
        name_format.set_bottom(5)
        name_format.set_top(5)
        return name_format

    def __person_format_init(self, person_format: xlsxwriter.Workbook.Format) -> xlsxwriter.Workbook.Format:
        person_format.set_left(5)
        person_format.set_right(5)
        person_format.set_bottom(5)
        person_format.set_top(5)
        person_format.set_bold(True)
        person_format.set_bg_color(self.light_blue)
        person_format.set_align("center")
        return person_format

    @staticmethod
    def __person_info_format_init(person_info_format: xlsxwriter.Workbook.Format) -> xlsxwriter.Workbook.Format:
        person_info_format.set_left(5)
        person_info_format.set_right(5)
        person_info_format.set_bottom(5)
        person_info_format.set_top(5)
        person_info_format.set_align("center")
        return person_info_format

    @staticmethod
    def __hour_format_init(hour_format: xlsxwriter.Workbook.Format) -> xlsxwriter.Workbook.Format:
        hour_format.set_left(2)
        hour_format.set_right(2)
        hour_format.set_num_format("h:mm")
        return hour_format

    @staticmethod
    def __end_hour_format_init(end_hour_format: xlsxwriter.Workbook.Format) -> xlsxwriter.Workbook.Format:
        end_hour_format.set_left(2)
        end_hour_format.set_right(2)
        end_hour_format.set_bottom(2)
        end_hour_format.set_num_format("h:mm")
        return end_hour_format

    @staticmethod
    def __sum_time_format_init(sum_time_format: xlsxwriter.Workbook.Format) -> xlsxwriter.Workbook.Format:
        sum_time_format.set_right(2)
        sum_time_format.set_num_format("h:mm")
        return sum_time_format

    def __holiday_hour_format_init(self, holiday_hour_format: xlsxwriter.Workbook.Format) -> xlsxwriter.Workbook.Format:
        holiday_hour_format.set_left(2)
        holiday_hour_format.set_right(2)
        holiday_hour_format.set_bg_color(self.light_gray)
        holiday_hour_format.set_align("center")
        return holiday_hour_format

    def __holiday_end_hour_format_init(self, holiday_end_hour_format: xlsxwriter.Workbook.Format) -> xlsxwriter.Workbook.Format:
        holiday_end_hour_format.set_left(2)
        holiday_end_hour_format.set_right(2)
        holiday_end_hour_format.set_bottom(2)
        holiday_end_hour_format.set_bg_color(self.light_gray)
        holiday_end_hour_format.set_align("center")
        return holiday_end_hour_format

    @staticmethod
    def __summary_format_init(summary_format: xlsxwriter.Workbook.Format) -> xlsxwriter.Workbook.Format:
        summary_format.set_left(2)
        summary_format.set_bottom(2)
        summary_format.set_align("right")
        return summary_format

    @staticmethod
    def __sum_value_format_init(sum_value_format: xlsxwriter.Workbook.Format) -> xlsxwriter.Workbook.Format:
        sum_value_format.set_bottom(2)
        sum_value_format.set_right(2)
        sum_value_format.set_num_format("h:mm")
        return sum_value_format

    def generate_excel(self):
        # Creating excel file
        month_name = self.dates[0].strftime("%B")
        workbook_name = f"{month_name}.xlsx"
        workbook = xlsxwriter.Workbook(workbook_name)
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
