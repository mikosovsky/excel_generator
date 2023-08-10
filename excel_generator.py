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
        self.holiday_list = []

    def change_month(self, month: int, year: int):
        num_days = calendar.monthrange(year, month)[1]
        # Making list of days in month
        for day in range(1, num_days + 1):
            date = datetime.datetime(year, month, day)
            self.dates.append(date)

    def change_people_list(self, people: [str]):
        self.people = people

    def __make_holidays_list(self):
        # Download holidays from Microsoft server
        num_end_date = len(self.dates) - 1
        end_date = self.dates[num_end_date]
        start_date = self.dates[0]
        hol = PublicHolidays(country_or_region='PL', start_date=start_date, end_date=end_date)
        hol_df = hol.to_pandas_dataframe()
        self.holiday_list = hol_df["date"].tolist()
        # Making list of rows with holiday in month
        for date in self.dates:
            if date.weekday() == 5 or date.weekday() == 6 or date in self.holiday_list:
                day = date.day
                self.holiday_rows.append(day + 1)
