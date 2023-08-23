#SINGLE WINDOW EXCEL GENERATOR APP

###Links to topics:
* [ExcelGenerator](##ExcelGenerator)
* [GUI](##GUI)
* [Holidays](##Holidays)



##ExcelGenerator
Class is responsible for creating excel form with list of people and marking all weekend days and holidays. All data is getting from GUI and is initialized in GUI class.

##GUI
Class is responsible for user interface and parsing data to ExcelGenerator class. GUI class is using PySimplyGUI library to generate gui.

##Holidays
Holidays is class which has offline access to all holidays dates in Poland until 2090. All data was downloaded by [azureml.opendatasets](https://learn.microsoft.com/en-us/azure/open-datasets/) and class [PublicHolidays](https://learn.microsoft.com/en-us/azure/open-datasets/dataset-public-holidays?tabs=azureml-opendatasets)
