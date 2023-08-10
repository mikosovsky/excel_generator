import ExcelGenerator

# Requesting for month and year
excel_generator = ExcelGenerator.ExcelGenerator()
month = int(input("Podaj miesiąc: "))
year = int(input("Podaj rok: "))


excel_generator.change_month(month, year)

# Ask for quantity of people in excel
people_quantity = int(input("Podaj ilość osób do dodania w excelu: "))
people = []

# Requesting for people and adding them to list
for i in range(people_quantity):
    person = input(f"Podaj dane {i+1}. osoby: ")
    people.append(person)

excel_generator.change_people_list(people)
excel_generator.generate_excel()
