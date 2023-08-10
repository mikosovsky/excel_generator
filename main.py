import ExcelGenerator
import PySimpleGUI as sg
import os

layout = [
    [
        sg.Text("Choose file:", background_color="#FFFFFF", text_color="#000000"),
        sg.Input(size=(50, 1), enable_events=True, key="-FILE-", background_color="#FFFFFF", text_color="#000000"),
        sg.FileBrowse(button_color="#C7C7CC", file_types=(("Text Files", "*.txt")), target="-FILE-")
    ],
    [
        sg.Multiline(
            enable_events=True, size=(500, 20), key = "-FILE LIST-", text_color="#000000", background_color="#FFFFFF"
        )
    ],
    [
        sg.Text("Choose month: ", background_color="#FFFFFF", text_color="#000000"),
        sg.Input(size = (25, 1), enable_events=True, key="-CAL-", background_color="#FFFFFF",text_color="#000000"),
        sg.CalendarButton(button_text= "Calendar",format="%m-%Y",button_color="#C7C7CC", target="-CAL-")
    ],
    [
      sg.Button(button_text="Save & Generate", button_color="#C7C7CC", key="-READY-")
    ],
    [
        sg.ProgressBar(max_value=100, orientation="h",expand_x=True, size=(20,20), key="-PBAR-")
    ]
]

window = sg.Window(title="Excel generator", layout=layout, background_color="#FFFFFF", size=(600, 500), element_justification="c", element_padding=15)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break

    if event == "-FILE-":
        file = values["-FILE-"]
        try:
            with open(file, "r+") as names_file:
                names_list = names_file.read()
                window["-FILE LIST-"].update(names_list)
        except FileNotFoundError:
            window["-FILE LIST-"].update("")

# # Requesting for month and year
# excel_generator = ExcelGenerator.ExcelGenerator()
# month = int(input("Podaj miesiąc: "))
# year = int(input("Podaj rok: "))
#
#
# excel_generator.change_month(month, year)
#
# # Ask for quantity of people in excel
# people_quantity = int(input("Podaj ilość osób do dodania w excelu: "))
# people = []
#
# # Requesting for people and adding them to list
# for i in range(people_quantity):
#     person = input(f"Podaj dane {i+1}. osoby: ")
#     people.append(person)
#
# excel_generator.change_people_list(people)
# excel_generator.generate_excel()
