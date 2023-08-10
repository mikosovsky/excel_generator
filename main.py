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
        sg.Text("Choose location to save excel: ", background_color="#FFFFFF", text_color="#000000"),
        sg.Input(size=(30, 1), enable_events=True, key="-FOLDER-", background_color="#FFFFFF", text_color="#000000"),
        sg.FolderBrowse(button_color="#C7C7CC", target="-FOLDER-")
    ],
    [
      sg.Button(button_text="Save & Generate", button_color="#C7C7CC", key="-READY-")
    ],
    [
        sg.ProgressBar(max_value=100, orientation="h",expand_x=True, size=(20,20), key="-PBAR-")
    ]
]

window = sg.Window(title="Excel generator", layout=layout, background_color="#FFFFFF", size=(600, 550), element_justification="c", element_padding=15)

excel_generator = ExcelGenerator.ExcelGenerator()

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

    if event == "-READY-":
        people = values["-FILE LIST-"]

        if people != "":

            if values["-FILE-"] != "":
                file = values["-FILE-"]
                with open(file, "w+") as names_file:
                    names_file.write(people)

            people = people.split("\n")
            excel_generator.change_people_list(people
                                               )
            date = values["-CAL-"]

            if date != "" and len(date.split("-")) == 2:
                date = date.split("-")
                excel_generator.change_month(int(date[0]), int(date[1]))
                if values["-FOLDER-"] != "":
                    path = values["-FOLDER-"]
                    if os.path.exists(path):
                        excel_generator.generate_excel(path)
                    else:
                        sg.popup_error("Choose right path to save excel!")
                else:
                    sg.popup_error("Set path to save excel!")
            else:
                sg.popup_error("Choose right date!")
        else:
            sg.popup_error("Choose the file or fill up textbox with content!")
