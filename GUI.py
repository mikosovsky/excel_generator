import PySimpleGUI as sg
import ExcelGenerator
import os


class GUI:
    # Initiator of class. Creating whole layout, window and ExcelGenerator class
    def __init__(self):
        self.layout = [
            [
                sg.Text("Choose file:", background_color="#FFFFFF", text_color="#000000"),
                sg.Input(size=(50, 1), enable_events=True, key="-FILE-", background_color="#FFFFFF",
                         text_color="#000000"),
                sg.FileBrowse(button_color="#C7C7CC", file_types=(("Text Files", "*.txt")), target="-FILE-")
            ],
            [
                sg.Multiline(
                    enable_events=True, size=(500, 20), key="-FILE LIST-", text_color="#000000",
                    background_color="#FFFFFF"
                )
            ],
            [
                sg.Text("Choose month: ", background_color="#FFFFFF", text_color="#000000"),
                sg.Input(size=(25, 1), enable_events=True, key="-CAL-", background_color="#FFFFFF",
                         text_color="#000000"),
                sg.CalendarButton(button_text="Calendar", format="%m-%Y", button_color="#C7C7CC", target="-CAL-")
            ],
            [
                sg.Text("Choose location to save excel: ", background_color="#FFFFFF", text_color="#000000"),
                sg.Input(size=(30, 1), enable_events=True, key="-FOLDER-", background_color="#FFFFFF",
                         text_color="#000000"),
                sg.FolderBrowse(button_color="#C7C7CC", target="-FOLDER-")
            ],
            [
                sg.Button(button_text="Save & Generate", button_color="#C7C7CC", key="-READY-")
            ],
            [
                sg.ProgressBar(max_value=100, orientation="h", expand_x=True, size=(20, 20), key="-PBAR-")
            ]
        ]

        self.window = sg.Window(title="Excel generator", layout=self.layout, background_color="#FFFFFF", size=(600, 550),
                           element_justification="c", element_padding=15)

        self.excel_generator = ExcelGenerator.ExcelGenerator()

    # Function to start showing window
    def window_loop(self):
        # Window loop
        while True:
            event, values = self.window.read()
            # Execute program after close window
            if event == sg.WIN_CLOSED:
                break

            # Path textbox edited
            if event == "-FILE-":
                file = values["-FILE-"]
                try:
                    with open(file, "r+") as names_file:
                        names_list = names_file.read()
                        self.window["-FILE LIST-"].update(names_list)
                except FileNotFoundError:
                    self.window["-FILE LIST-"].update("")

            # Button Save & Generate was clicked and generating excel
            if event == "-READY-":
                people = values["-FILE LIST-"]

                if people != "":

                    # Getting list of people in textBox "-FILE LIST-" and spliting it in 3 ways by new lines, commas and semicolons
                    people = people.split("\n")
                    people_backup = []
                    for person in people:
                        if person.find(";"):
                            person_splited = person.split(";")
                            for ps in person_splited:
                                if ps.find(","):
                                    person_splited1 = ps.split(",")
                                    people_backup += person_splited1
                                else:
                                    people_backup.append(ps)

                        elif person.find(","):
                            person_splited = person.split(",")
                            people_backup += person_splited
                        else:
                            people_backup.append(person)


                    # Clearing people list by deleting useless spaces
                    people = people_backup
                    people_backup = []
                    for person in people:
                        while person[0] == " ":
                            person = person.replace(" ", "", 1)
                        people_backup.append(person)

                    people = people_backup
                    people_file = ""
                    for person in people:
                        people_file +=f"{person}\n"

                    # Saving list of people in file (only if something is in "-FILE-" input)
                    if values["-FILE-"] != "":
                        file = values["-FILE-"]
                        with open(file, "w+") as names_file:
                            names_file.write(people_file)

                    # Setting people list
                    self.excel_generator.change_people_list(people)

                    # Getting date from calendar
                    date = values["-CAL-"]

                    if date != "" and len(date.split("-")) == 2:
                        date = date.split("-")
                        # Setting month and year in excel_generator
                        self.excel_generator.change_month(int(date[0]), int(date[1]))
                        if values["-FOLDER-"] != "":
                            path = values["-FOLDER-"]
                            if os.path.exists(path):
                                #Generating excel
                                self.excel_generator.generate_excel(path)
                            else:
                                sg.popup_error("Choose right path to save excel!")
                        else:
                            sg.popup_error("Set path to save excel!")
                    else:
                        sg.popup_error("Choose right date!")
                else:
                    sg.popup_error("Choose the file or fill up textbox with content!")