import PySimpleGUI as sg

#Theme
sg.theme("Dark2")
#inside window
layout = [
    [sg.Text("Welcome to the B.O.M Sniffer")],
    [sg.Text("Enter Syspro File"), sg.Button("Sniff File")],
    [sg.Text("Enter B.O.M File"), sg.Button("Sniff File")]
]

#Make the window
window = sg.Window("B.O.M Sniffer", layout)

#event loop to process "events' and get the 'Values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Cancel":
        break



window.close()
