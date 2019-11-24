import PySimpleGUI as sg

def main():
    layout = [
        [sg.Text("Welcome to PDFWork")],
        [sg.Text("What do you want to do?")],
        [sg.Button('Merge'), sg.Button('Split'), sg.Button('Word2PDF'), sg.Button('PDF2Image')]
    ]

    # window = sg.Window('PDF Work', layout)
    return sg.Window('PDF Work', layout)


