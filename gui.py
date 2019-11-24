import PySimpleGUI as sg

def Main():
    layout = [
        [sg.Text("Welcome to PDFWork")],
        [sg.Text("What do you want to do?")],
        [sg.Button('Merge'), sg.Button('Split'), sg.Button('Word2PDF'), sg.Button('PDF2Image')]
    ]

    # window = sg.Window('PDF Work', layout)
    return sg.Window('PDF Work', layout)



def Split():
    return [
        [sg.Text('Select PDF')],
        [sg.Input(key="path"), sg.FileBrowse()],
        [sg.Text('Settings')],
        [sg.Checkbox('Step', key='check_step'), sg.Input(key='in_step')],
        [sg.Checkbox('Parts', key='part_step'), sg.Input(key='in_parts')],
        [sg.Text('Save')],
        [sg.Input(key="folder"), sg.FileSaveAs(file_types=(("Pdf", "*.pdf"),))],
        [sg.Ok()]
    ]

def Merge():
    return [
        [sg.Text('Select Folder')],
        [sg .Input(key="input"), sg.FilesBrowse()],
        [sg.Text('Save')],
        [sg .Input(key="output"), sg.FileSaveAs(file_types=(("Pdf", "*.pdf"),))],
        [sg.Ok()]
    ]

def Word2PDF():
    return [
        [sg.Text('Select Folder')],
        [sg.Input(key="from"), sg.FilesBrowse(file_types=(("Word", "*.docx"),))],
        [sg.Text('Save')],
        [sg.Input(key="to"), sg.FileSaveAs(file_types=(("Pdf", "*.pdf"),))],
        [sg.Ok()]
    ]

def PDF2Image():
    return [
        [sg.Text('Not yet implemented')],
        [sg.Ok()]
    ]

def LayoutHandler():
    return {
        'Split': Split(),
        'Merge': Merge(),
        'Word2PDF': Word2PDF(),
        'PDF2Image': PDF2Image()
    }


if __name__ == '__main__':
    pass


