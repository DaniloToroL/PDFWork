import PySimpleGUI as sg 
from PyPDF2 import PdfFileReader, PdfFileWriter
import os
from subprocess import Popen, PIPE
docto = os.path.join("bin", "docto.exe")

def wordToPdf(data):

    print(data)
    from_ = data["from"]
    to_ =  data["to"]
    
    options = " -f " + from_ + " -O " + to_ + " -T wdFormatPDF"
    run_ = docto +  options
    con = Popen(run_, shell=True, stdout=PIPE, stderr=PIPE)
    a = con.communicate()


def mergePDF(data):
    output = data["output"]
    pdfs = data["input"].split(";")

    pdf_writer = PdfFileWriter()

    for pdf in pdfs:
        print(pdf)
        pdf_reader = PdfFileReader(pdf)
        for page in range(pdf_reader.getNumPages()):
            pdf_writer.addPage(pdf_reader.getPage(page))

    with open(output, 'wb') as fh:
        pdf_writer.write(fh)
    

def PDF2Image(data):
    pass


def splitPDF(data):
    path = data["path"]
    fname = os.path.splitext(os.path.basename(data["path"]))[0]
    pdf = PdfFileReader(path)
    page_number = pdf.getNumPages()
    start = 0

    if data["check_step"]:
        step = data["in_step"]
    elif data["part_step"]:
        step = page_number / int(data["in_parts"])
        if step == 4.5:
            step = 5
        else:
            step = round(step)
    while start < page_number:
        end = start + step
        if end > page_number:
            end = page_number
        pdf_writer = PdfFileWriter()
        for page in range(start, end):
            pdf_writer.addPage(pdf.getPage(page))
        output_filename = '{}_page_{}.pdf'.format(fname, start + 1)
        start = end
        with open(data["folder"] + os.sep + output_filename, 'wb') as out:
            pdf_writer.write(out)
            print('Created: {}'.format(output_filename))


if __name__ == '__main__':

    layout = [
        [sg.Text("Welcome to PDFWork")],
        [sg.Text("What do you want to do?")],
        [sg.Button('Merge'),sg.Button('Split'), sg.Button('Word2PDF'), sg.Button('PDF2Image')]
    ]
    window = sg.Window('PDF Work', layout)

    while True:
        event, values = window.read()
        print(event, values)

        if event is not "Ok":
            option = event
        if event is None or event == 'Exit':
            break
        if event == "Split":
            layout = [
                [sg.Text('Select PDF')],
                [sg .Input(key="path"), sg.FileBrowse()],
                [sg.Text('Settings')],
                [sg.Checkbox('Step', key='check_step'), sg.Input(key='in_step')],
                [sg.Checkbox('Parts', key='part_step'), sg.Input(key='in_parts')],
                [sg.Text('Save')],
                [sg .Input(key="folder"), sg.FileSaveAs(file_types=(("Pdf", "*.pdf"),))],
                [sg.Ok()]
            ]
        elif event == "Merge":
            layout = [
                [sg.Text('Select Folder')],
                [sg .Input(key="input"), sg.FilesBrowse()],
                [sg.Text('Save')],
                [sg .Input(key="output"), sg.FileSaveAs(file_types=(("Pdf", "*.pdf"),))],
                [sg.Ok()]
            ]
        elif event == "Word2PDF":
            layout = [
                [sg.Text('Select Folder')],
                [sg .Input(key="from"), sg.FilesBrowse(file_types=(("Word", "*.docx"),))],
                [sg.Text('Save')],
                [sg .Input(key="to"), sg.FileSaveAs(file_types=(("Pdf", "*.pdf"),))],
                [sg.Ok()]
            ]
        elif event == "Ok":
            if option == "Split":
                splitPDF(values)
                layout = [
                    [sg.Text("Success")],
                    [sg.Text("What do you want to do?")],
                    [sg.Button('Merge'),sg.Button('Split'), sg.Button('Word2PDF'), sg.Button('PDF2Image')]
                ]

            elif option == "Merge":
                mergePDF(values)
                layout = [
                    [sg.Text("Success")],
                    [sg.Text("What do you want to do?")],
                    [sg.Button('Merge'),sg.Button('Split'), sg.Button('Word2PDF'), sg.Button('PDF2Image')]
                ]

            elif option == "Word2PDF":
                wordToPdf(values)
                layout = [
                    [sg.Text("Success")],
                    [sg.Text("What do you want to do?")],
                    [sg.Button('Merge'),sg.Button('Split'), sg.Button('Word2PDF'), sg.Button('PDF2Image')]
                ]

        window.Close()
        window = sg.Window('PDF Work', layout)

    window.Close()

