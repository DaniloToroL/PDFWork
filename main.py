# Python
import os
from subprocess import Popen, PIPE

# Externals
import PySimpleGUI as sg
from PyPDF2 import PdfFileReader, PdfFileWriter
from PIL import Image

# PDFWork
import gui

docto = os.path.join("bin", "docto.exe")
global folder_path

# Utils functions

def getFolder(path):

    if os.sep in path:
        folder = os.sep.join(path.split(os.sep)[:-1])
    else:
        folder = "/".join(path.split("/")[:-1])
    
    return folder

def openFolder(path):

    path=os.path.realpath(path)
    os.startfile(path)


# Functions

def wordToPdf(data):
    """Convert Word files to PDF

        Parameters
        ----------
        data : dict
            The data obtained by the user
            {from: Word path, to: Path and name of PDF}

        Returns
        -------
        msg : str
            Success or failed message
    """
    
    try:
        from_ = data["from"]
        to_ =  data["to"]

        folder_path = getFolder(to_)
        
        options = " -f " + from_ + " -O " + to_ + " -T wdFormatPDF"
        run_ = docto +  options
        con = Popen(run_, shell=True, stdout=PIPE, stderr=PIPE)
        a = con.communicate()

        return "Success", folder_path
    
    except Exception as e:
        return "Failed", e
        


def mergePDF(data):
    
    output = data["output"]
    pdfs = data["input"].split(";")

    folder_path = getFolder(output)
     
    if not output.endswith(".pdf"):
        output += ".pdf"

    pdf_writer = PdfFileWriter()
    try:
        for pdf in pdfs:
            pdf_reader = PdfFileReader(pdf)
            for page in range(pdf_reader.getNumPages()):
                pdf_writer.addPage(pdf_reader.getPage(page))

        with open(output, 'wb') as fh:
            pdf_writer.write(fh)

        return "Success", folder_path
    except Exception as e:
        return "Failed", e
    

def PDF2Image(data):
    pass


def splitPDF(data):
    
    path = data["path"]
    folder_path = data["folder"]
    
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

    try:
        while start < page_number:
            end = start + step
            if end > page_number:
                end = page_number
            pdf_writer = PdfFileWriter()
            for page in range(start, end):
                pdf_writer.addPage(pdf.getPage(page))
            output_filename = '{}_page_{}.pdf'.format(fname, start + 1)
            start = end
            with open(folder + os.sep + output_filename, 'wb') as out:
                pdf_writer.write(out)
                print('Created: {}'.format(output_filename))
        
        return "Success", folder_path
    except Exception as e:
        return "Failed", e

def imageToPdf(data):

    pdf = data["output"]
    files = data["input"].split(";")

    folder_path = getFolder(pdf)

    if not pdf.endswith(".pdf"):
        output += ".pdf"

    try:
        images = []
        for file_ in files:
            if file_.endswith(".jpg"):
                jpg = Image.open(file_)
                images.append(jpg)

        print(images)

        images[0].save(pdf, save_all=True)
        for image in images[1:]:
            image.save(pdf, append=image)

        return "Success", folder_path
    except Exception as e:
        return "Failed", e



if __name__ == '__main__':
    window = gui.Main()
    while True:
        event, values = window.read()
        print(event, values)

        # if event is OK, on the navigation we are at the second level
        # if it is not OK, we are on the second level
        if event is not "Ok":
            option = event
        if event is None or event == 'Exit':
            break
        
        if event in ['Split', 'Merge', 'Word2PDF', 'PDF2Image', 'Image2PDF']:
            layout = gui.LayoutHandler()[event]
        elif event == "Ok":
            if option == "Split":
                event, msg = splitPDF(values)
                layout = gui.LayoutHandler(msg)[event]

            elif option == "Merge":
                event, msg = mergePDF(values)
                layout = gui.LayoutHandler(msg)[event]

            elif option == "Word2PDF":
                event, msg = wordToPdf(values)
                layout = gui.LayoutHandler(msg)[event]
            elif option == "Image2PDF":
                event, msg = imageToPdf(values)
                layout = gui.LayoutHandler(msg)[event]

        elif event == "Open Folder":
            openFolder(values["path"])
            layout = [
                [sg.Text("Welcome to PDFWork")],
                [sg.Text("What do you want to do?")],
                [sg.Button('Merge'),sg.Button('Split'), sg.Button('Word2PDF'), sg.Button('PDF2Image'), sg.Button('Image2PDF')]
            ]

        else:
            print("Error")

        window.Close()
        window = sg.Window('PDF Work', layout)

    window.Close()

