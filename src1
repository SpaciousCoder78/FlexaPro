import PySimpleGUI as sg
from docx2pdf import convert
from pdf2docx import Converter
import os


def docx2pdf():
    layout1 = [
               [sg.Text("Enter file path of the folder",size=(15,3)),sg.InputText(key="folder location")],
               [sg.Text("Enter file path of the docx file with name",size=(15,3)),sg.InputText(key="file location")],
               [sg.Text("Enter pdf file location",size=(15,3)),sg.InputText(key="pdf file location")],
               [sg.Button("submit"),sg.Button('Clear'),sg.Button("exit")]
            ]
    window1 = sg.Window("docx to pdf converter",layout1)
    while True:
        event,value = window1.read()
        if event=='exit':
            break
        if event=="Clear":
            for key in values:
                 window1[key]('')
        if event=="submit":
            pathfile = value["folder location"]
            path=os.listdir(pathfile)
            finalpath = value["file location"]
            pdfpath = value["pdf file location"]
            for i,files in enumerate(path):
                if files[-5:]==".docx":
                    convert(f"{finalpath}",f"{pdfpath}")
            sg.popup("Conversion complete!")
    window1.close()
        
        
def pdf2docx():
    layout2= [
               [sg.Text("Enter file path of pdf file",size=(15,3)),sg.InputText(key="file location")],
               [sg.Text("Enter docx file location with name",size=(15,3)),sg.InputText(key="docx file location")],
               [sg.Button("submit"),sg.Button('Clear'),sg.Button("exit")]
            ]
    window2 = sg.Window("pdf to docx converter",layout2)
    while True:
       event,value = window2.read()
       if event=='exit':
          break
       if event=="Clear":
          for key in values:
              window1[key]('')
       if event=="submit":
           pdfpath = value["file location"]
           docxpath = value["docx file location"]
           cv = Converter(pdfpath)
           cv.convert(docxpath)
           cv.close()
           sg.popup("Conversion complete!")
    window2.close()

sg.theme('DarkTeal9')

layout = [[sg.Text("Click 1 to convert docx to pdf and 2 to convert pdf to docx")],[sg.Button("1"),sg.Button("2")],[sg.Exit()]]

window = sg.Window("format converter",layout)

while True:
    events,values = window.read()
    if events== "1":
        docx2pdf()
    if events== "2":
        pdf2docx()
    if events == "Exit" or events == sg.WIN_CLOSED:
        break
window.close()

