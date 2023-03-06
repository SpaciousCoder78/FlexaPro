import PySimpleGUI as sg
from docx2pdf import convert
from pdf2docx import Converter
import os


def docx2pdf():
    #creating a layout with several buttons and inputs to get the file path of folder,docx file and final pdf file
    layout1 = [
               [sg.Text("Enter file path of the folder",size=(15,3)),sg.InputText(key="folder location")],
               [sg.Text("Enter file path of the docx file with name",size=(15,3)),sg.InputText(key="file location")],
               [sg.Text("Enter pdf file location",size=(15,3)),sg.InputText(key="pdf file location")],
               [sg.Button("submit"),sg.Button('Clear'),sg.Button("exit")]
            ]
    #creating a window with the above layout
    window1 = sg.Window("docx to pdf converter",layout1)
    while True:
        #the entered information is stored in values variable as a dictionary with keys as folder location,file location etc. as specified above and values as what ever u enter
        event,value = window1.read()
        if event=='exit':
            break
        #this clears the information entered wrongly
        if event=="Clear":
            for key in values:
                 window1[key]('')
        #converting docx file to pdf file.
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
    #creating a layout with several inputs to get the file path of the pdf file and final docx file
    layout2= [
               [sg.Text("Enter file path of pdf file",size=(15,3)),sg.InputText(key="file location")],
               [sg.Text("Enter docx file location with name",size=(15,3)),sg.InputText(key="docx file location")],
               [sg.Button("submit"),sg.Button('Clear'),sg.Button("exit")]
            ]
    #creating a window with the above layout
    window2 = sg.Window("pdf to docx converter",layout2)
    while True:
       #the entered information is stored in values variable as a dictionary with keys as folder location,file location etc. as specified above and values as what ever u enter
       event,value = window2.read()
       if event=='exit':
          break
       #this clears the information entered wrongly
       if event=="Clear":
          for key in values:
              window1[key]('')
       #converting pdf file to docx file
       if event=="submit":
           pdfpath = value["file location"]
           docxpath = value["docx file location"]
           cv = Converter(pdfpath)
           cv.convert(docxpath)
           cv.close()
           sg.popup("Conversion complete!")
    window2.close()

sg.theme('DarkTeal9')
#creating a layout with some buttons to choose conversion type
layout = [[sg.Button("docx to pdf")],[sg.Button("pdf to docx")],[sg.Exit()]]
#creating a window with the layout
window = sg.Window("format converter",layout)
#While loop allows user to convert files as many times as they want until they click exit
while True:
    events,values = window.read()
    if events== "docx to pdf":
        docx2pdf()
    if events== "pdf to docx":
        pdf2docx()
    #sg.WIN_CLOSED is the function used to close the file when the x button at the top right corner is clicked
    if events == "Exit" or events == sg.WIN_CLOSED:
        break
window.close()
