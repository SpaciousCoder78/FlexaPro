#importing packages
from docx2pdf import convert
from pdf2docx import Converter
import os

#docx to pdf function
def docx2pdf():
    print("---------------------------Docx to PDF-----------------------------")
    #asking the user for path
    pathfile=input("Enter the file path of the folder: ")
    #locating the file
    path=os.listdir(pathfile)
    for i,files in enumerate(path):
        if files[-5:]==".docx":
            #converting the file
            finalpath=input("Enter the path of the file:  ")
            convert(f"{finalpath}",f"C:/Users/aryan/Desktop/docx2pdf/test.pdf")
            print("File successfully converted")

#pdf to docx function
def pdf2docx():
    print("--------------------------PDF to docx------------------------------")
    #source file name
    pdf_file=input("Enter PDF File location: ")
    
    #output file directory
    docx_file=input("Enter docx file location and name: ")

    #converting the files
    cv=Converter(pdf_file)
    cv.convert(docx_file)
    cv.close()
    print("File successfully converted")

#menu module
menyoo=1
while menyoo==1:
    print("------------------------FlexaPro-----------------------------")
    print("1. Docx to pdf")
    print("2. Pdf to docx")
    print("3.Exit")
    menuoption=int(input("Enter your choice: "))
    if menuoption==1:
        docx2pdf()
    elif menuoption==2:
        pdf2docx()
    elif menuoption==3:
        pass
