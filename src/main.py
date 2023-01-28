#importing packages
from docx2pdf import convert
import os

#docx to pdf function
def docx2pdf():
    #asking the user for path
    pathfile=input("Enter the file path: ")
    #locating the file
    path=os.listdir(pathfile)
    for i,files in enumerate(path):
        if files[-5:]==".docx":
            #converting the file
            convert(f"pathfile",f"C:/Users/aryan/Desktop/docx2pdf/test.pdf")

#menu module
menyoo=1
while menyoo==1:
    print("------------------------Deimos Files and Images-----------------------------")
    print("1. Docx to pdf")
    print("2.Exit")
    menuoption=int(input("Enter your choice: "))
    if menuoption==1:
        docx2pdf()
    elif menuoption==2:
        pass
