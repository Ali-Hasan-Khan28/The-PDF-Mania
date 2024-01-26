from win32com import client
from flask import Flask, render_template, request, redirect, url_for, send_file
import os
import sys
import PyPDF2
import img2pdf
from pdf2image import convert_from_path
import pythoncom
from docx2pdf import convert
from PIL import Image
from fpdf import FPDF
from docx2pdf import convert
import os
from pdf2docx import Converter

app = Flask(__name__)

# @app.route("/")

# def hello_world():
#     sum=3+4
#     return render_template('index.html')


@app.route('/')
@app.route("/index1", methods=["GET", "POST"])
def index():
    print("fdghdhfdbcvbdfbdbn")
    if request.method == "POST":
        # storing image path
        VariableName = request.form['filename']
        print(VariableName)

        img_path = "C:/Users/HP/OneDrive - Higher Education Commission/Desktop/" + VariableName

        # storing pdf path
        pdf_path = "C:/Users/HP/OneDrive - Higher Education Commission/Desktop/" + \
            VariableName+".pdf"

        # opening image
        image = Image.open(img_path)

        # converting into chunks using img2pdf
        pdf_bytes = img2pdf.convert(image.filename)

        # opening or creating pdf file
        file = open(pdf_path, "wb")

        # writing pdf files with chunks
        file.write(pdf_bytes)

        # closing image file
        image.close()

        # closing pdf file
        file.close()

        # output
        print("Successfully made pdf file")
    return render_template("index1.html")


@app.route("/index2", methods=["GET", "POST"])
def Excel():
    print("dsdfds")
    if request.method == "POST":
        # Import Module
        # # Open Microsoft Excel
        pdf = FPDF()

        # Add a page
        pdf.add_page()

        # set style and size of font
        # that you want in the pdf
        pdf.set_font("Arial", size=15)

        # open the text file in read mode
        VariableName = request.form['filename']
        print(VariableName)
        f = open("C:/Users/HP/OneDrive - Higher Education Commission/Desktop/"+VariableName, "r")

        # insert the texts in pdf
        for x in f:
            pdf.cell(200, 10, txt=x, ln=1, align='C')

        # save the pdf with name .pdf
        pdf.output("C:/Users/HP/OneDrive - Higher Education Commission/Desktop/"+VariableName+".pdf")
    return render_template("index1.html")


@app.route("/index3", methods=["GET", "POST"])
def Word():
    if request.method == "POST":
        VariableName = request.form['filename']
        print(VariableName)
        xl=client.Dispatch("Word.Application",pythoncom.CoInitialize())
        convert("C:/Users/HP/OneDrive - Higher Education Commission/Desktop/"+VariableName)
        convert("C:/Users/HP/OneDrive - Higher Education Commission/Desktop/"+VariableName, "C:/Users/HP/OneDrive - Higher Education Commission/Desktop/"+VariableName+".pdf")
    return render_template("index1.html")

@app.route("/index4", methods=["GET", "POST"])
def Image():
    if request.method == "POST":
        VariableName = request.form['filename']
        print(VariableName)
        pdf="C:/Users/HP/OneDrive - Higher Education Commission/Desktop/"+VariableName
        docx="C:/Users/HP/OneDrive - Higher Education Commission/Desktop/"+VariableName+".docx"
        cv=Converter(pdf)
        cv.convert(docx)
    return render_template('index1.html')

@app.route("/index5", methods=["GET", "POST"])
def Imaging():
    if request.method == "POST":
        VariableName = request.form['filename']
        print(VariableName)
        poppler_path=r'D:/Python/Flask/poppler-22.12.0/Library/bin'
        images = convert_from_path('C:/Users/HP/OneDrive - Higher Education Commission/Desktop/'+VariableName,poppler_path=poppler_path)
        for i in range(len(images)):
            images[i].save('C:/Users/HP/OneDrive - Higher Education Commission/Desktop/'+VariableName+ str(i) +'.jpg', 'JPEG')
    return render_template('index1.html')

@app.route("/index6", methods=["GET", "POST"])
def Merging():
    if request.method == "POST":
        VariableName = request.form['filename']
        VariableName1 = request.form['filename1']
        print(VariableName1)
        print('\n')
        print(VariableName)
        first_pdf_file_location = 'C:/Users/HP/OneDrive - Higher Education Commission/Desktop/'+VariableName
        second_pdf_file_location = 'C:/Users/HP/OneDrive - Higher Education Commission/Desktop/'+VariableName1

        first_pdf_file_descriptor = open(first_pdf_file_location, 'rb')
        second_pdf_file_descriptor = open(second_pdf_file_location, 'rb')

        first_pdf = PyPDF2.PdfReader(first_pdf_file_descriptor)
        second_pdf = PyPDF2.PdfReader(second_pdf_file_descriptor)

# Create a new pdf instance
        merged_pdf = PyPDF2.PdfWriter()
  # Add the pages from the first pdf to the new pdf
        for page in first_pdf.pages:
            merged_pdf.add_page(page)
# Add the pages from the second pdf to the new pdf
        for page in second_pdf.pages:
            merged_pdf.add_page(page)

        merged_pdf_location = 'C:/Users/HP/OneDrive - Higher Education Commission/Desktop/merged.pdf'
        merged_pdf_file_descriptor = open(merged_pdf_location, 'wb')
        merged_pdf.write(merged_pdf_file_descriptor)

        first_pdf_file_descriptor.close()
        second_pdf_file_descriptor.close()
        merged_pdf_file_descriptor.close()
    return render_template('index1.html')


@app.route("/docx", methods=["GET", "POST"])
def greet():
    return render_template('greet.html', name=request.form.get("name", "world"))


app.run(debug=True)
