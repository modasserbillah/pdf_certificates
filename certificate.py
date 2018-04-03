from reportlab.lib.pagesizes import letter
from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.platypus import Image
import pandas as pd

#write filepath here
excel_file = 'names.xlsx'
#names column name
column_name = 'name'
# edit these two to change the certificate text
title = 'Certificate of Participation'
second_line = 'This certificate is awarded to'
#logo, change the logo file name here, use only JPG 
logo = 'logo.jpg'

def generate(file_name):
    data = pd.read_excel(file_name)
    # make sure the column name for the names in 'name'
    for name in data[column_name]:
        pdf_file_name = name + '.pdf'
        make_certificate(name, pdf_file_name)
        
def make_certificate(name, pdf_file_name):
    c = canvas.Canvas(pdf_file_name, pagesize=landscape(letter))
    
    #header
    c.setFont('Times-Italic', 50, leading=None)
    c.drawCentredString(415, 395, title)
    c.setFont('Times-Italic', 34, leading=None)
    c.drawCentredString(415, 350, second_line)
    #name
    c.setFont('Times-BoldItalic', 45, leading=None)
    c.drawCentredString(415, 295, name)
    
    c.drawImage(logo, 350, 180, width=100, height=100)
    
    #save pdf
    c.save()
    
if __name__ == '__main__':
    generate(excel_file)