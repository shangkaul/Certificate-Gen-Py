from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
from docx.enum.text import WD_LINE_SPACING
from docx2pdf import convert

import csv
import os.path


def createDoc(intern):
    document = Document('certi.docx')
    if 'Head' not in document.styles:
        style = document.styles.add_style('Head', WD_STYLE_TYPE.PARAGRAPH)
    else:
        style=document.styles['Head']
    font = style.font
    font.size = Pt(18)
    font.bold=True

    style = document.styles['Normal']
    font = style.font
    font.size = Pt(12)
    font.bold=False
    paragraph_format = document.styles['Normal'].paragraph_format
    paragraph_format.line_spacing = 1.5
    paragraph_format.space_before = Pt(18)


    head = document.add_paragraph("TO WHOMSOEVER IT MAY CONCERN")
    head.style=document.styles['Head']
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    head.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE


    content=[]
    name=intern[2]
    designation="Graphic Designer"
    para1="This is to certify that "+name+" has done his internship as a "+designation+"   at Inception Wave Pvt. Ltd, from 1st June to 31st July 2020."
    para2="During the internship, he has closely worked as a part of the Content Development Team. He made valuable contribution towards the publicity and digital marketing  ventures of Inception Wave Pvt. Ltd.."
    para3="His contribution to creative ideas are being implemented in our social media marketing campaigns. During the internship, he demonstrated good communication and designing skills with a self-motivated attitude to learn new things. His performance exceeded the expectations and was able to complete the assigned tasks successfully on time. "
    para4="We wish him all the best for his future endeavours."

    content.append(para1)
    content.append(para2)
    content.append(para3)
    content.append(para4)

    for i in content:
        para=document.add_paragraph(i)
        para.alignment=WD_ALIGN_PARAGRAPH.LEFT
        para.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

        para.styles=document.styles['Normal']

    empty=document.add_paragraph(' ')
    empty=document.add_paragraph(' ')



    stamp=document.add_picture('stamp.png')
    last_paragraph = document.paragraphs[-1] 
    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    count=0
    while os.path.isfile("Response/"+name+".docx"):
        count+=1
        name=name+str(count)
    filename="Response/"+name+".docx"
    document.save(filename)

def convertPDF():
    convert("Response/")

def rdCSV():
    with open('Intern_Response.csv', newline='') as f:
        reader = csv.reader(f)
        data = list(reader)
        for intern in data:
            createDoc(intern)


def main():
    rdCSV()
    convertPDF()


if __name__=="__main__":
    main()