from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
from docx.enum.text import WD_LINE_SPACING
from docx2pdf import convert
import PyPDF2

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
    dept=intern[4]
    gender=intern[5]
    if gender=="male" or "Male":
        s="m"
    elif gender=="female" or "Female":
        s="f"
    
    he={"m":"he","f":"she"}
    his={"m":"his","f":"her"}
    him={"m":"him","f":"her"}
    if dept=="GD":
        designation="Graphic Designer"
    elif dept=="HR":
        designation="HR Intern"
    elif dept=="Marketing":
        designation="Digital Marketeer"

    start=intern[6]
    end=intern[7]


    para1="This is to certify that "+name+" has done "+his[s]+" internship as a "+designation+"   at Inception Wave Pvt. Ltd, from "+start+" to "+end
    if dept=="Marketing":
        para2="During the internship, "+he[s]+" has closely worked as a part of the Operations team for our product : Grapido, "+he[s]+" made a valuable contribution towards the publicity and digital marketing  ventures of Inception Wave Pvt. Ltd.."
    elif dept=="GD":
        para2="During the internship, "+he[s]+" has closely worked as a part of the Design and marketing team for our product : Grapido, "+he[s]+" made a valuable contribution towards the publicity and digital marketing  ventures of Inception Wave Pvt. Ltd.."
    elif dept=="HR":
        para2="During the internship, "+he[s]+" has closely worked as a part of the Human Resource team for our product : Grapido, "+he[s]+" made a valuable contribution towards the overall team management of different domains of Inception Wave Pvt. Ltd.."
        
    para3="Throughout the internship, "+his[s]+" efforts and dedication towards the task assigned was praiseworthy. During the internship, "+he[s]+" demonstrated good communication and designing skills with a self-motivated attitude to learn new things.Further, "+his[s]+" performance exceeded the expectations and was able to complete the assigned tasks successfully on time. "
    para4="We wish "+him[s]+" all the best for "+his[s]+" future endeavours."

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

def convert(filename):
    watermark="Response_PDF/"+filename
    pdf_file="watermark.pdf"
    merged_file = "Response_PDF/"+filename+".pdf"
    input_file = open(pdf_file,'rb')
    input_pdf = PyPDF2.PdfFileReader(input_file)
    watermark_file = open(watermark,'rb')
    watermark_pdf = PyPDF2.PdfFileReader(watermark_file)
    pdf_page = input_pdf.getPage(0)
    watermark_page = watermark_pdf.getPage(0)
    pdf_page.mergePage(watermark_page)
    output = PyPDF2.PdfFileWriter()
    output.addPage(pdf_page)
    merged_file = open(merged_file,'wb')
    output.write(merged_file)
    merged_file.close()
    watermark_file.close()
    input_file.close()


def watermark():
    directory = os.fsencode("Response_PDF/")
    for file in os.listdir(directory):
     filename = os.fsdecode(file)
     if filename.endswith(".pdf"): 
        #  path=os.path.join("Response_PDF", filename)
         convert(filename)
         


def main():
    rdCSV()
    # convertPDF()
    # watermark()


if __name__=="__main__":
    main()