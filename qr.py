import os
from PIL import Image
from docx import Document
from docx.shared import Cm
from docx.enum.section import WD_ORIENTATION
from docx.oxml.ns import qn
from docx.shared import Inches
from docx2pdf import convert
from PIL import Image
import re
from generate_qr import generate_qr_with_label

#------------------------------------------------------------------------------------------------------------------#
#------------------------ GENERATE QR CODE FOR THE SET RANGES USING THE CUSTOM FUNCTION ---------------------------#
#------------------------------------------------------------------------------------------------------------------#
img = generate_qr_with_label('AEDCBD00', 12802, 12804)


#-------------------------------------------------------------------------------------------------------------#
#-------------------------------- MERGE GENERATED CODES AND LOCATION LOGO ------------------------------------#
#-------------------------------------------------------------------------------------------------------------#
location = "AEDC"
if location == 'AEDC':
    disco_logo = 'aedc-logo.png'
elif location == 'ECG':
    disco_logo = 'ecg-ghana.jpg'
else:
    raise ValueError('Invalid location')

asset = "BD"
dir_name = f"C:\\Users\\Moshood\\OneDrive\\dev\\qr_code_tag_generator\\QRCodes\\{location}"
image_files = [f for f in os.listdir(dir_name) if f.endswith(".png")] # get a list of all PNG files in the folder

global vtags_dir
vtags_dir = "VTags"

logo = os.path.join(os.getcwd(), disco_logo)
logo = Image.open(logo)
logo = logo.resize((332, 332))
logo_size = logo.size

for qc_name in image_files:
    qc = os.path.join(dir_name, qc_name)
    qc = Image.open(qc)
    qc_size = qc.size
    s2u = [i*2 for i in qc_size]
    qc = qc.resize(s2u)
    qc_size = qc.size

    logo_qc = Image.new('RGB',(332+290, 332), (250,250,250))
    logo_qc.paste(logo,(0,0))
    logo_qc.paste(qc,(332,0))

    if not os.path.exists(vtags_dir):
        os.makedirs(vtags_dir)

    logo_qc.save(os.path.join(vtags_dir, qc_name),"PNG")
    # logo_qc.show()

#-------------------------------------------------------------------------------------------------#
#------------------------ LOAD FILES INTO DOCX FILE AND CONVERT TO PDF ---------------------------#
#-------------------------------------------------------------------------------------------------#

pics = sorted(os.listdir(dir_name))
pics[:4], len(pics)

document = Document()
sections = document.sections
for section in sections:
    section.top_margin = Cm(1.27)
    section.bottom_margin = Cm(1.27)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)
    section.orientation = WD_ORIENTATION.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21)
    sectPr = section._sectPr
    cols = sectPr.xpath('./w:cols')[0]
    cols.set(qn('w:num'), '2')
    cols.set(qn('w:space'), '10')  # Set space between columns to 10 points ->0.01"

logo = os.path.join(os.getcwd(), disco_logo)

# Extract the beginning and end of the VTag range from the file names
vtags = [int(re.search(r'\d+', f).group()) for f in pics]
vtags_begin = min(vtags)
vtags_end = max(vtags)

for qrc in pics:
    paragraph = document.add_paragraph()
    pth = os.path.join(os.getcwd(), dir_name, qrc)
    qc = Image.open(pth)
    qc_size = qc.size
    s2u = [i*2 for i in qc_size]
    qc = qc.resize(s2u)
    qc_size = qc.size

    logo = os.path.join(os.getcwd(), disco_logo)
    logo = Image.open(logo)
    logo = logo.resize((332, 332))
    logo_size = logo.size

    logo_qc = Image.new('RGB',(332+290, 332), (250,250,250))
    logo_qc.paste(logo,(0,0))
    logo_qc.paste(qc,(332,0))

    # Construct the file name for the QR code image with the VTag range included
    qc_name = re.sub(r'[:\\]', '_', qrc)
    qc_range = f"{vtags_begin:05} - {vtags_end:05}"
    fpth = f"{vtags_dir}/{location} {vtags_dir} {asset} {qc_range} {qc_name}"
    logo_qc.save(fpth,"PNG")
    lqpth = os.path.join(os.getcwd(), fpth)

    run_2 = paragraph.add_run()
    run_2.add_picture(lqpth, height=Inches(1.25))

document_dir = "generated"
if not os.path.exists(document_dir):
    os.makedirs(document_dir)

docx_file_name = f'{location} {vtags_dir} {asset} {qc_range}.docx'
docx_file_name = re.sub(r'[:\\]', '_', docx_file_name)
docx_file_path = os.path.join(document_dir, docx_file_name)

# Save the Word document
document.save(docx_file_path)

# Convert the Word document to PDF
pdf_file_name = f'{location} {vtags_dir} {asset} {qc_range}.pdf'
pdf_file_name = re.sub(r'[:\\]', '_', pdf_file_name)
pdf_dir_path = os.path.join(document_dir, 'pdf')
if not os.path.exists(pdf_dir_path):
    os.makedirs(pdf_dir_path)
pdf_file_path = os.path.join(pdf_dir_path, pdf_file_name)
convert(docx_file_path, pdf_file_path)
