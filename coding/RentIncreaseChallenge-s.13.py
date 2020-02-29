import xlwings as xw
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
# from input import *

# =================================================
doc_filename = 'DepositDeduction-Challenge.docx'
para1_body = '{Date}\n{Taddress}\n{TPostcode}'
para2_body = 'Dear {LLname},\n\nRE: Deposit Deduction at {Taddress}\n\nI am the tenant at the above address and I am writing to inform you that the reason you are deducting my deposit is invalid for the following reason(s):\n\n'
# todo - Q2 or Q2.5         Apply logic based on Yes/No
b1_body = '{Q2} Landlords must protect tenancy deposits within 30 days of receiving them. If a landlord fails to protect the deposit, the tenant can be eligible for compensation 1-3 times the original deposit.'
b2_body = '{Q6} The property was left in a clean manner, that was similar to, or in better condition than when I moved in.'
b3_body = '{Q7} Landlords cannot deduct from a tenants deposit due to general wear and tear. I have considered the following factors:\n\n\t\t- The type of damages and the items material\n\t\t- What the item is and how long it is supposed to last\n\t\t- How old the item is and the length of your tenancy\n\t\t- Specifications of the item\n\t\t- What shape the item was in upon moving in\n\n'
para3_body = 'Please contact me as soon as possible to further discuss the matter.\n\nBest regards,\n\n\n{Name}'

# ================================================
wb = xw.Book('../data/Tenantchat.xlsx')
sht = wb.sheets['Deposit Deduction']

# -----------------------------------------------
# Define variables 
date = sht.range('A2').value
taddress = sht.range('F2').value
tpostcode = sht.range('G2').value
llname = sht.range('E2').value
name = sht.range('B2').value
q2 = sht.range('I2').value
q6 = sht.range('N2').value
q7 = sht.range('O2').value

# =================================================
# Create document from here
d = Document()

# -------------------------------------------------
para1 = d.add_paragraph(para1_body.format(
                                        Date= date,
                                        Taddress= taddress,
                                        TPostcode= tpostcode))

para1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
# -------------------------------------------------
para2 = d.add_paragraph(para2_body.format(
                                        LLname= llname,
                                        Taddress= taddress))

para2.alignment = WD_ALIGN_PARAGRAPH.LEFT
# -------------------------------------------------
if q2 == 'Yes':
    b1 = d.add_paragraph('\n\n\n')    # empty with 3 lines as per the preview img
elif q2 == 'No':
    b1 = d.add_paragraph(b1_body.format(Q2= ''), style= 'List Bullet')
else:
    b1 = d.add_paragraph(b1_body.format(Q2= q2), style= 'List Bullet')
# -------------------------------------------------
if q6 == 'Yes':
    b2 = d.add_paragraph('\n\n')    # empty with 2 lines as per the preview img
elif q6 == 'No':
    b2 = d.add_paragraph(b2_body.format(Q6= ''), style= 'List Bullet')
else:
    b2 = d.add_paragraph(b2_body.format(Q6= q6), style= 'List Bullet')
# -------------------------------------------------
if q7 == 'Yes':
    b3 = d.add_paragraph('\n\n\n\n\n\n\n\n\n')    # empty with 9 lines as per the preview img
elif q7 == 'No':
    b3 = d.add_paragraph(b3_body.format(Q7= ''), style= 'List Bullet')
else:
    b3 = d.add_paragraph(b3_body.format(Q7= q7), style= 'List Bullet')
# -------------------------------------------------
para3 = d.add_paragraph(para3_body.format(Name= name))

para3.alignment = WD_ALIGN_PARAGRAPH.LEFT

# -------------------------------------------------
d.save(doc_filename)