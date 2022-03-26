import csv
import time
from mailmerge import MailMerge
from docx2pdf import convert

with open('contacts.csv') as file:
    reader = csv.reader(file)
    next(reader)
    for Name, Email in reader:
        with MailMerge('Test.docx') as content_template:
            content_template.merge(Name=Name, Email=Email)
            content_template.write(f'E:\Private\Programming\Python\mail-merger\docx_files\{Name}.docx')

time.sleep(5)
convert("docx_files/")
        
        
