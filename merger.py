import csv
from mailmerge import MailMerge

with open('contacts.csv') as file:
    reader = csv.reader(file)
    next(reader)
    for Name, Email in reader:
        with MailMerge('Test.docx') as content_template:
            content_template.merge(Name=Name, Email=Email)
            content_template.write(f'E:\Private\Programming\Python\mail-merger\docx\{Name}.docx')
        
        
