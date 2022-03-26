import csv
from mailmerge import MailMerge

with open('contacts.csv') as file:
    reader = csv.reader(file)
    next(reader)
    for Name, Email in reader:
        with MailMerge('Test.docx') as content_template:
            print(content_template.get_merge_fields())
        
        
