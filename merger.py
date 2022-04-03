import csv
import time
import os
from tqdm import tqdm
from datetime import datetime
from mailmerge import MailMerge
from docx2pdf import convert

print('''
█▀▄▀█ ▄▀█ █ █░░   █▀▄▀█ █▀▀ █▀█ █▀▀ █▀▀
█░▀░█ █▀█ █ █▄▄   █░▀░█ ██▄ █▀▄ █▄█ ██▄
by 9akashnp8
''')

with open('recipients.csv') as row_file:
    reader = csv.reader(row_file)
    next(reader)
    num_of_rows = row_file.readlines()

with open('recipients.csv') as file:
    reader = csv.reader(file)
    next(reader)
    print('\n 1. Merging csv to docx \n')
    pbar = tqdm(total=len(num_of_rows))
    
    merged_date_time = datetime.now().strftime('%d-%m-%Y %I.%M.%S %p')
    if not os.path.exists(f"E:/Private/Programming/Python/mail-merger/docx_files/{merged_date_time}"):
        os.mkdir("E:/Private/Programming/Python/mail-merger/docx_files/" + merged_date_time)
        os.mkdir("E:/Private/Programming/Python/mail-merger/pdf_files/" + merged_date_time)

    for Name, Employee_ID, Designation, Email_ID in reader:
        '''Main loop that iterates through the rows in the csv. Data from the csv
        is then used in the next section to merge with the docx template'''

        pbar.update(1)

        with MailMerge('Template_File.docx') as content_template:
            '''Reads the docx file and automatically finds all the 'MergeFields' specified. It then merges
            the fields with the values fed through the for loop above. Next, it saves the docx to the specified
            'docx_files' folder.'''

            content_template.merge(Name=Name, Employee_ID=Employee_ID, Designation=Designation, Email_ID=Email_ID)
            content_template.write(f'E:\Private\Programming\Python\mail-merger\docx_files\{merged_date_time}\{Name}.docx')
    pbar.close()
    print('\n 2. Converting docx to pdf\n')
    convert(f"docx_files/{merged_date_time}", f"pdf_files/{merged_date_time}") #convert the docx files into pdf and save to the same
    print("\n Mail Merge Complete!")

        
        
