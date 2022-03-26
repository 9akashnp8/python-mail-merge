#Imports
import csv
import time
from mailmerge import MailMerge
from docx2pdf import convert

with open('recipients.csv') as file:

    #Read the csv and skip the header row using next(reader)
    reader = csv.reader(file)
    next(reader)

    for Name, Employee_ID, Designation, Email_ID in reader:
        '''Main loop that iterates through the rows in the csv. Data from the csv
        is then used in the next section to merge with the docx template'''

        with MailMerge('Template_File.docx') as content_template:
            '''Reads the docx file and automatically finds all the 'MergeFields' specified. It then merges
            the fields with the values fed through the for loop above. Next, it saves the docx to the specified
            'docx_files' folder.'''

            content_template.merge(Name=Name, Employee_ID=Employee_ID, Email_ID=Email_ID)
            content_template.write(f'E:\Private\Programming\Python\mail-merger\docx_files\{Name}.docx')

time.sleep(5) #Wait for the merging to be processed.

convert("docx_files/") #convert the docx files into pdf and save to the same
        
        
