import csv

message = """
Hi {Name} , your email is {Email}
"""

with open('contacts.csv') as file:
    reader = csv.reader(file)
    next(reader)
    for Name, Email in reader:
        print(message.format(Name=Name, Email=Email))
        
