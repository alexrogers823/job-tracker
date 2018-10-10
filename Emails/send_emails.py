import smtplib, openpyxl
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def send_email(tracker, names, emails, cells, message_template, MY_ADDRESS, s):
    # For each contact, send the email:
    for name, email, cell in zip(names, emails, cells):
        msg = MIMEMultipart()       # create a message

        # add in the actual person name to the message template
        message = message_template.substitute(PERSON_NAME=name.title())

        # setup the parameters of the message
        msg['From']=MY_ADDRESS
        msg['To']=email
        msg['Subject']="This is TEST"

        # add in the message body
        msg.attach(MIMEText(message, 'plain'))

        # send the message via the server set up earlier.
        s.send_message(msg)

        del msg

        # Change status of outcome from "ready" to "email sent"
        tracker[cell].value = "Email Sent"
        print(tracker[cell].value)


def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)


def get_contacts(worksheet, bottom_row): # Change this to match job tracker excel file
    print(worksheet)
    names = []
    emails = []
    cells = []
    print(bottom_row)
    for i in range(3, bottom_row+1):
        print(worksheet["J{}".format(i)].value)
        if worksheet["J{}".format(i)].value == "Ready":
            names.append(worksheet["F{}".format(i)].value)
            emails.append(worksheet["G{}".format(i)].value)
            cells.append("J{}".format(i))

    print(names)
    print(emails)
    print(cells)

    return names, emails, cells


def prepare_emails():

    wb = openpyxl.load_workbook("Job_Search_Tracker_Template.xlsx")
    tracker = wb["Application Tracker"]
    bottom_row = tracker.max_row

    MY_ADDRESS = "alex.rogers823@gmail.com"
    PASSWORD = input("Enter email password: ")
    s = smtplib.SMTP(host="smtp.gmail.com", port=587)
    s.starttls()
    s.login(MY_ADDRESS, PASSWORD)

    names, emails, cells = get_contacts(tracker, bottom_row)
    try:
        message_template = read_template('test_message.txt')
    except FileNotFoundError:
        message_template = read_template('Emails/test_message.txt')

    send_email(tracker, names, emails, cells, message_template, MY_ADDRESS, s)

    print("Emails sent!")
    wb.save("Job_Search_Tracker_Template.xlsx")

    s.quit() # Ends SMTP session and closes connection

# Main
# prepare_emails()
