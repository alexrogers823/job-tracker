import smtplib, openpyxl
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def send_email(tracker, template_variables, sender, message_template, MY_ADDRESS, s):
    # For each contact, send the email:
    for name, email, cell, company, position, in template_variables:
        msg = MIMEMultipart()       # create a message

        # add in the actual person name to the message template
        message = message_template.substitute(RECRUITER_FIRST_NAME=name.title(), MY_NAME=sender.title(), COMPANY_NAME=company, JOB_TITLE=position, MY_EMAIL=MY_ADDRESS)

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


def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)


def get_contacts(worksheet, bottom_row): # Change this to match job tracker excel file
    names = []
    emails = []
    cells = []
    company = []
    position = []
    for i in range(3, bottom_row+1):
        if worksheet["J{}".format(i)].value == "Ready":
            names.append(worksheet["F{}".format(i)].value)
            emails.append(worksheet["G{}".format(i)].value)
            company.append(worksheet["B{}".format(i)].value)
            position.append(worksheet["C{}".format(i)].value)
            cells.append("J{}".format(i))

    return zip(names, emails, cells, company, position)


def prepare_emails():

    wb = openpyxl.load_workbook("Job_Search_Tracker_Template.xlsx")
    tracker = wb["Application Tracker"]
    bottom_row = tracker.max_row

    MY_ADDRESS = "alex.rogers823@gmail.com"
    PASSWORD = input("Enter email password: ")
    sender = "Alex Rogers" # Change this to your own name
    s = smtplib.SMTP(host="smtp.gmail.com", port=587)
    s.starttls()
    s.login(MY_ADDRESS, PASSWORD)

    template_variables = get_contacts(tracker, bottom_row)
    try:
        message_template = read_template('recruiter_email.txt')
    except FileNotFoundError:
        message_template = read_template('Emails/recruiter_email.txt')

    send_email(tracker, template_variables, sender, message_template, MY_ADDRESS, s)

    print("Emails sent!")
    wb.save("Job_Search_Tracker_Template.xlsx")

    s.quit() # Ends SMTP session and closes connection

# Main
# prepare_emails()
