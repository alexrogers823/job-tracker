# This will save your life

import openpyxl, time, sys
from openpyxl_shift.shift import Shift
import Emails.send_emails as email

# hey this is important
# that is all

def menu():
    # Will add other things later
    choices = [
        "[send] introductory emails to recruiters",
        "[enter] job data",
        "[edit] job search tracker",
        "[list] information of active companies",
        "[exit] program"]
    print("What would you like to do?")
    for option in choices:
        print(option)
    print()
    while True:
        try:
            choice = input().lower()
            if choice.startswith("send") or choice.lower().startswith("email"):
                send_emails()
            elif choice.startswith("ent"):
                prepare_entry()
            elif choice.startswith("lis"):
                display_active()
            elif choice.startswith("edi"):
                edit_entry()
            elif choice.startswith("exi"):
                print("goodbye")
                sys.exit()
            else:
                raise ValueError
        except ValueError:
            print("Pick a valid option (or 'exit' to end program)")
        else:
            break




def prepare_entry():
    print("Separate requests with a comma")
    while True:
        company, title = tuple(input("Give the company name and work title\n").split(", "))
        recruiter, email = tuple(input("Now give me the name and email of contact person\n").split(", "))
        language = input("What's the primary langauge?\n")
        url = input("What's the job posting URL?\n")

        print("Making entry...")
        time.sleep(0.5)
        add_entry(company, title, recruiter, email, url, language)

        if input("Done. Another?\n").lower().startswith("n"):
            break

    menu()


def add_entry(company, title, recruiter, email, url, language):
    global bottom_row
    bottom_row += 1
    entries = [company, title, "not yet", None, recruiter, email, None, None, "Ready", url, "Open", language]
    column = 66

    for entry in entries:
        tracker['{}{}'.format(chr(column), bottom_row)] = entry
        column += 1

    wb.save("Job_Search_Tracker_Template.xlsx")



def send_emails():
    email.prepare_emails()
    menu()


def display_active():
    print("Active companies:")
    for i in range(2, bottom_row+1):
        if tracker["L"+str(i)].value == "Open":
            company = tracker["B"+str(i)].value
            contact = tracker["F"+str(i)].value
            status = tracker["J"+str(i)].value
            lang = tracker["M"+str(i)].value
            print("{} | {} | {} | {}".format(company, contact, status, lang))
    print()
    input('PRESS ENTER')
    menu()

def edit_entry():
    for i in range(2, bottom_row+1):
        position = '{}: {}, {}'
        if tracker["L"+str(i)] != "Open":
            position += ' (Closed)'
        print(position.format(i, tracker["B"+str(i)].value, tracker["C"+str(i)].value))
    print()
    num = (input("Which row to edit? (Input index number)\n"))

    for j in range(13):
        cell = chr(65+j) + num
        print('{}: {}'.format(cell, tracker[cell].value), end=" | ")

    print()
    change = input("Cell first, then change (Ex: I3, Found Recruiter)\n").split(", ")
    chosen_cell, edited = tuple(change)

    tracker[chosen_cell] = edited

    print("Logged")
    print(tracker[chosen_cell].value)
    wb.save("Job_Search_Tracker_Template.xlsx")
    menu()



# Variables
wb = openpyxl.load_workbook("Job_Search_Tracker_Template.xlsx")
tracker = wb["Application Tracker"]
bottom_row = tracker.max_row

menu()
wb.save("Job_Search_Tracker_Template.xlsx")
