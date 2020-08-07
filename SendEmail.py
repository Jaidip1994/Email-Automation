"""
Author: Jaidip Ghosh
Contact Details: jaidip1994@gmail.com
About This Module:
1. This project is for sending mail
2. This is only relevant when you want to send the mails from one Gmail Account to Another Email Account
3. Prior to running this project
    1. Please Ensure that Less secure app access is enabled: https://myaccount.google.com/lesssecureapps?pli=1
    2. Please run the setup.sh prior to running of the application
"""

import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import datetime
from texttable import Texttable
from datetime import time
from tqdm import tqdm
import time as timeprog

SENDER = "<sender_email_id>"
RECEIVER = "<received_email_id>"
PASSWORD = "<password>"


class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


print("---------------------------------------------------------------------------------------")
print(f"\n{bcolors.HEADER}Enter Enter the following details{bcolors.ENDC}\n")

SHIFT = input(f"{bcolors.OKBLUE}Handover Mail is for which Shift Details: {bcolors.ENDC}")

print(f"{bcolors.OKGREEN}Handover is for which Day \n1.Yesterday \n2.Today \n3.Tomorrow{bcolors.ENDC}")

HANDOVER_SENDING_TIME_CHOICE = int(input(f"{bcolors.OKBLUE}Enter you choice (1/2/3): {bcolors.ENDC}"))
DATE_TIME = datetime.date.today()
DATE = DATE_TIME.strftime("%d-%m-%y")
if HANDOVER_SENDING_TIME_CHOICE == 1:
    DATE_TIME -= datetime.timedelta(days=1)
    DATE = DATE_TIME.strftime("%d-%m-%Y")
elif HANDOVER_SENDING_TIME_CHOICE == 3:
    DATE_TIME += datetime.timedelta(days=1)
    DATE = DATE_TIME.strftime("%d-%m-%Y")


print(f"{bcolors.OKGREEN}When do you want to send the mail ? \n1.Today \n2.Tomorrow{bcolors.ENDC}")
DAY_CHOICE = int(input(f"{bcolors.OKBLUE}Enter you choice (1/2) {bcolors.ENDC}"))
TIME_AT_WHICH_TO_BE_SENT = input(f"{bcolors.OKBLUE}Enter the time in HH:MM format: {bcolors.ENDC} ")
if(len(TIME_AT_WHICH_TO_BE_SENT) != 5):
    print(f'{bcolors.FAIL}Please Try Again!!!{bcolors.ENDC}')
    assert 1 == 2
h, m = map(int, TIME_AT_WHICH_TO_BE_SENT.split(':'))
RES = time(hour=h, minute=m)

SEND_DATETIME = datetime.date.today()
if DAY_CHOICE == 2:
    SEND_DATETIME += datetime.timedelta(days=1)

HANDOVER_TO = input(f"{bcolors.OKBLUE}Handover To? {bcolors.ENDC}")

choice = 'y'
WORKED_ON_DET = list()
print(f"{bcolors.OKGREEN}Enter the Worked on / Working on details: {bcolors.ENDC}")

while choice == 'y':
    WORKED_ON_DET.append(input(f"{len(WORKED_ON_DET) + 1}. "))
    choice = input(f"{bcolors.WARNING}Do you have some more point, (y/n) {bcolors.ENDC}").lower()
choice = 'y'
TAKE_CARE = list()
print(f"{bcolors.OKGREEN}Things to Take Care of: {bcolors.ENDC}")

while choice == 'y':
    TAKE_CARE.append(input(f"{len(TAKE_CARE) + 1}. "))
    choice = input(f"{bcolors.WARNING}Do you have some more point, (y/n) {bcolors.ENDC}").lower()

print("---------------------------------------------------------------------------------------")

print(f'{bcolors.BOLD}Check the Final Shift Handover Details Template {bcolors.ENDC}')
print(f'{bcolors.OKBLUE}')
t = Texttable()
t.add_rows([['Handover Details', ''],
            ['From', SENDER],
            ['To', RECEIVER],
            ['Date', DATE],
            ['TIME', TIME_AT_WHICH_TO_BE_SENT],
            ['Handover To', HANDOVER_TO],
            ['Working On/Worked On', (str)('\n'.join(WORKED_ON_DET))],
            ['Things to Take Care of', (str)('\n'.join(TAKE_CARE))]])
print(t.draw())
print(f'{bcolors.ENDC}')

choice = input(
    f"{bcolors.WARNING}Does everything looks fine ? Do you want to proceed?(y/n) {bcolors.ENDC}").lower()
print("---------------------------------------------------------------------------------------")


if(choice == 'y'):
    # Determine how much time to wait
    final_date_time = datetime.datetime(
        SEND_DATETIME.year, SEND_DATETIME.month, SEND_DATETIME.day, RES.hour, RES.minute, 00, 000000)
    time_to_wait = round(
        (final_date_time - datetime.datetime.now()).total_seconds())
    print(
        f"{bcolors.WARNING}Time to wait is : { str(datetime.timedelta(seconds=time_to_wait))}{bcolors.ENDC}")

    # To show progress bar of wait
    interval_split = time_to_wait // 10
    for i in tqdm(range(10)):
        timeprog.sleep(interval_split)

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "Handover Details"
    msg["From"] = SENDER
    msg["To"] = RECEIVER
    msg["bcc"] = SENDER  # For own reference

    html = ''
    with open('testFile.html', 'r') as fh:
        for line in fh:
            html += line

    # Replacing the place holders
    html = html.replace('<SHIFT>', SHIFT)
    html = html.replace('<DATE>', DATE)
    html = html.replace('<HANDOVER_TO>', HANDOVER_TO)
    html = html.replace('<WORKED_DETAILS>', '\n'.join(
        [f'<li>{elem}</li>' for elem in WORKED_ON_DET]))
    html = html.replace('<TAKE_CARE>', '\n'.join(
        [f'<li>{elem}</li>' for elem in TAKE_CARE]))

    # Turn these into plain/html MIMEText objects
    part2 = MIMEText(html, "html")

    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    msg.attach(part2)

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            smtp.ehlo()     # Establishing connection with the mail server
            # Encryption of the traffic
            smtp.starttls(context=ssl.create_default_context())
            smtp.ehlo()     # Reidentify as the encrypted connectionl̥

            smtp.login(SENDER, PASSWORD)
            print(f'{bcolors.WARNING}Sending Message...{bcolors.ENDC}')
            smtp.sendmail(SENDER, RECEIVER, msg.as_string())
            print(f'{bcolors.OKGREEN}Message Sent!!!{bcolors.ENDC}')
    except Exception as e:
        print(e.message)
else:
    print(f'{bcolors.FAIL}Mission Aborted !!{bcolors.ENDC}')
