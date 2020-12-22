import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import base64
from email import encoders
from email.mime.base import MIMEBase
import sys
import signal

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

# def time_out(signum, frame):
#     raise Exception("Timeout, Adding to fail stack, and continuing to next...")

def send_paper(name, receiver_address, filename, sender_address, session, fail_stack, subject, success_stack):

    print(f"{bcolors.OKGREEN}Sending...{bcolors.ENDC}")

    # body
    mail_content = '''
    This mail contains question paper, for you endsem exam.
    Best of luck!!!
    '''
    #Setup the MIME
    message = MIMEMultipart("alternative")
    message['From'] = sender_address
    message['To'] = receiver_address
    message['Subject'] =  subject #The subject line
    #The body and the attachments for the mail

    # read image file
    try: 
        with open(filename, "rb") as attachment:
            # The content type "application/octet-stream" means that a MIME attachment is a binary file
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
    except:
        print(f"{bcolors.WARNING}\nFile not found, "+filename+f" doesnot exists, \n Continuing to next iteration...{bcolors.ENDC}")
        fail_stack.append(receiver_address)
        return
    # Encode to base64
    encoders.encode_base64(part)

    # Add header 
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )
    # attach file
    message.attach(part)

    # attach body
    message.attach(MIMEText(mail_content, 'plain'))
    text = message.as_string()
    try:
        session.sendmail(sender_address, receiver_address, text)
    except:
        fail_stack.append(receiver_address)
        print(f"{bcolors.WARNING}\nFailed to send email to "+name+f" continuing to next...{bcolors.ENDC}")
        return

    print(f"{bcolors.OKGREEN}Mail Sent To "+name+f"{bcolors.ENDC}")
    success_stack.append(receiver_address)

print(f"""{bcolors.OKGREEN}{bcolors.BOLD}
-----------     .                       .   --------------------------
|                   .               .                   |
|                       .       .                       |
|----------                 .                           |
|                       .      .                        |
|                   .               .                   |
-----------     .                        .              | REME
{bcolors.ENDC}""")

session = None
#The mail addresses and password
print("Enter Gmail's Email-Address")
sender_address = str(input())
print("Enter Password")
sender_pass = str(input())
print("Enter Mail Subject")
subject = str(input())

print(f"{bcolors.OKGREEN}logging in...{bcolors.ENDC}")
# login to account
try:
    sender_address = sender_address
    sender_pass = sender_pass

    #Create SMTP session for sending the mail
    session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
    session.starttls() #enable security
    session.login(sender_address, sender_pass) #login with mail_id and password
except:
    print(f"""
    {bcolors.FAIL}\nFailed to login :(, please check your email/pass and try again
    Note->You need to allow less secure app in your gmail account settings
    Exiting program...{bcolors.ENDC}""")
    sys.exit()



fail_stack = []
success_stack = []
# signal.signal(signal.SIGALRM, time_out)
# signal.alarm(60)

if session!=None:
    # emails = ["462_bt19@iiitkalyani.ac.in"]
    for i in range(462, 463):
        receiver_address = str(i)+"_bt19@iiitkalyani.ac.in"
    # for email in emails:
        # receiver_address = email
        filename = "papers/"+str(i)+".png"
        # try:
        send_paper(str(i) ,receiver_address, filename, sender_address, session, fail_stack, subject, success_stack)
        # except Exception as exc:
            # print(exc)
    session.quit()

while len(fail_stack)!=0:
    print(f"""{bcolors.FAIL}
    There was error in sending few emails.
    Do you want to try again for those email(s)??(y/n){bcolors.ENDC}""")

    ch = input()
    if ch=='n':
        break
    length = len(fail_stack)
    for i in range(0, length):

        receiver_address = fail_stack.pop()
        filename = "papers/"+str(i)+".png"
        send_paper(str(i) ,receiver_address, filename, sender_address, session, fail_stack, subject, success_stack)

if(len(fail_stack)==0):
    print(f"{bcolors.OKGREEN}\nFnished Successfully{bcolors.ENDC}")
    print(f"{bcolors.OKBLUE}\nSuccessfully sent to::{bcolors.ENDC}")
    for suc in success_stack:
        print(suc)

else:
    print(f"{bcolors.WARNING}\nFinished with few errors{bcolors.ENDC}")
    print(f"{bcolors.OKBLUE}\nSuccessfully sent to::{bcolors.ENDC}")
    for suc in success_stack:
        print(suc)
    
    print(f"{bcolors.FAIL}\nFailed to send to::{bcolors.ENDC}")
    for fail in fail_stack:
        print(fail)

print("\n")



