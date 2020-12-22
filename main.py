# Email sender
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import base64
from email import encoders
from email.mime.base import MIMEBase
import sys
# GUI
import tkinter as tk
# from tkinter import *
from PIL import Image, ImageTk
from tkinter import messagebox
from tkinter import ttk
from tkinter import scrolledtext 
from threading import Thread
import os
import re

class LoginWindow:

    def __init__(self, window):

        self.sender_address = None
        self.sender_pass = None
        self.session = None
        self.validate = False
        self.window = window
        self.window.geometry('600x600')
        self.window.resizable(width=False, height=False)
        self.window.title('Email-Sender')

        self.frame1 = tk.Frame(master=self.window,relief=tk.SUNKEN, width=600, height=400, bg='#cbcaca')
        self.frame1.pack(fill=tk.BOTH, side=tk.TOP, expand=True)

        self.frame2 = tk.Frame(master=self.window,relief=tk.SUNKEN , width=600, height=200, bg="black")
        self.frame2.pack(fill=tk.BOTH, side=tk.BOTTOM, expand=True)

        #Upper Frame
        file_name = os.getcwd()+"\email_sender.jpg"
        self.image = Image.open(file_name)
        self.backgroundImage=ImageTk.PhotoImage(self.image.resize((600,400)))
        self.label = tk.Label(master=self.frame1, image = self.backgroundImage)
        self.label.pack()

        # Lower Frame #
        # Login Labels
        login_label = tk.Label(master=self.frame2, text="Email:", bg="black", fg="white")
        # login_label.config(font=(5))
        login_label.pack(pady=10)
        self.login_entry = tk.Entry(master=self.frame2)
        self.login_entry.config(font=(5))
        self.login_entry.pack()
        
        pass_label = tk.Label(master=self.frame2, text="Password:", bg="black", fg="white")
        # pass_label.config(font=(5))
        pass_label.pack(pady=10)
        self.pass_entry = tk.Entry(master=self.frame2, show="*")
        self.pass_entry.config(font=(5))
        self.pass_entry.pack()

        # Login Button
        login_button = tk.Button(
        master=self.frame2,
        text="Login",
        command = self.login,
        bg="green",
        fg="white",
        )
        login_button.pack(pady=10)


    def login(self):
        try:
            self.sender_address = self.login_entry.get()
            self.sender_pass = self.pass_entry.get()

            #Create SMTP session for sending the mail
            self.session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
            self.session.starttls() #enable security
            self.session.login(self.sender_address, self.sender_pass) #login with mail_id and password
            self.validate = True
            self.window.destroy()


        except:
            messagebox.showerror('Error', 'Login Failed(Make sure "acces to less secure app" is turned on)')

            
class SenderWindows:

    def __init__(self, window, session_value, username_value, pass_value):

        self.success_stack = []
        self.fail_stack = []
        self.window = window
        self.session = session_value
        self.sender_address = username_value
        self.password = pass_value
        # self.window.geometry('1300x700')
        # self.window.resizable(width=False, height=False)
        self.window.title("Email-Sender")

        self.frame1 = tk.Frame(master=self.window,relief=tk.SUNKEN, width=200, height=600)
        self.frame1.grid(row=0, column=0, padx=70, pady=5)
        self.frame2 = tk.Frame(master=self.window,relief=tk.SUNKEN, width=800, height=600)
        self.frame2.grid(row=0, column=1)
        self.frame3 = tk.Frame(master=self.window,relief=tk.SUNKEN, width=300, height=600)
        self.frame3.grid(row=0, column=2, padx=70)

        # frame1 contains settings
        select_batch = tk.Label(master=self.frame1, text="Select*:")
        # select_batch.config(font=("Courier", 20))
        select_batch.pack()
        self.batch = ttk.Combobox(master=self.frame1,
                                    height=40,
                                    
                                    values=['From File'], 
                                    state='readonly',
                                    )
        # self.batch.bind("<<ComboboxSelected>>", self.load_gui)
        self.batch.pack(pady=8)

        select_type = tk.Label(master=self.frame1, text="Select File Type*:")
        # select_type.config(font=("Courier", 20))
        select_type.pack()
        self.type = ttk.Combobox(master=self.frame1,
                                    height=40,
                                    
                                    values=['Image', 'PDF'], #, '2nd Year', '3rd Year', '4th Year'
                                    state='readonly',
                                    )
        # self.batch.bind("<<ComboboxSelected>>", self.load_gui)
        self.type.pack(pady=8)

        type_label = tk.Label(master=self.frame1, text="Type the file extension*(example->JPG)")
        # login_label.config(font=(5))
        type_label.pack(pady=8)
        self.ext_entry = tk.Entry(master=self.frame1)
        self.ext_entry.pack()

        text_box = tk.Text(self.frame1, height=40, width=40)
        text_box.pack(pady=8)
        info = """
        ----  .   .  --------- 
        |      . .       |
        |---    .        |
        |      . .       |
        ----  .   .      |REME


            Instructions:

 Step-1-> Locate 'send_email.txt' in the main folder(where Run.bat is located)
and in each line write 'email-filename'
(ex-> ishan2198@hotmail.com-score_card), filename without extension

 Step-2-> locate 'papers' folder and 
paste all the files to be send in this
folder

 Step-3-> Choose the file format above 
and choose 'From File', and type the
file extension exactly as it is(ex jpg)

Step-4-> Write the subject, body of the email and hit start button

 All the information can be seen in the notification area, if some emails failed
than they will be stored in fail_stack
and can be resend after the execution
finishes
        """
        text_box.insert(tk.END, info)
        text_box.config(bg = "dark blue" ,state="disabled", foreground="white")

        # frame2 contains email info
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("green.Horizontal.TProgressbar", foreground='green', background='green')
        self.progress = ttk.Progressbar(master=self.frame2,
                                        style="green.Horizontal.TProgressbar",
                                        orient="horizontal", 
                                        mode='determinate',
                                        length=100)
        self.progress.pack(pady=8)

        sub_label = tk.Label(master=self.frame2, text="Email Subject:")
        # login_label.config(font=(5))
        sub_label.pack(pady=8)
        self.sub_entry = tk.Entry(master=self.frame2)
        self.sub_entry.pack()
        
        body_label = tk.Label(master=self.frame2, text="Email Body:")
        # login_label.config(font=(5))
        body_label.pack(pady=8)
        self.body = tk.Text(self.frame2, height=20, width=30)
        self.body.pack()

        self.start_button = tk.Button(
        master=self.frame2,
        text="Start",
        command = self.validate,
        bg="green",
        fg="white",
        width=20,
        )
        self.start_button.pack(pady=8)

        # send_again = tk.Button(
        # master=self.frame2,
        # text="Re-Initialize",
        # command = self.validate,
        # bg="green",
        # fg="white",
        # width=20,
        # state="disabled"
        # )
        # send_again.pack(pady=8)

        # frame3 contains the notification panel
        self.notification_box = scrolledtext.ScrolledText(self.frame3, wrap=tk.WORD ,height=45, width=90)
        self.notification_box.grid(row=0, column=0)
        self.notification_box.insert(tk.INSERT, "Notification Area:")
        self.notification_box.tag_add("notification", "1.0", "1.18")
        self.notification_box.tag_config("notification", foreground="blue", font=(5))
        self.notification_box.config(bg = "black" ,state="disabled")
        # all the notification tags are defined here
        self.notification_box.tag_config("green", foreground="green", font=(5))
        self.notification_box.tag_config("orange", foreground="orange", font=(5))
        self.notification_box.tag_config("cyan", foreground="cyan", font=(5))
        self.notification_box.tag_config("red", foreground="red", font=(5))

    def validate(self):

        flag = 0
        # check batch selected or not
        if not self.batch.get():
            flag = 1
        # check file type selected or not
        if not self.type.get():
            flag = 1
        # check extension type entered or not
        if not self.ext_entry.get():
            flag = 1
        # check for subject
        if not self.sub_entry.get():
            flag = 1
        # check body
        if len(self.body.get("1.0", "end")) == 0:
            flag = 1
        
        # throw error
        if flag == 1:
            messagebox.showerror('Error', 'You must fill all the required fields')
        # start process
        else:
            self.start_button['state'] = 'disabled'
            self.success_stack = []
            self.fail_stack = []
            self.progress['value'] = 0
            if self.is_connected() == False:
                self.login_util()
            self.start_thread()

    def start_thread(self):

        Thread(target=self.start, daemon=True).start() # deamon=True is important so that you can close the program correctly

    def validate_file(self, emails, papers):
        flag=0
        for email in emails:
            val = re.search(r".+@.+\.com", email)

            if val == None:
                messagebox.showerror('Error', 'Some Emails are wrong, recheck the text file')
                flag=1
                self.start_button['state'] = 'normal'
                self.session.quit()
                break
        if flag == 0:
            self.start_2(emails, papers)

        


    def start(self):

        # if self.batch.get() == "1st Year":
        #     batch_list = range(401,503)
        #     email_pattern = "_bt19@iiitkalyani.ac.in"
        #     papers = [str(i) for i in batch_list]
        #     batch_list = [str(i)+email_pattern for i in batch_list]

        #     self.start_2(batch_list, papers)
        # other batches 
        #########################################

        #########################################

        emails = []
        papers = []
        if self.batch.get() == "From File":

            with open('send_email.txt') as f:
                # creating list of lines
                lines = f.readlines()
                for line in lines:
                    emails.append(line[0:line.index("-")])
                    papers.append(line[line.index("-")+1:].strip("\n"))
            self.validate_file(emails, papers)
            


    def start_2(self, emails, papers):

        send_dict = dict(zip(emails, papers))     

        for i in range(0, len(emails)):
            receiver_address = emails[i]
            if '.' not in self.ext_entry.get():
                filename = os.getcwd()+"/papers/"+papers[i]+"."+self.ext_entry.get()
            else:
                filename = os.getcwd()+"/papers/"+papers[i]+self.ext_entry.get()
            
            self.send_paper(str(i), receiver_address, filename, 
                            self.sub_entry.get(), self.body.get("1.0", "end"))
            self.progress['value'] = int(((i+1)/len(emails))*100)
            self.frame2.update_idletasks() 
            
        while len(self.fail_stack)!=0:

            MsgBox = tk.messagebox.askquestion ('Fail stack non empty','There was error in sending few emails.\nDo you want to try again for those email(s)(No to exit)',icon = 'error')

            if MsgBox == 'no':
                break
            length = len(self.fail_stack)
            for i in range(0, length):

                receiver_address = self.fail_stack.pop()
                if '.' not in self.ext_entry.get():
                    filename = os.getcwd()+"/papers/"+send_dict[receiver_address]+"."+self.ext_entry.get()
                else:
                    filename = os.getcwd()+"/papers/"+send_dict[receiver_address]+self.ext_entry.get()

                self.send_paper(str(i), receiver_address, filename, 
                            self.sub_entry.get(), self.body.get("1.0", "end"))

        self.print_status()


    def print_status(self):

        if(len(self.fail_stack)==0):
            self.print_message("\n\nFnished Successfully", "green")
            self.print_message("\nSuccessfully sent to::\n", "cyan")
            for suc in self.success_stack:
                self.print_message(str(suc)+"\n", "green")

        else:
            self.print_message("\n\nFinished with few errors", "orange")
            self.print_message("\nSuccessfully sent to::\n", "cyan")
            for suc in self.success_stack:
                self.print_message(str(suc)+"\n", "green")
            
            self.print_message("\n\nFailed to send to::\n", "cyan")
            for fail in self.fail_stack:
                self.print_message(str(fail)+"\n", "red")

        self.start_button['state'] = 'normal'

        self.session.quit()




        
    def print_message(self, message, tag):

        self.notification_box.config(state="normal")
        self.notification_box.insert(tk.END, message, tag)
        self.notification_box.see(tk.END)
        self.notification_box.config(state="disabled")

    def is_connected(self):

        try:
            status = self.send_paper.noop()[0]
        except:  # smtplib.SMTPServerDisconnected
            status = -1
        return True if status == 250 else False

    def login_util(self):

        self.print_message("\nlogging in...", "green")
        try:
            self.session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
            self.session.starttls() #enable security
            self.session.login(self.sender_address, self.password) #login with mail_id and password
        except:
            self.login()

    def login(self):

        if tk.messagebox.askretrycancel("Connection Error!!!", 'You got disconnected, check your connection and try again or press cancel to cancel the whole operation',icon = 'error') == False:
            self.print_status()

        else:
            self.login_util()

    def send_paper(self, name, receiver_address, filename, subject, body):

        self.print_message("\n\nSending...", 'green')

        # body
        mail_content = body

        #Setup the MIME
        message = MIMEMultipart("alternative")
        message['From'] = self.sender_address
        message['To'] = receiver_address
        message['Subject'] =  subject #The subject line
        #The body and the attachments for the mail

        # read image file
        if self.type.get() == "Image":
            try: 
                with open(filename, "rb") as attachment:
                    # The content type "application/octet-stream" means that a MIME attachment is a binary file
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
            except:
                self.print_message("\n\nFile not found, "+filename+f" doesnot exists, \n Continuing to next iteration...", 'orange')
                self.fail_stack.append(receiver_address)
                return
            # Encode to base64
            encoders.encode_base64(part)

            # Add header
            send_name = filename[filename.index("papers")+6:] 
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {send_name}",
            )
            # attach file
            message.attach(part)

        if self.type.get() == "PDF":

             # Attach the pdf to the msg going by e-mail
            try:
                with open(filename, "rb") as f:
                    attach = MIMEApplication(f.read(),_subtype="pdf")
            except:
                self.print_message("\n\nFile not found, "+filename+f" doesnot exists, \n Continuing to next iteration...", 'orange')
                self.fail_stack.append(receiver_address)
                return

            send_name = filename[filename.index("papers")+6:]
            attach.add_header('Content-Disposition','attachment',filename=send_name)
            
            message.attach(attach)

           

        # attach body
        message.attach(MIMEText(mail_content, 'plain'))
        text = message.as_string()
        try:
            self.session.sendmail(self.sender_address, receiver_address, text)
        except Exception as e:
            if self.is_connected() == False:
                self.login()
            else:
                self.fail_stack.append(receiver_address)
                self.print_message("\n\nError message: "+str(e)+" \nFailed to send email to "+name+f" continuing to next...", 'orange')
                return
            

        self.print_message("\n\nMail Sent To "+receiver_address, 'green')
        self.success_stack.append(receiver_address)
        






        

root = tk.Tk()
session = LoginWindow(root)
root.mainloop()

if session.validate == True:
    root2 = tk.Tk()
    SenderWindows(root2, session.session, session.sender_address, session.sender_pass)
# SenderWindows(root2, None, None, None)
    root2.mainloop()