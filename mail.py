from tkinter import *
from tkinter import filedialog

import pandas as pd
import smtplib
from tkinter import messagebox
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import sys, os

application_path = os.path.dirname(sys.executable)


def send_message():
    head_info = head.get()
    body_info = bod.get()

    e = pd.read_excel("D:\mail.xlsx")
    emails = e['Mails'].values
    print('mails sent to:')
    print(emails)

    # Server Establishment
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login('#MAILID', '#PASSWORD')

    # Mail Contents
    subject = head_info
    msg = body_info

    body = "Subject:{}\n\n{}".format(subject, msg)

    for mail in emails:
        server.sendmail("MAILID", mail, body)

        print("message sent")

    messagebox.showinfo('Success', 'Mail Sent Successfully')


#02 Mail Attachment

def send_attachment():
    filename_info = filename.get()
    head_info = head.get()
    body_info = bod.get()

    global email
    e = pd.read_excel("D:\mail.xlsx")
    emails = e['Mails'].values
    print('mails sent to:')
    print(emails)

    msg = MIMEMultipart()

    msg['Subject'] = head_info
    msg['From'] = "#MAILID"
    email_body_info = body_info

    msg.attach(MIMEText(email_body_info, 'plain'))

    filename_info = filedialog.askopenfilename(initialdir="D:", title='open file')
    attachment = open(filename_info, 'rb')

    # instance of MIMEBase and named as p
    p = MIMEBase('application', 'octet-stream')

    p.set_payload(attachment.read())

    encoders.encode_base64(p)

    p.add_header('Content-Disposition', "attachment; filename= %s" % filename_info)

    msg.attach(p)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login("#MAILID", "#PASSWORD")
    text = msg.as_string()

    print("Login successful")

    for email in emails:
        server.sendmail("#MAILID", email, text)

    print("message sent")

    messagebox.showinfo('Success', 'Mail Sent Successfully')

# GUI
app = Tk()

app.geometry("500x250")
app.title("MAP")

heading = Label(text="MAP", bg="darkblue", fg="white", font="10", width="500", height="2")
heading.pack()

head = Label(text="Heading")
head.place(x=50, y=100)

Body = Label(text="Body")
Body.place(x=275, y=100)

filename = StringVar()
head = StringVar()
bod = StringVar()

head_entry = Entry(textvariable=head, width=25)
head_entry.place(x=50, y=130)

body_entry = Entry(textvariable=bod, width=25)
body_entry.place(x=275, y=130)

button1 = Button(app, text="Send Message", command=send_message, width="20", height="2", bg="grey")
button1.place(x=50, y=200)

button = Button(app, text="Send Attachment", command=send_attachment, width="20", height="2", bg="grey")
button.place(x=300, y=200)

mainloop()