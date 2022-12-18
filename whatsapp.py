import pywhatkit
import pandas as pd
from tkinter import *

def msg_send():
    msg = message.get()

    e = pd.read_excel("D:\mnumber.xlsx")
    number = e['Number'].values
    print('message sent to:')
    print(number)

    for num in number:
        pywhatkit.sendwhatmsg_instantly(num,msg)

        print("message sent")


#GUI
app = Tk()

app.geometry("500x250")
app.title("Whatsapp")

heading = Label(text="Whatsapp", bg="darkblue", fg="white", font="10", width="500", height="2")
heading.pack()

Body = Label(text="Message")
Body.place(x=220, y=100)

filename = StringVar()
message = StringVar()

message_entry = Entry(textvariable=message, width=50)
message_entry.place(x=50, y=130)

button1 = Button(app, text="Send Message", command=msg_send, width="20", height="2", bg="grey")
button1.place(x=170, y=175)

mainloop()