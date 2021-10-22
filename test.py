
import tkinter as tk
import webbrowser
from functools import partial
from tkinter.filedialog import askopenfile
import time
from tkinter.ttk import *
import pandas as pd
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from tkinter import messagebox


window= tk.Tk()
window.geometry('800x600')
#window.state('zoomed') #to open maximized tab automatically
window.title("Bulk Email Project")

#================FUNCTIONS=====================
def show_frame(frame):
    frame.tkraise() #to show frames

def openlink(url):
   webbrowser.open(url) #to open url link

def validateLogin(username, password):

    e = pd.read_excel("emails.xlsx")
    emailcol = e.get("emails")
    # print(emailcol)
    list_of_emails = list(emailcol)
    print(list_of_emails)
    try:
        server = sm.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(username.get(), password.get())
        from_ = username.get()
        to_ = list_of_emails
        message = MIMEMultipart("alternative")
        message['Subject'] = "TEST SUBJECT "
        message['From'] = username.get()

        html = '''
        This is a test email using python   
        '''
        text = MIMEText(html, "html")
        message.attach(text)
        server.sendmail(from_, to_, message.as_string())

        messagebox.showinfo("success","Message has been sent!")


    except Exception as ex:
        messagebox.showerror("Error",ex)

    return

def open_file():
    file_path = askopenfile(mode='r', filetypes=[('text files', '*txt')])
    if file_path is not None:
        pass

def uploadFiles():
    pb1 = Progressbar(
        window,
        orient=tk.HORIZONTAL,
        length=300,
        mode='determinate'
        )
    pb1.grid(row=4, columnspan=3, pady=20)
    for i in range(5):
        window.update_idletasks()
        pb1['value'] += 20
        time.sleep(1)
    pb1.destroy()
    Label(window, text='File Uploaded Successfully!', foreground='green').grid(row=4, columnspan=3, pady=10)


frame1= tk.Frame(window)
frame2= tk.Frame(window)
frame3= tk.Frame(window)
frame4= tk.Frame(window)

for frame in (frame1, frame2, frame3, frame4):
    frame.grid(row=0, column=0, sticky='nsew') #for the frame to stick when we expand the window
# The sticky attribute on a widget controls only where it will be placed if it doesn't completely fill the cell.

#--------------------------------------FRAME 1 INTRO ----------------------------------



frame1_l1 = tk.Label(frame1, text="Welcome!", font=("Arial Bold",50)).place(x=250,y=0)

frame1_instructions= tk.Label(frame1, text="Instructions: \n\n\n"
                           "1. You will need an excel file with the names and emails of the people who need to be emailed\n"
                           "2. Create notepad file with xyz placeholders\n"
                           "3. upload them\n\n"
                           "The next page will show you the demo video\n\n"
                            "For a demo video click this link -----> ", justify=tk.LEFT).place(x=20,y=100)

frame1_nextbt= tk.Button(frame1,text="Next page", padx=50,command=lambda: show_frame(frame2)).place(x=600,y=500)

link_bt= tk.Button(frame1, text="DEMO VIDEO", cursor="hand2", command= lambda: openlink("https://www.youtube.com/watch?v=6NMbJCmGaAg")).place(x=225,y=230)

#--------------------------------------FRAME 2 UPLOAD FILES---------------------------------------
frame2_l1 = tk.Label(frame2, text="UPLOAD REQUIRED FILES", font=("Arial Bold",25)).grid(row=0,column=0)

adhar = tk.Label(frame2,text='Upload file(text file) in .txt format ')
adhar.grid(row=1, column=0, padx=10)

adharbtn = tk.Button(frame2,text ='Choose File', command = lambda:open_file())
adharbtn.grid(row=1, column=1)

dl = tk.Label(frame2,text='Upload file(excel file) in .exe format ')
dl.grid(row=2, column=0, padx=10)

dlbtn = tk.Button(
    frame2,
    text ='Choose File ',
    command = lambda:open_file()
    )
dlbtn.grid(row=2, column=1)

upld = tk.Button(
    frame2,
    text='Upload Files',
    command=uploadFiles
    )
upld.grid(row=3, columnspan=3, pady=10)


frame2_nextbt= tk.Button(frame2,text="Next page", padx=50,command=lambda: show_frame(frame3)).place(x=600,y=500)
frame2_prevbt= tk.Button(frame2,text="previous page", padx=50,command=lambda: show_frame(frame1)).place(x=50,y=500)

#--------------------------------------FRAME 3 USERNAME AND PASSWORD---------------------------------------

blank_l1 = tk.Label(frame3, text="    ", font=("Arial Bold",25)).grid(row=0, column=0)
frame3_l1 = tk.Label(frame3, text="ENTER EMAIL AND PASSWORD", font=("Arial Bold",25)).place(x=0,y=0)

#username label and text entry box
usernameLabel = tk.Label(frame3, text="E-mail").grid(row=1, column=0)
username = tk.StringVar()
usernameEntry = tk.Entry(frame3, textvariable=username).grid(row=1, column=1)

#password label and password entry box
passwordLabel = tk.Label(frame3,text="Password").grid(row=2, column=0)
password = tk.StringVar()
passwordEntry = tk.Entry(frame3, textvariable=password, show='*').grid(row=2, column=1)

validateLogin = partial(validateLogin, username, password)

#login button
loginButton = tk.Button(frame3, text="send mail", command=validateLogin).grid(row=4, column=0)

#frame3_nextbt= tk.Button(frame3,text="Next page", padx=50,command=lambda: show_frame(frame3)).place(x=600,y=500)
frame3_prevbt= tk.Button(frame3,text="previous page", padx=50,command=lambda: show_frame(frame2)).place(x=50,y=500)




#--------------------------------------FRAME 4---------------------------------------

#frame4_nextbt= tk.Button(frame4,text="Next page", padx=50,command=lambda: show_frame(frame5).place(x=600,y=500)
#frame4_prevbt= tk.Button(frame4,text="previous page", padx=50,command=lambda: show_frame(frame3)).place(x=50,y=500)


show_frame(frame1)

window.rowconfigure(0,weight=1)
window.columnconfigure(0,weight=1)

window.mainloop()