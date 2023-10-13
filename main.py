# Importing required modules
import datetime
import time
import webbrowser
from tkinter import *
import tkinter.messagebox as msg
from win10toast import ToastNotifier
from playsound import playsound
from win32com.client import Dispatch

'''For fetching current date'''

current_time = datetime.datetime.now()

'''Function for Playing Alarm'''

def play_alarm(set_alarm_time, set_alarm_date, name):
    while True:
        time.sleep(1)
        current_time = datetime.datetime.now()
        now = current_time.strftime("%H:%M:%S")
        date = current_time.strftime("%d/%m/%Y")
        if (date == set_alarm_date) and (now == set_alarm_time):
            
            str = f"Today is {name}"
            speaker = Dispatch("SAPI.SpVoice")

            if "Birthday" in name or "birthday" in name or "BIRTHDAY" in name:
                playsound('music/happy_birthday.mp3')
                toast = ToastNotifier()
                toast.show_toast(f"{name}\U0001F37E", f"This is to notify you that\nToday is {name}!\nWish them your great wishes", icon_path="icon/alarm.ico", duration=10)


            elif "Anniversary" in name or "ANNIVERSARY" in name or "anniversary" in name:
                playsound('music/anniversary.mp3')
                toast = ToastNotifier()
                toast.show_toast(f"{name}\U0001F382", f"This is to notify you that\nToday is {name}!\nWish them your great wishes", icon_path="icon/alarm.ico", duration=10)

            else:
                playsound('music/Alarm.mp3')
                toaster = ToastNotifier()
                toaster.show_toast(f"{name}", f"This is to notify you that\nToday is {name}!", icon_path="icon/alarm.ico", duration=10)

            speaker.speak(str)
            break

'''Function for Processing input and set alarm'''

def set_alarm():
    name = eventname.get()
    userdate = datevalue.get()
    usermonth = monthvalue.get()
    useryear = yearvalue.get()
    userhour = hourvalue.get()
    userminute = minutevalue.get()
    usersecond = secondvalue.get()

    if not name:
        msg.showerror("Blank Field", "Event Name field is empty")

    if not userdate:
        msg.showerror("Blank Field", "Date field is empty")

    if not usermonth:
        msg.showerror("Blank Field", "Month field is empty")

    if not useryear:
        msg.showerror("Blank Field", "Year field is empty")

    if not userhour:
        msg.showerror("Blank Field", "Hour field is empty")

    if not userminute:
        msg.showerror("Blank Field", "Minute field is empty")

    if not usersecond:
        msg.showerror("Blank Field", "Second field is empty")

    if (int(userdate) < 1 or int(userdate) > 31):
        msg.showerror("Invalid Date","Date entered is either greater than 31 or less than 1")

    if (int(usermonth) < 1 or int(usermonth) > 12):
        msg.showerror("Invalid Month","Month entered is either greater than 12 or less than 1")

    if (int(useryear) < current_time.year):
        msg.showerror("Invalid Year", f"Input year is less than {current_time.year}")

    if (int(useryear) > (current_time.year+5)):
        msg.showerror("Invalid Year", f"Input year is too larger to set alarm")

    if (int(userhour) < 0 or int(userhour) > 23):
        msg.showerror("Invalid Hour","Hour entered is either greater than 23 or less than 00")

    if (int(userminute) < 0 or int(userminute) > 59):
        msg.showerror(
            "Invalid Minute", "Minute entered is either greater than 59 or less than 00")

    if (int(usersecond) < 0 or int(usersecond) > 59):
        msg.showerror(
            "Invalid Second", "Second entered is either greater than 59 or less than 00")

    if (name) and (userdate) and (usermonth) and (useryear) and (userhour) and (userminute) and (usersecond):
        setting()
        set_alarm_time = f"{userhour}:{userminute}:{usersecond}"
        set_alarm_date = f"{userdate}/{usermonth}/{useryear}"
        play_alarm(set_alarm_time, set_alarm_date, name)
        exit(0)


'''Instagram Linker'''

def insta():
    url = "https://www.instagram.com/tanishq_buddy/"
    webbrowser.open_new(url)


'''Linkedin Linker'''

def linkedin():
    url = "https://www.linkedin.com/in/tanish-gupta-85a440211/"
    webbrowser.open_new(url)

'''Setting Alarm Notification'''
def setting():
    status.set("Setting Alarm...")
    sbar.update()
    time.sleep(1)
    status.set("")


'''Check For update function'''

def checkup():
    status.set("Checking For Updates...")
    sbar.update()
    time.sleep(2)
    status.set("")
    msg.showinfo("Smart PC Alarm", "There are currently no updates available")


'''About Function'''

def about():
    msg.showinfo("About", "Smart PC Alarm : \n\n\n\nLanguage : Python\n\nDeveloped at : ITM GOI\n\nDevelopment year : 2022\n\nInstructed By : Mr. Vijay Boosa\n\nDeveloped By : Tanish Gupta\n\t          Utkarsh Sirohi\n\t          Sajal Nagaria\n\t          Tanishq Vyas")


'''Driver Code'''

root = Tk()
root.geometry("1265x670")
root.title(" Smart PC Alarm ")
root.wm_iconbitmap("icon/alarm.ico")
root.config(background="white")

'''Menu of the Interface'''

menubar = Menu(root)
help = Menu(menubar, tearoff=0)
help.add_command(label="Check for Updates...", command=checkup)
help.add_separator()
help.add_command(label="About", command=about)
menubar.add_cascade(label="Help", menu=help)
menubar.add_cascade(label="Exit", command=quit)
root.config(menu=menubar)

'''System's Main Interface'''

Label(text="Welcome to Smart PC Alarm", background="white",fg="black", font="Castellar 35 bold").place(x=180, y=10)

alarm_image = PhotoImage(file='images/alarm.png')
Label(root, image=alarm_image, background="white").place(x=550, y=80)

alarm_bell = PhotoImage(file='images/alarm_bell.png')
Label(root, image=alarm_bell, background="white").place(x=240, y=150)

alarm_clock = PhotoImage(file='images/alarm_clock.png')
Label(root, image=alarm_clock, background="white").place(x=950, y=450)

alarm_device = PhotoImage(file='images/alarm_device.png')
Label(root, image=alarm_device, background="white").place(x=230, y=450)

alarm_digital = PhotoImage(file='images/alarm_digital.png')
Label(root, image=alarm_digital, background="white").place(x=430, y=550)

alarm_noti = PhotoImage(file='images/alarm_noti.png')
Label(root, image=alarm_noti, background="white").place(x=950, y=150)

alarm_pc = PhotoImage(file='images/alarm_pc.png')
Label(root, image=alarm_pc, background="white").place(x=1100, y=300)

alarm_ring = PhotoImage(file='images/alarm_ring.png')
Label(root, image=alarm_ring, background="white").place(x=780, y=550)

alarm_snooz = PhotoImage(file='images/alarm_snooz.png')
Label(root, image=alarm_snooz, background="white").place(x=80, y=270)

Label(root, text="Event Name", background="white", fg="black", font="Forte 16").place(x=400, y=186)

Label(root, text="Date", background="white", fg="black", font="Forte 16").place(x=400, y=226)

Label(root, text="Month", background="white", fg="black", font="Forte 16").place(x=400, y=266)

Label(root, text="Year", background="white", fg="black", font="Forte 16").place(x=400, y=306)

Label(root, text="Hour", background="white", fg="black", font="Forte 16").place(x=400, y=366)

Label(root, text="Minute", background="white", fg="black", font="Forte 16").place(x=400, y=406)

Label(root, text="Second", background="white", fg="black", font="Forte 16").place(x=400, y=446)

'''variable declaration for different user inputs'''

eventname = StringVar()

datevalue = StringVar()
datevalue.set(current_time.strftime("%d"))

monthvalue = StringVar()
monthvalue.set(current_time.strftime("%m"))

yearvalue = StringVar()
yearvalue.set(current_time.strftime("%Y"))

hourvalue = StringVar()
hourvalue.set(current_time.strftime("%H"))

minutevalue = StringVar()
minutevalue.set(current_time.strftime("%M"))

secondvalue = StringVar()
secondvalue.set(current_time.strftime("%S"))

numbervalue = StringVar()

'''Entries to take input from user'''

Evententry = Entry(root, textvariable=eventname, border=1,borderwidth=1, background="white", fg="black", font="comicsans 16", insertbackground="grey", insertwidth=2).place(x=600, y=186)

dateentry = Entry(root, textvariable=datevalue, border=1,borderwidth=1, background="white", fg="black", font="comicsans 16", insertbackground="grey", insertwidth=2).place(x=600, y=226)

monthentry = Entry(root, textvariable=monthvalue, border=1,borderwidth=1,background="white", fg="black", font="comicsans 16", insertbackground="grey", insertwidth=2).place(x=600, y=266)

yearentry = Entry(root, textvariable=yearvalue, border=1,borderwidth=1, background="white", fg="black", font="comicsans 16", insertbackground="grey", insertwidth=2).place(x=600, y=306)

hourentry = Entry(root, textvariable=hourvalue, border=1,borderwidth=1, background="white", fg="black", font="comicsans 16", insertbackground="grey", insertwidth=2).place(x=600, y=366)

minuteentry = Entry(root, textvariable=minutevalue, border=1,borderwidth=1, background="white", fg="black", font="comicsans 16", insertbackground="grey", insertwidth=2).place(x=600, y=406)

secondeentry = Entry(root, textvariable=secondvalue, border=1,borderwidth=1, background="white", fg="black", font="comicsans 16", insertbackground="grey", insertwidth=2).place(x=600, y=446)

'''Set Alarm Button'''

Button(text="Set Alarm", background="black", cursor="hand2", fg="white", command=set_alarm, font="Forte 16").place(x=600, y=506)

'''Instagram Button'''

insta_btn = PhotoImage(file='images/instagram.png')
my_btn2 = Button(root, image=insta_btn, borderwidth=0, background="white", cursor="hand2", command=insta)
my_btn2.place(x=1170, y=625)

'''Linkedin Button'''

linkedin_btn = PhotoImage(file='images/linkedin.png')
my_btn3 = Button(root, image=linkedin_btn, borderwidth=0, background="white", cursor="hand2", command=linkedin)
my_btn3.place(x=1220, y=625)

'''Status Bar'''

status = StringVar()
status.set("")
sbar = Label(root, textvariable=status, width=165, background="white", fg="grey", font="comicsans 9")
sbar.pack(side=BOTTOM, anchor=NW)

root.resizable(False, False)
root.mainloop()