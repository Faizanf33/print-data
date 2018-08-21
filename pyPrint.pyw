### Imports go here
import ctypes
import logging, os, re
import time
try:
    import win32api, win32print
except:
    pass

import tkinter.messagebox as tkmb
import tkinter.simpledialog as tksd
from tkinter import *
from tkinter import ttk

## Imports from files
from Data.tool_tip import CreateToolTip
from Data.set_data import Delegate
from Data.send_mail import Mail

## To run the program explicitly
myappid = 'mycompany.myproduct.subproduct.version' # arbitrary string
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)


## Setting up debugging
# Level set to logging.Debug
try:
    Path = os.path.join('temp', 'info.log')
    logging.basicConfig(format="[NYMUN Forms]:[%(asctime)s]:%(levelname)s:%(message)s", datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.DEBUG, filename=Path)

except (IOError or FileNotFoundError):
    os.mkdir('temp')
    Path = os.path.join('temp', 'info.log')
    logging.basicConfig(format="[NYMUN Forms]:[%(asctime)s]:%(levelname)s:%(message)s", datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.DEBUG, filename=Path)


users_info = {}
admin = ""
status = False

def start_server():
    window.withdraw()
    id = tksd.askstring("Credentials", "EMAIL-ID :", parent=frame)
    if id:
        if not re.match(r"^[A-Za-z0-9\.\+_-]+@[A-Za-z0-9\._-]+\.[a-zA-Z]*$", id):
            terminal.insert(END, ">> Invalid EMAIL_ID Entered!")
            tkmb.showwarning("INVALID ID", "EMAIL_ID IS INVALID!\nTRY AGAIN...")
            window.deiconify()
            return

        else:
            passwrd = tksd.askstring("Credentials", "PASSWORD :", show="\u2022", parent=frame)
            if passwrd:
                terminal.insert(END, ">> Credentials received!")
                global admin, status
                admin = Mail(id, passwrd)
                window.deiconify()
                status = admin.server_status
                if status == True:
                    terminal.insert(END, ">> Logged in successfully")
                    terminal.insert(END, ">> Using Id {} at server {}".format(admin.user_id, admin.server))
                    b1.configure(text="RESTART SERVER", command=start_server)
                    b1_ttp = CreateToolTip(b1, "RESTART MAIL SERVER")

                elif admin.start_service() == True:
                    mainEntry.focus()
                    terminal.insert(END, ">> EMAIL_ID/PASSWORD INCORRECT")
                    return False

                else:
                    mainEntry.focus()
                    terminal.insert(END, ">> ____CONNECTION-TIMEDOUT_____")
                    return False
            else:
                window.deiconify()
                terminal.insert(END, ">> Credentials were not provided!")

    else:
        window.deiconify()
        terminal.insert(END, ">> Credentials were not provided!")

    mainEntry.focus()
    return False

def set_data():
    try:
        file = "NYMUN Multan.csv"
        logging.debug("File found => '{}'".format(file))
        new = Delegate(file)
        global users_info
        users_info = new.info
        terminal.insert(END, ">> File found => '{}'".format(file))
        terminal.insert(END, ">> {0} Files written in directory -> {1}".format(new.files_created, new.dir))
        terminal.insert(END, ">> On Your Desktop, Look For '{}' Directory".format(new.dir))
        terminal.insert(END, ">> Process completed with {} errors".format(new.errors))
        reply =  tkmb.askquestion("EMAIL", "START MAIL SERVER?")
        if reply == 'yes':
            terminal.insert(END, "\n>> Mail request is accepted!")
            start_server()
            if status == False:
                b1.configure(text="START SERVER", command=start_server)
                b1_ttp = CreateToolTip(b1, "START MAIL SERVER")
                return

        else:
            b1.configure(text="START SERVER", command=start_server)
            b1_ttp = CreateToolTip(b1, "START MAIL SERVER")
            terminal.insert(END, ">> Request for mail was rejected!")

        mainEntry.focus()
        return

    except FileNotFoundError:
        logging.error("File not found => '{}'".format(file))
        logging.info("No further proceedings!")
        terminal.insert(END, ">> File -> '{}' Not Found!".format(file))
        terminal.insert(END, ">> No further proceedings!")
        mainEntry.focus()
        return

def find_serial():
    if users_info == {}:
        tkmb.showwarning("DATA UN-ARRANGED", "Please arrange your data first!")
        return

    prov_info = tksd.askstring("FIND", "EMAIL_ID/PHONE NO.", parent=frame)
    if prov_info:
        terminal.insert(END, ">> Searching record for : {}".format(prov_info))
        if prov_info.isdigit():
            prov_info = prov_info[1:]

        for serial in users_info.keys():
            if prov_info in users_info[serial]:
                tkmb.showinfo("PASSED", "RECORD FOUND!\nSERIAL NO. : {}".format(serial))
                terminal.insert(END, ">> Record matched with serial no {}".format(serial))
                return

            else:
                pass

        terminal.insert(END, ">> No match found!")
        tkmb.showinfo("FAILED", "NO RECORD FOUND!")
        return

    else:
        terminal.insert(END, ">> Search request cancelled!")
        return

def send_mail(serial_no):
    if serial_no.upper() in users_info.keys():
        address = users_info[serial_no][7]
        name = users_info[serial_no][3]
        no_of_delegate = users_info[serial_no][2]
        if no_of_delegate == "Individual":
            amount = str(3500) + str(.0)

        else:
            amount = str(2500 + (int(no_of_delegate) * 2500)) + str(.0)

        msg = "Dear {0}! \n\nWe have received a payment of Rs {1} for NYMUN Registration against your team Id : {2}. \n\nIn case of any queries, reply here\n\nMr. Faraz Sheikh ( +92 306 8635633 )\n\n\n\nRegards\nRegistraton Team NYMUN".format(name, amount, serial_no.upper())
        terminal.insert(END, ">> Sending mail at {}".format(address))
        try:
            admin.sendmail(address, msg)
            terminal.insert(END, ">> Mail sent successfully.")
        except Exception:
            terminal.insert(END, ">> Unable to send mail, try restarting server!")
    else:
        terminal.insert(END, ">> Unable to send mail! serial_no {} not found".format(serial_no))

    return


def print_file(event=None):
    path = os.getcwd()
    serial_no = mainEntry.get()
    serial_no = serial_no.upper()

    if path == initial_path:
        mainEntry.delete(0, END)
        b1.focus()
        tkmb.showwarning("DATA UN-ARRANGED", "Please arrange your data first!")
        return

    if serial_no == "":
        tkmb.showwarning("NO SERIAL NO.", "PLEASE ENTER A SERIAL NO. FIRST!")
        return

    elif not serial_no.startswith("S#"):
        tkmb.showwarning("INVALID SERIAL NO.", "PLEASE ENTER A VALID SERIAL NO.\ne.g. s#0002 or S#0010 ...")
        return


    else:
        my_file = "{}.pdf".format(serial_no)
        directories = ["Individual", 3, 4, 5, 6]
        mainEntry.delete(0, END)
        mainEntry.focus()
        terminal.insert(END, ">> Searching for file with Serial No. {}".format(serial_no))
        for root, dirs, files in os.walk(path):
            if my_file in files:
                path_to_doc = os.path.join(root, my_file)
                terminal.insert(END, ">> File found at {}".format(path_to_doc))
                try:
                    win32api.ShellExecute (0, "print", path_to_doc, None, '/d:"{}"'.format(win32print.GetDefaultPrinter()), 0)
                    logging.info("File {} sent for printing".format(my_file))
                    terminal.insert(END, ">> File {} sent for printing".format(my_file))
                    os.chdir(path)
                    with open("PrintRecords.txt", "a") as inFile:
                        text = "File '{0}' , Print Time : '{1}' , Printed by : '{2}'\n".format(my_file, time.strftime('%c'), admin_name)
                        inFile.write(text)

                    if status == True:
                        reply =  tkmb.askquestion("EMAIL", "SEND CONFIRMATION MAIL?")
                        if reply == "yes":
                            send_mail(serial_no)

                        else:
                            pass

                except Exception as exc:
                    logging.warning("While printing file {} error was raised => {}".format(my_file, exc))
                    terminal.insert(END, ">> Unable to print file {}".format(my_file))
                    terminal.insert(END, ">> Please make sure your system has 'ADOBE READER' installed!")

                return

            else:
                pass
            os.chdir("..")

        terminal.insert(END, ">> No such file with Serial No. {}".format(serial_no))
        terminal.insert(END, ">> Recheck serial no. and try again..")
        os.chdir(path)
        return

def exit():
    path = os.getcwd()

    if path == initial_path:
        logging.info('exiting window')
        window.destroy()
        window.quit()

    else:
        window.withdraw()
        if status == True:
            msg = tkmb.askquestion("CONFIRM","ARE YOU SURE YOU WANT TO EXIT?\n\tServer still running...", icon='info')
        else:
            msg = tkmb.askquestion("CONFIRM","ARE YOU SURE YOU WANT TO EXIT?\n\tAll progress will be lost!", icon='info')

        if msg == "yes":
            logging.info('exiting window')
            window.destroy()
            window.quit()

        else:
            logging.info('window still running')
            window.deiconify()
            return

def tick():
    # get the current local time from the PC
    time2 = time.strftime('%I:%M:%S %p')
    # if time string has changed, update it
    clock.config(text=time2)
    # calls itself every 200 milliseconds
    # to update the time display as needed
    # could use >200 ms, but display gets jerky
    clock.after(200, tick)


### Creating windows
window = Tk()

window.title("NYMUN FORMS")

### Setting up windows geometry
w, h = 360, 360
## to open window in the centre of screen
ws = window.winfo_screenwidth()
hs = window.winfo_screenheight()
x_axis = (ws/2) - (w/2)
y_axis = (hs/2) - (h/2)

window.geometry('%dx%d+%d+%d' % (w, h, x_axis, y_axis))
window.resizable(0,0)

### remember initial_path
initial_path = os.getcwd()
### adding window icon
try:
    window.iconbitmap("nymun.ico")
except:
    pass

# if windows default cross button is pressed
window.protocol('WM_DELETE_WINDOW',exit)

### Creating Frames
frame = Frame(window, bg="bisque2", bd=4, relief=SUNKEN, colormap="new", height=h)
frame.pack(fill=BOTH, side=TOP, padx=0, pady=0, ipadx=0, ipady=0)
### Creating Frame2
frame2 = Frame(window, bg="bisque2", bd=4, relief=SUNKEN, colormap="new", height=h)
frame2.pack(fill=BOTH, side=TOP, padx=0, pady=0, ipadx=0, ipady=0)

### Creating Button
ttk.Style().configure("TButton", padding=4, borderwidth=8, relief=RAISED, background="cyan2", font=('times', 12, 'bold'))
b1 = ttk.Button(frame, text="ARRANGE DATA", command=set_data)
b1.pack(side=TOP, padx=6, ipadx=0, pady=6, ipady=0)
b1_ttp = CreateToolTip(b1, "ARRANGE YOUR DATA TO pdf's")

### Creating clock
clock = Label(frame, font=('times', 12), bg="bisque2", fg='grey1')
clock.place(x=5, y=12)
### calling time function
tick()

### Creating day
day = Label(frame, text=time.strftime('%A'), font=('times', 12), bg="bisque2", fg='grey1')
day.place(x=258, y=1)

### Creating date
date = Label(frame, text=time.strftime('%d-%b-%y'), font=('times', 12), bg="bisque2", fg='grey1')
date.place(x=258, y=22)

## Creating other Buttons
searchButton = Button(frame2, command=find_serial, relief=GROOVE,text="SEARCH", width=8, height=0, font=('times', 12, 'bold'), bg="light cyan", fg="black", bd=3, activebackground="pale turquoise", activeforeground="black")
searchButton.place(x=15, y=49)
prB_ttp = CreateToolTip(searchButton, "SEARCH SERIAL NO.")


printButton = Button(frame2, command=print_file, relief=GROOVE,text="PRINT", width=8, height=0, font=('times', 12, 'bold'), bg="light cyan", fg="black", bd=3, activebackground="pale turquoise", activeforeground="black")
printButton.place(x=165, y=49)
printButton.bind("<Return>", print_file)
prB_ttp = CreateToolTip(printButton, "PRINT YOUR FILE")

exitButton = Button(frame2, command=exit, relief=GROOVE, text="EXIT", width=8, height=1, font=('times', 12, 'bold'), bg="brown3", fg="white", bd=3, activebackground="brown4", activeforeground="white")
exitButton.pack(side=BOTTOM, anchor=E, padx=6, ipadx=0, pady=4, ipady=0)
exit_ttp = CreateToolTip(exitButton, "CLOSE WINDOW")

### Creating message box
scroll = Scrollbar(frame, bd=4, relief=RIDGE)
scroll.pack(side=RIGHT, anchor=E, fill=Y)
terminal = Listbox(frame, yscrollcommand = scroll.set, bg="grey22", fg="white", font=('times', 12), height=8, width=41, bd=5, relief=SUNKEN)
terminal.pack(side=LEFT, anchor=W, pady=0, ipadx=0, padx=0, ipady=0)
scroll.config(command=terminal.yview)

### Creating Labels
mainLabel = Label(frame2, text="Enter Serial No. :\t\t", font="times 12 bold", bg="bisque3", relief=RIDGE, bd=2)
mainLabel.pack(side=LEFT, anchor=SW, pady=8, ipadx=0, ipady=0, fil=BOTH)

### Creating entry
infoEntry = StringVar()
infoEntry.set("e.g. s#0000")
mainEntry = Entry(frame2, textvariable=infoEntry, font="times 12", bd=4, bg="white", fg="black", relief=RIDGE)
mainEntry.focus()
mainEntry.select_range(0, END)
mainEntry.pack(side=RIGHT, anchor=E, pady=8, ipadx=0, ipady=0)
mE_ttp = CreateToolTip(mainEntry, "TYPE HERE, e.g. S#0001")
mainEntry.bind("<Return>", print_file)


### Copyright label
cplabel = Label(window, text="NYMUN FORMS {} 2018".format(chr(0xa9)), font="chiller 12 bold", bg="bisque4", fg="grey2", relief=SUNKEN, bd=5)
cplabel.pack(side=BOTTOM, anchor=SW, fill=X, padx=0, pady=0, ipadx=0, ipady=0)

### Admin name
window.withdraw()
admin_name = tksd.askstring("Greets", "NAME PLEASE...", parent=window)
if not admin_name:
    admin_name = "ADMIN"
window.deiconify()

window.mainloop()
