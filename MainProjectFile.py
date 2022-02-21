# default ID pass are :
# UserID : GUEST
# Password : GUEST

# tkinter imports for basic GUI elements
from tkinter import *
from tkinter import simpledialog
import tkinter as tk
from tkinter import messagebox
# Pandas library to work with Excel as a database
import pandas as pd
# python imaging library(PIL) to process images for buttons and backgrounds of windows
from PIL import ImageTk, Image
# Random import to generate random OTPs and Account Numbers
import random
# SMTP server library to send emails
import smtplib
from email.mime.text import MIMEText
# date and time imports for transaction history
from pytz import timezone
from datetime import datetime
# to speak about transactions
from gtts import gTTS
import pyttsx3
# to remove the uncessary files
import os


# please specify the path where this folder is kept in :
# for example
# path = r'C:\Users\habha\AppData\Local\Programs\Python\Python310\CodeFiles\Tkinter\CSEprojectRelated\\'
# make sure to add an extra '\' character for the path to work properly

path = r'C:\Users\habha\Downloads\3-AU2120193-AU2140080-AU2140083-AU2140084-BankManagementSystem\\'

# main Tkinter window
root = Tk()

# to display icon on the window
root.iconbitmap(path+'bank.ico')

# we need to make a dynamic window to open in full screen mode according to the screen
heightWindow = root.winfo_screenheight()
widthWindow = root.winfo_screenwidth()
# here +0+0 specifies the window must be placed at the origin of the screen
strForGeometry = str(widthWindow) + "x" + str(heightWindow) + "+-10+0"

# setting the size of the calculator window
root.geometry(strForGeometry)
# root.attributes('-fullscreen',True) #used to fullscreen the window

# to change the title of your window you can use title function
root.title("BMS")

# image objects here
im = Image.open(path + 'backMain.png')
im = im.resize((widthWindow, heightWindow), Image.ANTIALIAS)
# the main background image will dynamically resize depending on the size of the laptop
ph = ImageTk.PhotoImage(im)
ph1 = PhotoImage(file=path + "history.png")
ph2 = PhotoImage(file=path + "deposit.png")
ph3 = PhotoImage(file=path + "accbal.png")
ph4 = PhotoImage(file=path + "withdraw.png")
ph5 = PhotoImage(file=path + "createNewButton.png")
ph6 = PhotoImage(file=path + "oldLogin.png")
ph7 = PhotoImage(file=path + "RewindButton.png")
ph8 = PhotoImage(file=path + "update.png")


# values dictionary to store all the users value in Excel sheet using python
values = {}


# method to create the frames as they were preloading in background which was annoying to deal with
def CreateFrame(frameNO):
    #global variables
    global frame1, frame2, frame3, frame4

    print("Creating Frame", frameNO)
    if (frameNO == 1):
        # Frame 1 for outer border line
        frame1 = LabelFrame(root, bg="#050a30",
                            width=widthWindow, height=heightWindow)
        frame1.pack(pady=5, padx=5)

    elif (frameNO == 2):

        # Frame 2 for Account Creation Window
        frame2 = LabelFrame(root, bg="#5cb6f9",
                            width=widthWindow, height=heightWindow)
        frame2.pack(pady=5, padx=5)

    elif (frameNO == 3):

        # Frame 3 for NewAccount Money deposition window
        frame3 = LabelFrame(root, bg="#5cb6f9",
                            width=widthWindow, height=heightWindow)
        frame3.pack(pady=5, padx=5)

    elif (frameNO == 4):

        # Frame 4 for Old user to do Verify its credentials
        frame4 = LabelFrame(root, bg="#5cb6f9",
                            width=widthWindow, height=heightWindow)
        frame4.pack(pady=5, padx=5)

    return

# function to switch between windows which are bascially independent and can be called anytime


def switchWindows(i):
    print("Switching windows", i)
    if(i == 1 or i == 2):
        frame2.destroy()
        loginWindow()
    elif(i == 3):
        frame4.destroy()
        mainTransactFun()
    elif(i == 4):
        frame1.destroy()
        mainTransactFun()
    elif(i == 5):
        frame4.destroy()
        loginWindow()
    elif(i == 6):
        frame3.destroy()
        loginWindow()
    return

# this function is called when the user logs in successfully the default variables need to be initialized by the application old user data stored in excel


def loadUserData(i):
    print("Loading user %d data from Excel" % (i))
    # Global variables
    global ii, firstName, lastName, middleName, DOB, Email, mobileNO, gender, aadharNO, AccNO, currentBal, passWord

    # Opening Excel to fetch and load Old users Data
    EXCEL_FILE = path+'BMS.xlsx'
    df = pd.read_excel(EXCEL_FILE)
    ii = i
    firstName = df.loc[i, 'firstName']
    lastName = df.loc[i, 'lastName']
    middleName = df.loc[i, 'middleName']
    DOB = df.loc[i, 'DOB']
    Email = df.loc[i, 'Email']
    mobileNO = df.loc[i, 'mobileNO']
    gender = df.loc[i, 'gender']
    aadharNO = df.loc[i, 'aadharNO']
    AccNO = df.loc[i, 'AccNO']
    currentBal = df.loc[i, 'currentBal']
    passWord = df.loc[i, 'passWord']
    print('Users index is', i)
    return


# Account Creation Window to get the users data(like Asking user Name,DOB ... etc)
def OnClickCreateAccount():
    print("Account Creation Started")
    # Global variables
    global nameEntryBox, DOBEntryBox, emailEntryBox, mobileEntryBox, addharEntryBox, gen, newPassEntrybox, AccNO, openingCreBox, currentBal

    # This will destroy the Exisitng frame1 which was for loading screen
    frame1.destroy()

    # creates the frame2 Again
    CreateFrame(2)
    # Creating the Title of the account details window
    accountOpeningLabel = Label(frame2, text="Welcome to BMS\n Please add the following details to continue",
                                width=100, height=5, font=("BankGothic Md BT ", 20, "bold"), bg="#cae8ff", fg="#233dff")
    accountOpeningLabel.place(y=-300, relx=0.5, rely=0.5, anchor=CENTER)

    # First Entry box to Enter Name
    nameEntryBox = Entry(frame2, font=("BankGothic Md BT", 15, "bold"))
    nameEntryBox.place(y=-50, x=50, relx=0.5, rely=0.5,
                       anchor=CENTER, width=500, height=30)
    nameEntryBox.insert(0, "FirstName FathersName LastName")

    # Name Input Label
    nameInputLabel = Label(frame2, text="NAME :", padx=10, pady=5, bg="#233dff", fg="#cae8ff",
                           font=("Tw Cen MT Condensed Extra Bold", 15))
    nameInputLabel.place(y=-50, x=-250, relx=0.5, rely=0.5, anchor=CENTER)

    # DOB Entry box to Enter Name
    DOBEntryBox = Entry(frame2, font=("BankGothic Md BT", 15, "bold"))
    DOBEntryBox.place(y=0, x=50, relx=0.5, rely=0.5,
                      anchor=CENTER, width=500, height=30)
    DOBEntryBox.insert(0, "DD/MM/YYYY")

    # DOB Input Label
    DOBInputLabel = Label(frame2, text="DOB :", padx=16, pady=5, bg="#233dff", fg="#cae8ff",
                          font=("Tw Cen MT Condensed Extra Bold", 15))
    DOBInputLabel.place(y=0, x=-250, relx=0.5, rely=0.5, anchor=CENTER)

    # Email Entry box to Enter Name
    emailEntryBox = Entry(frame2, font=("BankGothic Md BT", 15, "bold"))
    emailEntryBox.place(y=50, x=50, relx=0.5, rely=0.5,
                        anchor=CENTER, width=500, height=30)
    emailEntryBox.insert(0, "JohnDoe@mail.com")

    # Email Input Label
    emailEntryLabel = Label(frame2, text="Email :", padx=10, pady=5, bg="#233dff", fg="#cae8ff",
                            font=("Tw Cen MT Condensed Extra Bold", 15))
    emailEntryLabel.place(y=50, x=-250, relx=0.5, rely=0.5, anchor=CENTER)

    # MobileNo Entry box to Enter Name
    mobileEntryBox = Entry(frame2, font=("BankGothic Md BT", 15, "bold"))
    mobileEntryBox.place(y=100, x=50, relx=0.5, rely=0.5,
                         anchor=CENTER, width=500, height=30)
    mobileEntryBox.insert(0, "XXXXXXXXXX")

    # MobileNo Input Label
    mobileEntryLabel = Label(frame2, text="MobileNo :", padx=14, pady=5, bg="#233dff", fg="#cae8ff",
                             font=("Tw Cen MT Condensed Extra Bold", 15))
    mobileEntryLabel.place(y=100, x=-270, relx=0.5, rely=0.5, anchor=CENTER)

    # Gender radio button
    gen = StringVar(frame2)
    gen.set("Gender")
    genderList = OptionMenu(frame2, gen, "Male", "Female", "Other")
    genderList.place(y=150, x=-50, relx=0.5, rely=0.5,
                     anchor=CENTER, width=300, height=30)

    # Gender Input Label
    genderLabel = Label(frame2, text="Gender :", padx=5, pady=5, bg="#233dff", fg="#cae8ff",
                        font=("Tw Cen MT Condensed Extra Bold", 15))
    genderLabel.place(y=150, x=-250, relx=0.5, rely=0.5, anchor=CENTER)

    # aadharNO Entry box to Enter Name
    addharEntryBox = Entry(frame2, font=("BankGothic Md BT", 15, "bold"))
    addharEntryBox.place(y=200, x=50, relx=0.5, rely=0.5,
                         anchor=CENTER, width=500, height=30)
    addharEntryBox.insert(0, "XXXX XXXX XXXX")

    # aadharNO Input Label
    addharEntryLabel = Label(frame2, text="AadharNO :", padx=5, pady=5, bg="#233dff", fg="#cae8ff",
                             font=("Tw Cen MT Condensed Extra Bold", 15))
    addharEntryLabel.place(y=200, x=-261, relx=0.5, rely=0.5, anchor=CENTER)

    # password making field
    newPassWordLabel = Label(frame2, text="password :", padx=5, pady=5, bg="#233dff", fg="#cae8ff",
                             font=("Tw Cen MT Condensed Extra Bold", 15))
    newPassWordLabel.place(y=250, x=-261, relx=0.5, rely=0.5, anchor=CENTER)

    # password entry box
    newPassEntrybox = Entry(frame2, font=("BankGothic Md BT", 15, "bold"))
    newPassEntrybox.place(y=250, x=50, relx=0.5, rely=0.5,
                          anchor=CENTER, width=500, height=30)
    newPassEntrybox.insert(0, "min 8 digit password")

    # opening credit entry box
    openingCreBox = Entry(frame2, font=("BankGothic Md BT", 15, "bold"))
    openingCreBox.place(y=300, x=50, relx=0.5, rely=0.5,
                        anchor=CENTER, width=500, height=30)
    openingCreBox.insert(0, "deposit min Rs 2500 ")

    # Opening  Credit label
    openingCreLabel = Label(frame2, text="Opening Credit :", padx=5, pady=5, bg="#233dff", fg="#cae8ff",
                            font=("Tw Cen MT Condensed Extra Bold", 15))
    openingCreLabel.place(y=300, x=-285, relx=0.5, rely=0.5, anchor=CENTER)

    # Accout number  to be created randomly
    AccNO = random.randint(10 ** 12, (10 ** 13) - 1)

    # Submit Button after completely adding in the data
    submitButton = Button(frame2, text="SUBMIT", font=("Tw Cen MT Condensed Extra Bold", 15),
                          padx=30, pady=10, width=7, height=1, command=verifyNewAccountDetails)
    submitButton.place(y=370, x=-100, relx=0.5, rely=0.5, anchor=CENTER)
    # Back button
    RewindTimeButton = Button(frame2, text="BACK", command=lambda: switchWindows(1), image=ph7,
                              font=("Modern No.", 11, "bold"), bg="#5cb6f9",  fg="white")
    RewindTimeButton.place(height=100, width=100, relx=0.5,
                           rely=0.5, anchor=CENTER, x=500, y=300)

    # Resets the Value in the Entry boxes

    def Reset():
        nameEntryBox.delete(0, END)
        nameEntryBox.insert(0, "FirstName FathersName LastName")
        DOBEntryBox.delete(0, END)
        DOBEntryBox.insert(0, "DD/MM/YYYY")
        emailEntryBox.delete(0, END)
        emailEntryBox.insert(0, "JohnDoe@mail.com")
        mobileEntryBox.delete(0, END)
        mobileEntryBox.insert(0, "XXXXXXXXXX")
        gen.set("Gender")
        addharEntryBox.delete(0, END)
        addharEntryBox.insert(0, "XXXX XXXX XXXX")
        newPassEntrybox.delete(0, END)
        newPassEntrybox.insert(0, "min 8 digit password")
        openingCreBox.delete(0, END)
        openingCreBox.insert(0, "deposit min Rs 2500")

    # Reset button calls the reset function to reset all the input fields
    resetButton = Button(frame2, text="RESET", font=("Tw Cen MT Condensed Extra Bold", 15),
                         padx=30, pady=10, width=7, height=1, command=Reset)
    resetButton.place(y=370, x=100, relx=0.5, rely=0.5, anchor=CENTER)

    return

# the data stored in loaded OnClickCreateAccount function is to be verified by this function


def verifyNewAccountDetails():
    print("verifying NewAccount Details")
    # Global variables
    global firstName, lastName, middleName, DOB, Email, mobileNO, gender, aadharNO, passWord, currentBal, ii

    fullName = nameEntryBox.get()
    try:
        firstName, lastName, middleName = map(str, fullName.split(' '))
    except:
        messagebox.showerror("Error occured ", "Enter correct formatwise name")
        return
    DOB = DOBEntryBox.get()
    mobileNO = mobileEntryBox.get()
    aadharNO = addharEntryBox.get()
    passWord = newPassEntrybox.get()
    currentBal = openingCreBox.get()

    if(len(mobileNO) != 10):
        messagebox.showerror(
            "Error occured ", "Mobile number number length is incorrect")
        return

    try:
        curr = int(mobileNO)
    except:
        messagebox.showerror("Error occured ", "Enter correct MobileNO")
        return

    if(len(passWord) < 8):
        messagebox.showerror(
            "Error occured ", "Password number length is incorrect")
        return
    if(len(str(aadharNO)) != 14):
        messagebox.showerror(
            "Error occured ", "Addhar number number length is incorrect")
        return

    try:
        curr = int(currentBal)
    except:
        messagebox.showerror("Error occured ", "Enter deposit in Numbers")
        return

    if(int(currentBal) < 2500):
        messagebox.showerror(
            "Error occured ", "Minimum depoit of 2500Rs is required")
        return

    # After all checks are done this function is called to verify users email and then save the data
    onCLickSaveInfo()
    return


# Old Account Window   (Check Account Balance, Update , deposit and Delete ... etc)
def OnClickRedirectOldAcc():
    print("Welcome OLd User")

    # global variables
    global AccNOEntryBox, passWordEntrybox

    # This will destroy the Exisitng frame1 which was for loading screen
    frame1.destroy()
    CreateFrame(2)
    # Banner to welcome The old user
    TitleLabel = Label(frame2,
                       text="Welcome Again !!\n Please enter your details ",
                       width=100, height=5, font=("BankGothic Md BT ", 20, "bold"), bg="#cae8ff", fg="#233dff")
    TitleLabel.place(rely=0.5, relx=0.5, y=-250, anchor=CENTER)

    # Account no label to specify this to enter ID here
    AccNOlabel = Label(frame2, text="AccNO :", padx=20, pady=7, bg="#233dff", fg="#cae8ff",
                       font=("Tw Cen MT Condensed Extra Bold", 20))
    AccNOlabel.place(anchor=CENTER, relx=0.5, rely=0.5, x=-250, y=50)

    # Entry field to enter ID
    AccNOEntryBox = Entry(frame2, font=("BankGothic Md BT", 20, "bold"))
    AccNOEntryBox.place(y=50, x=25, relx=0.5, rely=0.5,
                        anchor=CENTER, width=400, height=50)
    AccNOEntryBox.insert(0, "8 digit ID")

    # passWord label to specify this to enter password
    passWordLabel = Label(frame2, text="Password :", padx=20, pady=7, bg="#233dff", fg="#cae8ff",
                          font=("Tw Cen MT Condensed Extra Bold", 20))
    passWordLabel.place(anchor=CENTER, relx=0.5, rely=0.5, x=-260, y=120)

    # PassWord EntryBox to enter Password
    passWordEntrybox = Entry(frame2, font=("BankGothic Md BT", 20, "bold"))
    passWordEntrybox.place(y=120, x=25, relx=0.5, rely=0.5,
                           anchor=CENTER, width=400, height=50)
    passWordEntrybox.insert(0, "8 digits characters")

    # Submit button to get the data and verify it
    submitButton = Button(frame2, text="SUBMIT", font=("Tw Cen MT Condensed Extra Bold", 15),
                          padx=30, pady=10, width=7, height=1, command=onCLickVerfiyAcc)
    submitButton.place(y=200, x=-100, relx=0.5, rely=0.5, anchor=CENTER)

    RewindTimeButton = Button(frame2, text="<--", command=lambda: switchWindows(2), image=ph7,
                              padx=300, pady=100, width=7, height=1, font=("Modern No.", 10, "bold"), fg="#5cb6f9", bg="#5cb6f9")
    RewindTimeButton.place(height=100, width=100, relx=0.5,
                           rely=0.5, anchor=CENTER, x=500, y=300)

    # Resets the Value in the Entry boxes

    def Reset():
        AccNOEntryBox.delete(0, END)
        AccNOEntryBox.insert(0, "8 digit ID")
        passWordEntrybox.delete(0, END)
        passWordEntrybox.insert(0, "8 digits characters")
        return

    resetButton = Button(frame2, text="RESET", font=("Tw Cen MT Condensed Extra Bold", 15),
                         padx=30, pady=10, width=7, height=1, command=Reset)
    resetButton.place(y=200, x=50, relx=0.5, rely=0.5, anchor=CENTER)
    return


# This function is called to send a secure email to the user about TranactionHistory,OTP,Updates etc.
def SendEmail(msgstr):
    print("Emailing")
    global Email
    # Emailing the user his ID and password
    # message to be sent
    msg = MIMEText(msgstr)
    msg['Subject'] = 'Welcome to BMS'
    msg['From'] = "pythonemailtestservice@gmail.com"
    msg['To'] = Email
    # creates SMTP session
    s = smtplib.SMTP('smtp.gmail.com', 587)
    # checking connections
    s.ehlo()
    # start TLS for security
    s.starttls()
    # Authentication
    s.login("pythonemailtestservice@gmail.com", "PythonOP123")
    # sending the mail
    s.sendmail("pythonemailtestservice@gmail.com", Email, msg.as_string())
    # terminating the session
    s.quit()
    print("Print Email")
    return

# This function is when depositButton or withdrawButton is clicked


def crdt_write(i):
    print("Deposit/Withdraw windows")
    # Creating a new deposit window
    global amt
    frame3.destroy()
    CreateFrame(1)
    if (i == 1):
        TitleLabel = Label(frame1,
                           text="DEPOSIT WINDOW",
                           width=100, height=5, font=("BankGothic Md BT ", 20, "bold"), bg="#cae8ff", fg="#1f15ad")
        TitleLabel.place(rely=0.5, relx=0.5, y=-250, anchor=CENTER)
    elif (i == 2):
        TitleLabel = Label(frame1,
                           text="WITHDRAW WINDOW",
                           width=100, height=5, font=("BankGothic Md BT ", 20, "bold"), bg="#cae8ff", fg="#1f15ad")
        TitleLabel.place(rely=0.5, relx=0.5, y=-250, anchor=CENTER)
    # amt contains the transaction amount
    amt = Entry(frame1, font=("BankGothic Md BT ", 20, "bold"))
    amt.place(relx=0.5, rely=0.5, anchor=CENTER, y=-100, width=400, height=50)

    if (i == 1):
        DepositButton = tk.Button(frame1, font=("BankGothic Md BT ", 10, "bold"),
                                  text="Deposit", command=lambda: cred(i))
        DepositButton.place(relx=0.5, rely=0.5, anchor=CENTER,
                            y=300, width=100, height=50)
    else:
        WithdrawButton = tk.Button(frame1, font=("BankGothic Md BT ", 10, "bold"),
                                   text="Withdraw", command=lambda: cred(i))
        WithdrawButton.place(relx=0.5, rely=0.5, anchor=CENTER,
                             y=300, width=100, height=50)

    RewindTimeButton = Button(frame1, text="BACK", command=lambda: switchWindows(4), image=ph7,
                              font=("Modern No.", 11, "bold"), bg="white", fg="white")
    RewindTimeButton.place(relx=0.5, rely=0.5, anchor=CENTER, x=500, y=300)
    return

# this function writes the new currentBalance of the user to Excel and the Transaction file.It aslo speaks the transaction.


def cred(i):
    print("Updating CurrentBalance")
    global currentBal, ii, amt
    try:
        amti = float(amt.get())
    except:
        messagebox.showerror("ValueError", "Please Enter correct amount")
        amt.delete(0, END)
        return
    amt.delete(0, END)
    print(currentBal)
    currentBal = float(currentBal)
    if (i == 1):
        currentBal += amti
        print("Amount added", amti, currentBal, i)
    elif (i == 2):
        if (amti > currentBal or currentBal < 0):
            messagebox.showerror("MoneyIssues", "you are Broke")
            return
        else:
            currentBal -= amti
            print("Amount subtracted", amti, currentBal, i)
    EXCEL_FILE = path + 'BMS.xlsx'
    df = pd.read_excel(EXCEL_FILE)
    df.loc[ii, 'currentBal'] = currentBal
    df.to_excel(EXCEL_FILE, index=False)
    print(amti)
    frec = open(path + str(AccNO) + "-rec.txt", 'a+')
    ind_time = datetime.now(timezone("Asia/Kolkata")
                            ).strftime('%Y-%m-%d %H:%M:%S')
    if (i == 1):
        frec.write(
            str(ind_time + "      " + str(amti) + "      " + str(currentBal) + "\n"))
    elif (i == 2):
        frec.write(
            str(ind_time + "      " + "-" + str(amti) + "      " + str(currentBal) + "\n"))
    frec.close()
    if(i == 1):
        messagebox.showinfo("Operation Successfull!!",
                            "Amount Deposited Successfully!!")
    elif(i == 2):
        messagebox.showinfo("Operation Successfull!!",
                            "Amount Withdrawed Successfully!!")

    # Speaking the transaction to the user
    engine = pyttsx3.init()
    if(i == 1):
        engine.say(str(amti) + "Amount Deposited Successfully")
    elif(i == 2):
        engine.say(str(amti) + "Amount Withdrawed Successfully")
    engine.runAndWait()
    return

# This function creates the transaction history window and displays the list of transaction history


def transactionHistory():
    print("Displaying transaction history")
    global AccNO
    frame3.destroy()
    CreateFrame(4)
    TitleLabel = Label(frame4,
                       text="TRANSACTION HISTORY OF USER : %s" % (AccNO),
                       width=100, height=5, font=("BankGothic Md BT ", 20, "bold"), bg="#cae8ff", fg="#1f15ad")
    TitleLabel.place(rely=0.5, relx=0.5, y=-350, anchor=CENTER)

    # Text box containing the transaction history of the user
    listOfTransactions = Listbox(frame4, width=150, height=25)

    # Creating a Scrollbar and
    # attaching it to root window
    scrollbar = Scrollbar(frame4)

    # Adding Scrollbar to the right
    # side of root window
    scrollbar.place(anchor=E)

    # Attaching Listbox to Scrollbar
    # Since we need to have a vertical
    # scroll we use yscrollcommand
    listOfTransactions.config(yscrollcommand=scrollbar.set)

    # setting scrollbar command parameter
    # to listbox.yview method its yview because
    # we need to have a vertical view
    scrollbar.config(command=listOfTransactions.yview)

    msgstr = 'Dear user your transaction history is :\n'
    listOfTransactions.place(relx=0.5, rely=0.5, anchor=CENTER)
    frec = open(path + str(AccNO) + "-rec.txt", 'r')
    for line in frec:
        msgstr += (line + "\n")
        listOfTransactions.insert(END, line)
    frec.close()
    # print(msgstr)
    # print("Transaction History")

    TransactMail = Button(frame4, text="MailTransactionHistory", command=lambda: SendEmail(msgstr),
                          font=("Modern No.", 11, "bold"))
    TransactMail.place(rely=0.5, relx=0, anchor=CENTER, x=500, y=300)

    RewindTimeButton = Button(frame4, text="BACK", command=lambda: switchWindows(3), image=ph7,
                              font=("Modern No.", 11, "bold"), bg="white", fg="white")
    RewindTimeButton.place(relx=0.5, rely=0.5, anchor=CENTER, x=500, y=300)
    return

# little popup window for displaying currentBalance


def popupCurrentBal():
    print("Showing Current Balance")
    global firstName, lastName, middleName
    EXCEL_FILE = path+'BMS.xlsx'
    df = pd.read_excel(EXCEL_FILE)
    messagebox.showinfo("Current Account Balance", "%s's Balance : %s" % (
        (firstName+' ' + middleName+' ' + lastName), df.loc[ii, 'currentBal']))
    return

# This function creates the updating info window and displays the old users data and it can be updated.


def EditInfo():
    print("Updating Info")
    global nameEntryBox, DOBEntryBox, emailEntryBox, mobileEntryBox, addharEntryBox, gen, newPassEntrybox
    global firstName, lastName, middleName, DOB, Email, mobileNO, gender, aadharNO, passWord, AccNO, currentBal
    # This will destroy the Exisitng frame1 which was for MAIN screen
    frame3.destroy()
    CreateFrame(4)
    # Creating the Title of the update account details window
    accountOpeningLabel = Label(frame4, text="UPDATE WINDOW FOR USER %s" % (AccNO),
                                width=100, height=5, font=("BankGothic Md BT ", 20, "bold"), bg="#cae8ff", fg="#1f15ad")
    accountOpeningLabel.place(y=-300, relx=0.5, rely=0.5, anchor=CENTER)

    # First Entry box to Enter Name
    nameEntryBox = Entry(frame4, font=("BankGothic Md BT", 15, "bold"))
    nameEntryBox.place(y=-50, x=50, relx=0.5, rely=0.5,
                       anchor=CENTER, width=500, height=30)
    nameEntryBox.insert(0, firstName + ' ' + middleName + ' ' + lastName)

    # Name Input Label
    nameInputLabel = Label(frame4, text="NAME :", padx=10, pady=5, bg="#233dff", fg="#cae8ff",
                           font=("Tw Cen MT Condensed Extra Bold", 15))
    nameInputLabel.place(y=-50, x=-250, relx=0.5, rely=0.5, anchor=CENTER)

    # DOB Entry box to Enter Name
    DOBEntryBox = Entry(frame4, font=("BankGothic Md BT", 15, "bold"))
    DOBEntryBox.place(y=0, x=50, relx=0.5, rely=0.5,
                      anchor=CENTER, width=500, height=30)
    DOBEntryBox.insert(0, DOB)

    # DOB Input Label
    DOBInputLabel = Label(frame4, text="DOB :", padx=16, pady=5, bg="#233dff", fg="#cae8ff",
                          font=("Tw Cen MT Condensed Extra Bold", 15))
    DOBInputLabel.place(y=0, x=-250, relx=0.5, rely=0.5, anchor=CENTER)

    # Email Entry box to Enter Name
    emailEntryBox = Entry(frame4, font=("BankGothic Md BT", 15, "bold"))
    emailEntryBox.place(y=50, x=50, relx=0.5, rely=0.5,
                        anchor=CENTER, width=500, height=30)
    emailEntryBox.insert(0, Email)

    # Email Input Label
    emailEntryLabel = Label(frame4, text="Email :", padx=10, pady=5, bg="#233dff", fg="#cae8ff",
                            font=("Tw Cen MT Condensed Extra Bold", 15))
    emailEntryLabel.place(y=50, x=-250, relx=0.5, rely=0.5, anchor=CENTER)

    # MobileNo Entry box to Enter Name
    mobileEntryBox = Entry(frame4, font=("BankGothic Md BT", 15, "bold"))
    mobileEntryBox.place(y=100, x=50, relx=0.5, rely=0.5,
                         anchor=CENTER, width=500, height=30)
    mobileEntryBox.insert(0, mobileNO)

    # MobileNo Input Label
    mobileEntryLabel = Label(frame4, text="MobileNo :", padx=14, pady=5, bg="#233dff", fg="#cae8ff",
                             font=("Tw Cen MT Condensed Extra Bold", 15))
    mobileEntryLabel.place(y=100, x=-270, relx=0.5, rely=0.5, anchor=CENTER)

    # Gender radio button
    gen = StringVar(frame4)
    gen.set(gender)
    genderList = OptionMenu(frame4, gen, "Male", "Female", "Other")
    genderList.place(y=150, x=-50, relx=0.5, rely=0.5,
                     anchor=CENTER, width=300, height=30)

    # Gender Input Label
    genderLabel = Label(frame4, text="Gender :", padx=5, pady=5, bg="#233dff", fg="#cae8ff",
                        font=("Tw Cen MT Condensed Extra Bold", 15))
    genderLabel.place(y=150, x=-250, relx=0.5, rely=0.5, anchor=CENTER)

    # aadharNO Entry box to Enter Name
    addharEntryBox = Entry(frame4, font=("BankGothic Md BT", 15, "bold"))
    addharEntryBox.place(y=200, x=50, relx=0.5, rely=0.5,
                         anchor=CENTER, width=500, height=30)
    addharEntryBox.insert(0, aadharNO)

    # aadharNO Input Label
    addharEntryLabel = Label(frame4, text="aadharNO :", padx=5, pady=5, bg="#233dff", fg="#cae8ff",
                             font=("Tw Cen MT Condensed Extra Bold", 15))
    addharEntryLabel.place(y=200, x=-261, relx=0.5, rely=0.5, anchor=CENTER)

    # password making field
    newPassWordLabel = Label(frame4, text="password :", padx=5, pady=5, bg="#233dff", fg="#cae8ff",
                             font=("Tw Cen MT Condensed Extra Bold", 15))
    newPassWordLabel.place(y=250, x=-261, relx=0.5, rely=0.5, anchor=CENTER)

    # password entry box
    newPassEntrybox = Entry(frame4, font=("BankGothic Md BT", 15, "bold"))
    newPassEntrybox.place(y=250, x=50, relx=0.5, rely=0.5,
                          anchor=CENTER, width=500, height=30)
    newPassEntrybox.insert(0, passWord)

    # Submit Button after completely adding in the data
    submitButton = Button(frame4, text="SUBMIT", font=("Tw Cen MT Condensed Extra Bold", 15),
                          padx=30, pady=10, width=7, height=1, command=onClickUpdateInfo)
    submitButton.place(y=370, x=-100, relx=0.5, rely=0.5, anchor=CENTER)

    # Account deletion button
    deleteAccount = Button(frame4, text="DELETE", font=("Tw Cen MT Condensed Extra Bold", 15),
                           padx=30, pady=10, width=7, height=1, command=delAccount)
    deleteAccount.place(y=370, x=300, relx=0.5, rely=0.5, anchor=CENTER)

    # Back button
    RewindTimeButton = Button(frame4, text="BACK", command=lambda: switchWindows(3), image=ph7,
                              font=("Modern No.", 11, "bold"), bg="#233dff", fg="white")
    RewindTimeButton.place(height=100, width=100, relx=0.5,
                           rely=0.5, anchor=CENTER, x=500, y=300)

    # Resets the Value in the Entry boxes
    def Reset():
        nameEntryBox.delete(0, END)
        nameEntryBox.insert(0, firstName + ' ' + middleName + ' ' + lastName)
        DOBEntryBox.delete(0, END)
        DOBEntryBox.insert(0, DOB)
        emailEntryBox.delete(0, END)
        emailEntryBox.insert(0, Email)
        mobileEntryBox.delete(0, END)
        mobileEntryBox.insert(0, mobileNO)
        gen.set(gender)
        addharEntryBox.delete(0, END)
        addharEntryBox.insert(0, aadharNO)
        newPassEntrybox.delete(0, END)
        newPassEntrybox.insert(0, passWord)
        return

    resetButton = Button(frame4, text="RESET", font=("Tw Cen MT Condensed Extra Bold", 15),
                         padx=30, pady=10, width=7, height=1, command=Reset)
    resetButton.place(y=370, x=100, relx=0.5, rely=0.5, anchor=CENTER)

    return

# this function delets the users data from the database and also removes the transaction history file from computer.


def delAccount():
    ask = messagebox.askokcancel(
        "Deleting Accout", "Are you sure you want to delete your account?")
    if(ask):
        print("Deleting Accout")
        global ii, AccNO
        i = ii
        # removes the users transaction history file
        os.remove(str(AccNO)+'-rec.txt')
        # removing everything from excel
        EXCEL_FILE = path+'BMS.xlsx'
        df = pd.read_excel(EXCEL_FILE)
        df.loc[i, 'firstName'] = ''
        df.loc[i, 'lastName'] = ''
        df.loc[i, 'middleName'] = ''
        df.loc[i, 'DOB'] = ''
        df.loc[i, 'Email'] = ''
        df.loc[i, 'mobileNO'] = ''
        df.loc[i, 'gender'] = ''
        df.loc[i, 'aadharNO'] = ''
        df.loc[i, 'passWord'] = '-1'
        df.loc[i, 'AccNO'] = '-1'
        df.loc[i, 'currentBal'] = ''
        df.to_excel(EXCEL_FILE, index=False)
        messagebox.showinfo("Deleted successfully",
                            "Account Successfully Removed")
        switchWindows(5)
    return

# When the users submits his updated info this function is called to update it in excel and also mail him that its done


def onClickUpdateInfo():
    print("Updating old users data in excel")
    global ii, firstName, lastName, middleName, DOB, Email, mobileNO, gender, aadharNO, passWord
    # Reading the Excel file
    EXCEL_FILE = path+'BMS.xlsx'
    df = pd.read_excel(EXCEL_FILE)
    i = ii
    fullName = nameEntryBox.get()
    firstName, lastName, surName = map(str, fullName.split(' '))
    df.loc[i, 'firstName'] = firstName
    df.loc[i, 'lastName'] = lastName
    df.loc[i, 'middleName'] = surName
    df.loc[i, 'DOB'] = DOBEntryBox.get()
    df.loc[i, 'Email'] = emailEntryBox.get()
    df.loc[i, 'mobileNO'] = mobileEntryBox.get()
    df.loc[i, 'gender'] = gen.get()
    df.loc[i, 'aadharNO'] = addharEntryBox.get()
    df.loc[i, 'passWord'] = newPassEntrybox.get()

    # writing to the Excel file
    df.to_excel(EXCEL_FILE, index=False)

    # updating the variables
    loadUserData(ii)

    # Sending user an email that his account was updated
    msgstr = """Dear user,\nThis mail is to inform you that you have updated your account details\n
    so now your updated details are:
    Name = %s %s %s
    DOB = %s
    Email = %s
    Mobile NO = %s
    gender = %s
    Addhar NO = %s
    password = %s
    """ % (firstName, lastName, surName, DOB, Email, mobileNO, gender, aadharNO, passWord)
    SendEmail(msgstr)

    messagebox.showinfo(
        "InfoUpdated", "Your account details have been updated please check your email")
    switchWindows(3)
    return

# The main transaction window with all the buttons to various functionalities


def mainTransactFun():
    print("Main Transaction Window")
    CreateFrame(3)
    # main image background
    mainBackGroundImg = Label(frame3, image=ph)
    mainBackGroundImg.place(anchor=CENTER, width=widthWindow,
                            height=heightWindow, rely=0.5, relx=0.5)

    # transact button
    transactButton = Button(frame3, image=ph1, command=transactionHistory)
    transactButton.place(anchor=CENTER, relx=0.5, rely=0.5, x=510, y=250)

    # deposit button
    depositButton = Button(frame3, image=ph2, command=lambda: crdt_write(1))
    depositButton.place(anchor=CENTER, relx=0.5, rely=0.5, x=-520, y=250)

    # withdraw button
    withdrawButton = Button(frame3, image=ph4, command=lambda: crdt_write(2))
    withdrawButton.place(anchor=CENTER, relx=0.5, rely=0.5, x=-520)

    # show account balance button
    accountBalButton = Button(frame3, image=ph3, command=popupCurrentBal)
    accountBalButton.place(anchor=CENTER, relx=0.5, rely=0.5, x=510)

    # update the users details
    editInfoButton = Button(frame3, text="Update Information",
                            command=EditInfo, image=ph8, bg="#2596be", fg="#2596be")
    editInfoButton.place(anchor=CENTER, relx=0.5, rely=0.5,
                         y=300, width=300, height=100)

    RewindTimeButton = Button(frame3, text="BACK", command=lambda: switchWindows(6), image=ph7,
                              font=("Modern No.", 11, "bold"), bg="white", fg="white")
    RewindTimeButton.place(relx=0.5, rely=0.5, anchor=CENTER, x=500, y=-300)

    return


# After Verfiying checks, this function sends the  user an OTP.
# if the OTP checks are successful then the users data is stored into excel and a mail of the data is sent to him
# Also a default transaction history file is creaded for each user with the minimum deposit as the first transaction.
def onCLickSaveInfo():
    print("Saving new users Data")
    global firstName, lastName, middleName, DOB, Email, mobileNO, gender, aadharNO, passWord, currentBal, ii

    ask = messagebox.askokcancel("Recheck Please!", "Are you sure?")
    if(not ask):
        return
    Email = emailEntryBox.get()
    # Prompt for OTP
    randOTP = random.randint(1000, 9999)
    msgstr = "Your OTP is : " + str(randOTP)
    SendEmail(msgstr)
    otp = simpledialog.askstring(
        title="OTP", prompt="Check your Email for OTP")
    print(otp)
    if(int(otp) == randOTP):
        print("Continue user is safe")
    else:
        messagebox.showwarning("OTP failed", "Please try again")
        print("Verififcation failed")
        return

    print("Account created moving to depositing in it")
    gender = gen.get()
    ii = 0

    # Dictionatry to be created to input values in Excel sheet
    values['firstName'] = firstName
    values['lastName'] = lastName
    values['middleName'] = middleName
    values['DOB'] = DOB
    values['Email'] = Email
    values['mobileNO'] = mobileNO
    values['gender'] = gender
    values['aadharNO'] = aadharNO
    values['passWord'] = passWord
    values['AccNO'] = AccNO
    values['currentBal'] = currentBal

    # Reading the Excel file
    EXCEL_FILE = path+'BMS.xlsx'
    df = pd.read_excel(EXCEL_FILE)
    df = df.append(values, ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)
    
    # reading the file and displaying the transaction history
    ind_time = datetime.now(timezone("Asia/Kolkata")
                            ).strftime('%Y-%m-%d %H:%M:%S')
    frec = open(str(AccNO) + "-rec.txt", 'w+')
    frec.write("      Date                Credit//Debit     Balance\n")
    frec.write(
        str(ind_time + "     " + currentBal + "     " + currentBal + "\n"))
    frec.close()

    # creating a string account creation confirmation
    # Sending user an email that his account was updated
    msgstr = """Dear user,\nThank you for creating an account\n
    your saved details are:
    Name = %s %s %s
    DOB = %s
    Email = %s
    Mobile NO = %s
    gender = %s
    Addhar NO = %s
    
    AccountNO = %s
    password = %s
    """ % (firstName, lastName, middleName, DOB, Email, mobileNO, gender, aadharNO, AccNO, passWord)
    SendEmail(msgstr)
    messagebox.showinfo("Verififcation Successful", "Please check your email")

    # loads the new users indeex so it can be accessed to perform some actions
    EXCEL_FILE = path+'BMS.xlsx'
    df = pd.read_excel(EXCEL_FILE)
    for i in range(0, len(df)):
        if (str(AccNO) == str(df.loc[i, 'AccNO'])):
            ii = i
            
    frame2.destroy()

    # Main function of the transactions window
    mainTransactFun()

    return


# this function is Verifying the old account Exists or not
def onCLickVerfiyAcc():
    print("Verifying the old account Exists")
    global iD, passWord

    iD = AccNOEntryBox.get()
    passWord = passWordEntrybox.get()
    EXCEL_FILE = path+'BMS.xlsx'
    df = pd.read_excel(EXCEL_FILE)
    passf = 0
    logf = 0
    index = 0
    #Chekcing if the Admin is logging in else just show an error
    if(iD == "GUEST" and passWord == "GUEST"):
        print("Welcome Admin")
        index = 0
        loadUserData(index)
        frame2.destroy()
        mainTransactFun()
        return
    try:
        for i in range(1, len(df)):
            # print(i, df.loc[i, 'AccNO'], df.loc[i, 'passWord'])
            if (iD == str(int(df.loc[i, 'AccNO'])) and iD != '-1'):
                # print(iD)
                logf = 1
                if (passWord == str(df.loc[i, 'passWord'])):
                    # print(passWord)
                    passf = 1
                    index = i
                    break
    except:
        messagebox.showerror(
            "Error occured ", "Account number length is incorrect")

    if (len(iD) != 13):
        messagebox.showerror(
            "Error occured ", "Account number length is incorrect")

    elif (logf == 0):
        messagebox.showerror("login ID Not found",
                             "Please Enter Correct AccNO")

    elif (passf == 0):
        messagebox.showerror("Password Not found",
                             "Please Enter Correct password")

    else:
        # Main function of the transactions window
        loadUserData(index)
        frame2.destroy()
        mainTransactFun()

    return


# First Window To Be Displayed
# the Basic window with options for the user to create a new account or login
def loginWindow():
    print("Welcome to BMS")
    CreateFrame(1)
    # main window Heading Label
    mainMenuLabel = Label(frame1, text="WELCOME TO BANK MANAGEMENT SYSTEM", padx=15, pady=15, bg="#587dfa",
                          fg="#8dd8e8",
                          font=("Tw Cen MT Condensed Extra Bold", 50))
    mainMenuLabel.place(y=-370, relx=0.5, rely=0.5, anchor=CENTER)

    # Old user Button
    oldUserButton = Button(frame1, text="Click here to login to old account", command=OnClickRedirectOldAcc,
                           font=("Modern No.", 11, "bold"), bg="red", fg="white", image=ph6)
    oldUserButton.place(y=50, x=300, relx=0.5, rely=0.5, anchor=CENTER)

    # new user Button
    newUserButton = Button(frame1, text="Click here to create a new account", command=OnClickCreateAccount,
                           font=("Modern No.", 11, "bold"), bg="red", fg="white", image=ph5)
    newUserButton.place(y=50, x=-300, relx=0.5, rely=0.5, anchor=CENTER)

    return


# Calling the Main Login window function to create the login window
loginWindow()

# loops around the root window until the user closes the application
root.mainloop()
