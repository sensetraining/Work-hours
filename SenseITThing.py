#-------------------------Import-------------------------#
# pip install gspread oauth2client
from asyncio.windows_events import NULL
from sys import flags
import tkinter as tk
import tkinter.messagebox as tkm
import datetime as tm
from datetime import datetime, timedelta
import os
import time

try:
    import gspread
except ImportError:
    os.system('pip install gspread oauth2client')
    import gspread
from oauth2client.service_account import ServiceAccountCredentials
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)
sheet = client.open("workhours").sheet1
data = sheet.get_all_records()
#-----------------------Functions------------------------#
def delete(canvas,frame):
    canvas.destroy()
    frame.destroy()

def login():
    window.geometry("500x280")

    frame = tk.Frame(window)
    frame.pack()

    canvas = tk.Canvas(frame, bg="#F57A22", width=500, height=280)
    canvas.pack()

    #window.geometry("500x285")
    #window.configure(bg="#F57A22")

    plainBox1 = tk.Label(bg="#113F8C",relief="solid")
    plainBox1.place(height=180,width=400,x=50,y=50)

    # helvetica, 18, bold height = 35

    nameLabel = tk.Label(text="Name:",fg="white",bg="#113F8C",font=("Helvetica", 18, "bold"))
    nameLabel.place(x=70,y=70,height=35)

    passLabel = tk.Label(text="Password:",fg="white",bg="#113F8C",font=("Helvetica", 18, "bold"))
    passLabel.place(x=70,y=110,height=35)

    nameInput = tk.Entry(font=("Helvetica", 18, "bold"))
    nameInput.place(x=200,y=70,width=230,height=35)

    passInput = tk.Entry(font=("Helvetica",18, "bold"),show="*")
    passInput.place(x=200,y=110,width=230,height=35)

    submitButton = tk.Button(text="Submit",command=lambda:loginCheck(nameInput.get(),passInput.get(),canvas,frame),bg="white",fg="black",font=("Helvetica", 12, "bold"))
    submitButton.place(x=175,y=165,height=40,width=150)

def loginCheck(userName,userPass,canvas,frame):
    position = NULL
    correctLogin = False
    names = sheet.row_values(1)
    passwords = sheet.row_values(2)

    print(names)
    print(len(names))

    for i in range (len(names)-1):
        if userName.capitalize() == (names[i+1]):
            position = i+1
            correctLogin = True

    print(position)

    if passwords[position] != userPass:
        correctLogin = False

    if correctLogin == False:
        inputError = tkm.showerror(title="Error", message="You entered an invalid name or password")
        delete(canvas,frame)
        login()
        window.mainloop()


    print("Correct login")
    delete(canvas,frame)
    main(userName.capitalize(),position+1)

def clockIn(name,position,canvas,frame):
    x = tm.datetime.now()
    currentTime = str(x.strftime("%X"))

    sheet.update_cell(3,position,"TRUE")

    sheet.update_cell(4,position,currentTime)

    canvas.destroy()
    frame.destroy()

    main(name,position)

def clockOut(name,position,canvas,frame):
    x = tm.datetime.now()
    currentTime = str(x.strftime("%X"))

    sheet.update_cell(3,position,"FALSE")

    sheet.update_cell(5,position,currentTime)

    sheet.update_cell(7,position,sheet.cell(6,position).value)

    sheet.update_cell(9,position,sheet.cell(8,position).value)

    

    canvas.destroy()
    frame.destroy()

    main(name,position)




def main(name,position):
    letters=["a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"]

    window.geometry("500x250")

    frame = tk.Frame(window)
    frame.pack()

    canvas = tk.Canvas(frame, bg="#F57A22", width=500, height=250)
    canvas.pack()

    plainBox1 = tk.Label(bg="#113F8C",relief="solid")
    plainBox1.place(height=150,width=400,x=50,y=50)

    welcomeLabel = tk.Label(text=f"Hello, {name}!",fg="white",bg="#113F8C",font=("Helvetica", 16, "bold"))
    welcomeLabel.place(x=70,y=65)

    col = sheet.col_values(position)

    textLabel = tk.Label(text="You are currently ",fg="white",bg="#113F8C",font=("Helvetica", 16, "bold"))
    textLabel.place(x=70,y=95)

    print(col[2])

    clockedStatus = (col[2])

    if clockedStatus == "FALSE":
        clockedLabel = tk.Label(text="not clocked in.",fg="red2",bg="#113F8C",font=("Helvetica", 16, "bold"))
        clockedLabel.place(x=250,y=95)

        clockInButton = tk.Button(text="Clock in",command=lambda:clockIn(name,position,canvas,frame),fg="lime green",bg="white",font=("Helvetica", 16, "bold"))
        clockInButton.place(x=175,y=135,height=40,width=150)

    else:
        clockedLabel = tk.Label(text="clocked in.",fg="lime green",bg="#113F8C",font=("Helvetica", 16, "bold"))
        clockedLabel.place(x=250,y=95)

        clockOutButton = tk.Button(text="Clock out",command=lambda:clockOut(name,position,canvas,frame),fg="red2",bg="white",font=("Helvetica", 16, "bold"))
        clockOutButton.place(x=175,y=135,height=40,width=150)

    column = letters[position-1]
    sheet.update_cell(6,position,f"=SUM({column}5-{column}4)")
    sheet.update_cell(8,position,f"=SUM({column}9+{column}7)")


    #hoursToday = col[7]
    #hoursLabel = tk.Label(text=f"You have done {hoursToday} hours today",fg="white",bg="#113F8C",font=("Helvetica", 16, "bold"))
    #hoursLabel.place(x=70,y=125)

    return

def weekCheck():
    
    # Getting last weeks dates
    #i=7
    #while i != 1:
    #    yesterdayDate = datetime.strftime(datetime.now() - timedelta(i), '%a - %d/%m/%y')
         
   
    # Checking last weeks dates
    
    yesterday = datetime.strftime(datetime.now() - timedelta(1), '%a - %d/%m/%y')
    if sheet.cell(9,1).value != yesterday:
        i=35
        while i != 0:
            i-=1
            yesterdayDate = datetime.strftime(datetime.now() - timedelta(i), '%a - %d/%m/%y')
            aYesterdayDate = [yesterdayDate]
            print(yesterdayDate)
            print(sheet.cell(9,1).value)
            print(i)
            if sheet.cell(9,1).value == yesterdayDate:
                print(i)
                i-=1
                while i != 0:
                    yesterdayDate = datetime.strftime(datetime.now() - timedelta(i), '%a - %d/%m/%y')
                    aYesterdayDate = [yesterdayDate]
                    yesterdayDay = datetime.strftime(datetime.now() - timedelta(i), '%a')
                    yesterdayMonth = datetime.strftime(datetime.now() - timedelta(i), '%d')
                    
                    sheet.insert_row(aYesterdayDate,9)
                    print(yesterdayDate)

                    print(yesterdayDay)
                    if yesterdayDay == "Sun":
                        print("It's Sunday")
                        weekTotal = ["Week Total","=SUM(B10:B16)","=SUM(C10:C16)","=SUM(D10:D16)","=SUM(E10:E16)","=SUM(F10:F16)","=SUM(G10:G16)","=SUM(H10:H16)","=SUM(I10:I16)","=SUM(J10:J16)","=SUM(K10:K16)","=SUM(L10:L16)","=SUM(M10:M16)","=SUM(N10:N16)","=SUM(O10:O16)","=SUM(P10:P16)","=SUM(Q10:Q16)","=SUM(R10:R16)","=SUM(S10:S16)","=SUM(T10:T16)","=SUM(U10:U16)","=SUM(V10:V16)","=SUM(W10:W16)","=SUM(X10:X16)","=SUM(Y10:Y16)"]
                        sheet.insert_row(weekTotal,9,value_input_option="USER_ENTERED")

                    
                    print(yesterdayMonth)
                    if yesterdayMonth == "01":
                        print("New month")
                        monthTotal = ["Month Total","Test"]
                        sheet.insert_row(monthTotal,9,value_input_option="USER_ENTERED")
                    time.sleep(5)

                    i-=1

                print("break")
                break

        # Checking for yesterday presense
        # if sheet.cell(9,1).value != yesterday:
        #     sheet.insert_row(yesterday,9)

    
    return

#--------------------------Main--------------------------#
x = tm.datetime.now()
currentDate = str(x.strftime("%a - %d/%m/%y"))

#sheet.update_cell(4,1,currentDate)
#row = ["Hello",2,"Bye"]
date = [currentDate,]

if sheet.cell(9,1).value != currentDate:
    weekCheck()
    sheet.insert_row(date,9)

#row = sheet.row_values(2)
#col = sheet.col_values(1)
#cell = sheet.cell(1,2).value


#sheet.update_cell(14,2,"=SUM(B7,B8,B9,B10,B11,B12)")

window = tk.Tk()
window.title("Hour tracker")

login()

window.mainloop()

