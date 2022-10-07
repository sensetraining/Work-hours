#-------------------------Import-------------------------#
# pip install gspread oauth2client
from asyncio.windows_events import NULL
from sys import flags
import tkinter as tk
import tkinter.messagebox as tkm
import datetime as tm
from datetime import datetime, timedelta
from calendar import monthrange
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
            print(i)
            if sheet.cell(9,1).value == yesterdayDate:
                i-=1
                while i != 0:
                    yesterdayDate = datetime.strftime(datetime.now() - timedelta(i), '%a - %d/%m/%y')
                    aYesterdayDate = [yesterdayDate]
                    yesterdayDay = datetime.strftime(datetime.now() - timedelta(i+1), '%a')
                    yesterdayMonth = datetime.strftime(datetime.now() - timedelta(i), '%d')
                    
                    print(yesterdayDate)
                    
                    print(yesterdayDay)
                    if yesterdayDay == "Sun":
                        print("It's Sunday")
                        addSUM1 = [10,11,12,13,14,15,16]
                        for j in range (7):
                            print(j)
                            if sheet.cell(9+j,1).value == "Month Total":
                                time.sleep(1)
                                print("Uh oh")
                                
                                for k in range (j,7):
                                    addSUM1[k] = addSUM1[k] + 1
                                    print(k)
                                    print(addSUM1)
                                addSUM2 = [f"B{addSUM1[0]}+",f"B{addSUM1[1]}+",f"B{addSUM1[2]}+",f"B{addSUM1[3]}+",f"B{addSUM1[4]}+",f"B{addSUM1[5]}+",f"B{addSUM1[6]}"]
                                addSUM = "".join(addSUM2)
                                print(addSUM)


                        print(f"=SUM(B{addSUM1[0]}+B{addSUM1[1]}+B{addSUM1[2]}+B{addSUM1[3]}+B{addSUM1[4]}+B{addSUM1[5]}+B{addSUM1[6]})")
                        weekTotal = ["Week Total",
                        f"=SUM(B{addSUM1[0]}+B{addSUM1[1]}+B{addSUM1[2]}+B{addSUM1[3]}+B{addSUM1[4]}+B{addSUM1[5]}+B{addSUM1[6]})",
                        f"=SUM(C{addSUM1[0]}+C{addSUM1[1]}+C{addSUM1[2]}+C{addSUM1[3]}+C{addSUM1[4]}+C{addSUM1[5]}+C{addSUM1[6]})",
                        f"=SUM(D{addSUM1[0]}+D{addSUM1[1]}+D{addSUM1[2]}+D{addSUM1[3]}+D{addSUM1[4]}+D{addSUM1[5]}+D{addSUM1[6]})",
                        f"=SUM(E{addSUM1[0]}+E{addSUM1[1]}+E{addSUM1[2]}+E{addSUM1[3]}+E{addSUM1[4]}+E{addSUM1[5]}+E{addSUM1[6]})",
                        f"=SUM(F{addSUM1[0]}+F{addSUM1[1]}+F{addSUM1[2]}+F{addSUM1[3]}+F{addSUM1[4]}+F{addSUM1[5]}+F{addSUM1[6]})",
                        f"=SUM(G{addSUM1[0]}+G{addSUM1[1]}+G{addSUM1[2]}+G{addSUM1[3]}+G{addSUM1[4]}+G{addSUM1[5]}+G{addSUM1[6]})",
                        f"=SUM(H{addSUM1[0]}+H{addSUM1[1]}+H{addSUM1[2]}+H{addSUM1[3]}+H{addSUM1[4]}+H{addSUM1[5]}+H{addSUM1[6]})",
                        f"=SUM(I{addSUM1[0]}+I{addSUM1[1]}+I{addSUM1[2]}+I{addSUM1[3]}+I{addSUM1[4]}+I{addSUM1[5]}+I{addSUM1[6]})",
                        f"=SUM(J{addSUM1[0]}+J{addSUM1[1]}+J{addSUM1[2]}+J{addSUM1[3]}+J{addSUM1[4]}+J{addSUM1[5]}+J{addSUM1[6]})",
                        f"=SUM(K{addSUM1[0]}+K{addSUM1[1]}+K{addSUM1[2]}+K{addSUM1[3]}+K{addSUM1[4]}+K{addSUM1[5]}+K{addSUM1[6]})",
                        f"=SUM(L{addSUM1[0]}+L{addSUM1[1]}+L{addSUM1[2]}+L{addSUM1[3]}+L{addSUM1[4]}+L{addSUM1[5]}+L{addSUM1[6]})",
                        f"=SUM(M{addSUM1[0]}+M{addSUM1[1]}+M{addSUM1[2]}+M{addSUM1[3]}+M{addSUM1[4]}+M{addSUM1[5]}+M{addSUM1[6]})",
                        f"=SUM(N{addSUM1[0]}+N{addSUM1[1]}+N{addSUM1[2]}+N{addSUM1[3]}+N{addSUM1[4]}+N{addSUM1[5]}+N{addSUM1[6]})",
                        f"=SUM(O{addSUM1[0]}+O{addSUM1[1]}+O{addSUM1[2]}+O{addSUM1[3]}+O{addSUM1[4]}+B{addSUM1[5]}+O{addSUM1[6]})",
                        f"=SUM(P{addSUM1[0]}+P{addSUM1[1]}+P{addSUM1[2]}+P{addSUM1[3]}+P{addSUM1[4]}+P{addSUM1[5]}+P{addSUM1[6]})",
                        f"=SUM(Q{addSUM1[0]}+Q{addSUM1[1]}+Q{addSUM1[2]}+Q{addSUM1[3]}+Q{addSUM1[4]}+Q{addSUM1[5]}+Q{addSUM1[6]})",
                        f"=SUM(R{addSUM1[0]}+R{addSUM1[1]}+R{addSUM1[2]}+R{addSUM1[3]}+R{addSUM1[4]}+R{addSUM1[5]}+R{addSUM1[6]})",
                        f"=SUM(S{addSUM1[0]}+S{addSUM1[1]}+S{addSUM1[2]}+S{addSUM1[3]}+S{addSUM1[4]}+S{addSUM1[5]}+S{addSUM1[6]})",
                        f"=SUM(T{addSUM1[0]}+T{addSUM1[1]}+T{addSUM1[2]}+T{addSUM1[3]}+T{addSUM1[4]}+T{addSUM1[5]}+T{addSUM1[6]})",
                        f"=SUM(B{addSUM1[0]}+U{addSUM1[1]}+U{addSUM1[2]}+U{addSUM1[3]}+U{addSUM1[4]}+U{addSUM1[5]}+U{addSUM1[6]})",
                        f"=SUM(B{addSUM1[0]}+V{addSUM1[1]}+V{addSUM1[2]}+V{addSUM1[3]}+V{addSUM1[4]}+V{addSUM1[5]}+V{addSUM1[6]})",
                        f"=SUM(B{addSUM1[0]}+W{addSUM1[1]}+W{addSUM1[2]}+W{addSUM1[3]}+W{addSUM1[4]}+W{addSUM1[5]}+W{addSUM1[6]})",
                        f"=SUM(B{addSUM1[0]}+X{addSUM1[1]}+X{addSUM1[2]}+X{addSUM1[3]}+X{addSUM1[4]}+X{addSUM1[5]}+X{addSUM1[6]})",
                        f"=SUM(B{addSUM1[0]}+Y{addSUM1[1]}+Y{addSUM1[2]}+Y{addSUM1[3]}+Y{addSUM1[4]}+Y{addSUM1[5]}+Y{addSUM1[6]})",
                        f"=SUM(B{addSUM1[0]}+Z{addSUM1[1]}+Z{addSUM1[2]}+Z{addSUM1[3]}+Z{addSUM1[4]}+Z{addSUM1[5]}+Z{addSUM1[6]})"]
                        sheet.insert_row(weekTotal,9,value_input_option="USER_ENTERED")
                    
                    print(yesterdayMonth)
                    if yesterdayMonth == "01":
                        print("End of month")
                        lastMonth = int(datetime.strftime(datetime.now() - timedelta(i+1), '%m'))
                        print(lastMonth)
                        addMonth = []
                        print(monthrange(2022, lastMonth)[1])
                        monthLength = (monthrange(2022, lastMonth)[1])
                        for l in range (monthLength):
                            addMonth.append(10+monthLength)
                        for m in range (monthLength):
                            if sheet.cell(9+m,1).value == "Week Total":
                                time.sleep(1)
                                print("Oh no")
                                for n in range(m,monthLength):
                                    addMonth[n] = addMonth[n] + 1
                                    print(n)
                                    print(addMonth)
                        
                        monthTotal = ["Month Total"]
                        monthRow = []
                        Letters = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
                        for p in range (25):
                            for o in range(monthLength):
                                monthRow.append(f"=SUM({Letters[p+1]}{10+o}+")

                            yes = "".join(monthRow)
                            monthTotal.append(yes)
                            print(monthTotal)

                        print(monthRow)


                        sheet.insert_row(monthTotal,9,value_input_option="USER_ENTERED")



                    time.sleep(5)

                    sheet.insert_row(aYesterdayDate,9)
                    i-=1
                sheet.insert_row(aYesterdayDate,9)
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

print(monthrange(2022, 10)[1])

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

