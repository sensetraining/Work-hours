#-------------------------Import-------------------------#
# pip install gspread oauth2client
import datetime as tm
import os
import time
import logging
import tkinter as tk
import tkinter.messagebox as tkm
from inspect import currentframe
from calendar import month, monthrange, week
from datetime import datetime, timedelta

x = tm.datetime.now()
time = str(x.strftime("%Y-%m-%d %H.%M.%S"))


try:
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    fh = logging.FileHandler(f'logs\{time}.txt')
    fh.setLevel(logging.INFO)
    logger.addHandler(fh)
except:
    os.mkdir('logs')
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    fh = logging.FileHandler(f'logs\{time}.txt')
    fh.setLevel(logging.INFO)
    logger.addHandler(fh)
    
def lineNum():
    x = tm.datetime.now()
    time = str(x.strftime("[%X-%x]"))
    cf = currentframe()
    return f"{time} Line {cf.f_back.f_lineno}:"

try:
    import gspread
except ImportError:
    print(f"{lineNum()} Installing gspread oauth2client")
    logger.info("Installing gspread oauth2client")
    os.system('pip install gspread oauth2client')
    import gspread

try:
    from gspread_formatting import *
except ImportError:
    print(f"{lineNum()} Installing gspread_formatting")
    logger.info("Installing gspread_formatting")
    os.system('pip install gspread_formatting')
    from gspread_formatting import *

from gspread.cell import Cell
from oauth2client.service_account import ServiceAccountCredentials

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)

while True:
    try:
        print(f"{lineNum()} Opening Staff Hours sheet")
        logger.info(f"{lineNum()} Opening Staff Hours sheet")
        sheet = client.open("Staff Hours").sheet1
        break
    except:
        print(f"{lineNum()} Waiting 0")
        logger.info(f"{lineNum()} Waiting 0")
        time.sleep(5)
#-----------------------Functions------------------------#
def delete(box):
    print(f"{lineNum()} Destroying box, going to login page")
    logger.info(f"{lineNum()} Destroying box, going to login page")
    box.destroy()
    login()

def login():
    print(f"{lineNum()} Running login page")
    logger.info(f"{lineNum()} Running login page")
    plainBox1 = tk.Label(bg="#113F8C",relief="solid")
    #plainBox1.place(height=180,width=400,x=50,y=50)
    plainBox1.place(height=180,width=400,relx=0.5,rely=0.5,anchor="center")

    # helvetica, 18, bold height = 35

    nameLabel = tk.Label(plainBox1,text="Name:",fg="white",bg="#113F8C",font=("Helvetica", 18, "bold"))
    nameLabel.place(x=20,y=20,height=35)

    passLabel = tk.Label(plainBox1,text="Password:",fg="white",bg="#113F8C",font=("Helvetica", 18, "bold"))
    passLabel.place(x=20,y=60,height=35)

    nameInput = tk.Entry(plainBox1,font=("Helvetica", 18, "bold"))
    nameInput.place(x=150,y=20,width=230,height=35)

    passInput = tk.Entry(plainBox1,font=("Helvetica",18, "bold"),show="*")
    passInput.place(x=150,y=60,width=230,height=35)

    submitButton = tk.Button(plainBox1,text="Submit",command=lambda:loginCheck(nameInput.get(),passInput.get(),plainBox1),bg="white",fg="black",font=("Helvetica", 12, "bold"))
    submitButton.place(x=125,y=115,height=40,width=150)

def loginCheck(userName,userPass,box):
    print(f"{lineNum()} loginCheck")
    logger.info(f"{lineNum()} loginCheck")
    position = 0
    correctLogin = False

    while True:
        try:
            print(f"{lineNum()} Collecting names and passwords from the sheets")
            logger.info(f"{lineNum()} Collecting names and passwords from the sheets")
            names = sheet.row_values(1)
            passwords = sheet.row_values(2)
            break
        except:
            print(f"{lineNum()} Waiting 1")
            logger.info(f"{lineNum()} Waiting 1")
            time.sleep(5)

    print(f"{lineNum()} Checking for a name matching {userName}")
    logger.info(f"{lineNum()} Checking for a name matching {userName}")
    for i in range (len(names)-1):
        if userName.capitalize() == (names[i+1]):
            position = i+1
            correctLogin = True

    print(f"{lineNum()} Checking if intputted password, {userPass} matches saved password, {passwords[position]} ")
    logger.info(f"{lineNum()} Checking if intputted password, {userPass} matches saved password, {passwords[position]} ")
    if passwords[position] != userPass:
        correctLogin = False

    if correctLogin == False:
        print(f"{lineNum()} Incorrect name or password was entered")
        logger.info(f"{lineNum()} Incorrect name or password was entered")
        inputError = tkm.showerror(title="Error", message="You entered an invalid name or password")
        delete(box)
        login()
        window.mainloop()

    position += 1

    letters=["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]

    x = tm.datetime.now()
    currentDate = str(x.strftime("%a - %d/%m/%y"))

    print(f"{lineNum()} Checking clocked status")
    logger.info(f"{lineNum()} Checking clocked status")
    while True:
            try:
                col = sheet.col_values(position)

                clockedStatus = (col[2])
                if col[5] != currentDate and clockedStatus == "TRUE":
                    print(f"{lineNum()} Clocked status TRUE, last clocked date {col[5]} does not match current date {currentDate}")
                    logger.info(f"{lineNum()} Clocked status TRUE, last clocked date {col[5]} does not match current date {currentDate}")
                    datesCol = sheet.col_values(1)
                    i = 9
                    print(f"{lineNum()} Searching for the clocked date")
                    logger.info(f"{lineNum()} Searching for the clocked date")
                    while datesCol[i] != col[5] and i < len(datesCol)-1:
                        i += 1

                    if i < len(datesCol)-1:
                        print(f"{lineNum()} Clocked date found on row {i+1}")
                        logger.info(f"{lineNum()} Clocked date found on row {i+1}")
                        sheet.update_cell(i+1,position,col[3])
                        print(lineNum(), "Adding clocked in time and highlighting red")
                        logger.info(lineNum(), "Adding clocked in time and highlighting red")
                        column = letters[position-1]
                        fmt = cellFormat(
                            backgroundColor=color(1, 0.5, 0.5),
                            )

                        format_cell_range(sheet, f'{column}{i+1}:{column}{i+1}', fmt)
                    else:
                        print(f"{lineNum()} Clocked date not found")
                        logger.info(f"{lineNum()} Clocked date not found")
                    print(f"{lineNum()} Setting clocked status to FALSE")
                    logger.info(f"{lineNum()} Setting clocked status to FALSE")
                    sheet.update_cell(3,position,"FALSE")
                break
            except Exception as error:
                print(error)
                logger.info(f"error")
                print(f"{lineNum()} Waiting 4")
                logger.info(f"{lineNum()} Waiting 4")
                time.sleep(5)

    main(userName.capitalize(),position,box)

def clockIn(name,position,box):
    print(f"{lineNum()} Clocking In")
    logger.info(f"{lineNum()} Clocking In")
    x = tm.datetime.now()
    currentTime = str(x.strftime("%X"))
    currentDate = str(x.strftime("%a - %d/%m/%y"))
    while True:
        try:
            print(f"{lineNum()} Setting clocked status to TRUE, updating clock in time and date on sheets")
            logger.info(f"{lineNum()} Setting clocked status to TRUE, updating clock in time and date on sheets")
            sheet.update_cell(3,position,"TRUE")
            sheet.update_cell(4,position,currentTime)
            sheet.update_cell(6,position,currentDate)

            break
        except:
            print(lineNum,"Waiting 2")
            logger.info(lineNum,"Waiting 2")
            time.sleep(5)

    main(name,position,box)

def clockOut(name,position,box):
    print(f"{lineNum()} Clocking out")
    logger.info(f"{lineNum()} Clocking out")
    x = tm.datetime.now()
    currentTime = str(x.strftime("%X"))

    while True:
        try:
            print(f"{lineNum()} Calculating clocked time")
            logger.info(f"{lineNum()} Calculating clocked time")
            clockInTime = sheet.cell(4,position).value
    
            clockInTime = clockInTime.split(":")

            clockInHour = int(clockInTime[0])
            clockInMin = int(clockInTime[1])
            clockInSec = int(clockInTime[2])

            currentHour = int(currentTime[0:2])
            currentMin = int(currentTime[3:5])
            currentSec = int(currentTime[6:8])

            clockedHour = currentHour - clockInHour
            clockedMin = currentMin - clockInMin
            clockedSec = currentSec - clockInSec

            clockedTime = (f"{clockedHour}:{clockedMin}:{clockedSec}")
            print(f"{lineNum()} Clocked time, {clockedTime}")
            logger.info(f"{lineNum()} Clocked time, {clockedTime}")

            cellNine = sheet.cell(9,position).value
            print(f"{lineNum()} Grabbing value of cell 9, {cellNine}")
            logger.info(f"{lineNum()} Grabbing value of cell 9, {cellNine}")
            cellNine = cellNine.split(":")
            if cellNine is not None:
                
                cellNineHour = int(cellNine[0])
                cellNineMin = int(cellNine[1])
                cellNineSec = int(cellNine[2])

                cellNineHourAdd = cellNineHour + clockedHour
                cellNineMinAdd = cellNineMin + clockedMin
                cellNineSecAdd = cellNineSec + clockedSec

                cellTotal = (f"{cellNineHourAdd}:{cellNineMinAdd}:{cellNineSecAdd}")
                print(f"{lineNum()} Adding previous clocked time to new clocked time, {':'.join(cellNine)} + {clockedTime} = {cellTotal}")
                logger.info(f"{lineNum()} Adding previous clocked time to new clocked time, {':'.join(cellNine)} + {clockedTime} = {cellTotal}")

                print(f"{lineNum()} Replacing cell 9 with new total")
                logger.info(f"{lineNum()} Replacing cell 9 with new total")
                sheet.update_cell(9,position,cellTotal)

            else:
                print(f"{lineNum()} Cell 9 is empty, inputting clocked time")
                logger.info(f"{lineNum()} Cell 9 is empty, inputting clocked time")
                sheet.update_cell(9,position,clockedTime)
            
            print(f"{lineNum()} Updating last clocked out time to current time, {currentTime}")
            logger.info(f"{lineNum()} Updating last clocked out time to current time, {currentTime}")
            sheet.update_cell(5,position,currentTime)

            break
        except Exception as error:
            print(error)
            logger.info(f"error")
            print(f"{lineNum()} Waiting 3")
            logger.info(f"{lineNum()} Waiting 3")
            time.sleep(5)

    print(f"{lineNum()} Updating clocked status to false")
    logger.info(f"{lineNum()} Updating clocked status to false")
    sheet.update_cell(3,position,"FALSE")
    main(name,position,box)


def main(name,position,box):
    box.destroy()
    print(f"{lineNum()} Running main")
    logger.info(f"{lineNum()} Running main")

    plainBox1 = tk.Label(bg="#113F8C",relief="solid")
    plainBox1.place(height=205,width=400,relx=0.5,rely=0.5,anchor="center")

    welcomeLabel = tk.Label(plainBox1,text=f"Hello, {name}!",fg="white",bg="#113F8C",font=("Helvetica", 16, "bold"))
    welcomeLabel.place(x=20,y=15)

    textLabel = tk.Label(plainBox1,text="You are currently ",fg="white",bg="#113F8C",font=("Helvetica", 16, "bold"))
    textLabel.place(x=20,y=45)
    print(f"{lineNum()} Checking clocked status")
    logger.info(f"{lineNum()} Checking clocked status")
    col = sheet.col_values(position)
    clockedStatus = (col[2])

    if clockedStatus == "FALSE":
        print(f"{lineNum()} Clocked status FALSE, displaying appropriate text")
        logger.info(f"{lineNum()} Clocked status FALSE, displaying appropriate text")
        clockedLabel = tk.Label(plainBox1,text="not clocked in.",fg="red2",bg="#113F8C",font=("Helvetica", 16, "bold"))
        clockedLabel.place(x=200,y=45)

        clockInButton = tk.Button(plainBox1,text="Clock in",command=lambda:clockIn(name,position,plainBox1),fg="lime green",bg="white",font=("Helvetica", 16, "bold"))
        clockInButton.place(x=125,y=90,height=40,width=150)

    else:
        print(f"{lineNum()} Clocked status TRUE, displaying appropriate text")
        logger.info(f"{lineNum()} Clocked status TRUE, displaying appropriate text")
        clockedLabel = tk.Label(plainBox1,text="clocked in.",fg="lime green",bg="#113F8C",font=("Helvetica", 16, "bold"))
        clockedLabel.place(x=200,y=45)

        clockOutButton = tk.Button(plainBox1,text="Clock out",command=lambda:clockOut(name,position,plainBox1),fg="red2",bg="white",font=("Helvetica", 16, "bold"))
        clockOutButton.place(x=125,y=90,height=40,width=150)

    logoutButton = tk.Button(plainBox1,text="Logout",command=lambda:delete(plainBox1),bg="white",fg="black",font=("Helvetica", 16, "bold"))
    logoutButton.place(x=125,y=145,height=40,width=150)

    return

def endWeek():
    print(f"{lineNum()} Running end of week function")
    logger.info(f"{lineNum()} Running end of week function")
    addSUM1 = [10,11,12,13,14,15,16]
    for j in range (7):
        while True:
            try:
                check = sheet.cell(9+j,1).value
                print(f"{lineNum()} Checking if {check} is Month Total")
                logger.info(f"{lineNum()} Checking if {check} is Month Total")
                if sheet.cell(9+j,1).value == "Month Total":
                    print(f"{lineNum()} Month Total found, skipping over it")
                    logger.info(f"{lineNum()} Month Total found, skipping over it")
                    for k in range (j,7):
                        addSUM1[k] = addSUM1[k] + 1
                break
            except:
                print(f"{lineNum()} Waiting 5")
                logger.info(f"{lineNum()} Waiting 5")
                time.sleep(5)

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
    f"=SUM(U{addSUM1[0]}+U{addSUM1[1]}+U{addSUM1[2]}+U{addSUM1[3]}+U{addSUM1[4]}+U{addSUM1[5]}+U{addSUM1[6]})",
    f"=SUM(V{addSUM1[0]}+V{addSUM1[1]}+V{addSUM1[2]}+V{addSUM1[3]}+V{addSUM1[4]}+V{addSUM1[5]}+V{addSUM1[6]})",
    f"=SUM(W{addSUM1[0]}+W{addSUM1[1]}+W{addSUM1[2]}+W{addSUM1[3]}+W{addSUM1[4]}+W{addSUM1[5]}+W{addSUM1[6]})",
    f"=SUM(X{addSUM1[0]}+X{addSUM1[1]}+X{addSUM1[2]}+X{addSUM1[3]}+X{addSUM1[4]}+X{addSUM1[5]}+X{addSUM1[6]})",
    f"=SUM(Y{addSUM1[0]}+Y{addSUM1[1]}+Y{addSUM1[2]}+Y{addSUM1[3]}+Y{addSUM1[4]}+Y{addSUM1[5]}+Y{addSUM1[6]})",
    f"=SUM(Z{addSUM1[0]}+Z{addSUM1[1]}+Z{addSUM1[2]}+Z{addSUM1[3]}+Z{addSUM1[4]}+Z{addSUM1[5]}+Z{addSUM1[6]})"]
    while True:
        try:
            print(f"{lineNum()} Adding week total to sheets")
            logger.info(f"{lineNum()} Adding week total to sheets")
            sheet.insert_row(weekTotal,9,value_input_option="USER_ENTERED")
            break
        except:
            sheet.delete_row(9)
            print(f"{lineNum()} Waiting 6")
            logger.info(f"{lineNum()} Waiting 6")
            time.sleep(5)

def endMonth():
    print(f"{lineNum()} Running end of month function")
    logger.info(f"{lineNum()} Running end of month function")
    lastMonth = int(datetime.strftime(datetime.now() - timedelta(1), '%m'))
    addMonth = []
    monthLength = (monthrange(2022, lastMonth)[1])
    monthLength1 = monthLength
    q = 0
    print(f"{lineNum()} Last month was {lastMonth}")
    logger.info(f"{lineNum()} Last month was {lastMonth}")
    print(f"{lineNum()} Last month had {monthLength} days")
    logger.info(f"{lineNum()} Last month had {monthLength} days")
    for l in range (monthLength):
        addMonth.append(10+l)
    for m in range (monthLength):
        while True:
            try:
                check = sheet.cell(9+m,1).value
                print(f"{lineNum()} Checking if {check} is Week Total")
                logger.info(f"{lineNum()} Checking if {check} is Week Total")
                if check == "Week Total":
                    print(f"{lineNum()} {check} is Week Total")
                    logger.info(f"{lineNum()} {check} is Week Total")
                    print(f"{lineNum()} Skipping over Week Total")
                    logger.info(f"{lineNum()} Skipping over Week Total")
                    for n in range(m,monthLength1):
                        addMonth[n-q] = addMonth[n-q] + 1
                    q += 1
                    monthLength1 += 1
                break
            except:
                print(f"{lineNum()} Waiting 7")
                logger.info(f"{lineNum()} Waiting 7")
                time.sleep(5)
    
    monthTotal = ["Month Total"]
    monthRow = ["=SUM("]
    Letters = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
    for p in range (25):
        for o in range(monthLength):
            monthRow.append(f"{Letters[p+1]}{addMonth[o]}")
            monthRow.append("+")

        monthRow.pop()
        monthRow.append(")")
        combine = "".join(monthRow)
        monthTotal.append(combine)
        print(f"{lineNum()} Month total, {monthTotal}")
        logger.info(f"{lineNum()} Month total, {monthTotal}")

    while True:
        try:
            print(f"{lineNum()} Adding month total to sheets")
            logger.info(f"{lineNum()} Adding month total to sheets")
            sheet.insert_row(monthTotal,9,value_input_option="USER_ENTERED")
            break
        except:
            print(f"{lineNum()} Waiting 8")
            logger.info(f"{lineNum()} Waiting 8")
            time.sleep(15)

def weekCheck():
    print(f"{lineNum()} Running week check function")
    logger.info(f"{lineNum()} Running week check function")
    yesterday = datetime.strftime(datetime.now() - timedelta(1), '%a - %d/%m/%y')
    while True:
        try:
            rowNine2 = sheet.cell(9,1).value
            print(f"{lineNum()} Check if row nine, {rowNine2} is yesterday date, {yesterday}")
            logger.info(f"{lineNum()} Check if row nine, {rowNine2} is yesterday date, {yesterday}")
            if rowNine2 != yesterday:
                print(f"{lineNum()} Dates do not match, running while loop")
                logger.info(f"{lineNum()} Dates do not match, running while loop")
                i=10
                while i != 0:
                    print(f"{lineNum()} i value is, {i}")
                    logger.info(f"{lineNum()} i value is, {i}")
                    i-=1
                    yesterdayDate = datetime.strftime(datetime.now() - timedelta(i), '%a - %d/%m/%y')
                    aYesterdayDate = [yesterdayDate]
                    while True:
                        try:
                            skipCell = 0
                            rowNine = sheet.cell(9,1).value
                            rowTen = sheet.cell(10,1).value
                            print(f"{lineNum()} Check for Month Total or Week Total on row 9 - {rowNine} and row 10 - {rowTen}")
                            logger.info(f"{lineNum()} Check for Month Total or Week Total on row 9 - {rowNine} and row 10 - {rowTen}")
                            if rowNine == "Month Total" or rowNine == "Week Total":
                                if rowTen == "Week Total" or rowTen == "Month Total":
                                    skipCell = 2
                                    print(f"{lineNum()} Both Week Total and Month Total found, skipCell is {skipCell}")
                                    logger.info(f"{lineNum()} Both Week Total and Month Total found, skipCell is {skipCell}")
                                else:
                                    skipCell = 1
                                    print(f"{lineNum()} Either Week Total or Month Total found, skipCell is {skipCell}")
                                    logger.info(f"{lineNum()} Either Week Total or Month Total found, skipCell is {skipCell}")
                            else:
                                print(f"{lineNum()} No total found, skipCell is {skipCell}")
                                logger.info(f"{lineNum()} No total found, skipCell is {skipCell}")
                            cell9Skip = sheet.cell(9+skipCell,1).value
                            print(f"{lineNum()} Check if {cell9Skip} and {yesterdayDate} match")
                            logger.info(f"{lineNum()} Check if {cell9Skip} and {yesterdayDate} match")
                            if cell9Skip == yesterdayDate:
                                print(f"{lineNum()} Match found")
                                logger.info(f"{lineNum()} Match found")
                                i-=1
                                while i != 0:
                                    yesterdayDate = datetime.strftime(datetime.now() - timedelta(i), '%a - %d/%m/%y')
                                    aYesterdayDate = [yesterdayDate,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
                                    yesterdayDay = datetime.strftime(datetime.now() - timedelta(i+1), '%a')
                                    yesterdayMonth = datetime.strftime(datetime.now() - timedelta(i), '%d')
                                    print(f"{lineNum()} Date check, {yesterdayDate}")
                                    logger.info(f"{lineNum()} Date check, {yesterdayDate}")
                                    if yesterdayDay == "Sun" and skipCell == 0:
                                        print(f"{lineNum()} End of week found")
                                        logger.info(f"{lineNum()} End of week found")
                                        endWeek()
                                    if yesterdayMonth == "01" and skipCell == 0:
                                        print(f"{lineNum()} End of month found")
                                        logger.info(f"{lineNum()} End of month found")
                                        endMonth()

                                    while True:
                                        try:
                                            print(f"{lineNum()} Inserting date, {yesterday} into sheets")
                                            logger.info(f"{lineNum()} Inserting date, {yesterday} into sheets")
                                            sheet.insert_row(aYesterdayDate,9)
                                            skipCell = 0
                                            i-=1
                                            break
                                        except:
                                            print(f"{lineNum()} Waiting 9")
                                            logger.info(f"{lineNum()} Waiting 9")
                                            time.sleep(15)
                                while True:
                                    try:
                                        currentDate = str(x.strftime("%a - %d/%m/%y"))

                                        rowNine = sheet.cell(9,1).value

                                        print(f"{lineNum()} Checking for end of week or end of month on final insert")
                                        logger.info(f"{lineNum()} Checking for end of week or end of month on final insert")
                                        if rowNine != currentDate:
                                            print(f"{lineNum()} Row 9, {rowNine} does not match current date, {currentDate}")
                                            logger.info(f"{lineNum()} Row 9, {rowNine} does not match current date, {currentDate}")
                                            if rowNine == datetime.strftime(datetime.now() - timedelta(1), '%a - %d/%m/%y'):
                                                if datetime.strftime(datetime.now() - timedelta(1), '%a') == "Sun":
                                                    print(f"{lineNum()} End of week found")
                                                    logger.info(f"{lineNum()} End of week found")
                                                    endWeek()
                                                elif datetime.strftime(datetime.now() - timedelta(0), '%d') == "01":
                                                    print(f"{lineNum()} End of month found")
                                                    logger.info(f"{lineNum()} End of month found")
                                                    endMonth()

                                        print(f"{lineNum()} Inserting current date. {currentDate}")
                                        logger.info(f"{lineNum()} Inserting current date. {currentDate}")
                                        date = [currentDate,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
                                        sheet.insert_row(date,9)
                                        break
                                    except:
                                        print(f"{lineNum()} Waiting 10")
                                        logger.info(f"{lineNum()} Waiting 10")
                                        time.sleep(15)
                                break
                            break
                        except:
                            print(f"{lineNum()} Waiting 11")
                            logger.info(f"{lineNum()} Waiting 11")
                            time.sleep(5)
            print(f"{lineNum()} Dates match, breaking loop")
            logger.info(f"{lineNum()} Dates match, breaking loop")
            break
        except:
            print(f"{lineNum()} Waiting 12")
            logger.info(f"{lineNum()} Waiting 12")
            time.sleep(5)


#--------------------------Main--------------------------#
x = tm.datetime.now()
currentDate = str(x.strftime("%a - %d/%m/%y"))
date = [currentDate,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]

rowNine = sheet.cell(9,1).value

print(f"{lineNum()} Current date is {currentDate}")
logger.info(f"{lineNum()} Current date is {currentDate}")
print(f"{lineNum()} Row 9 date is {rowNine}")
logger.info(f"{lineNum()} Row 9 date is {rowNine}")

if rowNine != currentDate:
    print(f"{lineNum()} Current date and row 9 date do not match")
    logger.info(f"{lineNum()} Current date and row 9 date do not match")
    if rowNine == datetime.strftime(datetime.now() - timedelta(1), '%a - %d/%m/%y'):
        print(f"{lineNum()} Row 9 is yesterdays date, {datetime.strftime(datetime.now() - timedelta(1), '%a - %d/%m/%y')}")
        logger.info(f"{lineNum()} Row 9 is yesterdays date, {datetime.strftime(datetime.now() - timedelta(1), '%a - %d/%m/%y')}")
        if datetime.strftime(datetime.now() - timedelta(1), '%a') == "Sun":
            print(f"{lineNum()} Yesterday day was sunday")
            logger.info(f"{lineNum()} Yesterday day was sunday")
            endWeek()
        elif datetime.strftime(datetime.now() - timedelta(0), '%d') == "01":
            print(f"{lineNum()} Today is the beginning of the month")
            logger.info(f"{lineNum()} Today is the beginning of the month")
            endMonth()
        while True:
            try:
                print(f"{lineNum()} Adding todays date to the sheets")
                logger.info(f"{lineNum()} Adding todays date to the sheets")
                sheet.insert_row(date,9)
                break
            except:
                print(f"{lineNum()} Waiting 13")
                logger.info(f"{lineNum()} Waiting 13")
                time.sleep(5)
    else:
        weekCheck()

window = tk.Tk()
window.title("Hour tracker")
window.geometry("500x305")
window['bg']="#F57A22"

login()

window.mainloop()

