import os
import sys
try:
    import requests
except ImportError:
    os.system('pip install requests')
    import requests
import subprocess
import tkinter as tk
from tkinter import messagebox

def update():
    file = requests.get("https://raw.githubusercontent.com/sensetraining/Work-hours/main/clock_in.pyw").content

    f = open("clock_in.pyw","wb")
    f.write(file)
    f.close()

    f = open("version.txt","w")
    f.write(version)
    f.close()

    window.destroy()

def cancel():
    response = messagebox.askquestion("Warning", "Are you sure you want to cancel?")
    if response == "yes":
        window.destroy()
    sys.exit()

def install():
    print("installing")
    os.system('winget install python.python.3.10')
    window.destroy()

version = requests.get("https://raw.githubusercontent.com/sensetraining/Work-hours/main/version.txt").text

f = open("version.txt","r").read()

print(f"Latest version: {version}")
print(f"Installed version: {f}")

try:
    subprocess.run(['python', '--version'], check=True, capture_output=True)
except subprocess.CalledProcessError:
    # Python is not installed
    print('Python is not installed')
    window = tk.Tk()
    window.title("Python Installer")
    window.geometry("400x270")
    window["bg"] = "#F57A22"

    plainBox1 = tk.Label(bg="#113F8C",relief="solid")
    plainBox1.place(height=195,width=325,relx=0.5,rely=0.5,anchor="center")

    nameLabel = tk.Label(plainBox1,text="Python needs installing to run",fg="white",bg="#113F8C",font=("Helvetica", 18, "bold"),wraplength=205)
    nameLabel.place(relx=0.5,y=40,height=60,anchor="center")

    updateButton = tk.Button(plainBox1,text="Install",command=install,bg="white",fg="black",font=("Helvetica", 12, "bold"))
    updateButton.place(relx=0.5,y=105,height=40,width=150,anchor="center")

    cancelButton = tk.Button(plainBox1,text="Cancel",command=cancel,bg="#ff4747",fg="white",font=("Helvetica", 12, "bold"))
    cancelButton.place(relx=0.5,y=155,height=40,width=150,anchor="center")

    window.mainloop()
else:
    print('Python is installed')


if f == version:
    print("Up to date")
    open("version.txt","r").close()

else:
    print("Available update")

    window = tk.Tk()
    window.title("Updater")
    window.geometry("400x250")
    window["bg"] = "#F57A22"

    plainBox1 = tk.Label(bg="#113F8C",relief="solid")
    plainBox1.place(height=175,width=325,relx=0.5,rely=0.5,anchor="center")

    nameLabel = tk.Label(plainBox1,text="Update available!",fg="white",bg="#113F8C",font=("Helvetica", 18, "bold"))
    nameLabel.place(relx=0.5,y=30,height=35,anchor="center")

    updateButton = tk.Button(plainBox1,text="Update",command=update,bg="white",fg="black",font=("Helvetica", 12, "bold"))
    updateButton.place(relx=0.5,y=85,height=40,width=150,anchor="center")

    cancelButton = tk.Button(plainBox1,text="Cancel",command=cancel,bg="#ff4747",fg="white",font=("Helvetica", 12, "bold"))
    cancelButton.place(relx=0.5,y=135,height=40,width=150,anchor="center")

    window.mainloop()

import clock_in