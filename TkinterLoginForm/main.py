import tkinter as tk
from tkinter import messagebox

BACKGROUD_COLOR='#333333'
FOREGROUND_COLOR='#ffffff'

window=tk.Tk()
window.title("Login Form")
window.geometry('340x340')
window.configure(bg=BACKGROUD_COLOR)

def clear_entries():
    username_entry.delete(0,tk.END)
    password_entry.delete(0,tk.END)


def login():
    username="Gorova"
    password="Today@123" #Not secure 😂
    if username_entry.get()==username and password_entry.get()==password:
        messagebox.showinfo(title="Login Success",message="You successfully logged in ")
        print("Logen in")
        clear_entries()
    else:
        messagebox.showerror(title="Login Error",message="Please Use valid credantials")
        print("Invalid login")
        clear_entries()


frame=tk.Frame(bg=BACKGROUD_COLOR)

#creeating widgets
login_label=tk.Label(frame,text="Login",bg=BACKGROUD_COLOR,fg='#ff3399',font=('Arial',30))
username_label=tk.Label(frame,text="Username",bg=BACKGROUD_COLOR,fg=FOREGROUND_COLOR,font=('Arial',16))
username_entry=tk.Entry(frame,font=('Arial',16))
password_entry=tk.Entry(frame,show="*",font=('Arial',16))
password_label=tk.Label(frame,text="Password",bg=BACKGROUD_COLOR,fg=FOREGROUND_COLOR,font=('Arial',16))
login_button=tk.Button(frame,text="Login",bg='#ff3399',fg=FOREGROUND_COLOR,font=('Arial',16),command=login)

#placing widgets on the screen

frame.pack()

login_label.grid(row=0,column=0,columnspan=2,sticky='news',pady=40)
username_label.grid(row=1,column=0)
username_entry.grid(row=1,column=1,pady=20)
password_entry.grid(row=2,column=1,pady=20)
password_label.grid(row=2,column=0)
login_button.grid(row=3,column=0,columnspan=2,pady=30)




window.mainloop()