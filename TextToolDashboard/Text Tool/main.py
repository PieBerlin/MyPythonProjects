from tkinter import *
from tkinter import ttk
from tkinter.ttk import Combobox
from tkinter import messagebox,filedialog 
import tkinter as tk 
import googletrans

import speech_recognition as sr 
#import pyPDF2
import pyttsx3
import os

import datetime
from gtts import gTTS
from playsound import playsound
import threading

FRAMEBG="#161d3f"
BODYBG="#84a1e8"

root=tk.Tk()
root.title("Text Tool")
root.geometry("1030x570+290+140")
root.resizable(False,False)
root.config(bg=BODYBG)

#icon
image_icon=PhotoImage(file=r"Images/icon.png")
root.iconphoto(False,image_icon)

#Top frame
Top_frame=Frame(root,bg=FRAMEBG,width=1100,height=130)
Top_frame.place(x=0,y=0)

logo_icon=PhotoImage(file=r"Images/logo.png")
logo_label=Label(Top_frame,image=logo_icon,bg=FRAMEBG).place(x=40,y=20)

Label(Top_frame,text="Text Tool",font="arial 20 bold",bg=FRAMEBG,fg="white").place(x=190,y=40)
#Text area
text_area=Text(root,font="arial 15",bg="#cbe7c2",fg="black",wrap=WORD,relief=GROOVE)
text_area.place(x=40,y=140,width=600,height=200)

#Language combobox
language=googletrans.LANGUAGES
languageV=list(language.values())


combo1=ttk.Combobox(root,values=languageV,font="Roboto 10 ",state="r",width=10)
combo1.place(x=200,y=100)
combo1.set("english")




combo=ttk.Combobox(root,values=languageV,font="Roboto 10 ",state="r",width=10)
combo.place(x=330,y=100)
combo.set("english")



text_area2=Text(root,font="arial 15",bg="#dddaf5",fg="black",wrap=WORD ,relief=GROOVE)
text_area2.place(x=40,y=360,width=600,height=200)

#voice and speech controls
Label(root,text="Voice",font="arial 15 bold",bg=BODYBG,fg="black").place(x=700,y=200)
Label(root,text="Speed",font="arial 15 bold",bg=BODYBG,fg="black").place(x=700,y=250)

gender_combobox=Combobox(root,values=['Male','Female'],font="arial 10",state="r",width=40)
gender_combobox.place(x=800,y=200)
gender_combobox.set("Male")


#slider for speed

currentValue=tk.DoubleVar()
def get_current_value():
    return '{:.2f}'.format(currentValue.get())

def slider_changed(event):
    value_label.config(text=get_current_value())

style=ttk.Style()
style.configure("TScale",background=BODYBG)

slider=ttk.Scale(root,from_=30,to=250,orient=HORIZONTAL,command=slider_changed,variable=currentValue)
slider.place(x=800,y=250)

value_label=ttk.Label(root,text=get_current_value())
value_label.place(x=905,y=255)


#buttons
image_icon=PhotoImage(file=r"Images/speak.png")
btn= Button(root,image=image_icon,compound=LEFT,width=130,bg=BODYBG,fg="black",bd=0)
btn.place(x=700,y=300)

image_icon2=PhotoImage(file=r"Images/download.png")
save= Button(root,image=image_icon2,compound=LEFT,width=130,bg=BODYBG,fg="black",bd=0)
save.place(x=850,y=300)

pdfupload=PhotoImage(file=r"Images/pdfimage.png")
upload_button= Button(root,image=pdfupload,bg=BODYBG,bd=0)
upload_button.place(x=700,y=57)

upload_audioimage=PhotoImage(file=r"Images/music.png")
upload_audio_button= Button(root,image=upload_audioimage,bg=BODYBG,bd=0)
upload_audio_button.place(x=630,y=57)

transimage=PhotoImage(file=r"Images/trans.png")
trans_button= Button(root,image=transimage,bg=BODYBG,bd=0)
trans_button.place(x=550,y=60)

speakimage=PhotoImage(file=r"Images/otherspeaker.png")
speak_button= Button(root,image=speakimage,bg=BODYBG,bd=0)
speak_button.place(x=50,y=525)

micimage=PhotoImage(file=r"Images/mic.png")
mic_button= Button(root,image=micimage,bg=BODYBG,bd=0)
mic_button.place(x=50,y=305)

#pdf and text mode
button_mode=True
choice="Text"

def changemode():
    global button_mode
    global choice
    if button_mode:
        choice="PDF"
        mode.config(image=pdfmode,activebackground="white")
        button_mode=False
    else:
        choice="Text"
        mode.config(image=textmode,activebackground="white")
        button_mode=True

textmode=PhotoImage(file=r"Images/modeText.png")
pdfmode=PhotoImage(file=r"Images/modepdf.png")
mode=Button(root,image=textmode,bg=FRAMEBG,bd=0,command=changemode)
mode.place(x=780,y=40)




root.mainloop()

