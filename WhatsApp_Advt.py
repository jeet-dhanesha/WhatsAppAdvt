from tkinter import *
import tkinter as tk
from tkinter.messagebox import *
from tkinter.filedialog import askopenfile
from PIL import ImageTk,Image
import xlrd
import time

from datetime import datetime
import os
import emoji
import pyautogui

default_bg = "grey70"

root = Tk()
root.title("WhatsApp Advt.")
root.geometry("788x600")
root.config(bg=default_bg)
root.resizable(0, 0)

photo = PhotoImage(file = r"myicon.png")
root.iconphoto(False, photo)

default_path = "No File Selected."
file_path = ""

try:
     import pywhatkit as kit
except:
     showinfo("Error...", "Network Issue...")
     root.destroy()

# Function responsible for the updation
# of the progress bar value

def Screen_3_Process(txt,txt1,select_file_button,start_advt_button,whatsapp_msg):
     global file_path
     txt.grid_forget()
     txt1.grid_forget()
     select_file_button.grid_forget()
     start_advt_button.grid_forget()

     wb = xlrd.open_workbook(file_path)
     sheet = wb.sheet_by_index(0)

     for i in range(1,sheet.nrows):
         now = datetime.now()
         current_time_hr = now.strftime("%H")
         current_time_min = now.strftime("%M")

         kit.sendwhatmsg("+91"+str(int(sheet.cell_value(i, 1))),whatsapp_msg+" *GOOD MORNING "+sheet.cell_value(i, 2).upper()+"*..."+emoji.emojize(":grinning_face_with_big_eyes:"),int(current_time_hr),int(current_time_min)+2)
         time.sleep(10)

         os.system("taskkill /im chrome.exe /f")


     showinfo("Message...", "Task Completed...")
     root.destroy()



def open_file(txt,txt1,select_file_button,start_advt_button):
    global default_path
    global file_path

    file = askopenfile(mode ='r', filetypes =[('Excel Sheet', '*.xlsx')])
    if file is not None:
        default_path = file.name
        file_path = file.name
        default_path = default_path.split("/")
        default_path = default_path[len(default_path)-1]
        if len(default_path) >= 18:
            default_path = default_path[0:15]+".."


        """
        txt = Label(root, text=default_path,bg=default_bg,foreground="black",font = ('Helvetica', 14, 'bold'))
        txt.grid(row=0,column=1,pady=(100,5),padx=(5,2))
        """
    txt.grid_forget()
    txt1.grid_forget()
    select_file_button.grid_forget()
    start_advt_button.grid_forget()
    Screen_2_Input()


def Screen_2_Input():
    global default_path
    global file_path
    if file_path == "":
        status_text_box = 'disabled'
    else:
        status_text_box = 'normal'
    txt = Label(root, text=default_path,bg=default_bg,foreground="black",font = ('Helvetica', 14, 'bold'))
    txt1 = Text(root, height=12, width=40,foreground="black",font = ('Comic Sans MS', 12),state=status_text_box)
    select_file_button = Button(text="Select File",bg="green",command=lambda: open_file(txt,txt1,select_file_button,start_advt_button),font = ('Arial', 14, 'bold'))
    start_advt_button = Button(text="Start Messaging",bg="green",command=lambda: Screen_3_Process(txt,txt1,select_file_button,start_advt_button,txt1.get("1.0",'end-1c')),font = ('Arial', 14, 'bold'))
    txt.grid(row=0,column=1,pady=(100,5),padx=(5,2))
    txt1.grid(row=1,column=0,columnspan=2,pady=(5,5),padx=(180,2))
    select_file_button.grid(row=0,column=0,pady=(100,5),padx=(180,2))
    start_advt_button.grid(row=2,column=0,columnspan=2,pady=(10,5),padx=(180,2))
    mainloop()


def Screen_1_Eraser(logo,txt,txt1,Password,submit_button,pwd):
    logo.place_forget()
    txt.place_forget()
    txt1.place_forget()
    Password.place_forget()
    submit_button.place_forget()
    if pwd == "123456":
        Screen_2_Input()
    else:
        showerror("Error...", "Wrong Pin Entered.")
        Screen_1_Pwd()


def Screen_1_Pwd():
    logo_path = PhotoImage(file = r"HEALTH OPTIONS LOGO.png")

    logo = Label(root, image=logo_path,bg=default_bg)

    txt = Label(root, text="WhatsApp Advt",bg=default_bg,foreground="green",font = ('Courier', 18, 'bold'))
    txt1 = Label(root, text="Enter Pin",bg=default_bg,foreground="black",font = ('Helvetica', 14, 'bold'))
    Password = Entry(root,show="*",foreground="black",font = ('Helvetica', 14, 'bold'))

    submit_button = Button(text="Submit",width=10,bg="green",command=lambda: Screen_1_Eraser(logo,txt,txt1,Password,submit_button,Password.get()),font = ('Arial', 14, 'bold'))
    logo.place(x=170,y=100)
    txt.place(x=310,y=320)
    txt1.place(x=230,y=360)
    Password.place(x=330,y=360)
    submit_button.place(x=335,y=407)
    mainloop()

if __name__ == "__main__":
     Screen_1_Pwd()
