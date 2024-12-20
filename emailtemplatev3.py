import win32com.client as win32
from datetime import datetime
from tkinter import ttk, filedialog
import tkinter as tk
from pathlib import Path
import os
from os.path import isfile

#Dictionary for templates
email_dict = {
    "Test Case" : {"Subject" : "A Test", 
                    "Body" : "Yay\n", 
                    "To" : "<Jarrod.Rudis@libertymutual.com>",
                    "From" : "<digitalmail_016C@libertymutual.com>"
                    }
}

signature = ""

main_window = tk.Tk()
main_window.config(width = 300, height = 200)
main_window.title("Mail Sending")

default_directory = tk.StringVar()
default_directory.set("C:/")

#In case of error when pushing a deprecated button, not currently used
#def callback():
    #key = var.get()
    #try:
        #value = email_dict[key]
        #print(value)
    #except KeyError:
        #print('Please choose an option')

#Prints entries in path selected, not currently used
#def path_printer():
    #directory = os.fsencode(chosen_directory["text"])
    
    #for entry in os.scandir(directory):
            #print(entry.path)

#Chooses directory, note that you cannot see files inside this directory
def directory_picker_method():
    default_directory.set(filedialog.askdirectory())

#Main method for sending emails with attachments
def send_email(fr, to, template, attachments, sig, manual = False):
    full_path = attachments.get()
    #x = 0 (variable for progress bar attempt)

    #Open up outlook
    outlook = win32.gencache.EnsureDispatch("Outlook.Application")
    
    progress_window = tk.Toplevel(main_window)
    progress_window.title("Sending Emails")
    progress_window.geometry("300x100")
    progress_window.resizable(0, 0)
    
    #email_progress = ttk.Progressbar(progress_window, variable = x, mode="determinate", length = len(next(os.walk(attachments.get()))[2])) (progress bar attempt)
    email_progress = ttk.Label(
        progress_window,
        text = "Sending ..."
    )
    email_progress.place(x = 25, y = 50, width = 250)

    #This is for sending emails with attachments in a folder, as long as they are pdfs
    for filename in os.listdir(attachments.get()):
        full_path = attachments.get() + "/" + filename
        progress_window.update()

        #For deciding if a file is a pdf
        if full_path.endswith(".pdf"):
            new_mail = outlook.CreateItem(0)

            new_mail.Subject = template["Subject"]

            new_mail.Body = template["Body"] + sig

            #For deciding if the template is to be used or manual input
            if manual == True: 
                #new_mail.CC = cc (placeholder from earlier as reminder)
                new_mail.To = to
                new_mail.SentOnBehalfOfName = fr
            
            else:
                #new_mail.CC = cc (placeholder from earlier as reminder)
                new_mail.To = template["To"]
                new_mail.SentOnBehalfOfName = template["From"]

            new_mail.Attachments.Add(Source = full_path)

            new_mail.Send()

        elif isfile(full_path) and full_path.endswith(".pdf") == False:
            print("File is not pdf")
        
        else:
             print("No file found")
  
    progress_window.destroy()


#Label for From: in email
from_label = tk.Label(
    text = "From:"
    )
from_label.grid(row = 0, column = 0, padx = 5, pady = 5)

#Dropdown box making for picking from email
from_email_picker = ttk.Combobox(
    state = "readonly",
    values = ["<digitalmail_016C@libertymutual.com>", "<BSC_Dover_West@libertymutual.com>"]
    )
from_email_picker.grid(row = 0, column = 1, padx = 5, pady = 5)

#Label for To: in email
to_label = tk.Label(
    text = "To:"
    )
to_label.grid(row = 1, column = 0, padx = 5, pady = 5)

#Dropdown box for picking to email
to_email_picker = ttk.Combobox(
    state = "readonly",
    values = ["<Jarrod.Rudis@libertymutual.com>"]
    )
to_email_picker.grid(row = 1, column = 1, padx = 5, pady = 5)

#Label for Choosing file directory
directory_choice = tk.Label(
    text = "Choose file directory:"
    )
directory_choice.grid(row = 2, column = 0, padx = 5, pady = 5)

#Button for picking directory
directory_picker = tk.Button(
    text = "Search",
    relief = tk.RAISED,
    command = directory_picker_method
)
directory_picker.grid(row = 2, column = 1, padx = 5, pady = 5)

#Keep track of directory chosen
chosen_directory = tk.Entry(
    textvariable = default_directory,
    state = "readonly",
    width = 40
    )
chosen_directory.grid(row = 3, column = 0, columnspan = 2, padx = 5, pady = 5)

#Button for printing picked directory
#choice_printer = tk.Button(
#    text = "Print choice",
#    relief = tk.RAISED,
#    command = path_printer
#)
#choice_printer.grid(row = 3, column = 1, padx = 5, pady = 5)

#Keep track of directory chosen
pick_email = tk.Label(
    text = "Pick an email:"
    )
pick_email.grid(row = 5, column = 0, padx = 5, pady = 5)

#Dropdown box for email templates
var = tk.StringVar()
email_type = ttk.Combobox(
    main_window,
    values = list(email_dict.keys()),
    textvariable = var,
    state = "readonly"
    )
email_type.current(0)
email_type.grid(row = 5, column = 1, padx = 5, pady = 5)

#Button to send email with selected information
emailer = tk.Button(
    text = "Send Email",
    relief = tk.RAISED,
    command = lambda: send_email(from_email_picker.get(), to_email_picker.get(), email_dict[var.get()], default_directory, signature)
)
emailer.grid(row = 6, column = 0, columnspan = 2, padx = 5, pady = 5)

main_window.resizable(0, 0)
main_window.geometry("300x200")
main_window.mainloop()