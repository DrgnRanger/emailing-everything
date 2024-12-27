import win32com.client as win32
from datetime import datetime
from tkinter import ttk, filedialog
import tkinter as tk
import os
from os.path import isfile
import json

#Chooses directory, note that you cannot see files inside this directory
def directory_picker_method():
    default_directory.set(filedialog.askdirectory())

#Method for updating various comboboxes from dictionaries
def update_combobox():
    template_type_delete.config(values = list(variable_email_dictionary.keys()))
    template_type_delete.update()

    template_type_send.config(values = list(variable_email_dictionary.keys()))
    template_type_send.update()

    send_sig_choice.config(values = list(variable_signature_dictionary.keys()))
    send_sig_choice.update()

#Main method for sending emails with attachments
def send_email(fr, to, template, attachments, sig, manual):
    full_path = attachments.get()

    #Open up outlook
    outlook = win32.dynamic.Dispatch("Outlook.Application")
    
    #Popup for determining if emails have been sent
    progress_window = tk.Toplevel(main_window)
    progress_window.title("Sending Emails")
    progress_window.geometry("300x100")
    progress_window.resizable(0, 0)
    
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
                new_mail.To = "<" + to + ">"
                new_mail.SentOnBehalfOfName = "<" + fr + ">"
            
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

#Load dictionary in json from path
def load_dict_json(path):
    with open(path) as readfile:
        initial_import_json = json.load(readfile)
    
    print(initial_import_json)
    return initial_import_json

#Send dictionary to json in same directory    
def send_to_json(path, edited_in):
    with open(path, "w") as ofile:
        json.dump(edited_in, ofile, indent = 4)

#Delete the specified entry in a dictionary
def delete_entry(dict, del_path, key):
    if key in dict:
        del dict[key]
    
    else:
        print("No key found")

    print(dict)
    send_to_json(del_path, dict)
    
    #Updating comboboxes
    update_combobox()


#Add an entry to the email dictionary
def add_email_entry(email_template_name, sub, bod, to, frm, email_in_dict):
    email_in_dict[email_template_name] = {"Subject" : sub,
                            "Body" : bod,
                            "To" : "<" + to + ">",
                            "From" : "<" + frm + ">"
                            }
    
    variable_email_dictionary = email_in_dict
    print(variable_email_dictionary)
    send_to_json(json_email_path, variable_email_dictionary)
    
    #Updating comboboxes
    update_combobox()

def add_signature_entry(sig_template_name, sig, sig_in_dict):
    sig_in_dict[sig_template_name] = sig
                            

    variable_signature_dictionary = sig_in_dict
    print(variable_signature_dictionary)
    send_to_json(json_signature_path, variable_signature_dictionary)

    #Updating comboboxes
    update_combobox()

#Global variables
json_email_path = "editable_email_dict.json"
json_signature_path = "editable_signature.json"
variable_email_dictionary = {}
variable_signature_dictionary = {}

#Initialize main window
main_window = tk.Tk()
main_window.config(width = 315, height = 490)
main_window.title("Mail Sending")
tab_control = ttk.Notebook(main_window)

#Loading the various dictionaries from jsons
variable_email_dictionary = load_dict_json(json_email_path)
variable_signature_dictionary = load_dict_json(json_signature_path)

#Setting default directory variable for changing later
default_directory = tk.StringVar()
default_directory.set("C:/")

#Settiing up tabs for the window
tab_1 = ttk.Frame(tab_control)
tab_2 = ttk.Frame(tab_control)

tab_control.add(tab_1, text = "Send Emails")
tab_control.add(tab_2, text = "Edit Templates")
tab_control.pack(expand = 1, fill ="both") 

#TAB 1

#Label for manual emailing option
manual_pick_label = tk.Label(
    tab_1,
    text = "Manually Emailing:"
)
manual_pick_label.grid(row = 0, column = 0, padx = 5, pady = 5)

#Dropdown for manual picking
manual_picker = ttk.Combobox(
    tab_1,
    state = "readonly",
    values = [True, False]
)
manual_picker.current(1)
manual_picker.grid(row = 0, column = 1, padx = 5, pady = 5)

#Label for From: in email
from_label = tk.Label(
    tab_1,
    text = "From:"
    )
from_label.grid(row = 1, column = 0, padx = 5, pady = 5)

#Dropdown box making for picking manual from email
from_email_picker = tk.Entry(
    tab_1
    )
from_email_picker.grid(row = 1, column = 1, padx = 5, pady = 5)

#Label for To: in email
to_label = tk.Label(
    tab_1,
    text = "To:"
    )
to_label.grid(row = 2, column = 0, padx = 5, pady = 5)

#Dropdown box for picking manual to email
to_email_picker = tk.Entry(
    tab_1
    )
to_email_picker.grid(row = 2, column = 1, padx = 5, pady = 5)

#Label for adding entry section
send_template_section = tk.Label(
    tab_1,
    text = "Send Via Template",
)
send_template_section.grid(row = 3, column = 0, columnspan = 2, padx = 5, pady = 5)

#Label for Choosing file directory
directory_choice = tk.Label(
    tab_1,
    text = "Choose file directory:"
    )
directory_choice.grid(row = 4, column = 0, padx = 5, pady = 5)

#Button for picking directory
directory_picker = tk.Button(
    tab_1,
    text = "Search",
    relief = tk.RAISED,
    command = directory_picker_method
)
directory_picker.grid(row = 4, column = 1, padx = 5, pady = 5)

#Keep track of directory chosen
chosen_directory = tk.Entry(
    tab_1,
    textvariable = default_directory,
    state = "readonly",
    width = 40
    )
chosen_directory.grid(row = 5, column = 0, columnspan = 2, padx = 5, pady = 5)

#Keep track of directory chosen
pick_email = tk.Label(
    tab_1,
    text = "Pick a template:"
    )
pick_email.grid(row = 6, column = 0, padx = 5, pady = 5)

#Dropdown box for email templates
var_email_choice = tk.StringVar()
template_type_send = ttk.Combobox(
    tab_1,
    values = list(variable_email_dictionary.keys()),
    textvariable = var_email_choice,
    state = "readonly"
    )
template_type_send.grid(row = 6, column = 1, padx = 5, pady = 5)

#Button to send email with selected information
emailer = tk.Button(
    tab_1,
    text = "Send Email",
    relief = tk.RAISED,
    command = lambda: send_email(from_email_picker.get(), to_email_picker.get(), variable_email_dictionary[var_email_choice.get()], default_directory, var_signature_choice.get(), manual_picker.get())
)
emailer.grid(row = 7, column = 0, columnspan = 2, padx = 5, pady = 5)

#Label for detailing how the sending email button works
send_template_section = tk.Label(
    tab_1,
    text = "(If manual is true, email addresses in template will be overwritten)",
)
send_template_section.grid(row = 8, column = 0, columnspan = 2, padx = 5, pady = 5)

#Label for signature picking
pick_signature = tk.Label(
    tab_1,
    text = "Pick a signature:"
    )
pick_signature.grid(row = 9, column = 0, padx = 5, pady = 5)

#Dropdown for signature picking
var_signature_choice = tk.StringVar()
send_sig_choice = ttk.Combobox(
    tab_1,
    values = list(variable_signature_dictionary.keys()),
    textvariable = var_signature_choice,
    state = "readonly"
    )
send_sig_choice.current(0)
send_sig_choice.grid(row = 9, column = 1, padx = 5, pady = 5)

#Button for deleting a dictionary entry
del_signature_entry_button = tk.Button(
    tab_1,
    text = "Delete Chosen Sig. Entry",
    relief = tk.RAISED,
    command = lambda: delete_entry(variable_signature_dictionary, json_signature_path, var_signature_choice.get())
)
del_signature_entry_button.grid(row = 10, column = 0, columnspan = 2, padx = 5, pady = 5)

#Text field for adding the signature
sig_full = tk.Text(
    tab_1,
    height = 5,
    width = 40
)
sig_full.grid(row = 11, column = 0, columnspan = 2, padx = 5, pady = 5)

#Label for name when adding signature
sig_entry_name = tk.Label(
    tab_1,
    text = "Sig. Name"
)
sig_entry_name.grid(row = 12, column = 0, padx = 5, pady = 5)

#Name of signature entry 
sig_entry_name_in = tk.Entry(
    tab_1,
    width = 15
)
sig_entry_name_in.grid(row = 12, column = 1, padx = 5, pady = 5)

#Button for adding to email dictionary
add_signature_entry_button = tk.Button(
    tab_1,
    text = "Add Entry",
    relief = tk.RAISED,
    command = lambda: add_signature_entry(sig_entry_name_in.get(), sig_full.get("1.0", "end"), variable_signature_dictionary)
)
add_signature_entry_button.grid(row = 13, column = 0, columnspan = 2, padx = 5, pady = 5)

#TAB 2

#Label for entry deletion
del_entry_label = tk.Label(
    tab_2,
    text = "Entry to Delete:"
    )
del_entry_label.grid(row = 0, column = 0, padx = 5, pady = 5)

#Combobox for selecting an entry
entry_var = tk.StringVar()
template_type_delete = ttk.Combobox(
    tab_2,
    values = list(variable_email_dictionary.keys()),
    textvariable = entry_var,
    state = "readonly"
    )
template_type_delete.grid(row = 0, column = 1, padx = 5, pady = 5)

#Button for deleting a dictionary entry
del_entry_button = tk.Button(
    tab_2,
    text = "Delete Entry",
    relief = tk.RAISED,
    command = lambda: delete_entry(variable_email_dictionary, json_email_path, entry_var.get())
)
del_entry_button.grid(row = 1, column = 0, columnspan = 2, padx = 5, pady = 5)

#Label for adding entry section
add_entry_label = tk.Label(
    tab_2,
    text = "Add an Entry",
)
add_entry_label.grid(row = 2, column = 0, columnspan = 2, padx = 5, pady = 5)

#Label for adding entry section
add_subject_label = tk.Label(
    tab_2,
    text = "Subject:",
)
add_subject_label.grid(row = 3, column = 0, padx = 5, pady = 5)

#Entry for subject line
subject_line = tk.Entry(
    tab_2,
    width = 30
    )
subject_line.grid(row = 3, column = 1, padx = 5, pady = 5)

#Label for adding to section
to_label = tk.Label(
    tab_2,
    text = "To:",
)
to_label.grid(row = 4, column = 0, padx = 5, pady = 5)

#Entry for to section
to_line = tk.Entry(
    tab_2,
    width = 30
    )
to_line.grid(row = 4, column = 1, padx = 5, pady = 5)

#Label for adding from section
fr_label = tk.Label(
    tab_2,
    text = "From:",
)
fr_label.grid(row = 5, column = 0, padx = 5, pady = 5)

#Entry for from section
fr_line = tk.Entry(
    tab_2,
    width = 30
    )
fr_line.grid(row = 5, column = 1, padx = 5, pady = 5)

#Label for body
fr_label = tk.Label(
    tab_2,
    text = "Body:"
)
fr_label.grid(row = 6, column = 0, columnspan= 2, padx = 5, pady = 5)

#Text field for adding the body
body_lines = tk.Text(
    tab_2,
    height = 10,
    width = 35
)
body_lines.grid(row = 7, column = 0, columnspan = 2, padx = 5, pady = 5)

#Label for name of email template
template_label = tk.Label(
    tab_2,
    text = "Name of Template:"
)
template_label.grid(row = 8, column = 0, padx = 5, pady = 5)

#Text field for adding the body
template_name = tk.Entry(
    tab_2,
    width = 30
)
template_name.grid(row = 8, column = 1, padx = 5, pady = 5)

#Button for adding to email dictionary
del_entry_button = tk.Button(
    tab_2,
    text = "Add Entry",
    relief = tk.RAISED,
    command = lambda: add_email_entry(template_name.get(), subject_line.get(), body_lines.get("1.0", "end"), to_line.get(), fr_line.get(), variable_email_dictionary)
)
del_entry_button.grid(row = 9, column = 0, columnspan = 2, padx = 5, pady = 5)

#END TABS

main_window.resizable(0, 0)
main_window.mainloop()