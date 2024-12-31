import win32com.client as win32
from datetime import datetime
from tkinter import ttk, filedialog
import tkinter as tk
import os
from os.path import isfile
import json

#Create a choice window that pops up and allows choices on button press
class choice_window(tk.Toplevel):
    def __init__(self, affirmative, closing, text):
        tk.Toplevel.__init__(self)
        self.title("")
        self.resizable(0, 0)

        for i in range(0, 1):
            self.grid_columnconfigure(i, weight = 1)
    
        for i in range(0, 1):
            self.grid_rowconfigure(i, weight = 1)

        #Loading description label
        self.load_choice = tk.Label(self)
        self.load_choice["text"] = text
        self.load_choice.grid(row = 0, column = 0, columnspan = 2, padx = 5, pady = 5)

        #Yes button
        self.yes_load = tk.Button(self)
        self.yes_load["text"] = "Yes"
        self.yes_load["relief"] = tk.RAISED
        self.yes_load["command"] = affirmative
        self.yes_load.grid(row = 1, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

        #Yes button
        self.no_load = tk.Button(self)
        self.no_load["text"] = "No"
        self.no_load["relief"] = tk.RAISED
        self.no_load["command"] = closing
        self.no_load.grid(row = 1, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))
        self.no_load.grid(row = 1, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Create a tooltip for a given widget, this was sourced on stackoverflow
class CreateToolTip(tk.Toplevel):
    def __init__(self, widget, text='widget info'):
        self.waittime = 200     #miliseconds
        self.wraplength = 180   #pixels
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None

    def enter(self, event = None):
        self.schedule()

    def leave(self, event = None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.waittime, self.showtip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def showtip(self, event=None):
        x = y = 0
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        #Creates a toplevel window
        self.tw = tk.Toplevel(self.widget)
        #Leaves only the label and removes the app window
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(self.tw, text = self.text, justify = 'left',
                       background ="#ffffff", relief = 'solid', borderwidth = 1,
                       wraplength = self.wraplength)
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tw
        self.tw= None
        if tw:
            tw.destroy()

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
    progress_window.geometry("100x50")
    progress_window.resizable(0, 0)
    
    email_progress = ttk.Label(
        progress_window,
        text = "Sending ..."
    )
    email_progress.grid(row = 0, column = 0, padx = 5, pady = 5)

    progress_window.grid_columnconfigure(0, weight = 1)
    progress_window.grid_rowconfigure(0, weight = 1)

    #This is for sending emails with attachments in a folder, as long as they are pdfs
    for filename in os.listdir(attachments.get()):
        full_path = attachments.get() + "/" + filename
        progress_window.update()

        #For deciding if a file is a pdf
        if full_path.endswith(".pdf"):
            new_mail = outlook.CreateItem(0)

            #Determining if the subject line should be today's date or not
            if template["Subject"] == "Today" or template["Subject"] == "today":
                new_mail.Subject = datetime.today().strftime("%m/%d/%y")

            else:    
                new_mail.Subject = template["Subject"]

            new_mail.Body = template["Body"] + sig

            #For deciding if the template is to be used or manual input
            if manual == True: 
                split_to = to.split()
                #split_cc = cc.split()
                split_fr = fr.split()

                email_to_list = ["<" + i + ">" for i in split_to]
                #email_cc_list = ["<" + i + ">" for i in split_cc]
                email_fr_list = ["<" + i + ">" for i in split_fr]

                combine_to = ""
                #combine_cc = ""
                combine_fr = ""

                combine_to = "; ".join(email_to_list)
                #combine_cc = "; ".join(email_cc_list)
                combine_fr = "; ".join(email_fr_list)

                #new_mail.CC = cc (this may be added, functionality above)
                new_mail.To = combine_to
                new_mail.SentOnBehalfOfName = combine_fr
            
            else:
                if template["CC"] != "":
                    new_mail.CC = template["CC"]
                
                else:
                    pass

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
    #Methods for using choice window
    def key_delete():
        if key in dict:
            del dict[key]
        
        else:
            print("No key found")

        #Cleanup files and finalize deletion
        print(dict)
        send_to_json(del_path, dict)
        deletion_window.destroy()

            #Updating comboboxes
        update_combobox()
    
    def no_delete():
        deletion_window.destroy()

    deletion_text = "Are you sure you want to delete this entry?\n(it will be unrecoverable)"

    deletion_window = choice_window(key_delete, no_delete, deletion_text)

#Add an entry to the email dictionary
def add_email_entry(email_template_name, sub, bod, to, cc, frm, email_in_dict):
    #Checking if to and from have <> so as not to stack
    split_to = to.split()
    split_cc = cc.split
    split_frm = frm.split()

    for char in to:
        if char == "<" or char == ">":
            add_to = to

        else:
            email_to_list = ["<" + i + ">" for i in split_to]
            add_to = "; ".join(email_to_list)

    for char in cc:
        if char == "<" or char == ">":    
            add_cc = cc

        else:
            email_cc_list = ["<" + i + ">" for i in split_cc]
            add_cc = "; ".join(email_cc_list)

    for char in frm:
        if char == "<" or char == ">":    
            add_frm = frm

        else:   
            email_frm_list = ["<" + i + ">" for i in split_frm]
            add_frm = "; ".join(email_frm_list)

    email_in_dict[email_template_name] = {"Subject" : sub,
                            "Body" : bod,
                            "To" : add_to,
                            "CC" : add_cc,
                            "From" : add_frm
                            }
    
    variable_email_dictionary = email_in_dict
    print(variable_email_dictionary)
    send_to_json(json_email_path, variable_email_dictionary)
    
    #Updating comboboxes
    update_combobox()

#Method for adding signature
def add_signature_entry(sig_template_name, sig, sig_in_dict):
    sig_in_dict[sig_template_name] = sig
                            
    variable_signature_dictionary = sig_in_dict
    send_to_json(json_signature_path, variable_signature_dictionary)

    #Updating comboboxes
    update_combobox()

#Method for loading signature entry
def load_selected_sig_entry(selected_sig_key):
    #Method definition of what buttons will do
    def true_button():
        sig_entry_name_in.delete(0, "end")
        sig_full.delete("1.0", "end")

        sig_entry_name_in.insert(0, selected_sig_key)
        sig_full.insert("end", variable_signature_dictionary[selected_sig_key])

        loading_signature.destroy()
    
    def false_button():
        loading_signature.destroy()

    #Text for button
    loading_text = "Do you want to load this entry?\n(this will delete what is written)"

    #Class for popup
    loading_signature = choice_window(true_button, false_button, loading_text)

#Method for loading email template
def load_selected_email_entry(selected_email_key):
    #Method definition of what buttons will do
    def true_button(): 
        subject_line.delete(0, "end")
        to_line.delete(0, "end")
        cc_line.delete(0, "end")
        fr_line.delete(0, "end")
        body_lines.delete("1.0", "end")
        template_name.delete(0, "end")
    
        subject_line.insert(0, variable_email_dictionary[selected_email_key]["Subject"])
        to_line.insert(0, variable_email_dictionary[selected_email_key]["To"])
        cc_line.insert(0, variable_email_dictionary[selected_email_key]["CC"])
        fr_line.insert(0, variable_email_dictionary[selected_email_key]["From"])
        body_lines.insert("1.0", variable_email_dictionary[selected_email_key]["Body"])
        template_name.insert(0, selected_email_key)
    
        loading_email.destroy()

    def false_button():
        loading_email.destroy()

    #Text for button
    loading_text = "Do you want to load this entry?\n(this will delete what is written)"
    
    #Class for popup
    loading_email = choice_window(true_button, false_button, loading_text)

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
tab_control.pack(expand = 1, fill = "both")

#TAB 1

#Label for manual emailing option
manual_pick_label = tk.Label(
    tab_1,
    text = "Manually Input Emails:"
)
manual_pick_label.grid(row = 0, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Dropdown for manual picking
manual_picker = ttk.Combobox(
    tab_1,
    state = "readonly",
    values = [True, False]
)
manual_picker.current(1)
manual_picker.grid(row = 0, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Label for From: in email
from_label = tk.Label(
    tab_1,
    text = "From:"
    )
from_label.grid(row = 1, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Dropdown box making for picking manual from email
from_email_picker = tk.Entry(
    tab_1
    )
from_email_picker.grid(row = 1, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))
frm_pick_ttp = CreateToolTip(from_email_picker, "For multiple emails, separate with a space.")

#Label for To: in email
to_label = tk.Label(
    tab_1,
    text = "To:"
    )
to_label.grid(row = 2, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Dropdown box for picking manual to email
to_email_picker = tk.Entry(
    tab_1
    )
to_email_picker.grid(row = 2, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))
to_pick_ttp = CreateToolTip(to_email_picker, "For multiple emails, separate with a space.")

#Label for adding entry section
send_template_section = tk.Label(
    tab_1,
    text = "Send Via Template",
)
send_template_section.grid(row = 3, column = 0, columnspan = 2, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Label for Choosing file directory
directory_choice = tk.Label(
    tab_1,
    text = "Choose file directory:"
    )
directory_choice.grid(row = 4, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Button for picking directory
directory_picker = tk.Button(
    tab_1,
    text = "Search",
    relief = tk.RAISED,
    command = directory_picker_method
)
directory_picker.grid(row = 4, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Keep track of directory chosen
chosen_directory = tk.Entry(
    tab_1,
    textvariable = default_directory,
    state = "readonly",
    width = 40
    )
chosen_directory.grid(row = 5, column = 0, columnspan = 2, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Keep track of directory chosen
pick_email = tk.Label(
    tab_1,
    text = "Pick a template:"
    )
pick_email.grid(row = 6, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Dropdown box for email templates
var_email_choice = tk.StringVar()
template_type_send = ttk.Combobox(
    tab_1,
    values = list(variable_email_dictionary.keys()),
    textvariable = var_email_choice,
    state = "readonly"
    )
template_type_send.grid(row = 6, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Button to send email with selected information
emailer = tk.Button(
    tab_1,
    text = "Send Email",
    relief = tk.RAISED,
    command = lambda: send_email(
        from_email_picker.get(), 
        to_email_picker.get(), 
        variable_email_dictionary[var_email_choice.get()], 
        default_directory, 
        variable_signature_dictionary[var_signature_choice.get()], 
        manual_picker.get()
        )
)
emailer.grid(row = 7, column = 0, columnspan = 2, padx = 5, pady = 5, sticky = (tk.E, tk.W))
emailer_ttp = CreateToolTip(emailer, "If manual is true, email addresses in template will be overwritten.")

#Label for signature picking
pick_signature = tk.Label(
    tab_1,
    text = "Pick a signature:"
    )
pick_signature.grid(row = 8, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Dropdown for signature picking
var_signature_choice = tk.StringVar()
send_sig_choice = ttk.Combobox(
    tab_1,
    values = list(variable_signature_dictionary.keys()),
    textvariable = var_signature_choice,
    state = "readonly"
    )
send_sig_choice.current(0)
send_sig_choice.grid(row = 8, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Button for deleting a dictionary entry
del_signature_entry_button = tk.Button(
    tab_1,
    text = "Delete Chosen Sig. Entry",
    relief = tk.RAISED,
    command = lambda: delete_entry(
        variable_signature_dictionary, 
        json_signature_path, 
        var_signature_choice.get()
        )
)
del_signature_entry_button.grid(row = 9, column = 0, columnspan = 2, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Text field for adding the signature
sig_full = tk.Text(
    tab_1,
    height = 5,
    width = 40
)
sig_full.grid(row = 10, column = 0, columnspan = 2, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Label for name when adding signature
sig_entry_name = tk.Label(
    tab_1,
    text = "Sig. Name"
)
sig_entry_name.grid(row = 11, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Name of signature entry 
sig_entry_name_in = tk.Entry(
    tab_1,
    width = 15
)
sig_entry_name_in.grid(row = 11, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Button for adding to email dictionary
add_signature_entry_button = tk.Button(
    tab_1,
    text = "Add Signature Entry",
    relief = tk.RAISED,
    command = lambda: add_signature_entry(
        sig_entry_name_in.get(), 
        sig_full.get("1.0", "end"), 
        variable_signature_dictionary
        )
)
add_signature_entry_button.grid(row = 12, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Button for loading signature from choice
load_signature_entry_button = tk.Button(
    tab_1,
    text = "Load Signature Entry",
    relief = tk.RAISED,
    command = lambda: load_selected_sig_entry(var_signature_choice.get())
)
load_signature_entry_button.grid(row = 12, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#TAB 2

#Label for entry deletion
del_entry_label = tk.Label(
    tab_2,
    text = "Entry to Delete/Load:"
    )
del_entry_label.grid(row = 0, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Combobox for selecting an entry
del_entry_var = tk.StringVar()
template_type_delete = ttk.Combobox(
    tab_2,
    values = list(variable_email_dictionary.keys()),
    textvariable = del_entry_var,
    state = "readonly"
    )
template_type_delete.current(0)
template_type_delete.grid(row = 0, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Button for deleting a dictionary entry
del_entry_button = tk.Button(
    tab_2,
    text = "Delete Entry",
    relief = tk.RAISED,
    command = lambda: delete_entry(
        variable_email_dictionary, 
        json_email_path, 
        del_entry_var.get()
        )
)
del_entry_button.grid(row = 1, column = 0, columnspan = 2, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Label for adding entry section
add_entry_label = tk.Label(
    tab_2,
    text = "Add an Entry"
)
add_entry_label.grid(row = 2, column = 0, columnspan = 2, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Label for adding entry section
add_subject_label = tk.Label(
    tab_2,
    text = "Subject:"
)
add_subject_label.grid(row = 3, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Entry for subject line
subject_line = tk.Entry(
    tab_2,
    width = 30
    )
subject_line.grid(row = 3, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))
subject_ttp = CreateToolTip(subject_line, "To make the template enter the date the email is sent, enter 'today' or 'Today' as the subject line.")

#Label for adding to section
to_label = tk.Label(
    tab_2,
    text = "To:"
)
to_label.grid(row = 4, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Entry for to section
to_line = tk.Entry(
    tab_2,
    width = 30
    )
to_line.grid(row = 4, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))
to_ttp = CreateToolTip(to_line, "For multiple emails, separate with a space.")

#Label for adding cc section
cc_label = tk.Label(
    tab_2,
    text = "cc:"
)
cc_label.grid(row = 5, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Entry for to section
cc_line = tk.Entry(
    tab_2,
    width = 30
    )
cc_line.grid(row = 5, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))
cc_ttp = CreateToolTip(to_line, "For multiple emails, separate with a space.")


#Label for adding from section
fr_label = tk.Label(
    tab_2,
    text = "From:"
)
fr_label.grid(row = 6, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Entry for from section
fr_line = tk.Entry(
    tab_2,
    width = 30
    )
fr_line.grid(row = 6, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))
fr_ttp = CreateToolTip(fr_line, "For multiple emails, separate with a space.")

#Label for body
fr_label = tk.Label(
    tab_2,
    text = "Body:"
)
fr_label.grid(row = 7, column = 0, columnspan = 2, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Text field for adding the body
body_lines = tk.Text(
    tab_2,
    height = 10,
    width = 40
)
body_lines.grid(row = 8, column = 0, columnspan = 2, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Label for name of email template
template_label = tk.Label(
    tab_2,
    text = "Name of Template:"
)
template_label.grid(row = 9, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Text field for adding the body
template_name = tk.Entry(
    tab_2,
    width = 30
)
template_name.grid(row = 9, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))
temp_name_ttp = CreateToolTip(template_name, "If you use the same name as a new template, or load a template, any changes you have made will overwrite the original.")

#Button for adding to email dictionary
del_entry_button = tk.Button(
    tab_2,
    text = "Add Email Template Entry",
    relief = tk.RAISED,
    command = lambda: add_email_entry(
        template_name.get(), 
        subject_line.get(), 
        body_lines.get("1.0", "end"), 
        to_line.get(),
        cc_line.get(), 
        fr_line.get(), 
        variable_email_dictionary
        )
)
del_entry_button.grid(row = 10, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Button for loading email template from choice
load_email_entry_button = tk.Button(
    tab_2,
    text = "Load Email Entry",
    relief = tk.RAISED,
    command = lambda: load_selected_email_entry(del_entry_var.get())
)
load_email_entry_button.grid(row = 10, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#For keeping grid in line
for i in range(0, 1):
    tab_1.columnconfigure(i, weight = 1)

for i in range(0, 13):  
    tab_1.rowconfigure(i, weight = 1)

for i in range(0, 1):
    tab_2.columnconfigure(i, weight = 1)

for i in range(0, 10):
    tab_2.rowconfigure(i, weight = 1)

#END TABS

main_window.resizable(0, 0)
main_window.mainloop()