import win32com.client as win32
from datetime import datetime
from tkinter import ttk, filedialog
import tkinter as tk
import os
import csv
from os.path import isfile
import json
import extract_msg as ext
import pymupdf as mu
import re

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

#Handles events when window is closed
def on_closing(event):
    if event.widget != main_window:
        return

    if not count_exported:
        export_count(variable_count_dictionary)

#Handles when specific count is selected in specific combobox
def update_count(event):
    number_of_type.set(variable_count_dictionary[var_count_choice.get()]["Count"])

#Handles when user wants to set count of emails sent manually
def change_count(number):
    
    #Functions for buttons in count change box
    def true_button():
        #Changes the count of selected type in combobox
        variable_count_dictionary[var_count_choice.get()]["Count"] = number
        print("Count changed for " + var_count_choice.get())
        count_change_w.destroy()

    def false_button():
        count_change_w.destroy()
    
    #Text for count change window
    count_change_text = "Are you sure you want \nto change this count?"

    #Double check if user wishes to change the count or not
    count_change_w = choice_window(true_button, false_button, count_change_text)

#Exports count data to csv file
def export_count(in_count):
    global count_exported
    
    #Opens csv and writes data to it
    with open(fname, "a", newline = '') as f:
        writer = csv.writer(f, delimiter = ",")
        for x in list(in_count):
            writer.writerows([[x, in_count[x]["Count"]]])
            #variable_count_dictionary[x]["Count"] = 0
    
    count_exported = True
    
    f.close()
    print("Count Exported")
    

#Debug printer for certain tools 
def print_debug():
    print("Sending emails with input from pdfs: " + str(take_pdf_input.get()))

#Method for updating various comboboxes from dictionaries
def update_combobox():
    template_type_delete.config(values = list(variable_email_dictionary.keys()))
    template_type_delete.update()

    template_type_send.config(values = list(variable_email_dictionary.keys()))
    template_type_send.update()

    send_sig_choice.config(values = list(variable_signature_dictionary.keys()))
    send_sig_choice.update()

    count_choice.config(values = list(variable_count_dictionary.keys()))
    count_choice.update()

#Main method for sending emails with attachments
def send_email(fr, to, template, attachments, sig, manual, is_pdf_input):
    #Function for search pdf text for specific user defined inputs
    def search_pdf_for(in_query, pdf_text):
        #User input followed by a number of any length
        if "##" in in_query:
            in_query = in_query.replace("##", "")
            completed_query = re.findall(rf"{in_query}\d+", pdf_text)

        #User input followed by a word of any length
        elif "ww" in in_query:
            in_query = in_query.replace("ww", "")
            completed_query = re.findall(rf"{in_query}\w+", pdf_text)

        #User input followed by a phone number
        elif "p#" in in_query:
            in_query = in_query.replace("p#", "")
            completed_query = re.findall(rf"{in_query}\(\d\{{3}}\).\d{{3}}.\d{{4}}|{in_query}\d{{3}}.\d{{3}}.\d{{4}}", pdf_text)

        #User input followed by a name (technically two capitalized words) in First Last or Last, First
        elif "nn" in in_query:
            in_query = in_query.replace("nn", "")
            completed_query = re.findall(rf"{in_query}[A-Z]\w+ [A-Z]\w+|{in_query}[A-Z]\w+, [A-Z]\w+|{in_query}[A-Z]\w+ [A-Z]\w+", pdf_text)

        else:
            completed_query = re.findall(query, pdf_text)
    
        return completed_query


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

    count_progress = 0
    no_sending_error = True
    error_count = 0

    #This is for sending emails with attachments in a folder, as long as they are pdfs
    for filename in os.listdir(attachments.get()):
        full_path = attachments.get() + "/" + filename
        progress_window.update()

        error = ""

        #For deciding if a file is a pdf
        if full_path.endswith(".pdf"):
            new_mail = outlook.CreateItem(0)

            #Determining if the subject line should be today's date or not
            if template["Subject"] == "Today" or template["Subject"] == "today":
                new_mail.Subject = datetime.today().strftime("%m/%d/%y")

            else:    
                new_mail.Subject = template["Subject"]

            #Determining if the body text should take PDF input, and then using the assigned queries
            find = []
            document_text = ""

            if is_pdf_input:
                #Opening the pdf
                document = mu.open(full_path)
                
                #Spliting the inputed queries
                finder_var = re.split("~", template["Queries"])

                #Getting text from pdf
                for page in document:
                    
                    document_text = page.get_text()

                find = []
                #Finding specified strings in text from the pdf
                for query in finder_var:
                    
                    try:
                        #Searches PDF for specified queries
                        end_query = search_pdf_for(query, document_text)

                    except:
                        #If search fails for any reason, ensures email is not sent for redunancy
                        no_sending_error = False
                        error = "Either queries are invalid due to syntax or template is taking input from PDF when it should not."

                    #Refining passing queries from templates
                    if len(end_query) > 1:
                        #If there are too many matches, ensures email is not sent for redunancy
                        no_sending_error = False
                        error += f" Multitple matches for {query} found, please refine.\n"

                    elif len(end_query) == 0:
                        #If there are no matches, ensures email is not sent for redundancy
                        no_sending_error = False
                        error += f" No matches found for {query}\n"  

                    else: 
                        find.append(end_query[0])

                if len(find) > 0:
                    try:
                        #Sets designated spots as places to send information in body of email
                        new_mail.Body = template["Body"].format(*find) + sig

                    except:
                        #Catches several errors of  trying to put queries into body
                        no_sending_error = False
                        error += f"Queries and template do not match, or there is an issue with the template."

                else:
                    #Sets body as empty template with queries but ensures email will not be sent
                    new_mail.Body = template["Body"] + sig
                    no_sending_error = False
                    error = "Could not find matching data from queries."
            
            else:
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
                #Sets CC line if it is not empty or None
                if template["CC"] != "" or template["CC"] != "None":
                    new_mail.CC = template["CC"]
                
                else:
                    pass
                
                #Sets to and from in email
                new_mail.To = template["To"]
                new_mail.SentOnBehalfOfName = template["From"]

            #Adds pdf atachment in folder
            new_mail.Attachments.Add(Source = full_path)

            try:
                #Attempts to send the email and update count
                if no_sending_error:
                    new_mail.Send()

                    count_progress += 1
                
                elif not no_sending_error:
                    #If an interim step fails, message will not be sent and error will be printed
                    error_count += 1
                    print(f"Email could not be sent. Error: {error}")
                
                else:
                    pass

            except:
                #Exception if above block fails
                print("Email could not be sent, check templates/folder/files.")
                error_count += 1

        elif isfile(full_path) and full_path.endswith(".pdf") == False:
            print("File is not pdf")
        
        else:
            print("No file found")
    
    #Cleanup for counting errors and updating the number of emails sent
    variable_count_dictionary[template["Type"] + ct_date]["Count"] += count_progress
    print(f"{count_progress} emails processed this batch. {error_count} errors.")

    e_in_batch.set(str(error_count))
    
    #Removes loading window
    progress_window.destroy()

#Load dictionary in json from path
def load_dict_json(path):
    with open(path) as readfile:
        initial_import_json = json.load(readfile)
    
    #print(initial_import_json)
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
        #print(dict)
        send_to_json(del_path, dict)
        deletion_window.destroy()

            #Updating comboboxes
        update_combobox()
    
    def no_delete():
        deletion_window.destroy()

    deletion_text = "Are you sure you want to delete this entry?\n(it will be unrecoverable)"

    deletion_window = choice_window(key_delete, no_delete, deletion_text)

#Add an entry to the email dictionary
def add_email_entry(email_template_name, sub, bod, to, cc, frm, email_in_dict, queries, count_type):
    #Initialize variables
    add_to = ""
    add_cc = ""
    add_frm = ""

    #Fixing multi-line emails and entries without formatting
    if "<" in to or ">" in to:
        add_to = to

    else:
        split_to = to.split()
        email_to_list = ["<" + i + ">" for i in split_to]
        add_to = "; ".join(email_to_list)

    if "<" in cc or ">" in cc:
        add_cc = cc

    else:
        split_cc = cc.split()
        email_cc_list = ["<" + i + ">" for i in split_cc]
        add_cc = "; ".join(email_cc_list)

    if "<" in frm or ">" in frm:
        add_frm = frm

    else:
        split_frm = frm.split()
        email_frm_list = ["<" + i + ">" for i in split_frm]
        add_frm = "; ".join(email_frm_list)

    email_in_dict[email_template_name] = {
                            "Subject" : sub,
                            "Body" : bod,
                            "To" : add_to,
                            "CC" : add_cc,
                            "From" : add_frm,
                            "Queries" : queries,
                            "Type" : count_type
                            }
    
    for item in list(variable_count_dictionary):
        if item.split(' - ')[0] is not count_type.split():
            variable_count_dictionary[count_type + ct_date] = {
                "Count": 0
            }
        
        else:
            pass

    variable_email_dictionary = email_in_dict
    #print(variable_email_dictionary)
    send_to_json(json_file_names[0], variable_email_dictionary)
    print(f"Template {email_template_name} added.")
    
    #Updating comboboxes
    update_combobox()

#Method for adding signature
def add_signature_entry(sig_template_name, sig, sig_in_dict):
    sig_in_dict[sig_template_name] = sig
                            
    variable_signature_dictionary = sig_in_dict
    send_to_json(json_file_names[1], variable_signature_dictionary)

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

#Delete entry template method
def delete_template_entries():
    subject_line.delete(0, "end")
    to_line.delete(0, "end")
    cc_line.delete(0, "end")
    fr_line.delete(0, "end")
    body_lines.delete("1.0", "end")
    query_line.delete(0, "end")
    count_type_line.delete(0, "end")
    template_name.delete(0, "end")

#Method for loading email template
def load_selected_email_entry(selected_email_key):
    #Method definition of what buttons will do
    def true_button():
        #Deleting anything in entry fields
        delete_template_entries()

        #Loading entry fields
        subject_line.insert(0, variable_email_dictionary[selected_email_key]["Subject"])
        to_line.insert(0, variable_email_dictionary[selected_email_key]["To"])
        cc_line.insert(0, variable_email_dictionary[selected_email_key]["CC"])
        fr_line.insert(0, variable_email_dictionary[selected_email_key]["From"])
        body_lines.insert("1.0", variable_email_dictionary[selected_email_key]["Body"])
        query_line.insert(0, variable_email_dictionary[selected_email_key]["Queries"])
        count_type_line.insert(0, variable_email_dictionary[selected_email_key]["Type"])
        template_name.insert(0, selected_email_key)
    
        loading_email.destroy()

    def false_button():
        loading_email.destroy()

    #Text for button
    loading_text = "Do you want to load this entry?\n(this will delete what is written)"
    
    #Class for popup
    loading_email = choice_window(true_button, false_button, loading_text)

#Method for loading an email template from a file
def import_email_from_file():
    
    def fix_email_string(email_string):
        split_string = ""
        finished_string = ""

        if email_string != "None":
            split_string = email_string.split()

            for i in split_string:
                if "<" in i or ">" in i:
                    finished_string += (i + " ")
                
                else:
                    pass
            
        else:
            print("No email address found")
            pass

        return(finished_string)
    
    #Method definition of what buttons will do
    def true_button():
        #Deleting entry fields 
        delete_template_entries()

        #Loading file for import
        imported_file_name = os.path.abspath(filedialog.askopenfilename(filetypes = (("Email Files", "*.msg"), ("All Files", "*.*"))))
        #outlook = win32.dynamic.Dispatch("Outlook.Application") (this was originally going to be used, instead of extract_msg, I don't like the module)
        print(imported_file_name)
        #imported_email = win32.OpenSharedItem(imported_file_name) (for some reason this gives an attribute error, I'm keeping this here in case I find a solution)
        imported_email = ext.Message(imported_file_name)
        
        fixed_to = fix_email_string(str(imported_email.to))
        fixed_cc = fix_email_string(str(imported_email.cc))
        fixed_fr = fix_email_string(str(imported_email.sender))

        #Setting entry fields as imported email
        subject_line.insert(0, str(imported_email.subject))
        to_line.insert(0, fixed_to)
        cc_line.insert(0, fixed_cc)
        fr_line.insert(0, fixed_fr)
        body_lines.insert("1.0", str(imported_email.body))
        template_name.insert(0, str(imported_email.subject))
    
        import_email.destroy()

    def false_button():
       import_email.destroy()

    #Text for button
    import_text = "Do you want to load a file?\n(this will delete what is written)"
    
    #Class for popup
    import_email = choice_window(true_button, false_button, import_text)

#Global variables, this does assume that both json files are in the same directory as this program
json_file_names = ["editable_email_dict.json", "editable_signature.json", "email_count_dictionary.csv"]
variable_email_dictionary = {}
variable_signature_dictionary = {}
variable_count_dictionary = {}
empty_data = {}
ct_date = ' - ' + datetime.today().strftime("%m/%d/%y")
count_exported = False
error_count = 0

#Checking if json files exist in directory where program is
for fname in json_file_names:
    if os.path.exists(fname):
        print(fname + " exists")
    
    elif fname.endswith(".json"):
        with open(fname, "w") as f:
            json.dump(empty_data, f)
    
    elif fname.endswith(".csv"):
        with open(fname, "w", newline = '') as f:
            writer = csv.writer(f, delimiter = ",")
            writer.writerows([["Type", "Count\n"]])
            f.close()

#Initialize main window, this could be done in a class for a more robust approach, but that will take a large rewrite and isn't entirely necessary
main_window = tk.Tk()
main_window.config(width = 315, height = 490)
main_window.title("Mail Sending")
tab_control = ttk.Notebook(main_window)

#Loading the various dictionaries from jsons
variable_email_dictionary = load_dict_json(json_file_names[0])
variable_signature_dictionary = load_dict_json(json_file_names[1])

#Creating dictionary for count values
try:
    for loaded_template in variable_email_dictionary:
        variable_count_dictionary[variable_email_dictionary[loaded_template]["Type"] + ct_date] = {
            "Count": 0
        }

except:
    print("No templates found for count.")

#Setting default directory variable for changing later
default_directory = tk.StringVar()
default_directory.set("C:/")

#Settiing up tabs for the window
tab_1 = ttk.Frame(tab_control)
tab_2 = ttk.Frame(tab_control)
tab_3 = ttk.Frame(tab_control)

tab_control.add(tab_1, text = "Send Emails")
tab_control.add(tab_2, text = "Edit Templates")
tab_control.add(tab_3, text = "Count")
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
        manual_picker.get(),
        take_pdf_input.get()
        )
)
emailer.grid(row = 7, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))
emailer_ttp = CreateToolTip(emailer, "If manual is true, email addresses in template will be overwritten.")

#Checkbutton to take input from PDF 
take_pdf_input = tk.BooleanVar()
take_pdf_input.set(False)
pdf_input_check = ttk.Checkbutton(
    tab_1,
    command = print_debug,
    text = "Use input from PDF",
    variable = take_pdf_input,
    onvalue = True,
    offvalue = False
)
pdf_input_check.grid(row = 7, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))
pdf_input_ttp = CreateToolTip(pdf_input_check, "If on, this will attempt to take input from a pdf and put it into the template you have provided.")

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
        json_file_names[1], 
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
        sig_full.get("1.0", "end-1c"), 
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
        json_file_names[0], 
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

#Label for queries field, this is for what the program will be searching for in pdfs
query_label = tk.Label(
    tab_2,
    text = "Queries:"
)
query_label.grid(row = 9, column = 0,  padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Entry for from section
query_line = tk.Entry(
    tab_2,
    width = 30
    )
query_line.grid(row = 9, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))
query_ttp = CreateToolTip(query_line, r"Input text that the program will query for, it will currently take any number after that text, spaces are important, separate each query by a tilda (~), add each input in order to body as {}.")

#Label for queries field, this is for what the program will be searching for in pdfs
count_type_label = tk.Label(
    tab_2,
    text = "Email Type (for count):"
)
count_type_label.grid(row = 10, column = 0,  padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Entry for from section
count_type_line = tk.Entry(
    tab_2,
    width = 30
    )
count_type_line.grid(row = 10, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))
count_type_ttp = CreateToolTip(count_type_line, "Specify the type of workflow this template pertains to, all templates of the same type will be added to a total count.")

#Label for entry field of name of template
template_label = tk.Label(
    tab_2,
    text = "Name of Template:"
)
template_label.grid(row = 11, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Entry field for naming template
template_name = tk.Entry(
    tab_2,
    width = 30
)
template_name.grid(row = 11, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))
temp_name_ttp = CreateToolTip(template_name, "If you use the same name as a new template, or load a template, any changes you have made will overwrite the original.")

#Button for adding to email dictionary
add_entry_button = tk.Button(
    tab_2,
    text = "Add Email Template Entry",
    relief = tk.RAISED,
    command = lambda: add_email_entry(
        template_name.get(), 
        subject_line.get(), 
        body_lines.get("1.0", "end-1c"), 
        to_line.get(),
        cc_line.get(), 
        fr_line.get(),
        variable_email_dictionary,
        query_line.get(),
        count_type_line.get()
        )
)
add_entry_button.grid(row = 12, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Button for loading email template from choice
load_email_entry_button = tk.Button(
    tab_2,
    text = "Load Email Entry",
    relief = tk.RAISED,
    command = lambda: load_selected_email_entry(del_entry_var.get())
)
load_email_entry_button.grid(row = 12, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Button for importing email for template
import_email_button = tk.Button(
    tab_2,
    text = "Import Email From File",
    relief = tk.RAISED,
    command = lambda: import_email_from_file()
)
import_email_button.grid(row = 13, column = 0, columnspan = 2, padx = 5, pady = 5, sticky = (tk.E, tk.W))
import_email_ttp = CreateToolTip(import_email_button, "You must save the template after importing, and should delete the signature from the email if there is one, you can rename the template as desired.")

#TAB 3

#Combobox for the type of email being sent relating to how many emails have been sent
var_count_choice = tk.StringVar()
count_choice = ttk.Combobox(
    tab_3,
    values = list(variable_count_dictionary.keys()),
    textvariable = var_count_choice,
    state = "readonly"
    )
count_choice.grid(row = 0, column = 0, padx = 5, pady = 5, sticky = (tk.E, tk.W, tk.N))

#Shows how many emails of a specific type have been sent, can be edited
number_of_type = tk.StringVar()
counted_type = tk.Entry(
    tab_3,
    width = 30,
    textvariable = number_of_type
    )
counted_type.grid(row = 0, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W, tk.N))

#Button for editing count of selected type
change_count_button = tk.Button(
    tab_3,
    text = "Change Selected Count",
    relief = tk.RAISED,
    command = lambda: change_count(counted_type.get())
)
change_count_button.grid(row = 1, columnspan = 2, padx = 5, pady = 5, sticky = (tk.E, tk.W, tk.N))

#Button for exporting the count of all emails to the csv file 
export_count_button = tk.Button(
    tab_3,
    text = "Export All Counts",
    relief = tk.RAISED,
    command = lambda: export_count(variable_count_dictionary)
)
export_count_button.grid(row = 2, columnspan = 2, padx = 5, pady = 5, sticky = (tk.E, tk.W, tk.N))

#Label for queries field, this is for what the program will be searching for in pdfs
error_amount_label = tk.Label(
    tab_3,
    text = "Errors In Last Batch:"
)
error_amount_label.grid(row = 3, column = 0,  padx = 5, pady = 5, sticky = (tk.E, tk.W))

#Entry for from section
e_in_batch = tk.StringVar()
errors_prev = tk.Entry(
    tab_3,
    width = 30,
    textvariable = e_in_batch, 
    state = "readonly"
    )
errors_prev.grid(row = 3, column = 1, padx = 5, pady = 5, sticky = (tk.E, tk.W))

#For keeping grid in line
for i in range(0, 1):
    tab_1.columnconfigure(i, weight = 1)

for i in range(0, 12):  
    tab_1.rowconfigure(i, weight = 1)

for i in range(0, 1):
    tab_2.columnconfigure(i, weight = 1)

for i in range(0, 13):
    tab_2.rowconfigure(i, weight = 1)

#END TABS

main_window.resizable(0, 0)
main_window.bind("<Destroy>", on_closing)
count_choice.bind("<<ComboboxSelected>>", update_count)
main_window.mainloop()