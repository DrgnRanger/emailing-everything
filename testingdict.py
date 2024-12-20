from datetime import datetime
from tkinter import ttk
import tkinter as tk
import json

email_dict = {
    "Test Case" : {"Subject" : "A Test", 
                   "Body" : "Yay\n", 
                   "To" : "An email", 
                   "From" : "An email"
                   },
    "1099 RTS Undeliverable" : {"Subject" : "1099 RTS-UNDELIVERABLE", 
                                "Body" : "\n", 
                                "To" : "An email", 
                                "From" : "An email"
                                }
}

json_path = "editable_email_dict.json"
variable_dictionary = {}

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
def delete_entry(key):
    if key in variable_dictionary:
        del variable_dictionary[key]
    
    else:
        print("No key found")

    print(variable_dictionary)
    send_to_json(json_path, variable_dictionary)
    email_type.config(values = list(variable_dictionary.keys()))
    email_type.update()

#Add an entry to the email dictionary
def add_entry(name, sub, bod, to, frm, in_dict):
    in_dict[name] = {"Subject" : sub,
                     "Body" : bod,
                     "To" : to,
                     "From" : frm
                     }
    
    variable_dictionary = in_dict
    print(variable_dictionary)
    send_to_json(json_path, variable_dictionary)
    email_type.config(values = list(variable_dictionary.keys()))
    email_type.update()

#Main window initalization 
main_window = tk.Tk()
main_window.config(width = 315, height = 470)
main_window.title("Editing Lists")

variable_dictionary = load_dict_json(json_path)

del_entry_label = tk.Label(
    text = "Entry to Delete:"
    )
del_entry_label.grid(row = 0, column = 0, padx = 5, pady = 5)

#Combobox for selecting an entry
entry_var = tk.StringVar()
email_type = ttk.Combobox(
    main_window,
    values = list(variable_dictionary.keys()),
    textvariable = entry_var,
    state = "readonly"
    )
email_type.grid(row = 0, column = 1, padx = 5, pady = 5)

#Button for deleting a dictionary entry
del_entry_button = tk.Button(
    text = "Delete Entry",
    relief = tk.RAISED,
    command = lambda: delete_entry(entry_var.get())
)
del_entry_button.grid(row = 1, column = 0, columnspan = 2, padx = 5, pady = 5)

#Label for adding entry section
add_entry_label = tk.Label(
    text = "Add an Entry",
)
add_entry_label.grid(row = 2, column = 0, columnspan = 2, padx = 5, pady = 5)

#Label for adding entry section
add_subject_label = tk.Label(
    text = "Subject:",
)
add_subject_label.grid(row = 3, column = 0, padx = 5, pady = 5)

#Entry for subject line
subject_line = tk.Entry(
    width = 30
    )
subject_line.grid(row = 3, column = 1, padx = 5, pady = 5)

#Label for adding to section
to_label = tk.Label(
    text = "To:",
)
to_label.grid(row = 4, column = 0, padx = 5, pady = 5)

#Entry for to section
to_line = tk.Entry(
    width = 30
    )
to_line.grid(row = 4, column = 1, padx = 5, pady = 5)

#Label for adding from section
fr_label = tk.Label(
    text = "From:",
)
fr_label.grid(row = 5, column = 0, padx = 5, pady = 5)

#Entry for from section
fr_line = tk.Entry(
    width = 30
    )
fr_line.grid(row = 5, column = 1, padx = 5, pady = 5)

#Label for body
fr_label = tk.Label(
    text = "Body:",
)
fr_label.grid(row = 6, column = 0, columnspan= 2, padx = 5, pady = 5)

#Text field for adding the body
body_lines = tk.Text(
    height = 10,
    width = 35
)
body_lines.grid(row = 7, column = 0, columnspan = 2, padx = 5, pady = 5)

#Label for name of email template
template_label = tk.Label(
    text = "Name of Template:",
)
template_label.grid(row = 8, column = 0, padx = 5, pady = 5)

#Text field for adding the body
template_name = tk.Entry(
    width = 30
)
template_name.grid(row = 8, column = 1, padx = 5, pady = 5)

#Button for adding to email dictionary
del_entry_button = tk.Button(
    text = "Add Entry",
    relief = tk.RAISED,
    command = lambda: add_entry(template_name.get(), subject_line.get(), body_lines.get("1.0", "end"), "<" + to_line.get() + ">", "<" + fr_line.get() + ">", variable_dictionary)
)
del_entry_button.grid(row = 9, column = 0, columnspan = 2, padx = 5, pady = 5)

main_window.resizable(0, 0)
main_window.geometry("315x470")
main_window.mainloop()