import tkinter
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl


def enter_data():
    accepted = accept_var.get()

    if accepted == "Accepted":
        # User info
        firstname = first_name_entry.get()
        lastname = last_name_entry.get()

        if firstname and lastname:
            title = title_combobox.get()
            age = age_spinbox.get()
            Wissen_Office_label = Wissen_Office_label_combobox.get()

            # Course info
            registration_status = reg_status_var.get()
            Completed_label = Completed_spinbox.get()
            Pending_label = Pending_spinbox.get()

            print("First name: ", firstname, "Last name: ", lastname)
            print("Title: ", title, "Age: ", age, "Wissen office: ", )
            print("# : ", Completed_label, "# Pending Courses: ", Pending_label)
            print("Registration status", registration_status)
            print("------------------------------------------")

            filepath = "https://docs.google.com/spreadsheets/d/12yI15gG6sXWDqPLwct1RRYy09lx4kZ8kbFDi_5nHuMU/edit#gid=0"

            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["First Name", "Last Name", "Title", "Age", "Office Location",
                           "# Courses", "# Semesters", "Registration status"]
                sheet.append(heading)
                workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([firstname, lastname, title, age, Wissen_Office_label, Completed_label,
                          Pending_label, registration_status])
            workbook.save(filepath)

        else:
            tkinter.messagebox.showwarning(title="Error", message="First name and last name are required.")
    else:
        tkinter.messagebox.showwarning(title="Error", message="You have not accepted the terms")


window = tkinter()
window.title("Wissen Office Record")

frame = tkinter.Frame(window)
frame.pack()

# Saving User Info
user_info_frame = tkinter.LabelFrame(frame, text="User Information")
user_info_frame.grid(row=0, column=0, padx=20, pady=10)

first_name_label = tkinter.Label(user_info_frame, text="First Name")
first_name_label.grid(row=0, column=0)
last_name_label = tkinter.Label(user_info_frame, text="Last Name")
last_name_label.grid(row=0, column=1)

first_name_entry = tkinter.Entry(user_info_frame)
last_name_entry = tkinter.Entry(user_info_frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

title_label = tkinter.Label(user_info_frame, text="Title")
title_combobox = ttk.Combobox(user_info_frame, values=["", "Mr.", "Ms.", "Dr."])
title_label.grid(row=0, column=2)
title_combobox.grid(row=1, column=2)

age_label = tkinter.Label(user_info_frame, text="Age")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=18, to=110)
age_label.grid(row=2, column=0)
age_spinbox.grid(row=3, column=0)

Wissen_Office_label = tkinter.Label(user_info_frame, text="Office Location")
Wissen_Office_label_combobox = ttk.Combobox(user_info_frame,
                                    values=["Bangalore", "Hyderabad", "Siddipet"])
Wissen_Office_label.grid(row=2, column=1)
Wissen_Office_label_combobox.grid(row=3, column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Saving Course Info
courses_frame = tkinter.LabelFrame(frame)
courses_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

registered_label = tkinter.Label(courses_frame, text="Registration Status")

reg_status_var = tkinter.StringVar(value="Not Registered")
registered_check = tkinter.Checkbutton(courses_frame, text="Currently Registered",
                                       variable=reg_status_var, onvalue="Registered", offvalue="Not registered")

registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)

Completed_label = tkinter.Label(courses_frame, text="# Completed Courses")
Completed_spinbox = tkinter.Spinbox(courses_frame, from_=0, to="infinity")
Completed_label.grid(row=0, column=1)
Completed_spinbox.grid(row=1, column=1)

Pending_label = tkinter.Label(courses_frame, text="# Pending Courses")
Pending_spinbox = tkinter.Spinbox(courses_frame, from_=0, to="infinity")
Pending_label.grid(row=0, column=2)
Pending_spinbox.grid(row=1, column=2)

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Accept terms
terms_frame = tkinter.LabelFrame(frame, text="Terms & Conditions")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

accept_var = tkinter.StringVar(value="Not Accepted")
terms_check = tkinter.Checkbutton(terms_frame, text="I accept the terms and conditions.",
                                  variable=accept_var, onvalue="Accepted", offvalue="Not Accepted")
terms_check.grid(row=0, column=0)

# Button
button = tkinter.Button(frame, text="Enter data", command=enter_data)
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)

window.mainloop()
