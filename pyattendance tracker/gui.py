import tkinter as tk
from tkinter import messagebox
from attendance import load_workbook, add_attendance, save_workbook
from email_utils import send_email


def submit_attendance():
    name = name_entry.get()
    date = date_entry.get()
    status = status_var.get()

    if not name or not date or not status:
        messagebox.showerror("Input Error", "All fields are required.")
        return

    file_path = 'attendance.xlsx'

    from_email = 'kkwc2ke@gmail.com'
    to_email = 'areynaoes@gmail.com'
    email_password = 'rcmt iljf zisr uvrn'  # Use the generated app password here
    email_subject = 'Attendance Update'
    email_body = f'Attendance has been recorded for {name} on {date}: {status}'

    try:
        workbook, sheet = load_workbook(file_path)
        add_attendance(sheet, name, date, status)
        save_workbook(workbook, file_path)
        send_email(email_subject, email_body, to_email, from_email, email_password)
        messagebox.showinfo("Success", "Attendance recorded and email sent.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


root = tk.Tk()
root.title("Attendance Tracker")

tk.Label(root, text="Name:").grid(row=0, column=0, padx=10, pady=5)
tk.Label(root, text="Date (MM-DD-YYYY):").grid(row=1, column=0, padx=10, pady=5)
tk.Label(root, text="Status:").grid(row=2, column=0, padx=10, pady=5)

name_entry = tk.Entry(root)
date_entry = tk.Entry(root)
status_var = tk.StringVar(value="Present")

name_entry.grid(row=0, column=1, padx=10, pady=5)
date_entry.grid(row=1, column=1, padx=10, pady=5)
tk.Radiobutton(root, text="Present", variable=status_var, value="Present").grid(row=2, column=1, padx=10, pady=5)
tk.Radiobutton(root, text="Absent", variable=status_var, value="Absent").grid(row=2, column=2, padx=10, pady=5)

tk.Button(root, text='Submit', command=submit_attendance).grid(row=3, column=1, pady=10)

root.mainloop()
