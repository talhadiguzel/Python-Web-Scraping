import tkinter as tk
from tkinter import filedialog, messagebox
import main
import pandas as pd
import os
from dotenv import load_dotenv

load_dotenv()

def browse_file():
    filename = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if filename:
        file_path.set(filename)
        update_sheet_listbox(filename)

def update_email_listbox():
    try:
        emails = os.getenv("TO_EMAIL")
        emails = emails.split(",")
        for email in emails:
            email_listbox.insert(tk.END, email)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while updating the email list: {str(e)}")

def update_sheet_listbox(file_path):
    global sheet_names
    try:
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        category_listbox.delete(0, tk.END)
        for sheet_name in sheet_names:
            category_listbox.insert(tk.END, sheet_name)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while loading the sheets: {str(e)}")

def run_process():
    file = file_path.get()
    selected_categories = [category_listbox.get(i) for i in category_listbox.curselection()]
    selected_emails = [email_listbox.get(i) for i in email_listbox.curselection()]

    if not file:
        messagebox.showerror("Error", "Please select an Excel file.")
        return
    try:
        if not selected_categories:
            success = main.job(file, sheet_names)
        else:
            success = main.job(file, selected_categories)

        if success:
            if checkbox_var.get():
                main.email_send(selected_emails)
            messagebox.showinfo("Success", "Operation completed successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while processing the file: {str(e)}")

root = tk.Tk()
root.title("Tkinter")
root.geometry("700x500")

file_path = tk.StringVar()

excel_frame = tk.Frame(root)
excel_frame.pack(pady=5, padx=5, fill=tk.X)

file_button = tk.Button(excel_frame, text="Open Excel File", command=browse_file, width=15)
file_button.pack(pady=5, side=tk.LEFT)

file_label = tk.Label(excel_frame, textvariable=file_path)
file_label.pack(side=tk.LEFT)

bot_frame = tk.Frame(root)
bot_frame.pack(pady=5, padx=5, fill=tk.BOTH, expand=True)

category_listbox = tk.Listbox(bot_frame, selectmode=tk.MULTIPLE, exportselection=False)
category_listbox.pack(pady=5, padx=5, side=tk.LEFT, fill=tk.BOTH, expand=True)

check_frame = tk.Frame(bot_frame)
check_frame.pack(pady=5, padx=5, side=tk.RIGHT, fill=tk.BOTH, expand=True)

mail_frame = tk.Frame(check_frame)
mail_frame.pack(padx=5, pady=5, side=tk.TOP, fill=tk.BOTH, expand=True)

email_listbox = tk.Listbox(mail_frame, selectmode=tk.MULTIPLE, exportselection=False)
email_listbox.pack(pady=2, fill=tk.BOTH, expand=True)

update_email_listbox()

button_frame = tk.Frame(check_frame)
button_frame.pack(pady=5, padx=5, side=tk.BOTTOM)

checkbox_var = tk.IntVar()

checkbox = tk.Checkbutton(button_frame, text="Send Email", variable=checkbox_var)
checkbox.pack(pady=5)

run_button = tk.Button(button_frame, text="Run", command=run_process)
run_button.pack(pady=5)

exit_button = tk.Button(button_frame, text="Exit", command=root.quit)
exit_button.pack(pady=5, padx=10)

root.mainloop()
