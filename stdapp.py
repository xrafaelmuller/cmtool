import tkinter as tk
import win32com.client as win32
from tkinter import ttk, messagebox

# Clipboard enablement
def add_copied_text():
    win32.OpenClipboard()
    clipboard_data = win32.GetClipboardData()
    win32.CloseClipboard()

    entry_sender.delete(0, tk.END)
    entry_sender.insert(tk.END, clipboard_data)

# Main Window
window = tk.Tk()
window.geometry("400x500")
window.title("Standard Change - Lazy Manager")

# Main Label
label = tk.Label(window, text="Welcome to Lazy Manager for Standard Changes!")
label.pack(pady=10)

#Function section: creating input fields and labels
def create_label(parent, text):
    label = tk.Label(parent, text=text, font=("Arial", 12))
    label.pack()

def create_entry(parent):
    entry = tk.Entry(parent, font=("Arial", 10), width=25)
    entry.pack(pady=10)
    return entry

def create_short(parent):
    entry = tk.Entry(parent, font=("Arial", 10), width=40)
    entry.pack(pady=10)
    return entry

def create_category(parent, options):
    combo = ttk.Combobox(parent, values=options, font=("Arial", 10))
    combo.pack(pady=10)
    return combo

# Input Fields
create_label(window, "Change Coordinator Email:")
entry_sender = create_entry(window)

create_label(window, "Request Item Number")
entry_change = create_entry(window)

create_label(window, "Standard Activity")
entry_short_description = create_short(window)

create_label(window, "Category")
entry_category_options = ["Application - Code","Application - Configuration","Application & Database - Code","Application & Database - Configuration","Database","Database - Code","Database - Configuration","Facilities - Building","Facilities - Data Center","Middleware","Network","Server","Security","Storage","Voice / Telecom"]
dropdown = create_category(window, entry_category_options)















# Iniciar o loop principal da janela
window.mainloop()
