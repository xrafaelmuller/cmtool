import tkinter as tk
import win32
from tkinter import ttk, scrolledtext
from tkcalendar import DateEntry
from tkinter import messagebox

# Clipboard enablement
def add_copied_text():
    win32.OpenClipboard()
    clipboard_data = win32.GetClipboardData()
    win32.CloseClipboard()

    entry_sender.delete(0, tk.END)
    entry_sender.insert(tk.END, clipboard_data)

# Function to create labels
def create_label(parent, text):
    label = tk.Label(parent, text=text, font=("Arial", 10), anchor="w")
    label.pack(fill="x", padx=10, pady=(5, 0))
    return label

# Function to create entry fields
def create_entry(parent, width):
    entry = tk.Entry(parent, font=("Arial", 10), width=width)
    entry.pack(fill="x", padx=10, pady=(0, 10))
    return entry

# Function to create a dropdown
def create_category(parent, options):
    combo = ttk.Combobox(parent, values=options, font=("Arial", 10))
    combo.pack(fill="x", padx=10, pady=(0, 10))
    return combo

# Function to create a date entry field
def create_date_entry(parent):
    date_entry = DateEntry(parent, width=12, foreground='white', borderwidth=2, font=("Arial", 10))
    date_entry.pack(fill="x", padx=10, pady=(0, 10))
    return date_entry

# Function to create a short entry field
def create_short(parent):
    entry = tk.Entry(parent, font=("Arial", 10), width=25)
    entry.pack(fill="x", padx=10, pady=(0, 10))
    return entry

# Function to handle the "Enviar" button click event
def send_button_click():
    # Get the content of all entry fields
    sender_email = entry_sender.get().strip()
    request_item_number = entry_change.get().strip()
    short_description = entry_short_description.get().strip()
    category = dropdown.get().strip()
    cab_approval_date = date_entry.get().strip()
    configuration_items = text_box.get("1.0", tk.END).strip()

    # Check if any of the fields is empty
    if not all([sender_email, request_item_number, short_description, category, cab_approval_date, configuration_items]):
        # Show an error message if any field is empty
        messagebox.showerror("Erro", "Todos os campos são obrigatórios. Preencha todos os campos antes de enviar.")
    else:
        # Proceed with the sending logic
        print("Change Coordinator Email:", sender_email)
        print("Request Item Number:", request_item_number)
        print("Standard Activity:", short_description)
        print("Category:", category)
        print("CAB Approval Date:", cab_approval_date)
        print("Configuration Items:", configuration_items)

        # Show a message box indicating that the content has been sent
        messagebox.showinfo("Success", "Texto enviado com sucesso!")

# Main Window
window = tk.Tk()
window.geometry("800x700")
window.title("Standard Change - Lazy Manager")

# Main Label
label = tk.Label(window, text="Welcome to Lazy Manager for Standard Changes!", font=("Arial", 12))
label.pack(pady=10)

# Create a frame to group related fields
frame_request_info = ttk.LabelFrame(window, text="Request Information", padding=(5, 5))
frame_request_info.pack(padx=5, pady=5, fill="both", expand=False)

# Input Fields
create_label(frame_request_info, "Change Coordinator Email:")
entry_sender = create_entry(frame_request_info, width=30)

create_label(frame_request_info, "Request Item Number:")
entry_change = create_entry(frame_request_info, width=30)

# Create a frame to group related fields
frame_request_details = ttk.LabelFrame(window, text="Request Details", padding=(10, 5), height=100)
frame_request_details.pack(padx=5, pady=5, fill="both", expand=False)

create_label(frame_request_details, "Standard Activity:")
entry_short_description = create_short(frame_request_details)

create_label(frame_request_details, "Category:")
entry_category_options = ["Application - Code", "Application - Configuration", "Application & Database - Code",
                           "Application & Database - Configuration", "Database", "Database - Code", "Database - Configuration",
                           "Facilities - Building", "Facilities - Data Center", "Middleware", "Network", "Server", "Security",
                           "Storage", "Voice / Telecom"]
dropdown = create_category(frame_request_details, entry_category_options)

create_label(frame_request_details, "CAB Approval Date:")
date_entry = create_date_entry(frame_request_details)

# Create a frame for the text box
frame_text_box = ttk.LabelFrame(window, text="Configuration Items", padding=(10, 5))
frame_text_box.pack(padx=5, pady=5, fill="both", expand=False)

# Create a scrolled text widget
text_box = scrolledtext.ScrolledText(frame_text_box, wrap=tk.WORD, width=40, height=10, font=("Arial", 10))
text_box.pack(fill="both", expand=True)

# Create the "Enviar" button
send_button = tk.Button(window, text="Enviar", command=send_button_click, font=("Arial", 12))
send_button.pack(pady=10)

# Start the main loop
window.mainloop()
