import tkinter as tk
import win32com.client as win32
from tkinter import Tk, Label, Entry, Checkbutton, BooleanVar, messagebox,PhotoImage

#Clipboard enablement#
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
label.pack(pady=10)  # Adiciona algum espaço ao redor do rótulo

#
def create_label(parent, text):
    label = tk.Label(parent, text=text, font=("Arial", 12))
    label.pack()

def create_entry(parent):
    entry = tk.Entry(parent, font=("Arial", 10), width=25)
    entry.pack(pady=10)
    return entry

#Input Fields
create_label(window, "Change Coordinator Email:", )
entry_sender = create_entry(window)

create_label(window, "Request Item Number")
entry_change = create_entry(window)

# Criar um botão na janela

# Iniciar o loop principal da janela
window.mainloop()
