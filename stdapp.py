import tkinter as tk
import win32
from tkinter import ttk, scrolledtext
from tkcalendar import DateEntry
from tkinter import messagebox
import win32com.client  # Import win32com.client for Outloo



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
    combo = ttk.Combobox(parent, values=options, font=("Arial", 10), state="readonly")
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
entry_ritm = create_entry(frame_request_info, width=30)

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

def send_email():
    mail_sender = entry_sender.get()
    if not mail_sender:
        messagebox.showerror("Error", "Please fill out the CM Email Field.")
        return

    username, domain = mail_sender.split("@")
    c_coordinator = username.split("_")

    request_item_number = entry_ritm.get()
    subject_mail = f"{request_item_number} - Review"
    body_mail = update_body_mail_email()
    body_mail = body_mail.replace("RITMXXXXXXX", request_item_number)
    body_mail = body_mail.replace("Change_Coordinator", " ".join([name.capitalize() for name in c_coordinator]))
    body_mail = body_mail.replace("XXSTDTYPEXX",  '\n'.join(entry_category_options))
                                    
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")
        caixa_saida = namespace.GetDefaultFolder(5)  # Output Box

        email = outlook.CreateItem(0)  # 0 is an email
        email.Subject = subject_mail
        email.HtmlBody = body_mail
        email.To = mail_sender

        email.Send()
        messagebox.showinfo("Success", "Email sent!")
    except Exception as e:
        messagebox.showerror("Error", f": {str(e)}")


#Update email body with selected option in the checkboxes#
def update_body_mail_email():
    body_mail = '''<html>
                        <head>
                            <meta charset="UTF-8">
                            <title>Your Email Title</title>
                            </head>
                            <body>
                            <table align="center" border="0" cellpadding="0" cellspacing="0" width="600">
                             <tr>
                               <td width="510" style="width:382.75pt;background:#0076CE;padding:0cm 5.4pt 0cm 5.4pt;
                                height:69.55pt">
                                <p class="MsoNormal"><a name="_MailAutoSig"><span style="font-size:15.0pt;
                                font-family:&quot;Arial&quot;,sans-serif;color:#F2F2F2;mso-no-proof:yes">Change
                                Enablement Notification</span><span style="mso-no-proof:yes"><o:p></o:p></span></a></p>
                             </td>
                                </tr>
                                <tr>
                                <td bgcolor="#ffffff" style="padding: 40px 30px 40px 30px;">
                                    <p style="font-size: 16px; color: #666666;">Dear Change_Coordinator,</p>
                                    <p style="font-size: 16px; color: #666666;">Your request RITMXXXXXXX for a new Standard Change, XXActivityXX was approved by CAB as a Standard Change on XXXDATEXXX</p>
                                    <p style="font-size: 16px; color: #666666;">Please refer to KB0912448: How To: Submit a Standard Change / Standard Change Job Aid for information on how to use your new Standard Change. Use the below information to locate your Standard Change in the Catalog</p>
                                    <li style="font-size: 16px; color: #666666;">Standard Change type: XXSTDTYPEXX</li>
                                    <li style="font-size: 16px; color: #666666;">Activity: XXActivityXX</li>
                                    <p id="paragrafo-invisivel3" style="display: none;">Configuration_item</p>
                                    <p id="paragrafo-invisivel4" style="display: none;">Category</p>
                                    <p style="font-size: 16px; color: #666666;">Best regards,<br>IT-Change-Managers@dell.com</p>
                                </td>
                                </tr>
                            </table>
                            </body>
                            </html>
    '''    
    return body_mail

# Create a scrolled text widget
text_box = scrolledtext.ScrolledText(frame_text_box, wrap=tk.WORD, width=40, height=10, font=("Arial", 10))
text_box.pack(fill="both", expand=True)

# Create the "Enviar" button
send_button = tk.Button(window, text="Send", command=send_email, font=("Arial", 12))
send_button.pack(pady=10)

# Start the main loop
window.mainloop()
