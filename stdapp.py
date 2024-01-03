import tkinter as tk
import win32
import babel.numbers
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
    short = tk.Entry(parent, font=("Arial", 10), width=25)
    short.pack(fill="x", padx=10, pady=(0, 10))
    return short

# Function to create a short entry field
def create_config(parent):
    config = tk.Entry(parent, font=("Arial", 10), width=25)
    config.pack(fill="x", padx=10, pady=(0, 10))
    return config

def create_hyperlink(parent):
    config = tk.Entry(parent, font=("Arial", 10), width=25)
    config.pack(fill="x", padx=10, pady=(0, 10))
    return config

# Main Window
window = tk.Tk()
window.geometry("600x600")
window.title("Standard Change - CM Tool")

# Main Label
label = tk.Label(window, text="Welcome to CM Tool for Standard Changes!", font=("Arial", 12))
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

create_label(frame_request_details, "Activity Link")
entry_activity_hyperlink = create_hyperlink(frame_request_details)

create_label(frame_request_details, "Category:")
entry_category_options = ["Application - Code", "Application - Configuration", "Application & Database - Code",
                           "Application & Database - Configuration", "Database", "Database - Code", "Database - Configuration",
                           "Facilities - Building", "Facilities - Data Center", "Middleware", "Network", "Server", "Security",
                           "Storage", "Voice / Telecom"]

dropdown = create_category(frame_request_details, entry_category_options)

create_label(frame_request_details, "CAB Approval Date:")
date_entry = create_date_entry(frame_request_details)

frame_aditional_info = ttk.LabelFrame(window, text="Configuration Items", padding=(10, 5), height=100)
frame_aditional_info.pack(padx=5, pady=5, fill="both", expand=False)
config_item = create_config(frame_aditional_info)


def send_email():
    mail_sender = entry_sender.get()
    if not mail_sender:
        messagebox.showerror("Error", "Please fill out the CM Email Field.")
        return


    ## Replacement code block
    username, domain = mail_sender.split("@")
    c_coordinator = username.split("_")
    
    ##Validation code block
    request_item_number = entry_ritm.get()
    if not request_item_number:
        messagebox.showerror("Error", "Please fill out the Request Item Field.")
        return
    
    inputed_activity = entry_short_description.get()
    if not inputed_activity:
        messagebox.showerror("Error", "Please fill out the Short Description Field.")
        return
    
    selected_category = dropdown.get()
    selected_date = date_entry.get()

    inputed_configuration_items = config_item.get()
    inputed_hyperlink = entry_activity_hyperlink.get()
    if not inputed_hyperlink:
        messagebox.showerror("Error", "Please fill out the Activity Link Field.")
        return

    subject_mail = f"{request_item_number} - "f"{inputed_activity}"
    body_mail = update_body_mail_email()
    body_mail = body_mail.replace("RITMXXXXXXX", request_item_number)
    body_mail = body_mail.replace("Change_Coordinator", " ".join([name.capitalize() for name in c_coordinator]))
    body_mail = body_mail.replace("XXSTDTYPEXX",  selected_category)
    body_mail = body_mail.replace("XXActivityXX", inputed_activity)
    body_mail = body_mail.replace("XXXDATEXXX", selected_date)
    body_mail = body_mail.replace("XXXHYPERLINKXXX", inputed_hyperlink)

    if inputed_configuration_items is None:
        body_mail = body_mail.replace('<p style="font-size: 16px; display: none; font-family: \'Arial\'; ">It has been authorized to be used with the following Configuration Items:</p>', '<p style="font-size: 16px; font-family: \'Arial\'; ">It has been authorized to be used with the following Configuration Items:</p>')
        body_mail = body_mail.replace('<p style="font-size: 16px; display: none; font-family: \'Arial\';"><strong>XXCONFIGITEMSXX</strong></p>', '<p style="font-size: 16px; font-family: \'Arial\';"><strong>XXCONFIGITEMSXX</strong></p>')
        body_mail = body_mail.replace('XXCONFIGITEMSXX', inputed_configuration_items)

        


    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")
        caixa_saida = namespace.GetDefaultFolder(5)  # Output Box

        email = outlook.CreateItem(0)  # 0 is an new email
        email.Subject = subject_mail
        email.HtmlBody = body_mail
        email.To = mail_sender
        ##email.bcc = "rsyn@live.com"

        email.Send()
        messagebox.showinfo("Success", "Email sent!")
    except Exception as e:
        messagebox.showerror("Error", f": {str(e)}")


#EMAIL
def update_body_mail_email():
    body_mail = '''<html>
                        <head>
                            <meta charset="UTF-8">
                            <title>Change Enablement Notification</title>
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
                                <td bgcolor="#ffffff" style="padding: 50px 30px 40px 30px;">
                                    <p style="font-size: 16px; font-family: 'Arial';">Dear Change_Coordinator,</p>
                                    <p style="font-size: 16px; font-family: 'Arial'; ">Your request <strong> RITMXXXXXXX </strong> for a new Standard Change, <strong> XXActivityXX </strong> was approved by CAB as a Standard Change on <strong> XXXDATEXXX </strong></p>
                                    <p style="font-size: 16px; font-family: 'Arial'; ">Link to <a href="XXXHYPERLINKXXX" target="_blank"> ServiceNow Standard Change Activity</a></p>
                                    <p style="font-size: 16px; font-family: 'Arial'; ">Please refer to <a href="https://dell.service-now.com/esc?id=kb_article&table=kb_knowledge&sys_kb_id=KB0912448" target="_blank">KB0912448: How To: Submit a Standard Change / Standard Change Job Aid</a> for information on how to use your new Standard Change. Use the below information to locate your Standard Change in the Catalog</p>
                                    <li style="font-size: 14px; font-family: 'Arial', sans-serif; color: #666666;"><strong>Standard Change Type:</strong> XXSTDTYPEXX</li>
                                    <li style="font-size: 16px; font-family: 'Arial';"><strong>Change Activity: </strong> XXActivityXX</li>
                                    <p style="font-size: 16px; display: none; font-family: 'Arial'; "> It has been authorized to be used with the following Configuration Items: </p>
                                    <p style="font-size: 16px; display: none; font-family: 'Arial';"><strong> XXCONFIGITEMSXX </strong> </p>
                                    <p style="font-size: 16px; font-family: 'Arial';">For any questions, please contact <a href="mailto:IT-Change-Managers@dell.com">IT-Change-Managers@dell.com</a></p>                                  
                                    <p style="font-size: 16px; font-family: 'Arial';">Note: As the owner, you are accountable for the proper usage of this Standard Change activity. Please monitor this activity frequently for the following: </p>
                                    <li style="font-size: 16px; font-family: 'Arial';"> Who is using this Standard Change activity </li>
                                    <li style="font-size: 16px; font-family: 'Arial';"> Did they use it for its intended purpose (strictly adhered to the implementation steps associated with this Standard Change)</li>
                                </td>
                                </tr>
                            </table>
                            </body>
                            </html>
    '''    
    return body_mail



# Create the "Enviar" button
send_button = tk.Button(window, text="Send", command=send_email, font=("Arial", 12))
send_button.pack(pady=10)

# Start the main loop
window.mainloop()
