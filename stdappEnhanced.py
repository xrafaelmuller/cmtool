# Start the Tkinter event loop
import tkinter as tk
import win32
import babel.numbers
from tkinter import ttk, scrolledtext
from tkcalendar import DateEntry
from tkinter import messagebox
import win32com.client  # Import win32com.client for Outloo

def on_tab_change(event):
    selected_tab = tab_control.index(tab_control.select())
    

# Clipboard enablement
def add_copied_text():
    win32.OpenClipboard()
    clipboard_data = win32.GetClipboardData()
    win32.CloseClipboard()

    entry_cm_std.delete(0, tk.END)
    entry_cm_inc.delete(0, tk.END)
    entry_cm_std.insert(tk.END, clipboard_data)
    entry_cm_inc.insert(tk.END, clipboard_data)

# Function to create labels
def create_label(parent, text):
    label = tk.Label(parent, text=text, font=("Arial", 10), anchor="w")
    label.pack(fill="x", padx=10, pady=(5, 0))
    return label

# Function to create entry fields
def create_box(parent, width):
    entry = tk.Entry(parent, font=("Arial", 10))
    entry.pack(fill="x", padx=5, pady=(5, 5))
    entry.configure(width=width)  # Set the width using configure
    return entry

# Function to create a dropdown
def create_dropdown(parent, options):
    combo = ttk.Combobox(parent, values=options, font=("Arial", 10), state="readonly")
    combo.pack(fill="x", padx=10, pady=(5, 10))
    return combo

# Function to create a date entry field
def create_date_entry(parent):
    date_entry = DateEntry(parent, width=12, foreground='white', borderwidth=2, font=("Arial", 10))
    date_entry.pack(fill="x", padx=10, pady=(5, 10))
    return date_entry

# Main Window
cmtool_window = tk.Tk()
cmtool_window.geometry("400x650")
cmtool_window.title("CM Tool")
cmtool_window.resizable(False, False)


# Main Label
cmtool_label = tk.Label(cmtool_window, text="Welcome to CM Tool!", font=("Arial", 12))
cmtool_label.pack(pady=10)

# Create a Tab Control
tab_control = ttk.Notebook(cmtool_window)
# Create tabs
tab1 = ttk.Frame(tab_control)
tab2 = ttk.Frame(tab_control)
# Add tabs to the Tab Control
tab_control.add(tab1, text="Standard Change Creation")
tab_control.add(tab2, text="Incident Caused By Change")
# Bind the tab change event to a function
tab_control.bind("<<NotebookTabChanged>>", on_tab_change)
# Pack the Tab Control
tab_control.pack(expand=1, fill="both")


# Create a frame to group related fields for Std Change Creation Tab
frame_request_info = ttk.LabelFrame(tab1, text="Request Information", padding=(5, 5))
frame_request_info.pack(padx=5, pady=5, fill="both", expand=False)

# Input Fields
create_label(frame_request_info, "Change Coordinator Email:")
entry_cm_std = create_box(frame_request_info, width=20)
create_label(frame_request_info, "Request Item Number:")
entry_ritm = create_box(frame_request_info, width=20)
frame_request_details = ttk.LabelFrame(tab1, text="Request Details", padding=(5, 5), height=100)
frame_request_details.pack(padx=5, pady=5, fill="both", expand=False)
create_label(frame_request_details, "Standard Activity:")
entry_short_description = create_box(frame_request_details, width=20)
create_label(frame_request_details, "Activity Link")
entry_activity_hyperlink = create_box(frame_request_details, width=20)

create_label(frame_request_details, "Category:")
entry_category_options = ["Application - Code", "Application - Configuration", "Application & Database - Code",
                           "Application & Database - Configuration", "Database", "Database - Code", "Database - Configuration",
                           "Facilities - Building", "Facilities - Data Center", "Middleware", "Network", "Server", "Security",
                           "Storage", "Voice / Telecom"]
dropdown = create_dropdown(frame_request_details, entry_category_options)

create_label(frame_request_details, "CAB Approval Date:")
date_entry = create_date_entry(frame_request_details)
frame_aditional_info = ttk.LabelFrame(tab1, text="Configuration Items", padding=(5, 5), height=100)
frame_aditional_info.pack(padx=5, pady=5, fill="both", expand=False)
config_item = create_box(frame_aditional_info, width=5)

################################################################
# Create a frame to group related fields for Incident Caused Tab Tab
frame_incident_caused = ttk.LabelFrame(tab2, text="Change Activity Details", padding=(10, 5), height=100)
frame_incident_caused.pack(padx=5, pady=5, fill="both", expand=False)
create_label(frame_incident_caused, "Change Coordinator Email:")
entry_cm_inc = create_box(frame_incident_caused, width=10)
create_label(frame_incident_caused, "Change Activity:")
entry_change_activity = create_box(frame_incident_caused, width=5)
create_label(frame_incident_caused, "Change Record:")
entry_change_number = create_box(frame_incident_caused, width=5)
create_label(frame_incident_caused, "Incident(s):")
entry_incident_number = create_box(frame_incident_caused, width=5)



def send_std_email():
    mail_sender = entry_cm_std.get()
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
    body_mail = std_creation_html()
    body_mail = body_mail.replace("RITMXXXXXXX", request_item_number)
    body_mail = body_mail.replace("Change_Coordinator", " ".join([name.capitalize() for name in c_coordinator]))
    body_mail = body_mail.replace("XXSTDTYPEXX",  selected_category)
    body_mail = body_mail.replace("XXXACTIVITYXXX", inputed_activity)
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
def std_creation_html():
    body_mail = '''<html>
                        <head>
                            <meta charset="UTF-8">
                            <title>Change Enablement Notification</title>
                            </head>
                            <body>
                            <table align="center" border="0" cellpadding="0" cellspacing="0" width="900">
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
                                    <p style="font-size: 16px; font-family: 'Arial', text-allign: justify; ">Your request <strong> RITMXXXXXXX </strong> for a new Standard Change, <strong> XXXACTIVITYXXX </strong> was approved by CAB as a Standard Change on <strong> XXXDATEXXX </strong></p>
                                    <p style="font-size: 16px; font-family: 'Arial' , text-allign: justify; ">Link to <a href="XXXHYPERLINKXXX" target="_blank"> ServiceNow Standard Change Activity</a></p>
                                    <p style="font-size: 16px; font-family: 'Arial', text-allign: justify; ">Please refer to <a href="https://dell.service-now.com/esc?id=kb_article&table=kb_knowledge&sys_kb_id=KB0912448" target="_blank">KB0912448: How To: Submit a Standard Change / Standard Change Job Aid</a> for information on how to use your new Standard Change. Use the below information to locate your Standard Change in the Catalog</p>
                                    <li style="font-size: 14px; font-family: 'Arial', text-allign: justify; color: #666666;"><strong>Standard Change Type:</strong> XXSTDTYPEXX</li>
                                    <li style="font-size: 16px; font-family: 'Arial', text-allign: justify;"><strong>Change Activity: </strong> XXXACTIVITYXXX</li>
                                    <p style="font-size: 16px; display: none; font-family: 'Arial' , text-allign: justify; "> It has been authorized to be used with the following Configuration Items: </p>
                                    <p style="font-size: 16px; display: none; font-family: 'Arial , text-allign: justify';"><strong> XXCONFIGITEMSXX </strong> </p>
                                    <p style="font-size: 16px; font-family: 'Arial' , text-allign: justify;">For any questions, please contact <a href="mailto:IT-Change-Managers@dell.com">IT-Change-Managers@dell.com</a></p>                                  
                                    <p style="font-size: 16px; font-family: 'Arial' , text-allign: justify;">Note: As the owner, you are accountable for the proper usage of this Standard Change activity. Please monitor this activity frequently for the following: </p>
                                    <li style="font-size: 16px; font-family: 'Arial' , text-allign: justify;"> Who is using this Standard Change activity </li>
                                    <li style="font-size: 16px; font-family: 'Arial' , text-allign: justify;"> Did they use it for its intended purpose (strictly adhered to the implementation steps associated with this Standard Change)</li>
                                </td>
                                </tr>
                            </table>
                            </body>
                            </html>
    '''    
    return body_mail



###INCIDENT CAUSED BY CHANGE EMAIL SINTAX ###
def send_incident_email():
    mail_sender = entry_cm_inc.get()
    if not mail_sender:
        messagebox.showerror("Error", "Please fill out the CM Email Field.")
        return

    ## Replacement code block
    username, domain = mail_sender.split("@")
    c_coordinator = username.split("_")
    
    ##Validation code block    
    inc_activity = entry_change_activity.get()
    if not inc_activity:
        messagebox.showerror("Error", "eld.")
        return

    inc_change_record = entry_change_number.get()
    if not inc_change_record:
        messagebox.showerror("Error", "Please fill out the Change Number Field.")
        return
    
    inc_incident_number = entry_incident_number.get()
    if not inc_incident_number:
        messagebox.showerror("Error", "Please fill out the Incident(s) Field.")
        return


    subject_mail = "Action Required: Review Standard Change "f"{inc_change_record}" " that caused an Incident"
    inc_body_mail = inc_caused_html()
    inc_body_mail = inc_body_mail.replace("Change_Coordinator", " ".join([name.capitalize() for name in c_coordinator]))
    inc_body_mail = inc_body_mail.replace("XXXCHANGENUMBERXXX", inc_change_record)
    inc_body_mail = inc_body_mail.replace("XXXACTIVITYXXX", inc_activity)
    inc_body_mail = inc_body_mail.replace("XXXINCIDENTXXX", inc_incident_number)



    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")
        caixa_saida = namespace.GetDefaultFolder(5)  # Output Box

        email = outlook.CreateItem(0)  # 0 is an new email
        email.Subject = subject_mail
        email.HtmlBody = inc_body_mail
        email.To = mail_sender
        ##email.bcc = "rsyn@live.com"

        email.Send()
        messagebox.showinfo("Success", "Email sent!")
    except Exception as e:
        messagebox.showerror("Error", f": {str(e)}")


def inc_caused_html():
    inc_body_mail = '''<html>
                        <head>
                            <meta charset="UTF-8">
                            <title>Change Enablement Notification</title>
                            </head>
                            <body>
                            <table align="center" border="0" cellpadding="0" cellspacing="0" width="900">
                             <tr>
                               <td width="600" style="width:382.75pt;background:#0076CE;padding:0cm 5.4pt 0cm 5.4pt;
                                height:69.55pt">
                                <p class="MsoNormal"><a name="_MailAutoSig"><span style="font-size:15.0pt;
                                font-family:&quot;Arial&quot;,sans-serif;color:#F2F2F2;mso-no-proof:yes">Change
                                Enablement Notification</span><span style="mso-no-proof:yes"><o:p></o:p></span></a></p>
                             </td>
                                </tr>
                                <tr>
                                <td bgcolor="#ffffff" style="padding: 50px 30px 40px 30px;">
                                    <p style="font-size: 16px; font-family: 'Arial'; text-align: justify; ">Dear Change_Coordinator,</p>
                                    <p style="font-size: 16px; font-family: 'Arial'; text-align: justify;">Your Change record XXXCHANGENUMBERXXX has caused the below Incident:</p>
                                    <p style="font-size: 16px; font-family: 'Arial'; text-align: justify;">Change Short Description: XXXACTIVITYXXX</a></p>
                                    <p style="font-size: 16px; font-family: 'Arial'; text-align: justify;">Incident(s): XXXINCIDENTXXX</p>
                                    <p style="font-size: 16px; font-family: 'Arial'; text-align: justify;"><strong>What to do Next...</strong></p>
                                    <p style="font-size: 16px; font-family: 'Arial'; text-align: justify;"> Review the details of the Incident and determine why your Change was identified as the cause.</p>
                                    <p style="font-size: 16px; font-family: 'Arial'; text-align: justify;"> If you agree that your Change did cause the Incident, then you will need to update the state of your Change to ‘Closed Incomplete’. </p>
                                    <p style="font-size: 16px; font-family: 'Arial'; text-align: justify;">If you do not agree that your Change caused the Incident: </p>
                                    <li style="font-size: 16px; font-family: 'Arial'; text-align: justify;">Work with the individual who associated your Change to the Incident for a better understanding and potential removal of the association. You can check in the RFC Notes tab activities on who had associated the INC to the Change. There is a statement with the individual’s name as illustrated below: 
                  <Individual Name>
                      incident INCxxxxxxxx has been added from the 'incidents caused by change' related list </li>
                                    <li style="font-size: 16px; font-family: 'Arial'; text-align: justify;"> If there is a consensus that the Change did not cause any Incident, the individual who associated the Incident to the Change should remove the association </li>
                                    <li style="font-size: 16px; font-family: 'Arial'; text-align: justify;"> If the ‘Caused by Change’ field in the Incident is locked (greyed out), then you can reach out to the Change Manager assigned to your Change for assistance </li>
                                    <li style="font-size: 16px; font-family: 'Arial'; text-align: justify;"> If you are successful in removing the association of the Incident to your Change, the Change may be closed as ‘Closed Complete’
Was your Change a Standard Change? </li>
                                    <p style="font-size: 16px; font-family: 'Arial'; text-align: justify;"><strong> Was your Change a Standard Change? </strong></p>
                                    <p style="font-size: 16px; font-family: 'Arial'; text-align: justify;">Your Standard Change will immediately be deactivated from the Service Catalog (“Revoked”). </p>
                                    <p style="font-size: 16px; font-family: 'Arial'; text-align: justify;">If you were successful in removing the association of your Change to the Incident (as noted above), your Standard Change may be reactivated upon providing that evidence the Change Enablement team.</p>                                  
                                    <p style="font-size: 16px; font-family: 'Arial'; text-align: justify;">If you were not successful in removing the association of your Change to the Incident and you wish to have the Standard Change reactivated, you must follow the instructions provided in the following KB article:</p>
                                    <p style="font-size: 16px; font-family: 'Arial'; text-align: justify;"> KB1068741 <a href=     target="_blank">“Standard Change Activity Maintenance and New Proposals”. </a> </p>
                                    <p style="font-size: 16px; font-family: 'Arial'; text-align: justify;"> If approved by CAB, your Standard Change will be reactivated and once again be visible in the Service Catalog.</p>
                                    <p style="font-size: 16px; font-family: 'Arial'; text-align: justify;"> Until reactivated, you will now need to follow the Normal Change process for these Change activities.</p>
                                </td>
                                </tr>
                            </table>
                            </body>
                            </html>
    '''    
    return inc_body_mail






# Create the "Send" button for first window
send_button = tk.Button(tab1, text="Send", command=send_std_email, font=("Arial", 12))
send_button.pack(pady=10)

# Create the "Send" button for second window
send_button = tk.Button(tab2, text="Send", command=send_incident_email, font=("Arial", 12))
send_button.pack(pady=10)

# Start the main loop
cmtool_window.mainloop()

