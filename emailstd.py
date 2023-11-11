from stdapp import entry_change, messagebox,entry_sender, win32

def send_email():
    mail_sender = entry_sender.get()
    if not mail_sender:
        messagebox.showerror("Error", "Please fill out the CM Email Field.")
        return

    username, domain = mail_sender.split("@")
    c_coordinator = username.split("_")

    request_item_number = entry_change.get()
    subject_mail = f"{request_item_number} - Review"
    body_mail = update_body_mail_email()
    body_mail = body_mail.replace("RITMXXXXXXX", request_item_number)
    body_mail = body_mail.replace("Change Coordinator", " ".join([name.capitalize() for name in c_coordinator]))
 

    try:
        outlook = win32.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")
        caixa_saida = namespace.GetDefaultFolder(5)  #Output Box

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
    body_mail = ''


