import win32com.client as win32
import os
import pandas as pd

outlook = win32.Dispatch('outlook.application')




# reading spreasheet
email_list = pd.read_excel(r"C:\Users\ADSIMAS\Downloads\Safe_owners_2.xlsx")

#getting the names and the emails
names = email_list['Current Owner']
emails = email_list['Email']
safes = email_list['Safe Name']

#creating a dictionary to group the safes by email
safes_by_email = {}
for i in range(len(emails)):
    email = emails[i]
    safe = safes[i]
    if email in safes_by_email:
        safes_by_email[email].append(safe)
    
    else:
        safes_by_email[email] = [safe]

# iterating through the email address and safe of the dictionary
for email, safe_list in safes_by_email.items():
    # get the name of the owners
    name = names[emails == email].iloc[0]
    print(email)

    #setting the message of the body
    safe_str = "<br>".join(safe_list)        
     
    #configure the mail to send it
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = F'Safe ownership Confirmation - {name}'
    mail.HTMLBody = F"""
    Hello, {name}, <br><br>
    
    As part of the current ITPA process for Safes Ownership update,
    we have checked that currently you have the safe(s) below under your ownership that were last updated in Emurald in 2021.<br><br>

    
    <b>{safe_str}</b> <br><br>
     
    We would like to confirm that you are the current safe owner for the safe(s) mentioned above. In case you shouldn't be the safe owner for the safe(s) mentioned, kindly provide to whom this should be transfered, so we can raise a request for this update.
 
    Thank you, <br><br>

  
    
    """   
    
    mail.Send()

print("todos os emails foram enviados")