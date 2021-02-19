import pyperclip
import win32com.client as client

# Variables:
outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)

print("Please insert your CR:")
CR = input()

# Functions
def mail(w, x, y, z):
    message.To = w
    message.CC = x
    message.Subject = y
    message.Body = z 


# Recipients mails
recipients = {
    "elk": "aip.elk@accenture.com",
    "cp": "aip.engineering.global@accenture.com;PDC.IS.AIP.MSS.SOC@accenture.com;AIP.L2@accenture.com;AIP.L2.Region_Leads@accenture.com",
    "enablement": "AIP_Enablement@accenture.com",
    "vault": "AIP.Cloud.Key.Management.team@accenture.com",
    "qualys": "aip.security@accenture.com;AIP_NESSUS@accenture.com",
    "patching": "aip.patchmgt.global@accenture.com"
}

start_menu="Please select an email option:\n \
    a-Deployment. \n \
    b-Decommission"

Deployment_menu = "Please select an option:\n \
    a-Communication Plan For newly launched Instance/s\n \
    b-Update Patching list\n \
    c-ELK Signoff"
subjects = {
    "a": str("Hi Team,\n\nBelow instance has been provisioned for H365 Dev as part of "+CR+":"),
    "b": str("Hi Team,\n\nBelow instance has been provisioned for Sun Chemicals Prod as part of "+CR+", please add it to patching list."),
    "c": "Hi elk team,\n\nCan you please check if logs are reaching to Kibana from this instance:"
}
Decommission_menu = "Please select an option:\n \
    a-Communication Plan For newly decommisioned Instance/s\n \
    b-Remove entries from vault\n \
    c-Remove instance from Qualys\n \
    d-Remove from patching list"
dec_options = {
    "a": str("Hi Team,\n\nBelow instances are being decommissioned as part of "+CR+":"),
    "b": "Hi Team,\n\nPlease remove below entries from vault:",
    "c": str("Hi Team,\n\nCould please remove below instances from Qualys as requested on "+CR+":"),
    "d": "Please remove below instances from patching list:"
}

# START MENU
print(start_menu)
start_option = input()

# Deployment mails
if start_option == "a":
    print(Deployment_menu)
    dep_option = input()
    if dep_option == "a":
        pyperclip.copy(subjects["a"])
    elif dep_option == "b":
        pyperclip.copy(subjects["b"])
    elif dep_option == "c":
        pyperclip.copy(subjects["c"])
    else: 
        print("Invalid option.\n", Deployment_menu)
# Decommission mails
elif start_option == "b":
    print(Decommission_menu)
    dec_option = input()
    if dec_option == "a":
        pyperclip.copy(dec_options["a"])
    elif dec_option == "b":
        pyperclip.copy(dec_options["b"])
    elif dec_option == "c":
        pyperclip.copy(dec_options["c"])
    elif dec_option == "d":
        pyperclip.copy(dec_options["d"])
    else:
        print("Invalid option.\n", Decommission_menu)
else:
    print("Invalid option", start_menu)

 # 0 is the code for a mail item (see the enumerations)
message.Display()