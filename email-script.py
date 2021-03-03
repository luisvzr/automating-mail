import win32com.client as client

# Variables:
outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
signature = "\n\nBest regards,\nLuis Villamizar"
CR = input("Please insert your CR:\n")
aip_client = input("Please insert your Client:\n")
aip_env = input("Which environment is this?\n")
instances_lst = []
aip_clenv = aip_client + " " + aip_env
instances_n = input("How many instances have been affected?\n")
instances_n = int(instances_n)

if instances_n != 0:
    for x in range(0,instances_n):
        print(str("Insert instance"+ str(x+1) +" name+tab+ip:")) 
        instance = input()
        instances_lst.append(str("\n" + instance))
else:
    pass

# Functions
def mail(w, x, y, z):
    message.Display()
    message.To = w
    message.CC = x
    message.Subject = y
    message.Body = z + str_instances_lst +signature

# Function to convert a list to string
def listToString(list):  
    string = ""
    for i in list:
        str(i+ "\n")
    return (string.join(list))


str_instances_lst = listToString(instances_lst)

# Recipients mails
recipients = {
    "elk": "aip.elk@accenture.com",
    "cp": "aip.engineering.global@accenture.com;PDC.IS.AIP.MSS.SOC@accenture.com;AIP.L2@accenture.com;AIP.L2.Region_Leads@accenture.com",
    "enablement": "AIP_Enablement@accenture.com;",
    "vault": "AIP.Cloud.Key.Management.team@accenture.com",
    "qualys": "aip.security@accenture.com;AIP_NESSUS@accenture.com",
    "patching": "aip.patchmgt.global@accenture.com"
}

start_menu="Please select an email option:\n \
    a-Deployment. \n \
    b-Decommission\n"

Deployment_menu = "Please select an option:\n \
    a-Communication Plan For newly launched Instance/s\n \
    b-Update Patching list\n \
    c-ELK Signoff"
dep_msgs = {
    "a": str("Hi Team,\n\nBelow instance has been provisioned for "+ aip_clenv +" as part of "+CR+":"),
    "b": str("Hi Team,\n\nBelow instance has been provisioned for "+ aip_clenv +" as part of "+CR+", please add it to patching list."),
    "c": str("Hi elk team,\n\nCan you please check if logs are reaching to Kibana from this instance:\n"+ aip_clenv),
    "d": str("Hi patching team,\n\nPlease provide sign off for below servers:\n"+ aip_clenv)
}
Decommission_menu = "Please select an option:\n \
    a-Communication Plan For newly decommisioned Instance/s\n \
    b-Remove entries from vault\n \
    c-Remove instance from Qualys\n \
    d-Remove from patching list"
dec_msgs = {
    "a": str("Hi Team,\n\nBelow instances are being decommissioned as part of "+ CR +":"),
    "b": "Hi Team,\n\nPlease remove below entries from vault:",
    "c": str("Hi Team,\n\nBelow instance has been decommissioned from "+ aip_clenv +" as part of "+CR+", please remove from Qualys."),
    "d": "Hi Team,\n\nPlease remove below instances from patching list:"
}
subjects = {
    "subj-dp-a": str("Communication Plan For newly launched Instance || "+ CR),
    "subj-dp-b": str("Patching list || "+ CR),
    "subj-dp-c": str("ELK sign off || "+ CR),
   #"subj-dp-d": str("Patching sign off || "+aip_client+" "+aip_env),
    "subj-dc-a": str("Communication Plan For Instance Decommission || "+ CR),
    "subj-dc-b&c": str("Servers Decommission || "+ CR),
    "subj-dc-d": str("Patching list || "+ CR)
}

# Communication Plan For Instance Decommission


# START MENU
start_option = input(start_menu)


# Deployment mails
if start_option == "a":
    print(Deployment_menu)
    dep_option = input()
    if dep_option == "a":
        mail(recipients["cp"], recipients["enablement"], subjects["subj-dp-a"], dep_msgs["a"])
    elif dep_option == "b":
        mail(recipients["patching"], recipients["enablement"], subjects["subj-dp-b"], dep_msgs["b"])
    elif dep_option == "c":
        mail(recipients["elk"], recipients["enablement"], subjects["subj-dp-c"], dep_msgs["c"])
    else: 
        print("Invalid option.\n", Deployment_menu)
# Decommission mails
elif start_option == "b":
    print(Decommission_menu)
    dec_option = input()
    if dec_option == "a":
        mail(recipients["cp"], recipients["enablement"], subjects["subj-dc-a"], dec_msgs["a"])
    elif dec_option == "b":
        mail(recipients["vault"], recipients["enablement"], subjects["subj-dc-b&c"], dec_msgs["b"])
    elif dec_option == "c":
        mail(recipients["qualys"], recipients["enablement"], subjects["subj-dc-b&c"], dec_msgs["c"])
    elif dec_option == "d":
        mail(recipients["patching"], recipients["enablement"], subjects["subj-dc-d"], dec_msgs["d"])
    else:
        print("Invalid option.\n", Decommission_menu)
else:
    print("Invalid option", start_menu)