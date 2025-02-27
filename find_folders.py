import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Udskift dette med din e-mailkonto
account_name = "your-email@domain.com"

# Hent den specifikke e-mailkonto
account_folder = namespace.Folders.Item(account_name)

# List alle mapper i kontoen
for i in range(1, account_folder.Folders.Count + 1):
    print(f"Mappe {i}: {account_folder.Folders.Item(i).Name}")