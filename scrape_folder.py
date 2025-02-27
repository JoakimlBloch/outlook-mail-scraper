import win32com.client
import pandas as pd
import re

# Opret forbindelse til Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Vælg en specifik konto
account_name = "your-email@domain.com"
account_folder = namespace.Folders.Item(account_name)

# Vælg den specifikke Inbox du vil bruge (indsæt præcist navn fra Trin 1)
inbox_name = "Shared Inbox"  # Udskift med den rigtige Inbox

# Hent den ønskede Inbox
inbox = account_folder.Folders.Item(inbox_name)

# Regex for at filtrere gyldige e-mails
email_pattern = re.compile(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$")

senders = set()

# Loop gennem alle e-mails i den valgte Inbox
for message in inbox:
    senders.add(message.SenderEmailAddress)

# Gem e-mails i en Excel-fil
df = pd.DataFrame(sorted(senders), columns=["emails:"])
df.to_excel("outlook_senders.xlsx", index=False)

print(f"Ekstraherede {len(senders)} unikke e-mailadresser fra '{inbox_name}'. Gemt som 'outlook_senders.xlsx'")