import win32com.client
import pandas as pd

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)  # Folder 6 is the inbox

senders = set()

# Loop through all emails in the inbox
for message in inbox.Items:
    try:
        senders.add(message.SenderEmailAddress)
    except AttributeError:
        pass  # Skip items without a sender

# Convert to DataFrame and save to Excel
df = pd.DataFrame(sorted(senders), columns=["emails:"])
df.to_excel("outlook_senders.xlsx", index=False)

print("Unique senders saved to 'outlook_senders.xlsx'")