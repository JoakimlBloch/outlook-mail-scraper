import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# List alle konti
for i in range(namespace.Folders.Count):
    print(f"Konto {i+1}: {namespace.Folders.Item(i+1).Name}")