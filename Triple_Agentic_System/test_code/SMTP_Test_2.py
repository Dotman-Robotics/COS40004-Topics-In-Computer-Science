import win32com.client as win32

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

# 1. Define the email address you want to send FROM
send_from_address = "lateralus.lateralus.40004@outlook.com"

outlook = win32.Dispatch('Outlook.Application').GetNamespace("MAPI")

# 2. Find the corresponding account object
target_account = None
for account in outlook.Session.Accounts:
    if account.DisplayName == send_from_address:
        target_account = account
        break
print(f"Total accounts found: {outlook.Session.Accounts.Count}")
for i in range(1, outlook.Session.Accounts.Count + 1):
    acc = outlook.Session.Accounts.Item(i)
    print(f"Index {i}: {acc.DisplayName} ({acc.SmtpAddress})")

for store in outlook.Stores:
    print(f"Store Name: {store.DisplayName}")
    
# 3. Configure the email
mail.To = 'aaron_jt05@hotmail.com'
mail.Subject = 'Testing multiple accounts'
mail.Body = 'This message was sent from a specific account.'

# 4. Set the sending account and send
if target_account:
    # Use the account object directly
    mail.SendUsingAccount = target_account
    mail.Send()
else:
    print(f"Account {send_from_address} not found in Outlook.")
