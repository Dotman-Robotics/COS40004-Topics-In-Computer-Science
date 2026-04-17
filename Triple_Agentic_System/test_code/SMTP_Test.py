import win32com.client as win32

olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

# construct email item object
mailItem = olApp.CreateItem(0)
mailItem.Subject = 'Hello 123'
mailItem.BodyFormat = 1
mailItem.Body = 'Hello There'
mailItem.To = ''
mailItem.Sensitivity  = 2
# optional (account you want to use to send the email)
mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('lateralus.lateralus.40004@outlook.com')))
mailItem.Display()
#mailItem.Save()
#mailItem.Send()
