Attribute VB_Name = "modSharedFolderItems"

'This function will return the Trust ops
'product management inbox items
Function setItems(sharedAddress as string) As items


'first we need to set the applicaiton and namespace
Dim App As Outlook.Application
Dim NS As Outlook.NameSpace
Set App = Outlook.Application
Set NS = olApp.GetNamespace("MAPI")

'then we want to generate a recipient from the namespace and resolve it
Dim rec As Recipient
Set rec = NS.CreateRecipient(sharedAddress)
rec.Resolve

'now we will get our items from the default folder of a shared account
'where our recipient is the account owner and set it to return
Dim inbox As folder
Set inbox = objNS.GetSharedDefaultFolder(rec, olFolderInbox)
Dim items As items
Set items = inbox.items
Set setItems = items
End Function


