Attribute VB_Name = "modSharedFolderItems"

'This function will return the Trust ops
'product management inbox items
Function setItems() As items


'first we need to set the applicaiton and namespace
Dim olApp As Outlook.Application
Dim objNS As Outlook.NameSpace
Set olApp = Outlook.Application
Set objNS = olApp.GetNamespace("MAPI")

'then we want to generate a recipient from the namespace and resolve it
Dim rec As Recipient
Set rec = objNS.CreateRecipient("SharedInbox@Email.Com")
rec.Resolve

'now we will get our items from the default folder of a shared account
'where our recipient is the account owner and set it to return
Dim inbox As folder
Set inbox = objNS.GetSharedDefaultFolder(rec, olFolderInbox)
Dim str As String
Dim items As items
Set items = inbox.items
Set setItems = items

End Function


