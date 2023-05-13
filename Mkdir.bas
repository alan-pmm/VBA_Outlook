Sub Mkdir_()

'Documentation officielle de Microsoft: https://docs.microsoft.com/en-us/office/vba/outlook/How-
to/Rules/create-a-rule-to-move-specific-e-mails-to-a-folder

'CALL VBA CLASSES
Dim myNamespace As Outlook.NameSpace
Dim myFolder As Outlook.Folder
Dim myNewFolder As Outlook.Folder

'SET CLASSES
Set myNamespace = Application.GetNamespace("MAPI")
Set myFolder = Application.Session.GetDefaultFolder(olFolderInbox)

'SET HEAD DIRECTORY
On Error GoTo ErrorHandler
Set myNewFolder = myFolder.Folders.Add("APPLIS")
Resume Next

'SET SUB DIRECTORY
Set myFolder = Application.Session.GetDefaultFolder(olFolderInbox).Folders.item("APPLIS")
'IF DIR EXIST GOTO MSG
On Error GoTo ErrorHandler

'MK ALL DIRS 
Set myNewFolder = myFolder.Folders.Add("TANDEM")
Set myNewFolder = myFolder.Folders.Add("DSN")
Set myNewFolder = myFolder.Folders.Add("DUE")
Set myNewFolder = myFolder.Folders.Add("GLD")
Set myNewFolder = myFolder.Folders.Add("GED")

ErrorHandler:
Resume Next
On Error GoTo 0
End Sub
