Sub DeleteOlderEmails()
    Dim objStores As Outlook.Stores
    Dim objStore As Outlook.Store
    Dim objOutlookFile As Outlook.Folder
    Dim objFolder As Outlook.Folder
 
    'Process all Outlook files
    Set objStores = Application.Session.Stores
  
    For Each objStore In objStores
        Set objOutlookFile = objStore.GetRootFolder
        If objOutlookFile = "<EMAIL>" Then
            Set objFolder = objOutlookFile.folders.Item("Inbox").folders.Item("<FOLDER>")
            Call ProcessFolders(objFolder)
        End If
    Next
End Sub

Sub ProcessFolders(ByVal objCurFolder As Outlook.Folder)
    Dim olderDate As Date
    Dim objOlderItems As Outlook.Items
    Dim objOlderItem As Object
    Dim i As Integer
    Dim objSubFolder As Outlook.Folder
 
    olderDate = DateAdd("m", <NUM>, Date)
    olderDate = DateSerial(Year(olderDate), Month(olderDate), 1)
    Set objOlderItems = objCurFolder.Items.Restrict("[ReceivedTime]< '" & olderDate & "'")
    Total = objOlderItems.Count
    For i = Total To 1 Step -1
        Set objOlderItem = objOlderItems.Item(i)
        objOlderItem.Delete
    Next
 
    'Process subfolders recursively
    If objCurFolder.folders.Count > 0 Then
       For Each objSubFolder In objCurFolder.folders
           Call ProcessFolders(objSubFolder)
       Next
    End If
End Sub
