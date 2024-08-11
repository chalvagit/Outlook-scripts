Sub MarkAllItemsAsRead()
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
    Dim objUnreadItems As Outlook.Items
    Dim objUnreadItem As Object
    Dim i As Integer
    Dim objItem As Object
    Dim objSubFolder As Outlook.Folder
 
    Set objUnreadItems = objCurFolder.Items.Restrict("[Unread]=True")
    Set objUnreadItem = objCurFolder.Items.Find("[Unread]=True")
    Total = objUnreadItems.Count
    If Not (objUnreadItem Is Nothing) Then
        For i = 1 To Total
            objUnreadItem.UnRead = False
            Set objUnreadItem = objCurFolder.Items.Find("[Unread]=True")
        Next
    End If
 
    'Process subfolders recursively
    If objCurFolder.folders.Count > 0 Then
       For Each objSubFolder In objCurFolder.folders
           Call ProcessFolders(objSubFolder)
       Next
    End If
End Sub
