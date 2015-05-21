Option Explicit

'Both of these folders are required;
'you will need to add these folders to your inbox
'WIP means "Work in Progress"

Const FOCUS_FOLDER_NAME = "0_FOCUS"
Const WIP_FOLDER_NAME = "1_WIP"

Dim focus_folder_base As Folder
Dim WithEvents folder_items_base As Items
Dim WithEvents inbox_folder_items_base As Items
Dim WithEvents wip_folder_items_base As Items
Dim WithEvents sent_folder_items_base As Items


Private Sub Class_Initialize()
Set FocusFolderItems = GetFolder(FOCUS_FOLDER_NAME).Items
Set InboxFolderItems = Outlook.Session.GetDefaultFolder(olFolderInbox).Items
Set WIPFolderItems = GetFolder(WIP_FOLDER_NAME).Items
Set SentFolderItems = Outlook.Session.GetDefaultFolder(olFolderSentMail).Items
    UnReadWIP
End Sub

Private Sub Class_Terminate()
On Error Resume Next
Set FocusFolderItems = Nothing
Set InboxFolderItems = Nothing
Set WIPFolderItems = Nothing
Set SentFolderItems = Nothing
End Sub



'*****************PROPERTIES***********************
'FocusFolderItems (Read-Only)
'InboxFolderItems (Read-Only)
'WIPFolderItems (Read-Only)
'SendFolderItems (Read-Only)

    Public Property Get FocusFolderItems() As Items
    Set FocusFolderItems = folder_items_base
    End Property
    
    Private Property Set FocusFolderItems(i As Items)
    Set folder_items_base = i
    End Property
    Public Property Get InboxFolderItems() As Items
    Set InboxFolderItems = inbox_folder_items_base
    End Property
    
    Private Property Set InboxFolderItems(i As Items)
    Set inbox_folder_items_base = i
    End Property
    
    Public Property Get WIPFolderItems() As Items
    Set WIPFolderItems = wip_folder_items_base
    End Property
    
    Private Property Set WIPFolderItems(i As Items)
    Set wip_folder_items_base = i
    End Property
    Public Property Get SentFolderItems() As Items
    Set SentFolderItems = sent_folder_items_base
    End Property
    Private Property Set SentFolderItems(i As Items)
    Set sent_folder_items_base = i
    End Property

'Private Functions
'GetFirstItem-Gets first item from default inbox

    Private Function GetFirstItem() 'As Item, to include appointments, etc.
    Set GetFirstItem = InboxFolderItems.GetFirst
    End Function
'GetFolder-(param FolderName as string)
    Private Function GetFolder(fName As String) As Folder
    Set GetFolder = Outlook.Session.GetDefaultFolder(olFolderInbox).Parent.Folders(fName)
    End Function
'Actually responsible for moving the first item to the focus folder
    Private Sub MoveFirstToFocus()
    Dim f As Folder
    Set f = GetFolder(FOCUS_FOLDER_NAME)
    If f.Items.Count < 1 Then
        GetFirstItem.Move f
    End If
    End Sub
'UnReadWIP makes all items in WIP box unread
    Private Sub UnReadWIP()
    Dim i As Items
    Dim m
    Set i = GetFolder(WIP_FOLDER_NAME).Items
    For Each m In i
        m.UnRead = True
    Next
    End Sub

'*****Child events
    
    Private Sub folder_items_base_ItemChange(ByVal Item As Object)
    UnReadWIP
    MoveFirstToFocus
    End Sub
    
    Private Sub folder_items_base_ItemRemove()
    UnReadWIP
    MoveFirstToFocus
    End Sub
    
    Private Sub inbox_folder_items_ItemAdd(ByVal Item As Object)
    UnReadWIP
    If IsSenderVIP(Item) = True Then
        Item.Move FocusFolder
    End If
    End Sub
    
    Private Sub wip_folder_items_base_ItemAdd(ByVal Item As Object)
    UnReadWIP
    Item.UnRead = True
    End Sub
    
    
    Private Sub sent_folder_items_base_ItemAdd(ByVal Item As Object)
    UnReadWIP
    Item.Move GetFolder(FOCUS_FOLDER_NAME)
    End Sub
    
    Private Sub wip_folder_items_base_ItemChange(ByVal Item As Object)
    UnReadWIP
    Item.UnRead = True
    End Sub

