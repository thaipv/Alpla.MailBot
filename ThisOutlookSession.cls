VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private mcolAutoFileWatchers As Collection

Private Sub Application_Startup()
    InitAutoFileWatchers
End Sub

'Fires right when the user sends a message (new, reply, or forward), before it is actually transmitted.
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    On Error GoTo ErrorHandler

    If "MailItem" = TypeName(Item) Then mAutoFile.QueueSendActivity Item

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

Public Sub InitAutoFileWatchers()
    On Error GoTo ErrorHandler

    Dim objNamespace As Outlook.NameSpace
    Dim strInboxId As String
    Dim strSentId As String

    Set objNamespace = session
    Set mcolAutoFileWatchers = New Collection

    strInboxId = objNamespace.GetDefaultFolder(olFolderInbox).entryID
    strSentId = objNamespace.GetDefaultFolder(olFolderSentMail).entryID

    AddWatchersRecursive objNamespace.GetDefaultFolder(olFolderInbox).Parent, strInboxId, strSentId

ExitHere:
    Exit Sub

ErrorHandler:
    MsgBox "Could not start auto-file watchers: " & Err.Description, vbExclamation
    Resume ExitHere
End Sub

' Walks every folder in the mailbox and starts a watcher on each mail
' folder found (Calendar/Contacts/Tasks/etc are skipped automatically).
' Inbox and Sent Items additionally feed the conversation auto-file queue.
Private Sub AddWatchersRecursive(ByVal ParentFolder As Outlook.folder, _
                                 ByVal InboxId As String, _
                                 ByVal SentId As String)
    On Error GoTo ErrorHandler

    Dim objSubfolder As Outlook.folder
    Dim objWatcher As cFolderWatcher
    Dim bolEnableRouting As Boolean

    If olMailItem = ParentFolder.DefaultItemType Then
        If Not IsExcludedFolder(ParentFolder) Then
            bolEnableRouting = (ParentFolder.entryID = InboxId) Or (ParentFolder.entryID = SentId)

            Set objWatcher = New cFolderWatcher
            objWatcher.Init ParentFolder, bolEnableRouting, mAutoFile.GetTracker()
            mcolAutoFileWatchers.Add objWatcher
        End If
    End If

    For Each objSubfolder In ParentFolder.folders
        AddWatchersRecursive objSubfolder, InboxId, SentId
    Next objSubfolder

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

' Edit this list to skip noisy folders (add folder names as needed).
Private Function IsExcludedFolder(ByVal SourceFolder As Outlook.folder) As Boolean
    Select Case SourceFolder.name
        Case "Conversation History", "Sync Issues", "RSS Feeds"
            IsExcludedFolder = True
        Case Else
            IsExcludedFolder = False
    End Select
End Function


