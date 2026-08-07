VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFolderWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Watches one folder's Items collection for new mail (ItemAdd), read and
' flag changes (ItemChange), and removals (ItemRemove). One instance per
' watched folder - ThisOutlookSession creates one for every mail folder
' in the mailbox. EnableRouting is True only for Inbox and Sent Items,
' the two folders that feed the conversation-based auto-file queue;
' every folder still reports to cActivityTracker for activity logging.

Private WithEvents mItems As Outlook.items
Attribute mItems.VB_VarHelpID = -1
Private mbolEnableRouting As Boolean
Private mobjTracker As cActivityTracker

Public Sub Init(ByVal SourceFolder As Outlook.folder, _
                ByVal EnableRouting As Boolean, _
                ByVal Tracker As cActivityTracker)
    Set mItems = SourceFolder.items
    mbolEnableRouting = EnableRouting
    Set mobjTracker = Tracker

    Tracker.PrimeFolder SourceFolder
End Sub

Private Sub mItems_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrorHandler

    If "MailItem" = TypeName(Item) Then
        If mbolEnableRouting Then mAutoFile.QueueItemForRouting Item
        mobjTracker.HandleItemAdd Item
    End If

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

Private Sub mItems_ItemChange(ByVal Item As Object)
    On Error GoTo ErrorHandler

    If "MailItem" = TypeName(Item) Then mobjTracker.HandleItemChange Item

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

Private Sub mItems_ItemRemove()
    On Error GoTo ErrorHandler

    mobjTracker.HandleItemRemove mItems

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

Private Sub Class_Terminate()
    Set mItems = Nothing
    Set mobjTracker = Nothing
End Sub
