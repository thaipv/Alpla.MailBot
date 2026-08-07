VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cThreadRouter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Finds the business folder a mail item's conversation already lives in
' (based on where earlier or related messages were filed) and moves the
' item there. No separate mapping table is required - the Outlook
' Conversation index is the source of truth. If Logger is supplied, a
' successful move is recorded in the shared Events log. If Tracker is
' supplied, it is told about the move beforehand so its own generic
' "Moved" detection does not log the same move a second time.

Private Const MAX_CONVERSATION_DEPTH As Long = 200

' -------------------------------------------------------------------
' Moves Item to its conversation's home folder if one is found and is
' different from its current folder. Returns True if a move happened.
' -------------------------------------------------------------------
Public Function RouteItem(ByVal Item As Outlook.mailItem, _
                          ByVal EventSource As String, _
                          ByVal Logger As cEventLogger, _
                          ByVal Tracker As cActivityTracker) As Boolean
    On Error GoTo ErrorHandler

    Dim fldHome As Outlook.folder
    Dim fldOld As Outlook.folder
    Dim dicFields As Object

    RouteItem = False
    Set fldHome = FindHomeFolder(Item)

    If Not fldHome Is Nothing Then
        Set fldOld = Item.Parent

        If fldHome.entryID <> fldOld.entryID Then
            ' Capture loggable fields before Move - EntryID and some
            ' properties are not safe to re-read off Item afterwards.
            If Not Logger Is Nothing Then Set dicFields = Logger.CaptureMailFields(Item)

            ' Let the activity tracker know this move is already being
            ' logged, so its own ItemAdd handling does not log it again.
            If Not Tracker Is Nothing Then Tracker.SuppressNextArrival Item

            Item.Move fldHome
            RouteItem = True

            If Not Logger Is Nothing Then
                dicFields("FolderFrom") = fldOld.FolderPath
                dicFields("FolderTo") = fldHome.FolderPath
                Logger.LogEvent EventSource, "Moved to conversation's existing folder", dicFields
            End If
        End If
    End If

ExitHere:
    Set fldHome = Nothing
    Set fldOld = Nothing
    Exit Function

ErrorHandler:
    RouteItem = False
    Resume ExitHere
End Function

' -------------------------------------------------------------------
' Walks the item's whole conversation tree across all folders and
' returns the folder of the most recent sibling that is not a default
' system folder. Returns Nothing if the conversation was never filed.
' -------------------------------------------------------------------
Private Function FindHomeFolder(ByVal Item As Outlook.mailItem) As Outlook.folder
    On Error GoTo ErrorHandler

    Dim objConversation As Outlook.Conversation
    Dim itmsRoot As Outlook.SimpleItems
    Dim dicVisited As Object
    Dim fldBest As Outlook.folder
    Dim dtBest As Date
    Dim i As Long

    Set objConversation = Item.GetConversation()

    If objConversation Is Nothing Then GoTo ExitHere

    Set dicVisited = CreateObject("Scripting.Dictionary")
    Set itmsRoot = objConversation.GetRootItems()
    dtBest = 0

    For i = 1 To itmsRoot.Count
        WalkConversationItem itmsRoot.Item(i), objConversation, Item, dicVisited, 0, fldBest, dtBest
    Next i

ExitHere:
    Set FindHomeFolder = fldBest
    Exit Function

ErrorHandler:
    Set fldBest = Nothing
    Resume ExitHere
End Function

' Recursively visits every item in the conversation tree (all folders).
Private Sub WalkConversationItem(ByVal ConvItem As Object, _
                                 ByVal ParentConversation As Outlook.Conversation, _
                                 ByVal ExcludeItem As Object, _
                                 ByVal dicVisited As Object, _
                                 ByVal lngDepth As Long, _
                                 ByRef fldBest As Outlook.folder, _
                                 ByRef dtBest As Date)
    On Error GoTo ErrorHandler

    Dim strEntryId As String
    Dim fldCurrent As Outlook.folder
    Dim dtItem As Date
    Dim itmsChild As Outlook.SimpleItems
    Dim i As Long

    If MAX_CONVERSATION_DEPTH < lngDepth Then GoTo ExitHere

    strEntryId = ConvItem.entryID

    If vbNullString <> strEntryId Then
        If dicVisited.Exists(strEntryId) Then GoTo ExitHere
        dicVisited.Add strEntryId, True
    End If

    If Not IsExcludedItem(ConvItem, ExcludeItem) Then
        Set fldCurrent = ConvItem.Parent

        If Not fldCurrent Is Nothing Then
            If Not IsSystemFolder(fldCurrent) Then
                dtItem = GetBestItemDate(ConvItem)

                If dtItem >= dtBest Then
                    dtBest = dtItem
                    Set fldBest = fldCurrent
                End If
            End If
        End If
    End If

    Set itmsChild = ParentConversation.GetChildren(ConvItem)

    If Not itmsChild Is Nothing Then
        For i = 1 To itmsChild.Count
            WalkConversationItem itmsChild.Item(i), ParentConversation, ExcludeItem, dicVisited, lngDepth + 1, fldBest, dtBest
        Next i
    End If

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

Private Function IsExcludedItem(ByVal ConvItem As Object, _
                                ByVal ExcludeItem As Object) As Boolean
    If ExcludeItem Is Nothing Then
        IsExcludedItem = False
    Else
        IsExcludedItem = (ConvItem.entryID = ExcludeItem.entryID)
    End If
End Function

' Best-effort timestamp: received time for incoming mail, sent time for
' outgoing mail, creation time as a last resort.
Private Function GetBestItemDate(ByVal ConvItem As Object) As Date
    On Error GoTo ErrorHandler

    Dim dtResult As Date
    dtResult = ConvItem.ReceivedTime

    If 0 = dtResult Then dtResult = ConvItem.SentOn
    If 0 = dtResult Then dtResult = ConvItem.CreationTime

    GetBestItemDate = dtResult

ExitHere:
    Exit Function

ErrorHandler:
    GetBestItemDate = 0
    Resume ExitHere
End Function

' True for Inbox, Sent Items, Outbox, Drafts, Deleted Items and Junk -
' i.e. any folder that is not a business folder a person filed mail into.
Private Function IsSystemFolder(ByVal SourceFolder As Outlook.folder) As Boolean
    On Error GoTo ErrorHandler

    Dim objNamespace As Outlook.NameSpace
    Dim strEntryId As String

    Set objNamespace = Application.session
    strEntryId = SourceFolder.entryID

    Select Case strEntryId
        Case objNamespace.GetDefaultFolder(olFolderInbox).entryID, _
             objNamespace.GetDefaultFolder(olFolderSentMail).entryID, _
             objNamespace.GetDefaultFolder(olFolderOutbox).entryID, _
             objNamespace.GetDefaultFolder(olFolderDrafts).entryID, _
             objNamespace.GetDefaultFolder(olFolderDeletedItems).entryID, _
             objNamespace.GetDefaultFolder(olFolderJunk).entryID
            IsSystemFolder = True
        Case Else
            IsSystemFolder = False
    End Select

ExitHere:
    Set objNamespace = Nothing
    Exit Function

ErrorHandler:
    IsSystemFolder = False
    Resume ExitHere
End Function


