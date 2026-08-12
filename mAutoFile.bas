Option Explicit

' Entry point macros (ScanAndRefileExisting) and the deferred processing
' queue for everything reported by cAutoFileWatcher and Application_
' ItemSend. Outlook has no Application.OnTime like Excel or Word, so a
' short Win32 timer batches work and lets rules/sync settle before a
' routing decision is made - this avoids the tool and a client-side rule
' moving the same item at the same time. The same delay also gives
' cActivityTracker's Deleted-vs-Moved resolution a chance to see the
' final state of a move before deciding. Three kinds of pending work
' share one timer and one Excel batch:
'   - "Route": an item that landed in Inbox or Sent Items and needs the
'     conversation-based auto-file check (cConversationRouter.RouteItem).
'   - Removal candidate: an item that left some folder's Items
'     collection, held by cActivityTracker until ResolvePendingRemovals
'     can tell whether it was actually deleted or just moved elsewhere.
'   - Activity: a ready-to-write log row for any other tracked event
'     (Received/Sent/Moved/MarkedRead/MarkedUnread/FlagChanged/Deleted),
'     built immediately by cActivityTracker and queued for the next
'     batched Excel write.

#If VBA7 Then
    Private mlngTimerId As LongPtr
    
    Private Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal lngHwnd As LongPtr, _
        ByVal lngIdEvent As LongPtr, _
        ByVal lngElapse As Long, _
        ByVal lngTimerProc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal lngHwnd As LongPtr, _
        ByVal lngIdEvent As LongPtr) As Long
#Else
    Private mlngTimerId As Long
    
    Private Declare Function SetTimer Lib "user32" ( _
        ByVal lngHwnd As Long, _
        ByVal lngIdEvent As Long, _
        ByVal lngElapse As Long, _
        ByVal lngTimerProc As Long) As Long
    Private Declare Function KillTimer Lib "user32" ( _
        ByVal lngHwnd As Long, _
        ByVal lngIdEvent As Long) As Long
#End If

Private Const TIMER_DELAY_MS As Long = 4000

Private mdicPendingRoutes As Object         ' EntryID -> StoreID
Private mcolPendingActivities As Collection ' Dictionary {EventType, EventDescription, Fields}
Private mobjEventLogger As cEventLogger
Private mobjActivityTracker As cActivityTracker

' -------------------------------------------------------------------
' Called from cAutoFileWatcher.mItems_ItemAdd (Inbox / Sent Items only).
' -------------------------------------------------------------------
Public Sub QueueItemForRouting(ByVal Item As Outlook.mailItem)
    On Error GoTo ErrorHandler

    Dim strEntryId As String
    Dim strStoreId As String

    strEntryId = Item.entryID
    strStoreId = Item.Parent.StoreID

    If vbNullString = strEntryId Then GoTo ExitHere

    If mdicPendingRoutes Is Nothing Then Set mdicPendingRoutes = CreateObject("Scripting.Dictionary")

    If Not mdicPendingRoutes.Exists(strEntryId) Then
        mdicPendingRoutes.Add strEntryId, strStoreId
    End If

    ScheduleQueueProcessing

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

' -------------------------------------------------------------------
' Called from cActivityTracker once it has already resolved an event
' into a ready-to-write set of fields, and from QueueSendActivity.
' -------------------------------------------------------------------
Public Sub QueueActivity(ByVal EventType As String, _
                         ByVal EventDescription As String, _
                         ByVal Fields As Object)
    On Error GoTo ErrorHandler

    Dim dicWork As Object
    Set dicWork = CreateObject("Scripting.Dictionary")
    dicWork("EventType") = EventType
    dicWork("EventDescription") = EventDescription
    Set dicWork("Fields") = Fields

    If mcolPendingActivities Is Nothing Then Set mcolPendingActivities = New Collection
    mcolPendingActivities.Add dicWork

    ScheduleQueueProcessing

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

' Called from ThisOutlookSession's Application_ItemSend.
Public Sub QueueSendActivity(ByVal Item As Outlook.mailItem)
    On Error GoTo ErrorHandler

    Dim dicFields As Object
    Set dicFields = GetLogger().CaptureMailFields(Item)
    dicFields("FolderFrom") = vbNullString
    dicFields("FolderTo") = vbNullString
    dicFields("SentAt") = Now

    QueueActivity "Sent", "User sent a message", dicFields

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

' Lets cActivityTracker request a batch run for its own deferred
' Deleted-vs-Moved resolution, without exposing ScheduleQueueProcessing
' or the timer implementation to it directly.
Public Sub RequestQueueProcessing()
    ScheduleQueueProcessing
End Sub

Private Sub ScheduleQueueProcessing()
    On Error GoTo ErrorHandler

    If 0 <> mlngTimerId Then GoTo ExitHere

    mlngTimerId = SetTimer(0, 0, TIMER_DELAY_MS, AddressOf TimerProc)

    If 0 = mlngTimerId Then ProcessPendingQueue

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

' Win32 callback for SetTimer - must stay in a standard module (AddressOf
' cannot target a class method).
#If VBA7 Then
Public Sub TimerProc(ByVal lngHwnd As LongPtr, _
                     ByVal lngMsg As Long, _
                     ByVal lngIdEvent As LongPtr, _
                     ByVal lngSysTime As Long)
    On Error GoTo ErrorHandler

    KillTimer 0, lngIdEvent
    mlngTimerId = 0
    ProcessPendingQueue

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub
#Else
Public Sub TimerProc(ByVal lngHwnd As Long, _
                     ByVal lngMsg As Long, _
                     ByVal lngIdEvent As Long, _
                     ByVal lngSysTime As Long)
    On Error GoTo ErrorHandler
    
    KillTimer 0, lngIdEvent
    mlngTimerId = 0
    ProcessPendingQueue

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub
#End If

Private Sub ProcessPendingQueue()
    On Error GoTo ErrorHandler

    Dim objLogger As cEventLogger
    Dim bolHasWork As Boolean

    'bolHasWork = (Not mdicPendingRoutes Is Nothing) Or (Not mcolPendingActivities Is Nothing) Or GetTracker().HasPendingRemovals()
    bolHasWork = (Not mdicPendingRoutes Is Nothing) Or (Not mcolPendingActivities Is Nothing) Or GetTracker().HasPendingVerdicts()
    If Not bolHasWork Then GoTo ExitHere

    Set objLogger = GetLogger()
    objLogger.BeginBatch

    ProcessPendingRoutes objLogger

    ' Resolve Deleted-vs-Moved verdicts now that any ItemAdd from the
    ' same move (including the routes just processed above) has already
    ' run. Genuine deletes are written straight through Logger, the same
    ' way ProcessPendingRoutes writes its own events.
    GetTracker().ResolvePendingRemovals objLogger

    ProcessPendingActivities objLogger

ExitHere:
    If Not objLogger Is Nothing Then objLogger.EndBatch
    Set objLogger = Nothing
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

'Private Sub ProcessPendingRoutes(ByVal Logger As cEventLogger)
'    On Error GoTo ErrorHandler
'
'    Dim objNamespace As Outlook.NameSpace
'    Dim objRouter As cThreadRouter
'    Dim varKey As Variant
'    Dim objItem As Object
'
'    If mdicPendingRoutes Is Nothing Then GoTo ExitHere
'
'    Set objNamespace = Application.session
'    Set objRouter = New cThreadRouter
'
'    For Each varKey In mdicPendingRoutes.Keys
'        Set objItem = Nothing
'
'        ' A queued item can have been deleted or moved again by the time
'        ' the timer fires - skip it and keep processing the rest.
'        On Error Resume Next
'        Set objItem = objNamespace.GetItemFromID(CStr(varKey), CStr(mdicPendingRoutes(varKey)))
'        On Error GoTo ErrorHandler
'
'        If Not objItem Is Nothing Then
'            If "MailItem" = TypeName(objItem) Then
'                If objRouter.RouteItem(objItem, "AutoFile", Logger, GetTracker()) Then
'                    ShowAutoFileNotification objItem
'                End If
'            End If
'        End If
'    Next varKey
'
'ExitHere:
'    Set mdicPendingRoutes = Nothing
'    Set objRouter = Nothing
'    Set objNamespace = Nothing
'    Exit Sub
'
'ErrorHandler:
'    Resume ExitHere
'End Sub

Private Sub ProcessPendingRoutes(ByVal Logger As cEventLogger)
    On Error GoTo ErrorHandler

    Dim objNamespace As Outlook.NameSpace
    Dim objRouter As cThreadRouter
    Dim varKey As Variant
    Dim objItem As Object
    Dim fldDestination As Outlook.folder
    Dim strSubject As String

    If mdicPendingRoutes Is Nothing Then GoTo ExitHere

    Set objNamespace = Application.session
    Set objRouter = New cThreadRouter

    For Each varKey In mdicPendingRoutes.Keys
        Set objItem = Nothing

        ' A queued item can have been deleted or moved again by the time
        ' the timer fires - skip it and keep processing the rest.
        On Error Resume Next
        Set objItem = objNamespace.GetItemFromID(CStr(varKey), CStr(mdicPendingRoutes(varKey)))
        On Error GoTo ErrorHandler

        If Not objItem Is Nothing Then
            If "MailItem" = TypeName(objItem) Then
                ' Read Subject before RouteItem may Move the item - once
                ' moved, the original Item reference is stale and can
                ' silently report pre-move values instead of erroring.
                strSubject = objItem.Subject
                Set fldDestination = Nothing

                If objRouter.RouteItem(objItem, "AutoFile", Logger, GetTracker(), fldDestination) Then
                    ShowAutoFileNotification strSubject, fldDestination.name
                End If
            End If
        End If
    Next varKey

ExitHere:
    Set mdicPendingRoutes = Nothing
    Set fldDestination = Nothing
    Set objRouter = Nothing
    Set objNamespace = Nothing
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub



Private Sub ProcessPendingActivities(ByVal Logger As cEventLogger)
    On Error GoTo ErrorHandler

    Dim dicWork As Object

    If mcolPendingActivities Is Nothing Then GoTo ExitHere

    For Each dicWork In mcolPendingActivities
        Logger.LogEvent dicWork("EventType"), dicWork("EventDescription"), dicWork("Fields")
    Next dicWork

ExitHere:
    Set mcolPendingActivities = Nothing
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

' Simple heads-up popup for the live auto-file path only. Bulk refiling
' (ScanAndRefileExisting) already shows one summary MsgBox at the end,
' so it is not repeated here to avoid one dialog per item.
'Private Sub ShowAutoFileNotification(ByVal Item As Outlook.mailItem)
'    On Error GoTo ErrorHandler
'
'    MsgBox Item.Subject & vbCrLf & "-> " & Item.Parent.name, vbInformation, "Email auto-filed"
'
'ExitHere:
'    Exit Sub
'
'ErrorHandler:
'    Resume ExitHere
'End Sub

Private Sub ShowAutoFileNotification(ByVal Subject As String, ByVal FolderName As String)
    On Error GoTo ErrorHandler

    MsgBox Subject & vbCrLf & "-> " & FolderName, vbInformation, "Email auto-filed"

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

' Lazily creates a single cEventLogger reused for the life of the
' Outlook session (keeps the session/event ID counters consistent and
' avoids re-launching Excel logic on every batch). Public so
' cActivityTracker can reach it too.
Public Function GetLogger() As cEventLogger
    If mobjEventLogger Is Nothing Then Set mobjEventLogger = New cEventLogger
    Set GetLogger = mobjEventLogger
End Function

' Lazily creates the single cActivityTracker shared by every watched
' folder and by cConversationRouter's move suppression.
Public Function GetTracker() As cActivityTracker
    If mobjActivityTracker Is Nothing Then Set mobjActivityTracker = New cActivityTracker
    Set GetTracker = mobjActivityTracker
End Function

' -------------------------------------------------------------------
' Manual cleanup macro for mail that was already scattered before this
' tool was installed. Run with Alt+F8 -> mAutoFile.ScanAndRefileExisting
' -------------------------------------------------------------------
Public Sub ScanAndRefileExisting()
    On Error GoTo ErrorHandler

    Dim objNamespace As Outlook.NameSpace
    Dim objRouter As cThreadRouter
    Dim objLogger As cEventLogger
    Dim lngTotal As Long
    Dim lngMoved As Long

    If vbYes <> MsgBox("This will scan Inbox and Sent Items, and move emails that " & _
        "belong to already-classified conversations into their business folder." & vbCrLf & vbCrLf & _
        "Continue?", vbQuestion + vbYesNo, "Refile conversations") Then
        GoTo ExitHere
    End If

    Set objNamespace = Application.session
    Set objRouter = New cThreadRouter
    Set objLogger = GetLogger()
    objLogger.BeginBatch

    RefileFolderItems objNamespace.GetDefaultFolder(olFolderInbox), objRouter, objLogger, lngTotal, lngMoved
    RefileFolderItems objNamespace.GetDefaultFolder(olFolderSentMail), objRouter, objLogger, lngTotal, lngMoved

    MsgBox "Scanned " & lngTotal & " emails, moved " & lngMoved & " to their conversation's folder.", _
        vbInformation, "Done"

ExitHere:
    If Not objLogger Is Nothing Then objLogger.EndBatch
    Set objRouter = Nothing
    Set objLogger = Nothing
    Set objNamespace = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error while refiling: " & Err.Description, vbExclamation, "Refile conversations"
    Resume ExitHere
End Sub

Private Sub RefileFolderItems(ByVal SourceFolder As Outlook.folder, _
                              ByVal Router As cThreadRouter, _
                              ByVal Logger As cEventLogger, _
                              ByRef TotalCount As Long, _
                              ByRef MovedCount As Long)
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim objItem As Object

    ' Iterate backwards: Move changes the index of the remaining items.
    For i = SourceFolder.items.Count To 1 Step -1
        Set objItem = Nothing

        On Error Resume Next
        Set objItem = SourceFolder.items(i)
        On Error GoTo ErrorHandler

        If Not objItem Is Nothing Then
            If "MailItem" = TypeName(objItem) Then
                TotalCount = TotalCount + 1

                If Router.RouteItem(objItem, "BulkRefile", Logger, GetTracker()) Then MovedCount = MovedCount + 1
            End If
        End If
    Next i

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub
