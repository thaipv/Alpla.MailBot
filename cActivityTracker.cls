VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cActivityTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Tracks mailbox items by PR_SEARCH_KEY, a MAPI identifier that (unlike
' EntryID) stays the same when an item is moved to another folder. This
' lets ItemAdd tell a genuinely new item (Received) apart from one that
' simply arrived from another folder (Moved), and it lets a deleted item
' still be logged with useful details even though Outlook's ItemRemove
' event carries no reference to what was removed - the cached details
' from the last time the item was seen are used instead.
'
' cConversationRouter calls SuppressNextArrival right before moving an
' item itself, so that move is not logged twice: once as "AutoFile" or
' "BulkRefile" by cConversationRouter, and again as a generic "Moved" by
' this tracker's own ItemAdd handling.
'
' ItemRemove fires whenever an item leaves a folder's Items collection -
' for a real delete AND for a move - and Outlook does not guarantee
' whether the matching ItemAdd on the destination folder fires before or
' after it. So HandleItemRemove never logs "Deleted" itself; it only
' records a candidate. ResolvePendingRemovals, called later from
' mAutoFile's existing batch timer, checks whether the candidate's cache
' entry still points at the folder it disappeared from. If it does, the
' item is genuinely gone and "Deleted" is logged. If the cache entry now
' points somewhere else, an ItemAdd already moved it and logged that
' move (as "Moved", or as "AutoFile"/"BulkRefile" by cThreadRouter) -
' so nothing further is logged here.

Private Const PR_SEARCH_KEY     As String = "http://schemas.microsoft.com/mapi/proptag/0x300B0102"
Private Const PRIME_YIELD_EVERY As Long = 200

Private mdicCache As Object           ' StableKey -> Dictionary {FolderPath, UnRead, FlagStatus, LightFields}
Private mdicSuppressed As Object      ' StableKey -> True, consumed by the next matching ItemAdd
Private mdicPendingRemovals As Object ' StableKey -> Dictionary {FolderPath, LightFields}, awaiting ResolvePendingRemovals

' =====================================================================
' Startup / setup
' =====================================================================

' One-time scan of a folder's current items, so ItemRemove has a
' baseline to diff against and ItemAdd can recognize items that already
' existed before this add-in started watching. Yields with DoEvents
' periodically so a large mailbox does not freeze Outlook while priming.
Public Sub PrimeFolder(ByVal SourceFolder As Outlook.folder)
    On Error GoTo ErrorHandler

    Dim objItem As Object
    Dim i As Long

    If mdicCache Is Nothing Then Set mdicCache = CreateObject("Scripting.Dictionary")

    For i = 1 To SourceFolder.items.Count
        On Error Resume Next
        Set objItem = SourceFolder.items(i)
        On Error GoTo ErrorHandler

        If Not objItem Is Nothing Then
            If "MailItem" = TypeName(objItem) Then CacheItem objItem, SourceFolder.FolderPath
        End If

        If 0 = (i Mod PRIME_YIELD_EVERY) Then DoEvents
    Next i

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

' Called by cConversationRouter immediately before it moves an item, so
' the resulting ItemAdd on the destination folder is not logged again.
Public Sub SuppressNextArrival(ByVal Item As Outlook.mailItem)
    On Error GoTo ErrorHandler

    Dim strKey As String
    strKey = GetStableKey(Item)

    If vbNullString <> strKey Then
        If mdicSuppressed Is Nothing Then Set mdicSuppressed = CreateObject("Scripting.Dictionary")
        mdicSuppressed(strKey) = True
    End If

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

' =====================================================================
' Event handling (called from cAutoFileWatcher)
' =====================================================================

'Public Sub HandleItemAdd(ByVal Item As Outlook.mailItem)
'    On Error GoTo ErrorHandler
'
'    Dim strKey As String
'    Dim strNewFolder As String
'    Dim dicPrevious As Object
'    Dim dicFields As Object
'
'    strKey = GetStableKey(Item)
'    If vbNullString = strKey Then GoTo ExitHere
'
'    strNewFolder = Item.Parent.FolderPath
'
'    If IsSuppressed(strKey) Then
'        ClearSuppression strKey
'    ElseIf mdicCache.Exists(strKey) Then
'        Set dicPrevious = mdicCache(strKey)
'
'        If dicPrevious("FolderPath") <> strNewFolder Then
'            Set dicFields = mAutoFile.GetLogger().CaptureMailFields(Item)
'            dicFields("FolderFrom") = dicPrevious("FolderPath")
'            dicFields("FolderTo") = strNewFolder
'            mAutoFile.QueueActivity "Moved", "Item moved between folders", dicFields
'        End If
'    ElseIf Not IsSentItemsFolder(Item.Parent) Then
'        If Item.Sent Then
'            Set dicFields = mAutoFile.GetLogger().CaptureMailFields(Item)
'            dicFields("FolderFrom") = vbNullString
'            dicFields("FolderTo") = strNewFolder
'            mAutoFile.QueueActivity "Received", "New item arrived", dicFields
'        End If
'    End If
'
'    CacheItem Item, strNewFolder
'
'ExitHere:
'    Exit Sub
'
'ErrorHandler:
'    Resume ExitHere
'End Sub
Public Sub HandleItemAdd(ByVal Item As Outlook.mailItem)
    On Error GoTo ErrorHandler

    Dim strKey As String
    Dim strNewFolder As String
    Dim dicPrevious As Object
    Dim dicFields As Object

    strKey = GetStableKey(Item)
    If vbNullString = strKey Then GoTo ExitHere

    strNewFolder = Item.Parent.FolderPath

    If IsSuppressed(strKey) Then
        ClearSuppression strKey
    ElseIf HasPendingRemoval(strKey) Then
        ' The item that seemed to vanish earlier has just reappeared in
        ' another folder - ItemRemove fired before ItemAdd for this
        ' Move, so the earlier "disappearance" candidate is now
        ' confirmed as a Move rather than a Deleted item.
        Set dicPrevious = mdicPendingRemovals(strKey)
        ClearPendingRemoval strKey

        Set dicFields = mAutoFile.GetLogger().CaptureMailFields(Item)
        dicFields("FolderFrom") = dicPrevious("FolderPath")
        dicFields("FolderTo") = strNewFolder
        mAutoFile.QueueActivity "Moved", "Item moved between folders", dicFields
    ElseIf mdicCache.Exists(strKey) Then
        Set dicPrevious = mdicCache(strKey)

        If dicPrevious("FolderPath") <> strNewFolder Then
            Set dicFields = mAutoFile.GetLogger().CaptureMailFields(Item)
            dicFields("FolderFrom") = dicPrevious("FolderPath")
            dicFields("FolderTo") = strNewFolder
            mAutoFile.QueueActivity "Moved", "Item moved between folders", dicFields
        End If
    ElseIf Not IsSentItemsFolder(Item.Parent) Then
        If Item.Sent Then
            Set dicFields = mAutoFile.GetLogger().CaptureMailFields(Item)
            dicFields("FolderFrom") = vbNullString
            dicFields("FolderTo") = strNewFolder
            mAutoFile.QueueActivity "Received", "New item arrived", dicFields
        End If
    End If

    CacheItem Item, strNewFolder

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

Public Sub HandleItemChange(ByVal Item As Outlook.mailItem)
    On Error GoTo ErrorHandler

    Dim strKey As String
    Dim dicPrevious As Object
    Dim dicFields As Object
    Dim bolCurrentUnread As Boolean
    Dim lngCurrentFlag As Long

    strKey = GetStableKey(Item)
    If vbNullString = strKey Then GoTo ExitHere
    If Not mdicCache.Exists(strKey) Then GoTo CacheAndExit

    Set dicPrevious = mdicCache(strKey)
    bolCurrentUnread = Item.UnRead
    lngCurrentFlag = Item.flagStatus

    If dicPrevious("UnRead") <> bolCurrentUnread Then
        Set dicFields = mAutoFile.GetLogger().CaptureMailFields(Item)
        dicFields("FolderFrom") = vbNullString
        dicFields("FolderTo") = vbNullString

        If bolCurrentUnread Then
            mAutoFile.QueueActivity "MarkedUnread", "Item marked as unread", dicFields
        Else
            mAutoFile.QueueActivity "MarkedRead", "Item marked as read", dicFields
        End If
    End If

    If dicPrevious("FlagStatus") <> lngCurrentFlag Then
        Set dicFields = mAutoFile.GetLogger().CaptureMailFields(Item)
        dicFields("FolderFrom") = vbNullString
        dicFields("FolderTo") = vbNullString
        dicFields("Remark") = "Flag status is now " & mAutoFile.GetLogger().GetFlagStatusText(lngCurrentFlag)
        mAutoFile.QueueActivity "FlagChanged", "Flag status changed", dicFields
    End If

CacheAndExit:
    CacheItem Item, Item.Parent.FolderPath

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

' Items.ItemRemove carries no reference to the removed item, so the
' folder's current contents are re-enumerated and compared against the
' cache; whatever cached key for this folder is now missing is a
' candidate for removal. The candidate is only queued here, not logged -
' see the class-level comment above for why the verdict is deferred to
' ResolvePendingRemovals.
Public Sub HandleItemRemove(ByVal FolderItems As Outlook.items)
    On Error GoTo ErrorHandler

    Dim strThisFolderPath As String
    Dim dicCurrentKeys As Object
    Dim objItem As Object
    Dim i As Long
    Dim varKey As Variant
    Dim dicEntry As Object
    Dim bolQueuedAny As Boolean

    strThisFolderPath = FolderItems.Parent.FolderPath
    Set dicCurrentKeys = CreateObject("Scripting.Dictionary")
    bolQueuedAny = False

    For i = 1 To FolderItems.Count
        On Error Resume Next
        Set objItem = FolderItems(i)
        On Error GoTo ErrorHandler

        If Not objItem Is Nothing Then
            If "MailItem" = TypeName(objItem) Then dicCurrentKeys(GetStableKey(objItem)) = True
        End If
    Next i

    For Each varKey In mdicCache.Keys
        Set dicEntry = mdicCache(varKey)

        If dicEntry("FolderPath") = strThisFolderPath Then
            If Not dicCurrentKeys.Exists(varKey) Then
                ' Do not log "Deleted" yet - Outlook can fire ItemRemove
                ' on the source folder before ItemAdd on the destination
                ' folder for an ordinary Move. Park this item as a
                ' candidate and let HandleItemAdd upgrade it to "Moved"
                ' if a matching arrival shows up; otherwise it is
                ' resolved as a genuine deletion once the batch timer
                ' processes the queue (see ResolvePendingVerdicts).
                QueuePendingRemoval CStr(varKey), dicEntry
                mdicCache.Remove varKey
                bolQueuedAny = True
            End If
        End If
    Next varKey

    ' Ask mAutoFile to run its batch timer even if nothing else is
    ' pending, so this candidate eventually gets a verdict.
    If bolQueuedAny Then mAutoFile.RequestQueueProcessing

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

' True while at least one removal candidate is still awaiting a verdict.
' Checked by mAutoFile so a batch run is not skipped when this is the
' only pending work.
Public Function HasPendingRemovals() As Boolean
    HasPendingRemovals = False
    If mdicPendingRemovals Is Nothing Then Exit Function
    HasPendingRemovals = (0 < mdicPendingRemovals.Count)
End Function

' Called by mAutoFile.ProcessPendingQueue, on the same delayed timer
' already used for routing decisions - by the time this runs, any
' ItemAdd triggered by the same move (including cThreadRouter's own
' moves) has already run and updated the cache. A candidate whose cache
' entry still points at the folder it disappeared from was genuinely
' deleted and is logged as "Deleted" here. A candidate whose FolderPath
' has since changed was actually moved; that move was already logged
' elsewhere, so it is dropped silently.
Public Sub ResolvePendingRemovals(ByVal Logger As cEventLogger)
    On Error GoTo ErrorHandler

    Dim varKey As Variant
    Dim dicCandidate As Object
    Dim dicCurrentEntry As Object
    Dim dicFields As Object

    If mdicPendingRemovals Is Nothing Then GoTo ExitHere

    For Each varKey In mdicPendingRemovals.Keys
        Set dicCandidate = mdicPendingRemovals(varKey)

        If mdicCache.Exists(varKey) Then
            Set dicCurrentEntry = mdicCache(varKey)

            If dicCurrentEntry("FolderPath") = dicCandidate("FolderPath") Then
                Set dicFields = dicCandidate("LightFields")
                dicFields("FolderFrom") = dicCandidate("FolderPath")
                dicFields("FolderTo") = vbNullString
                Logger.LogEvent "Deleted", "Item removed from folder", dicFields

                mdicCache.Remove varKey
            End If
        End If

        mdicPendingRemovals.Remove varKey
    Next varKey

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

' =====================================================================
' Deferred Moved/Deleted verdicts
' =====================================================================
' Outlook does not guarantee ItemAdd (destination folder) fires before
' ItemRemove (source folder) during an ordinary Move. HandleItemRemove
' parks a "disappeared" item here instead of logging Deleted right
' away, giving HandleItemAdd a chance to reclassify it as Moved if the
' item turns up elsewhere before the batch timer processes the queue.
' Whatever is still pending when the timer fires is a genuine deletion.

Private Sub QueuePendingRemoval(ByVal StableKey As String, ByVal CacheEntry As Object)
    On Error GoTo ErrorHandler

    Dim dicCandidate As Object
    Set dicCandidate = CreateObject("Scripting.Dictionary")
    dicCandidate("FolderPath") = CacheEntry("FolderPath")
    Set dicCandidate("LightFields") = CacheEntry("LightFields")

    If mdicPendingRemovals Is Nothing Then Set mdicPendingRemovals = CreateObject("Scripting.Dictionary")
    Set mdicPendingRemovals(StableKey) = dicCandidate

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

Private Function HasPendingRemoval(ByVal StableKey As String) As Boolean
    If mdicPendingRemovals Is Nothing Then
        HasPendingRemoval = False
    Else
        HasPendingRemoval = mdicPendingRemovals.Exists(StableKey)
    End If
End Function

Private Sub ClearPendingRemoval(ByVal StableKey As String)
    On Error Resume Next
    mdicPendingRemovals.Remove StableKey
End Sub

' True if any disappeared item is still awaiting a verdict, so
' mAutoFile knows to open a batch and call ResolvePendingVerdicts even
' when its own route/activity queues are empty.
Public Function HasPendingVerdicts() As Boolean
    If mdicPendingRemovals Is Nothing Then
        HasPendingVerdicts = False
    Else
        HasPendingVerdicts = (0 < mdicPendingRemovals.Count)
    End If
End Function

' Called by mAutoFile.ProcessPendingQueue once per batch, inside an
' already-open batch. Anything still parked here never got reclaimed by
' a matching ItemAdd, so it really was deleted.
Public Sub ResolvePendingVerdicts(ByVal Logger As cEventLogger)
    On Error GoTo ErrorHandler

    Dim varKey As Variant
    Dim dicCandidate As Object
    Dim dicFields As Object

    If mdicPendingRemovals Is Nothing Then GoTo ExitHere

    For Each varKey In mdicPendingRemovals.Keys
        Set dicCandidate = mdicPendingRemovals(varKey)
        Set dicFields = dicCandidate("LightFields")
        dicFields("FolderFrom") = dicCandidate("FolderPath")
        dicFields("FolderTo") = vbNullString
        Logger.LogEvent "Deleted", "Item removed from folder", dicFields
    Next varKey

    Set mdicPendingRemovals = Nothing

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub


' =====================================================================
' Cache management
' =====================================================================

Private Sub CacheItem(ByVal Item As Outlook.mailItem, ByVal FolderPath As String)
    On Error GoTo ErrorHandler

    Dim strKey As String
    Dim dicEntry As Object

    strKey = GetStableKey(Item)
    If vbNullString = strKey Then GoTo ExitHere

    Set dicEntry = CreateObject("Scripting.Dictionary")
    dicEntry("FolderPath") = FolderPath
    dicEntry("UnRead") = Item.UnRead
    dicEntry("FlagStatus") = Item.flagStatus
    Set dicEntry("LightFields") = mAutoFile.GetLogger().CaptureLightMailFields(Item)

    Set mdicCache(strKey) = dicEntry

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

Private Function IsSuppressed(ByVal StableKey As String) As Boolean
    If mdicSuppressed Is Nothing Then
        IsSuppressed = False
    Else
        IsSuppressed = mdicSuppressed.Exists(StableKey)
    End If
End Function

Private Sub ClearSuppression(ByVal StableKey As String)
    On Error Resume Next
    mdicSuppressed.Remove StableKey
End Sub

Private Function IsSentItemsFolder(ByVal SourceFolder As Outlook.folder) As Boolean
    On Error GoTo ErrorHandler

    IsSentItemsFolder = (SourceFolder.entryID = Application.session.GetDefaultFolder(olFolderSentMail).entryID)

ExitHere:
    Exit Function

ErrorHandler:
    IsSentItemsFolder = False
    Resume ExitHere
End Function

' =====================================================================
' Stable identity (PR_SEARCH_KEY) - unlike EntryID this does not change
' when an item moves to another folder within the same mailbox.
' =====================================================================

Private Function GetStableKey(ByVal Item As Outlook.mailItem) As String
    On Error GoTo ErrorHandler

    GetStableKey = BytesToHex(Item.PropertyAccessor.GetProperty(PR_SEARCH_KEY))

ExitHere:
    Exit Function

ErrorHandler:
    GetStableKey = vbNullString
    Resume ExitHere
End Function

Private Function BytesToHex(ByVal Bytes As Variant) As String
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim strResult As String

    strResult = vbNullString

    For i = LBound(Bytes) To UBound(Bytes)
        strResult = strResult & Right$("0" & Hex$(Bytes(i)), 2)
    Next i

    BytesToHex = strResult

ExitHere:
    Exit Function

ErrorHandler:
    BytesToHex = vbNullString
    Resume ExitHere
End Function

