VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEventLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Writes rows to the shared Events log workbook using hidden Excel COM
' automation (Visible is never set to True). Built for a workbook on a
' network share that several people may write to at the same time:
'   - BeginBatch/EndBatch open and close the workbook once per burst of
'     activity instead of once per row, which matters for bulk refiling.
'   - TryOpenWorkbook retries with a short non-blocking wait if the file
'     is locked open (read-only) by another writer.
'   - If the shared file cannot be reached at all, rows are appended to
'     a local fallback file and flushed into the workbook the next time
'     a batch opens successfully, so nothing is lost.

Private Const LOG_FILE_PATH         As String = "H:\Mailbox Tracker\OutlookEvents.xlsx"
Private Const LOG_SHEET_NAME        As String = "Events"
Private Const APP_VERSION           As String = "1.0.0"
Private Const MAX_OPEN_RETRIES      As Long = 5
Private Const RETRY_DELAY_BASE_MS   As Long = 400
Private Const BODY_PREVIEW_LENGTH   As Long = 200
Private Const XL_UP                 As Long = -4162
Private Const XL_WORKBOOK_XLSX      As Long = 51
Private Const FALLBACK_FILE_NAME    As String = "OutlookEvents_Pending.log"
Private Const UNSET_DATE_YEAR       As Integer = 4501

Private mobjExcel           As Object
Private mobjWorkbook        As Object
Private mbolBatchOpen       As Boolean
Private mstrSessionId       As String
Private mlngEventCounter    As Long

' =====================================================================
' Public interface
' =====================================================================

' Opens Excel hidden and the shared workbook once for a burst of LogEvent
' calls. Returns True if the workbook is ready to write to. Always call
' EndBatch afterwards, even if this returns False.
Public Function BeginBatch() As Boolean
    On Error GoTo ErrorHandler

    mbolBatchOpen = TryOpenWorkbook()

    If mbolBatchOpen Then FlushLocalFallback

    BeginBatch = mbolBatchOpen

ExitHere:
    Exit Function

ErrorHandler:
    mbolBatchOpen = False
    BeginBatch = False
    Resume ExitHere
End Function

' Saves and closes the workbook, then quits the hidden Excel instance.
' Safe to call even if BeginBatch returned False or was never called.
Public Sub EndBatch()
    On Error GoTo ErrorHandler

    If mbolBatchOpen Then
        mobjWorkbook.Save
        mobjWorkbook.Close False
    End If

    If Not mobjExcel Is Nothing Then mobjExcel.Quit

ExitHere:
    Set mobjWorkbook = Nothing
    Set mobjExcel = Nothing
    mbolBatchOpen = False
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

' Extracts the mail-related columns from Item. Call this BEFORE moving
' or otherwise changing the item, then add FolderFrom/FolderTo/Remark to
' the returned dictionary as needed and pass it to LogEvent.
Public Function CaptureMailFields(ByVal Item As Outlook.mailItem) As Object
    On Error GoTo ErrorHandler

    Dim dicFields As Object
    Set dicFields = CreateObject("Scripting.Dictionary")

    dicFields("SenderName") = SafeGetString(Item, "SenderName")
    dicFields("ToRecip") = SafeGetString(Item, "To")
    dicFields("CcRecip") = SafeGetString(Item, "CC")
    dicFields("BccRecip") = SafeGetString(Item, "BCC")
    dicFields("MailSubject") = SafeGetString(Item, "Subject")
    dicFields("BodyPrev") = GetBodyPreview(Item)
    dicFields("AttachFiles") = GetAttachmentList(Item)
    dicFields("StoreName") = GetStoreName(Item)
    dicFields("ReceivedAt") = SafeGetDate(Item, "ReceivedTime")
    dicFields("SentAt") = SafeGetDate(Item, "SentOn")
    dicFields("SizeBytes") = Item.Size
    dicFields("Importance") = GetImportanceText(Item.importance)
    dicFields("FlagStatus") = GetFlagStatusText(Item.flagStatus)

    Set CaptureMailFields = dicFields

ExitHere:
    Exit Function

ErrorHandler:
    Set CaptureMailFields = dicFields
    Resume ExitHere
End Function

' Lightweight version of CaptureMailFields used by cActivityTracker's
' cache (avoids reading Body/Attachments for every item when priming a
' whole folder or refreshing the cache on every ItemAdd/ItemChange).
Public Function CaptureLightMailFields(ByVal Item As Outlook.mailItem) As Object
    On Error GoTo ErrorHandler

    Dim dicFields As Object
    Set dicFields = CreateObject("Scripting.Dictionary")

    dicFields("SenderName") = SafeGetString(Item, "SenderName")
    dicFields("MailSubject") = SafeGetString(Item, "Subject")
    dicFields("ReceivedAt") = SafeGetDate(Item, "ReceivedTime")

    Set CaptureLightMailFields = dicFields

ExitHere:
    Exit Function

ErrorHandler:
    Set CaptureLightMailFields = dicFields
    Resume ExitHere
End Function

' Writes one row. Requires an open batch (see BeginBatch); if none is
' open, or the write fails, the row goes to the local fallback instead
' of being lost.
Public Sub LogEvent(ByVal EventType As String, _
                    ByVal EventDescription As String, _
                    ByVal Fields As Object)
    
    On Error GoTo ErrorHandler

    If Fields Is Nothing Then Set Fields = CreateObject("Scripting.Dictionary")

    Fields("Timestamp") = Now
    Fields("SessionID") = GetSessionId()
    Fields("AppVersion") = APP_VERSION
    Fields("UserID") = Environ$("USERNAME")
    Fields("ComputerID") = Environ$("COMPUTERNAME")
    Fields("EventID") = GetNextEventId()
    Fields("EventType") = EventType
    Fields("EventDescription") = EventDescription

    If mbolBatchOpen Then
        On Error Resume Next
        WriteRow Fields
        
        If 0 <> Err.Number Then AppendToLocalFallback Fields
        On Error GoTo ErrorHandler
    Else
        AppendToLocalFallback Fields
    End If

ExitHere:
    Exit Sub

ErrorHandler:
    AppendToLocalFallback Fields
    Resume ExitHere
End Sub

' =====================================================================
' Workbook / worksheet handling
' =====================================================================

Private Function TryOpenWorkbook() As Boolean
    On Error GoTo ErrorHandler

    Dim i As Long

    TryOpenWorkbook = False

    Set mobjExcel = CreateObject("Excel.Application")
    mobjExcel.Visible = False
    mobjExcel.DisplayAlerts = False
    mobjExcel.ScreenUpdating = False
    mobjExcel.AutomationSecurity = 3 ' msoAutomationSecurityForceDisable

    For i = 1 To MAX_OPEN_RETRIES
        Set mobjWorkbook = Nothing

        If FileExists(LOG_FILE_PATH) Then
            On Error Resume Next
            Set mobjWorkbook = mobjExcel.Workbooks.Open(LOG_FILE_PATH, False, False)
            On Error GoTo ErrorHandler
        Else
            Set mobjWorkbook = CreateNewLogWorkbook()
        End If

        If Not mobjWorkbook Is Nothing Then
            If mobjWorkbook.ReadOnly Then
                mobjWorkbook.Close False
                Set mobjWorkbook = Nothing
            Else
                TryOpenWorkbook = True
                Exit For
            End If
        End If

        WaitMs RETRY_DELAY_BASE_MS * i
    Next i

    If Not TryOpenWorkbook Then
        mobjExcel.Quit
        Set mobjExcel = Nothing
    End If

ExitHere:
    Exit Function

ErrorHandler:
    TryOpenWorkbook = False
    Resume ExitHere
End Function

Private Function CreateNewLogWorkbook() As Object
    On Error GoTo ErrorHandler

    Dim objWorkbook As Object
    Dim objSheet As Object
    Dim varKeys As Variant
    Dim i As Long

    EnsureFolderExists

    Set objWorkbook = mobjExcel.Workbooks.Add
    Set objSheet = objWorkbook.Worksheets(1)
    objSheet.name = LOG_SHEET_NAME

    varKeys = GetColumnKeys()

    For i = LBound(varKeys) To UBound(varKeys)
        objSheet.Cells(1, i + 1).Value = varKeys(i)
    Next i

    objSheet.Rows(1).Font.Bold = True

    objWorkbook.SaveAs LOG_FILE_PATH, XL_WORKBOOK_XLSX

    Set CreateNewLogWorkbook = objWorkbook

ExitHere:
    Set objSheet = Nothing
    Set objWorkbook = Nothing
    Exit Function

ErrorHandler:
    Set CreateNewLogWorkbook = Nothing
    Resume ExitHere
End Function

Private Sub WriteRow(ByVal Fields As Object)
    On Error GoTo ErrorHandler

    Dim objSheet As Object
    Dim lngRow As Long
    Dim varKeys As Variant
    Dim i As Long

    Set objSheet = mobjWorkbook.Worksheets(LOG_SHEET_NAME)
    lngRow = GetNextRow(objSheet)
    varKeys = GetColumnKeys()

    For i = LBound(varKeys) To UBound(varKeys)
        If Fields.Exists(varKeys(i)) Then
            objSheet.Cells(lngRow, i + 1).Value = Fields(varKeys(i))
        End If
    Next i

ExitHere:
    Set objSheet = Nothing
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

Private Function GetNextRow(ByVal Sheet As Object) As Long
    On Error GoTo ErrorHandler

    GetNextRow = Sheet.Cells(Sheet.Rows.Count, 1).End(XL_UP).Row + 1

ExitHere:
    Exit Function

ErrorHandler:
    GetNextRow = 2
    Resume ExitHere
End Function

Private Function GetColumnKeys() As Variant
    GetColumnKeys = Array("Timestamp", "SessionID", "AppVersion", "UserID", "ComputerID", _
        "EventID", "EventType", "EventDescription", "SenderName", "ToRecip", "CcRecip", _
        "BccRecip", "MailSubject", "BodyPrev", "AttachFiles", "FolderFrom", "FolderTo", _
        "StoreName", "ReceivedAt", "SentAt", "SizeBytes", "Importance", "FlagStatus", "Remark")
End Function

' =====================================================================
' Local fallback queue (used when the shared file cannot be reached)
' =====================================================================

Private Function GetFallbackFilePath() As String
    GetFallbackFilePath = Environ$("TEMP") & "\" & FALLBACK_FILE_NAME
End Function

Private Sub AppendToLocalFallback(ByVal Fields As Object)
    On Error GoTo ErrorHandler

    Dim intFile As Integer
    Dim varKeys As Variant
    Dim strLine As String
    Dim i As Long

    varKeys = GetColumnKeys()
    strLine = vbNullString

    For i = LBound(varKeys) To UBound(varKeys)
        If 0 < i Then strLine = strLine & vbTab

        If Fields.Exists(varKeys(i)) Then
            strLine = strLine & SanitizeForFallback(CStr(Fields(varKeys(i))))
        End If
    Next i

    intFile = FreeFile
    Open GetFallbackFilePath() For Append As #intFile
    Print #intFile, strLine
    Close #intFile

ExitHere:
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

' Writes any locally-queued rows into the workbook that is currently
' open, then clears the local file. Requires mbolBatchOpen already True.
Private Sub FlushLocalFallback()
    On Error GoTo ErrorHandler

    Dim objFso As Object
    Dim strPath As String
    Dim intFile As Integer
    Dim strLine As String
    Dim varKeys As Variant
    Dim varParts As Variant
    Dim dicRow As Object
    Dim i As Long

    strPath = GetFallbackFilePath()

    Set objFso = CreateObject("Scripting.FileSystemObject")
    If Not objFso.FileExists(strPath) Then GoTo ExitHere

    varKeys = GetColumnKeys()
    intFile = FreeFile
    Open strPath For Input As #intFile

    Do While Not EOF(intFile)
        Line Input #intFile, strLine

        If vbNullString <> Trim(strLine) Then
            varParts = Split(strLine, vbTab)
            Set dicRow = CreateObject("Scripting.Dictionary")

            For i = LBound(varKeys) To UBound(varKeys)
                If i <= UBound(varParts) Then dicRow(varKeys(i)) = varParts(i)
            Next i

            WriteRow dicRow
        End If
    Loop

    Close #intFile
    objFso.DeleteFile strPath

ExitHere:
    Set objFso = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Close #intFile
    On Error GoTo 0
    Resume ExitHere
End Sub

Private Function SanitizeForFallback(ByVal Value As String) As String
    SanitizeForFallback = Replace(Replace(Value, vbTab, " "), vbCrLf, " ")
End Function

' =====================================================================
' Field extraction helpers
' =====================================================================
Private Function SafeGetString(ByVal Item As Object, _
                               ByVal PropertyName As String) As String
    On Error GoTo ErrorHandler

    SafeGetString = CStr(CallByName(Item, PropertyName, VbGet))

ExitHere:
    Exit Function

ErrorHandler:
    SafeGetString = vbNullString
    Resume ExitHere
End Function

' Outlook returns 1/1/4501 (not an error, not 0) for Date properties such
' as SentOn or ReceivedTime that have never been set - e.g. a draft that
' has not been sent yet, or a mail item whose ReceivedTime is not
' applicable. Without the UNSET_DATE_YEAR check below, that sentinel gets
' treated as a real date and written into the log as a bogus year-4501
' value instead of being left blank.


Private Function SafeGetDate(ByVal Item As Object, _
                             ByVal PropertyName As String) As Variant
    On Error GoTo ErrorHandler

    Dim dtValue As Date
    dtValue = CDate(CallByName(Item, PropertyName, VbGet))

    If 0 <> dtValue And UNSET_DATE_YEAR <> Year(dtValue) Then
        SafeGetDate = dtValue
    Else
        SafeGetDate = vbNullString
    End If

ExitHere:
    Exit Function

ErrorHandler:
    SafeGetDate = vbNullString
    Resume ExitHere
End Function

Private Function GetBodyPreview(ByVal Item As Outlook.mailItem) As String
    On Error GoTo ErrorHandler

    Dim strBody As String
    strBody = Item.Body
    strBody = Replace(Replace(strBody, vbCrLf, " "), vbLf, " ")
    strBody = Trim(Replace(strBody, vbTab, " "))

    If BODY_PREVIEW_LENGTH < Len(strBody) Then
        GetBodyPreview = Left(strBody, BODY_PREVIEW_LENGTH) & "..."
    Else
        GetBodyPreview = strBody
    End If

ExitHere:
    Exit Function

ErrorHandler:
    GetBodyPreview = vbNullString
    Resume ExitHere
End Function

Private Function GetAttachmentList(ByVal Item As Outlook.mailItem) As String
    On Error GoTo ErrorHandler

    Dim objAttachment As Outlook.Attachment
    Dim strResult As String

    strResult = vbNullString

    For Each objAttachment In Item.Attachments
        If vbNullString <> strResult Then strResult = strResult & "; "
        strResult = strResult & objAttachment.DisplayName
    Next objAttachment

    GetAttachmentList = strResult

ExitHere:
    Exit Function

ErrorHandler:
    GetAttachmentList = vbNullString
    Resume ExitHere
End Function

Private Function GetStoreName(ByVal Item As Outlook.mailItem) As String
    On Error GoTo ErrorHandler

    GetStoreName = Item.Parent.Store.DisplayName

ExitHere:
    Exit Function

ErrorHandler:
    GetStoreName = vbNullString
    Resume ExitHere
End Function

Private Function GetImportanceText(ByVal ImportanceValue As Long) As String
    Select Case ImportanceValue
        Case olImportanceLow
            GetImportanceText = "Low"
        Case olImportanceHigh
            GetImportanceText = "High"
        Case Else
            GetImportanceText = "Normal"
    End Select
End Function

Public Function GetFlagStatusText(ByVal FlagValue As Long) As String
    Select Case FlagValue
        Case olFlagComplete
            GetFlagStatusText = "Complete"
        Case olFlagMarked
            GetFlagStatusText = "Flagged"
        Case Else
            GetFlagStatusText = "Not flagged"
    End Select
End Function

Private Function FileExists(ByVal FilePath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim objFso As Object
    Set objFso = CreateObject("Scripting.FileSystemObject")

    FileExists = objFso.FileExists(FilePath)

ExitHere:
    Set objFso = Nothing
    Exit Function

ErrorHandler:
    FileExists = False
    Resume ExitHere
End Function

Private Sub EnsureFolderExists()
    On Error GoTo ErrorHandler

    Dim objFso As Object
    Dim strFolder As String

    Set objFso = CreateObject("Scripting.FileSystemObject")
    strFolder = objFso.GetParentFolderName(LOG_FILE_PATH)

    If Not objFso.FolderExists(strFolder) Then objFso.CreateFolder strFolder

ExitHere:
    Set objFso = Nothing
    Exit Sub

ErrorHandler:
    Resume ExitHere
End Sub

' =====================================================================
' Session / event identity
' =====================================================================

Private Function GetSessionId() As String
    If vbNullString = mstrSessionId Then
        Randomize
        mstrSessionId = Environ$("COMPUTERNAME") & "-" & Format(Now, "yyyymmddhhnnss") & _
            "-" & Format(Int(Rnd() * 9000) + 1000, "0000")
    End If

    GetSessionId = mstrSessionId
End Function

Private Function GetNextEventId() As String
    mlngEventCounter = mlngEventCounter + 1
    GetNextEventId = GetSessionId() & "-" & Format(mlngEventCounter, "00000")
End Function

Private Sub WaitMs(ByVal Milliseconds As Long)
    Dim dblEnd As Double
    dblEnd = Timer + (Milliseconds / 1000)

    Do While Timer < dblEnd
        DoEvents
    Loop
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If mbolBatchOpen Then EndBatch
End Sub

