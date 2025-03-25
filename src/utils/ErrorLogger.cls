VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' ErrorLogger Class - Centralized error logging system

' Constants
Private Const LOG_SHEET_NAME As String = "ErrorLog"
Private Const MAX_LOG_ENTRIES As Long = 1000
Private Const LOG_LEVEL_INFO As Integer = 1
Private Const LOG_LEVEL_WARNING As Integer = 2
Private Const LOG_LEVEL_ERROR As Integer = 3
Private Const LOG_LEVEL_CRITICAL As Integer = 4

' Column constants
Private Const COL_TIMESTAMP As Long = 1
Private Const COL_LEVEL As Long = 2
Private Const COL_SOURCE As Long = 3
Private Const COL_ERROR_NUM As Long = 4
Private Const COL_MESSAGE As Long = 5
Private Const COL_USER As Long = 6
Private Const COL_DETAILS As Long = 7

' Private variables
Private m_LogSheet As Worksheet
Private m_IsInitialized As Boolean
Private m_ErrorCount As Long
Private m_WarningCount As Long
Private m_CurrentLogLevel As Integer

' ===== Properties =====

' Current log level - determines what gets logged
Public Property Get LogLevel() As Integer
    If Not m_IsInitialized Then Initialize
    LogLevel = m_CurrentLogLevel
End Property

Public Property Let LogLevel(value As Integer)
    If value >= LOG_LEVEL_INFO And value <= LOG_LEVEL_CRITICAL Then
        m_CurrentLogLevel = value
    Else
        m_CurrentLogLevel = LOG_LEVEL_INFO ' Default to INFO if invalid
    End If
End Property

' Error count
Public Property Get ErrorCount() As Long
    If Not m_IsInitialized Then Initialize
    ErrorCount = m_ErrorCount
End Property

' Warning count
Public Property Get WarningCount() As Long
    If Not m_IsInitialized Then Initialize
    WarningCount = m_WarningCount
End Property

' Access to the log sheet
Public Property Get LogSheet() As Worksheet
    If Not m_IsInitialized Then Initialize
    Set LogSheet = m_LogSheet
End Property

' ===== Methods =====

' Initialize the logger
Public Sub Initialize()
    On Error Resume Next
    
    ' Set default log level if not set
    If m_CurrentLogLevel = 0 Then
        m_CurrentLogLevel = LOG_LEVEL_INFO
    End If
    
    ' Get or create log sheet
    Set m_LogSheet = GetOrCreateLogSheet()
    
    ' Count existing errors and warnings
    CountExistingLogs
    
    m_IsInitialized = True
    
    ' Log initialization
    LogInfo "ErrorLogger", "Error logging system initialized", "LogLevel=" & m_CurrentLogLevel
    
    On Error GoTo 0
End Sub

' Log an informational message
Public Sub LogInfo(source As String, message As String, Optional details As String = "")
    ' Only log if our level is info or lower
    If m_CurrentLogLevel <= LOG_LEVEL_INFO Then
        LogEntry LOG_LEVEL_INFO, source, 0, message, details
    End If
End Sub

' Log a warning
Public Sub LogWarning(source As String, message As String, Optional errorNumber As Long = 0, Optional details As String = "")
    ' Only log if our level is warning or lower
    If m_CurrentLogLevel <= LOG_LEVEL_WARNING Then
        LogEntry LOG_LEVEL_WARNING, source, errorNumber, message, details
        m_WarningCount = m_WarningCount + 1
    End If
End Sub

' Log an error
Public Sub LogError(source As String, errorNumber As Long, errorDescription As String, Optional details As String = "")
    ' Only log if our level is error or lower
    If m_CurrentLogLevel <= LOG_LEVEL_ERROR Then
        LogEntry LOG_LEVEL_ERROR, source, errorNumber, errorDescription, details
        m_ErrorCount = m_ErrorCount + 1
        
        ' Also output to immediate window for developers
        Debug.Print "ERROR in " & source & ": " & errorNumber & " - " & errorDescription
    End If
End Sub

' Log a critical error
Public Sub LogCritical(source As String, errorNumber As Long, errorDescription As String, Optional details As String = "")
    ' Critical errors are always logged
    LogEntry LOG_LEVEL_CRITICAL, source, errorNumber, errorDescription, details
    m_ErrorCount = m_ErrorCount + 1
    
    ' Also output to immediate window for developers
    Debug.Print "CRITICAL ERROR in " & source & ": " & errorNumber & " - " & errorDescription
End Sub

' Clear all logs
Public Sub ClearLogs()
    On Error Resume Next
    
    If Not m_IsInitialized Then Initialize
    
    ' Keep the header row
    If Not m_LogSheet Is Nothing Then
        If m_LogSheet.Cells(1, 1).value <> "" Then
            Dim lastRow As Long
            lastRow = GetLastRow(m_LogSheet, 1)
            
            If lastRow > 1 Then
                m_LogSheet.Rows("2:" & lastRow).Delete
            End If
        End If
    End If
    
    ' Reset counters
    m_ErrorCount = 0
    m_WarningCount = 0
    
    ' Log the clearing action
    LogInfo "ErrorLogger", "Logs cleared", "User initiated clear operation"
    
    On Error GoTo 0
End Sub

' Export logs to CSV file
Public Function ExportLogs(filePath As String) As Boolean
    On Error Resume Next
    
    If Not m_IsInitialized Then Initialize
    
    Dim lastRow As Long
    lastRow = GetLastRow(m_LogSheet, 1)
    
    ' Check if we have any logs
    If lastRow <= 1 Then
        ExportLogs = False
        Exit Function
    End If
    
    ' Export to CSV
    m_LogSheet.Range("A1:G" & lastRow).Copy
    
    ' Create a new workbook for export
    Dim tempWB As Workbook
    Set tempWB = Workbooks.Add
    
    tempWB.Sheets(1).Range("A1").PasteSpecial
    
    ' Save as CSV
    tempWB.SaveAs filePath, xlCSV
    
    If Err.Number <> 0 Then
        LogError "ErrorLogger.ExportLogs", Err.Number, "Failed to export logs: " & Err.Description
        ExportLogs = False
    Else
        ExportLogs = True
    End If
    
    ' Close temp workbook
    tempWB.Close False
    Set tempWB = Nothing
    
    ' Log the export action
    LogInfo "ErrorLogger", "Logs exported to " & filePath, "Rows exported: " & (lastRow - 1)
    
    On Error GoTo 0
End Function

' Get errors filtered by source
Public Function GetErrorsBySource(source As String) As Collection
    On Error Resume Next
    
    If Not m_IsInitialized Then Initialize
    
    Dim result As New Collection
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = GetLastRow(m_LogSheet, 1)
    
    ' Check if we have any logs
    If lastRow <= 1 Then
        Set GetErrorsBySource = result
        Exit Function
    End If
    
    ' Collect matching errors
    For i = 2 To lastRow
        If m_LogSheet.Cells(i, COL_SOURCE).value = source And _
           (m_LogSheet.Cells(i, COL_LEVEL).value = "ERROR" Or _
            m_LogSheet.Cells(i, COL_LEVEL).value = "CRITICAL") Then
            
            Dim errorInfo As Object
            Set errorInfo = CreateObject("Scripting.Dictionary")
            
            errorInfo.Add "Timestamp", m_LogSheet.Cells(i, COL_TIMESTAMP).value
            errorInfo.Add "Level", m_LogSheet.Cells(i, COL_LEVEL).value
            errorInfo.Add "Source", m_LogSheet.Cells(i, COL_SOURCE).value
            errorInfo.Add "ErrorNumber", m_LogSheet.Cells(i, COL_ERROR_NUM).value
            errorInfo.Add "Message", m_LogSheet.Cells(i, COL_MESSAGE).value
            errorInfo.Add "User", m_LogSheet.Cells(i, COL_USER).value
            errorInfo.Add "Details", m_LogSheet.Cells(i, COL_DETAILS).value
            
            result.Add errorInfo
        End If
    Next i
    
    Set GetErrorsBySource = result
    
    On Error GoTo 0
End Function

' ===== Private Methods =====

' Get or create the log sheet
Private Function GetOrCreateLogSheet() As Worksheet
    On Error Resume Next
    
    Dim ws As Worksheet
    
    ' Check if sheet exists
    If SheetExists(LOG_SHEET_NAME) Then
        Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    Else
        ' Create new hidden sheet
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = LOG_SHEET_NAME
        ws.Visible = xlSheetVeryHidden
        
        ' Set up headers
        ws.Cells(1, COL_TIMESTAMP).value = "Timestamp"
        ws.Cells(1, COL_LEVEL).value = "Level"
        ws.Cells(1, COL_SOURCE).value = "Source"
        ws.Cells(1, COL_ERROR_NUM).value = "ErrorNumber"
        ws.Cells(1, COL_MESSAGE).value = "Message"
        ws.Cells(1, COL_USER).value = "User"
        ws.Cells(1, COL_DETAILS).value = "Details"
        
        ' Format headers
        ws.Range("A1:G1").Font.Bold = True
    End If
    
    Set GetOrCreateLogSheet = ws
    
    On Error GoTo 0
End Function

' Count existing errors and warnings
Private Sub CountExistingLogs()
    On Error Resume Next
    
    If m_LogSheet Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    Dim i As Long
    
    m_ErrorCount = 0
    m_WarningCount = 0
    
    lastRow = GetLastRow(m_LogSheet, 1)
    
    ' Count existing errors and warnings
    For i = 2 To lastRow
        Dim levelStr As String
        levelStr = m_LogSheet.Cells(i, COL_LEVEL).value
        
        Select Case levelStr
            Case "WARNING"
                m_WarningCount = m_WarningCount + 1
            Case "ERROR", "CRITICAL"
                m_ErrorCount = m_ErrorCount + 1
        End Select
    Next i
    
    On Error GoTo 0
End Sub

' Add a log entry
Private Sub LogEntry(level As Integer, source As String, errorNumber As Long, message As String, details As String)
    On Error Resume Next
    
    If Not m_IsInitialized Then Initialize
    
    If m_LogSheet Is Nothing Then Exit Sub
    
    ' Get next row
    Dim newRow As Long
    newRow = GetLastRow(m_LogSheet, 1) + 1
    
    ' Check if we need to trim the log
    If newRow > MAX_LOG_ENTRIES + 1 Then
        TrimLog
        newRow = GetLastRow(m_LogSheet, 1) + 1
    End If
    
    ' Add entry
    m_LogSheet.Cells(newRow, COL_TIMESTAMP).value = Now
    m_LogSheet.Cells(newRow, COL_TIMESTAMP).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    
    m_LogSheet.Cells(newRow, COL_LEVEL).value = level
    m_LogSheet.Cells(newRow, COL_SOURCE).value = source
    m_LogSheet.Cells(newRow, COL_ERROR_NUM).value = errorNumber
    m_LogSheet.Cells(newRow, COL_MESSAGE).value = message
    m_LogSheet.Cells(newRow, COL_USER).value = Application.UserName
    m_LogSheet.Cells(newRow, COL_DETAILS).value = details
    
    ' Format based on level
    Select Case level
        Case LOG_LEVEL_WARNING
            m_LogSheet.Cells(newRow, COL_LEVEL).value = "WARNING"
            m_LogSheet.Cells(newRow, COL_LEVEL).Interior.Color = RGB(255, 255, 0) ' Yellow
            
        Case LOG_LEVEL_ERROR
            m_LogSheet.Cells(newRow, COL_LEVEL).value = "ERROR"
            m_LogSheet.Cells(newRow, COL_LEVEL).Interior.Color = RGB(255, 0, 0) ' Red
            
        Case LOG_LEVEL_CRITICAL
            m_LogSheet.Cells(newRow, COL_LEVEL).value = "CRITICAL"
            m_LogSheet.Cells(newRow, COL_LEVEL).Interior.Color = RGB(192, 0, 0) ' Dark Red
            m_LogSheet.Cells(newRow, COL_LEVEL).Font.Bold = True
            
        Case Else ' INFO
            m_LogSheet.Cells(newRow, COL_LEVEL).value = "INFO"
    End Select
    
    On Error GoTo 0
End Sub

' Trim the log if it gets too large
Private Sub TrimLog()
    On Error Resume Next
    
    If m_LogSheet Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = GetLastRow(m_LogSheet, 1)
    
    ' Keep header plus the most recent MAX_LOG_ENTRIES entries
    If lastRow > MAX_LOG_ENTRIES + 1 Then
        m_LogSheet.Rows("2:" & (lastRow - MAX_LOG_ENTRIES)).Delete
        
        ' Recount after trimming
        CountExistingLogs
    End If
    
    On Error GoTo 0
End Sub

' Get the last used row in a column
Private Function GetLastRow(ws As Worksheet, col As Long) As Long
    On Error Resume Next
    
    If ws Is Nothing Then
        GetLastRow = 0
        Exit Function
    End If
    
    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    
    If Err.Number <> 0 Then
        GetLastRow = 0
    End If
    
    On Error GoTo 0
End Function

' Check if sheet exists
Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    SheetExists = Not ws Is Nothing
    Set ws = Nothing
End Function

' Class initialize
Private Sub Class_Initialize()
    m_IsInitialized = False
    m_ErrorCount = 0
    m_WarningCount = 0
    m_CurrentLogLevel = LOG_LEVEL_INFO ' Default to INFO level
End Sub

' Class terminate
Private Sub Class_Terminate()
    ' Ensure we clean up references
    Set m_LogSheet = Nothing
End Sub