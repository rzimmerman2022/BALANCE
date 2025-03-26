' File: src/utils/ErrorLogger.cls
'-----------------------------------
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
' ErrorLogger Class - Implements IErrorLogger
'
' Purpose: Provides centralized error logging functionality
' The logger captures errors, warnings, and informational messages
' and stores them in a hidden worksheet for later review.
'
' Design Decisions:
' - Uses PredeclaredId = True to enable Singleton pattern
' - Implements IErrorLogger for dependency injection
' - Uses Excel Tables for structured log storage
' - Provides multiple severity levels with filtering
' - Includes export capability for log analysis

' Interface implementation
Implements IErrorLogger

' Log level enum definition
Public Enum LogLevel
    LogLevel_Error = 1
    LogLevel_Warning = 2
    LogLevel_Info = 3
    LogLevel_Debug = 4
End Enum

' Constants for log sheet and table
Private Const LOG_SHEET_NAME As String = "ErrorLog"
Private Const LOG_TABLE_NAME As String = "LogTable"
Private Const MAX_LOG_ENTRIES As Long = 1000 ' Limit log size to prevent excessive growth

' Column indices for the log table
Private Const COL_TIMESTAMP As Long = 1
Private Const COL_LEVEL As Long = 2
Private Const COL_SOURCE As Long = 3
Private Const COL_MESSAGE As Long = 4
Private Const COL_ERROR_NUMBER As Long = 5
Private Const COL_USER As Long = 6

' Private member variables
Private m_LogSheet As Worksheet
Private m_LogTable As ListObject
Private m_IsInitialized As Boolean
Private m_EnabledLevels As Object ' Dictionary to store enabled log levels
Private m_MaxEntries As Long ' Maximum number of log entries to keep

'=========================================================================
' Initialization and Setup
'=========================================================================

Private Sub Class_Initialize()
    ' Set default enabled log levels
    Set m_EnabledLevels = CreateObject("Scripting.Dictionary")
    m_EnabledLevels(LogLevel_Error) = True
    m_EnabledLevels(LogLevel_Warning) = True
    m_EnabledLevels(LogLevel_Info) = True
    
    ' Debug logging is off by default in production
    #If DebugMode Then
        m_EnabledLevels(LogLevel_Debug) = True
    #Else
        m_EnabledLevels(LogLevel_Debug) = False
    #End If
    
    m_MaxEntries = MAX_LOG_ENTRIES
    m_IsInitialized = False
End Sub

' Initialize the logger and set up log storage
Private Sub IErrorLogger_Initialize()
    On Error Resume Next
    
    ' Check if already initialized
    If m_IsInitialized Then Exit Sub
    
    ' Get or create log sheet
    Set m_LogSheet = GetOrCreateLogSheet()
    
    ' Get or create log table
    Set m_LogTable = GetOrCreateLogTable()
    
    ' Trim log if needed
    TrimLogIfNeeded
    
    m_IsInitialized = True
    
    ' Log initialization
    Me.LogInfo "ErrorLogger.Initialize", "Error logger initialized"
    
    On Error GoTo 0
End Sub

' Public wrapper for Initialize
Public Sub Initialize()
    IErrorLogger_Initialize
End Sub

' Get or create the log sheet
Private Function GetOrCreateLogSheet() As Worksheet
    On Error Resume Next
    
    Dim ws As Worksheet
    
    ' Try to get existing sheet
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    
    ' If sheet doesn't exist, create it
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = LOG_SHEET_NAME
        
        ' Hide the sheet (very hidden so it can't be unhidden through the UI)
        ws.Visible = xlSheetVeryHidden
    End If
    
    Set GetOrCreateLogSheet = ws
    
    On Error GoTo 0
End Function

' Get or create the log table
Private Function GetOrCreateLogTable() As ListObject
    On Error Resume Next
    
    Dim tbl As ListObject
    
    ' Check if table exists
    For Each tbl In m_LogSheet.ListObjects
        If tbl.Name = LOG_TABLE_NAME Then
            Set GetOrCreateLogTable = tbl
            Exit Function
        End If
    Next tbl
    
    ' Create table headers
    m_LogSheet.Cells(1, COL_TIMESTAMP).Value = "Timestamp"
    m_LogSheet.Cells(1, COL_LEVEL).Value = "Level"
    m_LogSheet.Cells(1, COL_SOURCE).Value = "Source"
    m_LogSheet.Cells(1, COL_MESSAGE).Value = "Message"
    m_LogSheet.Cells(1, COL_ERROR_NUMBER).Value = "ErrorNumber"
    m_LogSheet.Cells(1, COL_USER).Value = "User"
    
    ' Create table
    Set tbl = m_LogSheet.ListObjects.Add(xlSrcRange, _
                           m_LogSheet.Range(m_LogSheet.Cells(1, COL_TIMESTAMP), _
                                           m_LogSheet.Cells(1, COL_USER)), , xlYes)
                                           
    ' Set table name
    tbl.Name = LOG_TABLE_NAME
    
    ' Format the table
    With tbl.HeaderRowRange
        .Font.Bold = True
        .Interior.Color = RGB(192, 0, 0) ' Red header for error log
        .Font.Color = RGB(255, 255, 255) ' White text
    End With
    
    ' Set column widths for readability
    m_LogSheet.Columns(COL_TIMESTAMP).ColumnWidth = 20
    m_LogSheet.Columns(COL_LEVEL).ColumnWidth = 10
    m_LogSheet.Columns(COL_SOURCE).ColumnWidth = 30
    m_LogSheet.Columns(COL_MESSAGE).ColumnWidth = 50
    m_LogSheet.Columns(COL_ERROR_NUMBER).ColumnWidth = 15
    m_LogSheet.Columns(COL_USER).ColumnWidth = 20
    
    Set GetOrCreateLogTable = tbl
    
    On Error GoTo 0
End Function

'=========================================================================
' Interface Implementation - Logging Methods
'=========================================================================

' Log an error message
Private Sub IErrorLogger_LogError(ByVal source As String, ByVal errorNumber As Long, ByVal description As String)
    ' Check if error logging is enabled
    If Not m_IsInitialized Then IErrorLogger_Initialize
    If Not m_EnabledLevels(LogLevel_Error) Then Exit Sub
    
    ' Add log entry
    AddLogEntry LogLevel_Error, source, description, errorNumber
End Sub

' Public wrapper for LogError
Public Sub LogError(ByVal source As String, ByVal errorNumber As Long, ByVal description As String)
    IErrorLogger_LogError source, errorNumber, description
End Sub

' Log a warning message
Private Sub IErrorLogger_LogWarning(ByVal source As String, ByVal message As String)
    ' Check if warning logging is enabled
    If Not m_IsInitialized Then IErrorLogger_Initialize
    If Not m_EnabledLevels(LogLevel_Warning) Then Exit Sub
    
    ' Add log entry
    AddLogEntry LogLevel_Warning, source, message
End Sub

' Public wrapper for LogWarning
Public Sub LogWarning(ByVal source As String, ByVal message As String)
    IErrorLogger_LogWarning source, message
End Sub

' Log an informational message
Private Sub IErrorLogger_LogInfo(ByVal source As String, ByVal message As String)
    ' Check if info logging is enabled
    If Not m_IsInitialized Then IErrorLogger_Initialize
    If Not m_EnabledLevels(LogLevel_Info) Then Exit Sub
    
    ' Add log entry
    AddLogEntry LogLevel_Info, source, message
End Sub

' Public wrapper for LogInfo
Public Sub LogInfo(ByVal source As String, ByVal message As String)
    IErrorLogger_LogInfo source, message
End Sub

' Log a debug message
Private Sub IErrorLogger_LogDebug(ByVal source As String, ByVal message As String)
    ' Check if debug logging is enabled
    If Not m_IsInitialized Then IErrorLogger_Initialize
    If Not m_EnabledLevels(LogLevel_Debug) Then Exit Sub
    
    ' Add log entry
    AddLogEntry LogLevel_Debug, source, message
End Sub

' Public wrapper for LogDebug
Public Sub LogDebug(ByVal source As String, ByVal message As String)
    IErrorLogger_LogDebug source, message
End Sub

'=========================================================================
' Interface Implementation - Log Configuration
'=========================================================================

' Enable or disable a specific log level
Private Sub IErrorLogger_SetLogLevel(ByVal level As Long, ByVal enabled As Boolean)
    If Not m_IsInitialized Then IErrorLogger_Initialize
    
    ' Make sure level is valid
    If level < LogLevel_Error Or level > LogLevel_Debug Then Exit Sub
    
    ' Set level enabled state
    m_EnabledLevels(level) = enabled
End Sub

' Public wrapper for SetLogLevel
Public Sub SetLogLevel(ByVal level As Long, ByVal enabled As Boolean)
    IErrorLogger_SetLogLevel level, enabled
End Sub

' Check if a specific log level is enabled
Private Function IErrorLogger_IsLogLevelEnabled(ByVal level As Long) As Boolean
    If Not m_IsInitialized Then IErrorLogger_Initialize
    
    ' Make sure level is valid
    If level < LogLevel_Error Or level > LogLevel_Debug Then
        IErrorLogger_IsLogLevelEnabled = False
        Exit Function
    End If
    
    ' Return level enabled state
    IErrorLogger_IsLogLevelEnabled = m_EnabledLevels(level)
End Function

' Public wrapper for IsLogLevelEnabled
Public Function IsLogLevelEnabled(ByVal level As Long) As Boolean
    IsLogLevelEnabled = IErrorLogger_IsLogLevelEnabled(level)
End Function

' Clear all log entries
Private Sub IErrorLogger_ClearLog()
    On Error Resume Next
    
    If Not m_IsInitialized Then IErrorLogger_Initialize
    
    ' Check if table has data
    If m_LogTable.ListRows.Count > 0 Then
        m_LogTable.DataBodyRange.Delete
    End If
    
    ' Log clearance (will be the first new entry)
    Me.LogInfo "ErrorLogger.ClearLog", "Log cleared"
    
    On Error GoTo 0
End Sub

' Public wrapper for ClearLog
Public Sub ClearLog()
    IErrorLogger_ClearLog
End Sub

' Export log to a file
Private Function IErrorLogger_ExportLog(ByVal filePath As String) As Boolean
    On Error Resume Next
    
    If Not m_IsInitialized Then IErrorLogger_Initialize
    
    ' Check if table has data
    If m_LogTable.ListRows.Count = 0 Then
        IErrorLogger_ExportLog = False
        Exit Function
    End If
    
    ' Make a copy of the log sheet
    Dim tempSheet As Worksheet
    m_LogSheet.Copy
    Set tempSheet = ActiveSheet
    
    ' Make the sheet visible
    tempSheet.Visible = xlSheetVisible
    
    ' Save as CSV
    tempSheet.Parent.SaveAs filePath, xlCSV
    
    ' Close the temporary workbook without saving changes
    tempSheet.Parent.Close False
    
    ' Log export action
    Me.LogInfo "ErrorLogger.ExportLog", "Log exported to " & filePath
    
    IErrorLogger_ExportLog = (Err.Number = 0)
    
    On Error GoTo 0
End Function

' Public wrapper for ExportLog
Public Function ExportLog(ByVal filePath As String) As Boolean
    ExportLog = IErrorLogger_ExportLog(filePath)
End Function

'=========================================================================
' Helper Methods
'=========================================================================

' Add a log entry with specified level, source, and message
Private Sub AddLogEntry(ByVal level As LogLevel, ByVal source As String, _
                        ByVal message As String, Optional ByVal errorNumber As Long = 0)
    On Error Resume Next
    
    If Not m_IsInitialized Then IErrorLogger_Initialize
    
    ' Get level name
    Dim levelName As String
    Select Case level
        Case LogLevel_Error
            levelName = "ERROR"
        Case LogLevel_Warning
            levelName = "WARNING"
        Case LogLevel_Info
            levelName = "INFO"
        Case LogLevel_Debug
            levelName = "DEBUG"
        Case Else
            levelName = "UNKNOWN"
    End Select
    
    ' Trim log if needed before adding new entry
    TrimLogIfNeeded
    
    ' Add a new row to the table
    Dim newRow As ListRow
    Set newRow = m_LogTable.ListRows.Add
    
    ' Populate the row
    newRow.Range(1, COL_TIMESTAMP).Value = Now
    newRow.Range(1, COL_LEVEL).Value = levelName
    newRow.Range(1, COL_SOURCE).Value = source
    newRow.Range(1, COL_MESSAGE).Value = message
    newRow.Range(1, COL_ERROR_NUMBER).Value = errorNumber
    newRow.Range(1, COL_USER).Value = GetCurrentUser()
    
    ' Format based on level
    Select Case level
        Case LogLevel_Error
            newRow.Range.Interior.Color = RGB(255, 200, 200) ' Light red
        Case LogLevel_Warning
            newRow.Range.Interior.Color = RGB(255, 255, 200) ' Light yellow
    End Select
    
    ' Debug output in immediate window
    #If DebugMode Then
        Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & _
                  levelName & " | " & source & " | " & message & _
                  IIf(errorNumber <> 0, " | Error #" & errorNumber, "")
    #End If
    
    On Error GoTo 0
End Sub

' Trim log if it exceeds maximum entries
Private Sub TrimLogIfNeeded()
    On Error Resume Next
    
    ' Check if log needs trimming
    If m_LogTable.ListRows.Count > m_MaxEntries Then
        ' Calculate how many rows to remove (remove oldest 20% of entries)
        Dim rowsToRemove As Long
        rowsToRemove = m_LogTable.ListRows.Count - m_MaxEntries + CInt(m_MaxEntries * 0.2)
        
        ' Delete oldest rows (they're at the top of the table)
        m_LogTable.DataBodyRange.Rows("1:" & rowsToRemove).Delete
    End If
    
    On Error GoTo 0
End Sub

' Get current username
Private Function GetCurrentUser() As String
    On Error Resume Next
    
    GetCurrentUser = Application.UserName
    
    ' If Excel username not available, try Windows username
    If Len(GetCurrentUser) = 0 Then
        GetCurrentUser = Environ("USERNAME")
    End If
    
    ' If still not available, use "Unknown"
    If Len(GetCurrentUser) = 0 Then
        GetCurrentUser = "Unknown"
    End If
    
    On Error GoTo 0
End Function

'=========================================================================
' Properties
'=========================================================================

' Get/set the maximum number of log entries to keep
Public Property Get MaxEntries() As Long
    MaxEntries = m_MaxEntries
End Property

Public Property Let MaxEntries(ByVal value As Long)
    ' Ensure minimum of 100 entries
    If value < 100 Then value = 100
    
    m_MaxEntries = value
    
    ' Trim log if needed after changing max entries
    If m_IsInitialized Then TrimLogIfNeeded
End Property

'=========================================================================
' Cleanup
'=========================================================================

Private Sub Class_Terminate()
    ' Clean up object references
    Set m_LogSheet = Nothing
    Set m_LogTable = Nothing
    Set m_EnabledLevels = Nothing
End Sub