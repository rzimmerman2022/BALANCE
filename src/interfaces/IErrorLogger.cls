' File: src/interfaces/IErrorLogger.cls
'----------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IErrorLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' IErrorLogger Interface
'
' Purpose: Defines the contract for error logging functionality.
' This interface allows for dependency injection of logging capabilities
' and enables different logging implementations (e.g., file-based, 
' worksheet-based, or mock loggers for testing).

' Log levels enum (would be in a separate module)
' Public Enum LogLevel
'     LogLevel_Error = 1
'     LogLevel_Warning = 2
'     LogLevel_Info = 3
'     LogLevel_Debug = 4
' End Enum

' Log an error message with error number and description
Public Sub LogError(ByVal source As String, ByVal errorNumber As Long, ByVal description As String)
End Sub

' Log a warning message
Public Sub LogWarning(ByVal source As String, ByVal message As String)
End Sub

' Log an informational message
Public Sub LogInfo(ByVal source As String, ByVal message As String)
End Sub

' Log a debug message (only logged in debug mode)
Public Sub LogDebug(ByVal source As String, ByVal message As String)
End Sub

' Enable or disable a specific log level
Public Sub SetLogLevel(ByVal level As Long, ByVal enabled As Boolean)
End Sub

' Check if a specific log level is enabled
Public Function IsLogLevelEnabled(ByVal level As Long) As Boolean
End Function

' Clear all logged messages
Public Sub ClearLog()
End Sub

' Export log to a file
Public Function ExportLog(ByVal filePath As String) As Boolean
End Function

' Initialize the logger
Public Sub Initialize()
End Sub