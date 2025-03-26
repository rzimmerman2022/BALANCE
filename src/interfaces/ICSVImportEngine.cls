' File: src/interfaces/ICSVImportEngine.cls
'------------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICSVImportEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' ICSVImportEngine Interface
'
' Purpose: Defines the contract for the CSV import engine.
' The engine coordinates the import process, selecting appropriate strategies
' and handling the overall import workflow.

' Register a strategy for a particular CSV format
Public Sub RegisterStrategy(ByVal strategy As ICSVImportStrategy)
End Sub

' Unregister a strategy by name
Public Function UnregisterStrategy(ByVal strategyName As String) As Boolean
End Function

' Get a list of all registered strategies
Public Function GetRegisteredStrategies() As Collection
End Function

' Import transactions from a specific CSV file
Public Function ImportTransactionsFromFile(ByVal filePath As String, ByVal owner As String, _
                                         Optional ByVal specificStrategy As String = "") As Collection
End Function

' Import transactions from all CSV files in a directory
Public Function ImportTransactionsFromDirectory(ByVal directoryPath As String, ByVal owner As String, _
                                              Optional ByVal filePattern As String = "*.csv") As Collection
End Function

' Auto-detect the appropriate strategy for a CSV file
Public Function DetectStrategy(ByVal filePath As String) As ICSVImportStrategy
End Function

' Initialize the engine
Public Sub Initialize(Optional ByVal repository As ITransactionRepository = Nothing, _
                     Optional ByVal logger As IErrorLogger = Nothing)
End Sub