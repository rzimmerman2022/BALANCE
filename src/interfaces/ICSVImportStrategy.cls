' File: src/interfaces/ICSVImportStrategy.cls
'------------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICSVImportStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' ICSVImportStrategy Interface
'
' Purpose: Defines the contract for strategies that import specific CSV formats.
' Each strategy knows how to parse a particular CSV format and convert it
' into standardized Transaction objects.

' Check if this strategy can handle the given CSV file
Public Function CanProcessFile(ByVal filePath As String) As Boolean
End Function

' Get a name for this strategy (e.g., "Bank of America", "Chase Credit Card")
Public Property Get StrategyName() As String
End Property

' Get a description for this strategy
Public Property Get StrategyDescription() As String
End Property

' Parse the CSV file and return a collection of Transaction objects
Public Function ImportTransactions(ByVal filePath As String, ByVal owner As String, _
                                  Optional ByVal logger As IErrorLogger = Nothing) As Collection
End Function

' Get the column mapping for this strategy
Public Function GetColumnMapping() As Object ' Dictionary
End Function

' Get expected headers for this strategy
Public Function GetExpectedHeaders() As Variant ' Array of header names
End Function