' File: src/interfaces/ITransactionReader.cls
'----------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITransactionReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' ITransactionReader Interface
'
' Purpose: Defines the contract for reading transaction data.
' This interface is focused solely on read operations, following
' the Interface Segregation Principle by separating read and write concerns.

' Get all transactions
Public Function GetTransactions() As Collection
End Function

' Get transaction count
Public Property Get Count() As Long
End Property

' Get transactions by date range
Public Function GetTransactionsByDateRange(ByVal startDate As Date, ByVal endDate As Date) As Collection
End Function

' Get transactions by category
Public Function GetTransactionsByCategory(ByVal category As String) As Collection
End Function

' Get transactions by owner
Public Function GetTransactionsByOwner(ByVal owner As String) As Collection
End Function

' Initialize the reader
Public Sub Initialize(Optional ByVal logger As IErrorLogger = Nothing)
End Sub
