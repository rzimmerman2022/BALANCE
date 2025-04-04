' File: src/interfaces/ITransactionRepository.cls
'----------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITransactionRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' ITransactionRepository Interface
'
' Purpose: Defines the contract for the transaction repository.
' This interface combines both read and write operations by extending
' ITransactionReader and ITransactionWriter interfaces, following
' the Interface Segregation Principle.

' This interface inherits all methods from ITransactionReader and ITransactionWriter
' It serves as a convenience interface for components that need both read and write access

' Get all transactions (from ITransactionReader)
Public Function GetTransactions() As Collection
End Function

' Get transaction count (from ITransactionReader)
Public Property Get Count() As Long
End Property

' Get transactions by date range (from ITransactionReader)
Public Function GetTransactionsByDateRange(ByVal startDate As Date, ByVal endDate As Date) As Collection
End Function

' Get transactions by category (from ITransactionReader)
Public Function GetTransactionsByCategory(ByVal category As String) As Collection
End Function

' Get transactions by owner (from ITransactionReader)
Public Function GetTransactionsByOwner(ByVal owner As String) As Collection
End Function

' Add a single transaction (from ITransactionWriter)
Public Function AddTransaction(ByVal transaction As ITransaction) As Boolean
End Function

' Add multiple transactions (from ITransactionWriter)
Public Function AddTransactions(ByVal transactions As Collection) As Long
End Function

' Remove a transaction by ID (from ITransactionWriter)
Public Function RemoveTransaction(ByVal transactionId As String) As Boolean
End Function

' Save changes to storage (from ITransactionWriter)
Public Function SaveChanges() As Boolean
End Function

' Clear all transactions (from ITransactionWriter)
Public Sub ClearAll()
End Sub

' Initialize the repository
Public Sub Initialize(Optional ByVal logger As IErrorLogger = Nothing)
End Sub
