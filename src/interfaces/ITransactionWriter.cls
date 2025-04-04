' File: src/interfaces/ITransactionWriter.cls
'----------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITransactionWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' ITransactionWriter Interface
'
' Purpose: Defines the contract for writing transaction data.
' This interface is focused solely on write operations, following
' the Interface Segregation Principle by separating read and write concerns.

' Add a single transaction
Public Function AddTransaction(ByVal transaction As ITransaction) As Boolean
End Function

' Add multiple transactions
Public Function AddTransactions(ByVal transactions As Collection) As Long
End Function

' Remove a transaction by ID
Public Function RemoveTransaction(ByVal transactionId As String) As Boolean
End Function

' Save changes to storage
Public Function SaveChanges() As Boolean
End Function

' Clear all transactions
Public Sub ClearAll()
End Sub

' Initialize the writer
Public Sub Initialize(Optional ByVal logger As IErrorLogger = Nothing)
End Sub
