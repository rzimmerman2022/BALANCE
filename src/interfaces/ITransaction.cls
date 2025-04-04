' File: src/interfaces/ITransaction.cls
'----------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' ITransaction Interface
'
' Purpose: Defines the contract for transaction objects.
' This interface abstracts the transaction data and behavior,
' allowing for different implementations or specialized transaction types.

' Properties
Public Property Get ID() As String
End Property

Public Property Let ID(value As String)
End Property

Public Property Get TransactionDate() As Date
End Property

Public Property Let TransactionDate(value As Date)
End Property

Public Property Get Merchant() As String
End Property

Public Property Let Merchant(value As String)
End Property

Public Property Get Category() As String
End Property

Public Property Let Category(value As String)
End Property

Public Property Get Amount() As Currency
End Property

Public Property Let Amount(value As Currency)
End Property

Public Property Get Account() As String
End Property

Public Property Let Account(value As String)
End Property

Public Property Get Owner() As String
End Property

Public Property Let Owner(value As String)
End Property

Public Property Get IsShared() As Boolean
End Property

Public Property Let IsShared(value As Boolean)
End Property

Public Property Get Notes() As String
End Property

Public Property Let Notes(value As String)
End Property

Public Property Get SourceFile() As String
End Property

Public Property Let SourceFile(value As String)
End Property

' Methods
Public Function IsExpense() As Boolean
End Function

Public Function IsIncome() As Boolean
End Function

Public Function GetAbsoluteAmount() As Currency
End Function

Public Sub CalculateSplit(ByRef user1Amount As Currency, ByRef user2Amount As Currency)
End Sub

Public Function ToDictionary() As Object
End Function

Public Sub InitFromRow(dateVal As Variant, merchant As Variant, category As Variant, _
                      amount As Variant, account As Variant, owner As Variant, _
                      isShared As Variant, notes As Variant, sourceFile As Variant)
End Sub
