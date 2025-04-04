' File: src/interfaces/IBalanceCalculator.cls
'----------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBalanceCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' IBalanceCalculator Interface
'
' Purpose: Defines the contract for calculating core balance functionality.
' This interface has been refactored to focus only on balance calculations,
' with other analysis functions moved to more specialized interfaces
' (ICategoryAnalyzer, ITimeAnalyzer, ITrendAnalyzer) to better follow
' the Interface Segregation Principle.

' Calculate the overall balance between users
Public Function CalculateBalance(ByVal transactions As Collection) As Double
End Function

' Calculate who owes whom and how much
Public Function CalculateBalanceSummary(ByVal transactions As Collection) As Object ' Dictionary
End Function

' Initialize the calculator with dependencies
Public Sub Initialize(Optional ByVal logger As IErrorLogger = Nothing)
End Sub
