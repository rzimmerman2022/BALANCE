' File: src/interfaces/ITrendAnalyzer.cls
'----------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITrendAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' ITrendAnalyzer Interface
'
' Purpose: Defines the contract for trend analysis functionality.
' This interface is responsible for analyzing spending trends over time,
' allowing for more focused implementation of the Interface Segregation Principle.

' Calculate spending trends (month-over-month changes)
Public Function CalculateSpendingTrends(ByVal transactions As Collection, _
                                       Optional ByVal monthCount As Integer = 12) As Object ' Dictionary
End Function

' Initialize the analyzer with dependencies
Public Sub Initialize(Optional ByVal logger As IErrorLogger = Nothing)
End Sub
