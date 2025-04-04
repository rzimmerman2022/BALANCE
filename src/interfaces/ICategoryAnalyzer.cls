' File: src/interfaces/ICategoryAnalyzer.cls
'----------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICategoryAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' ICategoryAnalyzer Interface
'
' Purpose: Defines the contract for category-based financial analysis.
' This interface is responsible for analyzing spending by category,
' allowing for more focused implementation of the Interface Segregation Principle.

' Calculate spending by category
Public Function CalculateCategorySummary(ByVal transactions As Collection, _
                                         Optional ByVal startDate As Date, _
                                         Optional ByVal endDate As Date) As Object ' Dictionary
End Function

' Initialize the analyzer with dependencies
Public Sub Initialize(Optional ByVal logger As IErrorLogger = Nothing)
End Sub
