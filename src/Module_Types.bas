Attribute VB_Name = "Module_Types"
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' Module_Types - Contains public Type definitions used across the project.

' Analysis result type (Moved from TransactionAnalyzer.cls)
Public Type AnalysisInsight
    Category As String
    Title As String
    Description As String
    Value As Variant
    Trend As String ' "Up", "Down", "Neutral"
    Importance As Integer ' 1 (Low) to 5 (High)
End Type

' Add other public Type definitions here as needed...

' Example of other potential types (commented out)
'Public Type CategorySummary
'    Category As String
'    TotalAmount As Currency
'    Percentage As Double
'End Type
'
'Public Type MonthSummary
'    MonthKey As String ' yyyy-mm
'    TotalAmount As Currency
'    IncomeAmount As Currency
'    ExpenseAmount As Currency
'End Type
'
'Public Type DayOfWeekSummary
'    DayName As String
'    TotalAmount As Currency
'    Percentage As Double
'End Type
