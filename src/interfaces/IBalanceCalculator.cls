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
' Purpose: Defines the contract for calculating balances and financial summaries.
' This interface abstracts the balance calculation logic, allowing for different
' implementations or testing versions to be substituted as needed.

' Calculate the overall balance between users
Public Function CalculateBalance(ByVal transactions As Collection) As Double
End Function

' Calculate who owes whom and how much
Public Function CalculateBalanceSummary(ByVal transactions As Collection) As Object ' Dictionary
End Function

' Calculate spending by category
Public Function CalculateCategorySummary(ByVal transactions As Collection, _
                                         Optional ByVal startDate As Date, _
                                         Optional ByVal endDate As Date) As Object ' Dictionary
End Function

' Calculate spending by day of week
Public Function CalculateDayOfWeekSummary(ByVal transactions As Collection, _
                                         Optional ByVal startDate As Date, _
                                         Optional ByVal endDate As Date) As Object ' Dictionary
End Function

' Calculate spending by month
Public Function CalculateMonthSummary(ByVal transactions As Collection, _
                                     Optional ByVal monthCount As Integer = 12) As Object ' Dictionary
End Function

' Calculate daily average spending
Public Function CalculateDailyAverageSpending(ByVal transactions As Collection, _
                                             Optional ByVal startDate As Date, _
                                             Optional ByVal endDate As Date) As Double
End Function

' Calculate spending trends (month-over-month changes)
Public Function CalculateSpendingTrends(ByVal transactions As Collection, _
                                       Optional ByVal monthCount As Integer = 12) As Object ' Dictionary
End Function

' Initialize the calculator with dependencies
Public Sub Initialize(Optional ByVal logger As IErrorLogger = Nothing)
End Sub