' File: src/interfaces/ITimeAnalyzer.cls
'----------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITimeAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' ITimeAnalyzer Interface
'
' Purpose: Defines the contract for time-based financial analysis.
' This interface is responsible for analyzing spending patterns over time,
' allowing for more focused implementation of the Interface Segregation Principle.

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

' Initialize the analyzer with dependencies
Public Sub Initialize(Optional ByVal logger As IErrorLogger = Nothing)
End Sub
