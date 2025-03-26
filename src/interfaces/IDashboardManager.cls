' File: src/interfaces/IDashboardManager.cls
'----------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDashboardManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' IDashboardManager Interface
'
' Purpose: Defines the contract for the dashboard management functionality.
' This interface abstracts the dashboard UI operations, allowing for
' decoupling of UI logic from business logic and data access.

' Initialize the dashboard
Public Sub Initialize(Optional ByVal repository As ITransactionRepository = Nothing, _
                     Optional ByVal calculator As IBalanceCalculator = Nothing, _
                     Optional ByVal logger As IErrorLogger = Nothing)
End Sub

' Refresh the entire dashboard
Public Sub RefreshDashboard()
End Sub

' Update the balance summary card
Public Sub UpdateBalanceSummary()
End Sub

' Update category summary and charts
Public Sub UpdateCategorySummary()
End Sub

' Update monthly spending chart
Public Sub UpdateMonthlySpending()
End Sub

' Update day of week spending chart
Public Sub UpdateDayOfWeekSpending()
End Sub

' Handle transaction data changes
Public Sub HandleTransactionsChanged()
End Sub

' Handle balance updates
Public Sub HandleBalanceUpdated(ByVal balanceData As Variant)
End Sub

' Handle category data changes
Public Sub HandleCategoryDataChanged(ByVal categoryData As Variant)
End Sub

' Handle settings changes
Public Sub HandleSettingsChanged()
End Sub

' Apply date filter
Public Sub ApplyDateFilter(ByVal startDate As Date, ByVal endDate As Date)
End Sub

' Get the current filter settings
Public Function GetCurrentFilter() As Object ' Dictionary
End Function