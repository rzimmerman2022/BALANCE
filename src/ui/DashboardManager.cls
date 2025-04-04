' File: src/ui/DashboardManager.cls
'---------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DashboardManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' DashboardManager Class - Implements IDashboardManager
'
' Purpose: Manages the dashboard UI, handling the display and update of
' all dashboard elements based on data changes and user interactions.
'
' Design Decisions:
' - Uses PredeclaredId = True to enable Singleton pattern
' - Implements IDashboardManager for loose coupling
' - Uses event-driven architecture to respond to data changes
' - Subscribes to events rather than being directly called
' - Optimizes UI updates to minimize Excel recalculation

' Interface implementation
Implements IDashboardManager

' Constants
Private Const DASHBOARD_SHEET_NAME As String = "Dashboard"
Private Const BALANCE_CARD_RANGE As String = "B4:D8"
Private Const CATEGORY_CHART_NAME As String = "CategoryChart"
Private Const MONTHLY_CHART_NAME As String = "MonthlyChart"
Private Const DAILY_CHART_NAME As String = "DailyChart"
Private Const DEFAULT_DATE_RANGE_MONTHS As Integer = 6

' Private member variables
Private m_Repository As ITransactionRepository
Private m_Calculator As IBalanceCalculator
Private m_ChartFactory As ChartFactory
Private m_Logger As IErrorLogger
Private m_DashboardSheet As Worksheet
Private m_IsInitialized As Boolean
Private m_StartDate As Date
Private m_EndDate As Date
Private m_IsRefreshing As Boolean ' Flag to prevent event loops

'=========================================================================
' Initialization
'=========================================================================

Private Sub Class_Initialize()
    ' Set default values
    m_IsInitialized = False
    m_IsRefreshing = False
    
    ' Set default date range (last 6 months)
    m_EndDate = Date
    m_StartDate = DateAdd("m", -DEFAULT_DATE_RANGE_MONTHS, m_EndDate)
End Sub

' Initialize the dashboard manager with dependencies
Private Sub IDashboardManager_Initialize(Optional ByVal repository As ITransactionRepository = Nothing, _
                                        Optional ByVal calculator As IBalanceCalculator = Nothing, _
                                        Optional ByVal logger As IErrorLogger = Nothing)
    On Error GoTo ErrorHandler
    
    ' Set dependencies if provided
    If Not repository Is Nothing Then
        Set m_Repository = repository
    End If
    
    If Not calculator Is Nothing Then
        Set m_Calculator = calculator
    End If
    
    If Not logger Is Nothing Then
        Set m_Logger = logger
    End If
    
    ' Get or create the dashboard sheet
    Set m_DashboardSheet = GetOrCreateDashboardSheet()
    
    ' Initialize chart factory (Use PredeclaredId instance)
    Set m_ChartFactory = ChartFactory ' Changed New ChartFactory to ChartFactory
    m_ChartFactory.Initialize m_Logger
    
    ' Subscribe to events
    SubscribeToEvents
    
    m_IsInitialized = True
    
    ' Log initialization
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "DashboardManager.Initialize", "Dashboard Manager initialized"
    End If
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "DashboardManager.Initialize", Err.Number, Err.Description
    End If
End Sub

' Public wrapper for Initialize
Public Sub Initialize(Optional ByVal repository As ITransactionRepository = Nothing, _
                     Optional ByVal calculator As IBalanceCalculator = Nothing, _
                     Optional ByVal logger As IErrorLogger = Nothing)
    IDashboardManager_Initialize repository, calculator, logger
End Sub

' Get or create the dashboard sheet
Private Function GetOrCreateDashboardSheet() As Worksheet
    On Error Resume Next
    
    Dim ws As Worksheet
    
    ' Try to get existing sheet
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET_NAME)
    
    ' If sheet doesn't exist, create it
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = DASHBOARD_SHEET_NAME
        
        ' Set up basic dashboard structure
        SetupDashboardLayout ws
    End If
    
    Set GetOrCreateDashboardSheet = ws
    
    On Error GoTo 0
End Function

' Set up the basic dashboard layout
Private Sub SetupDashboardLayout(ByVal ws As Worksheet)
    On Error Resume Next
    
    ' Clear the sheet
    ws.Cells.Clear
    
    ' Add title
    With ws.Range("B2")
        .Value = "BALANCE - Financial Dashboard"
        .Font.Size = 16
        .Font.Bold = True
    End With
    
    ' Add balance card
    With ws.Range(BALANCE_CARD_RANGE)
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(240, 240, 240)
    End With
    
    ws.Range("B4").Value = "Balance Summary"
    ws.Range("B4").Font.Bold = True
    
    ' Add section titles
    ws.Range("B10").Value = "Spending by Category"
    ws.Range("B10").Font.Bold = True
    
    ws.Range("F10").Value = "Monthly Spending"
    ws.Range("F10").Font.Bold = True
    
    ws.Range("B25").Value = "Day of Week Spending"
    ws.Range("B25").Font.Bold = True
    
    ' Add date filter controls
    ws.Range("G2").Value = "Start Date:"
    ws.Range("G3").Value = "End Date:"
    
    ws.Range("H2").Value = m_StartDate
    ws.Range("H2").NumberFormat = "mm/dd/yyyy"
    
    ws.Range("H3").Value = m_EndDate
    ws.Range("H3").NumberFormat = "mm/dd/yyyy"
    
    ' Add refresh button
    ws.Buttons.Add(670, 5, 80, 25).Select
    With Selection
        .OnAction = "RefreshDashboardAction"
        .Caption = "Refresh"
        .Name = "RefreshButton"
    End With
    
    ' Deselect
    ws.Range("A1").Select
    
    On Error GoTo 0
End Sub

' Subscribe to events from data sources
Private Sub SubscribeToEvents()
    On Error GoTo ErrorHandler
    
    ' Subscribe to transaction changes
    EventManager.Subscribe EventType.TransactionsChanged, Me, "HandleTransactionsChanged"
    
    ' Subscribe to balance updates
    EventManager.Subscribe EventType.BalanceUpdated, Me, "HandleBalanceUpdated"
    
    ' Subscribe to category data changes
    EventManager.Subscribe EventType.CategoryDataChanged, Me, "HandleCategoryDataChanged"
    
    ' Subscribe to settings changes
    EventManager.Subscribe EventType.SettingsChanged, Me, "HandleSettingsChanged"
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "DashboardManager.SubscribeToEvents", Err.Number, Err.Description
    End If
End Sub

'=========================================================================
' Interface Implementation - Dashboard Updates
'=========================================================================

' Refresh the entire dashboard
Private Sub IDashboardManager_RefreshDashboard()
    On Error GoTo ErrorHandler
    
    ' Ensure dashboard is initialized
    If Not m_IsInitialized Then IDashboardManager_Initialize
    
    ' Prevent refresh loops
    If m_IsRefreshing Then Exit Sub
    
    m_IsRefreshing = True
    
    ' Log refresh start
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "DashboardManager.RefreshDashboard", "Refreshing dashboard"
    End If
    
    ' Get updated date filter from UI
    UpdateDateFilterFromUI
    
    ' Performance optimizations
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Update all dashboard components
    IDashboardManager_UpdateBalanceSummary
    IDashboardManager_UpdateCategorySummary
    IDashboardManager_UpdateMonthlySpending
    IDashboardManager_UpdateDayOfWeekSpending
    
    ' Restore Excel settings
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' Trigger dashboard refreshed event
    EventManager.PublishEvent EventType.DashboardRefreshed
    
    ' Reset refreshing flag
    m_IsRefreshing = False
    
    ' Log refresh completion
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "DashboardManager.RefreshDashboard", "Dashboard refresh completed"
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Restore Excel settings even on error
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' Reset refreshing flag
    m_IsRefreshing = False
    
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "DashboardManager.RefreshDashboard", Err.Number, Err.Description
    End If
End Sub

' Public wrapper for RefreshDashboard
Public Sub RefreshDashboard()
    IDashboardManager_RefreshDashboard
End Sub

' Update the balance summary card
Private Sub IDashboardManager_UpdateBalanceSummary()
    On Error GoTo ErrorHandler
    
    ' Ensure dashboard is initialized
    If Not m_IsInitialized Then IDashboardManager_Initialize
    
    ' Check for dependencies
    If m_Repository Is Nothing Or m_Calculator Is Nothing Then
        If Not m_Logger Is Nothing Then
            m_Logger.LogWarning "DashboardManager.UpdateBalanceSummary", _
                "Cannot update balance summary - missing dependencies"
        End If
        Exit Sub
    End If
    
    ' Get transactions
    Dim transactions As Collection
    Set transactions = m_Repository.GetTransactions()
    
    ' Calculate balance summary
    Dim balanceSummary As Object
    Set balanceSummary = m_Calculator.CalculateBalanceSummary(transactions)
    
    ' Clear previous data
    m_DashboardSheet.Range(BALANCE_CARD_RANGE).Cells.ClearContents
    
    ' Add header
    m_DashboardSheet.Range("B4").Value = "Balance Summary"
    m_DashboardSheet.Range("B4").Font.Bold = True
    
    ' Update balance card
    With m_DashboardSheet
        .Range("B5").Value = "Total Spent:"
        .Range("C5").Value = balanceSummary("TotalSpent")
        .Range("C5").NumberFormat = "$#,##0.00"
        
        .Range("B6").Value = "Shared Expenses:"
        .Range("C6").Value = balanceSummary("TotalShared")
        .Range("C6").NumberFormat = "$#,##0.00"
        
        .Range("B7").Value = balanceSummary("User1") & " Spent:"
        .Range("C7").Value = balanceSummary("User1Spent")
        .Range("C7").NumberFormat = "$#,##0.00"
        
        .Range("B8").Value = balanceSummary("User2") & " Spent:"
        .Range("C8").Value = balanceSummary("User2Spent")
        .Range("C8").NumberFormat = "$#,##0.00"
    End With
    
    ' Display who owes whom
    If balanceSummary("OwedAmount") > 0 Then
        m_DashboardSheet.Range("D7:D8").Merge
        m_DashboardSheet.Range("D7").Value = balanceSummary("OwingUser") & " owes " & _
                                          balanceSummary("OwedUser") & " $" & _
                                          Format(balanceSummary("OwedAmount"), "#,##0.00")
        
        ' Highlight with a different color
        m_DashboardSheet.Range("D7").Interior.Color = RGB(255, 235, 156)
        m_DashboardSheet.Range("D7").Font.Bold = True
    Else
        m_DashboardSheet.Range("D7:D8").Merge
        m_DashboardSheet.Range("D7").Value = "Expenses are balanced"
        m_DashboardSheet.Range("D7").Interior.Color = RGB(198, 239, 206)
    End If
    
    ' Publish balance updated event with the summary data
    EventManager.PublishEventWithData EventType.BalanceUpdated, balanceSummary
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "DashboardManager.UpdateBalanceSummary", Err.Number, Err.Description
    End If
End Sub

' Public wrapper for UpdateBalanceSummary
Public Sub UpdateBalanceSummary()
    IDashboardManager_UpdateBalanceSummary
End Sub

' Update category summary and chart
Private Sub IDashboardManager_UpdateCategorySummary()
    On Error GoTo ErrorHandler
    
    ' Ensure dashboard is initialized
    If Not m_IsInitialized Then IDashboardManager_Initialize
    
    ' Check for dependencies
    If m_Repository Is Nothing Or m_Calculator Is Nothing Then
        If Not m_Logger Is Nothing Then
            m_Logger.LogWarning "DashboardManager.UpdateCategorySummary", _
                "Cannot update category summary - missing dependencies"
        End If
        Exit Sub
    End If
    
    ' Get transactions
    Dim transactions As Collection
    Set transactions = m_Repository.GetTransactions()
    
    ' Calculate category summary for the current date range
    Dim categorySummary As Object
    Set categorySummary = m_Calculator.CalculateCategorySummary(transactions, m_StartDate, m_EndDate)
    
    ' Clear any existing chart
    DeleteChartIfExists CATEGORY_CHART_NAME
    
    ' Create data for chart
    Dim chartData As Object
    Set chartData = CreateObject("Scripting.Dictionary")
    
    Dim category As Variant
    For Each category In categorySummary.Keys
        ' Only include categories, not metadata like percentages and totals
        If Right(category, 8) <> "_Percent" And category <> "Total" Then
            chartData.Add category, categorySummary(category)
        End If
    Next category
    
    ' Create the chart (Use correct method name)
    m_ChartFactory.CreateCategoryPieChart m_DashboardSheet, "B11:E24", chartData, "Spending by Category", CATEGORY_CHART_NAME
    
    ' Publish category data changed event with the summary data
    EventManager.PublishEventWithData EventType.CategoryDataChanged, categorySummary
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "DashboardManager.UpdateCategorySummary", Err.Number, Err.Description
    End If
End Sub

' Public wrapper for UpdateCategorySummary
Public Sub UpdateCategorySummary()
    IDashboardManager_UpdateCategorySummary
End Sub

' Update monthly spending chart
Private Sub IDashboardManager_UpdateMonthlySpending()
    On Error GoTo ErrorHandler
    
    ' Ensure dashboard is initialized
    If Not m_IsInitialized Then IDashboardManager_Initialize
    
    ' Check for dependencies
    If m_Repository Is Nothing Or m_Calculator Is Nothing Then
        If Not m_Logger Is Nothing Then
            m_Logger.LogWarning "DashboardManager.UpdateMonthlySpending", _
                "Cannot update monthly spending - missing dependencies"
        End If
        Exit Sub
    End If
    
    ' Get transactions
    Dim transactions As Collection
    Set transactions = m_Repository.GetTransactions()
    
    ' Calculate monthly spending
    Dim monthCount As Integer
    monthCount = DateDiff("m", m_StartDate, m_EndDate) + 1
    
    Dim monthlySummary As Object
    Set monthlySummary = m_Calculator.CalculateMonthSummary(transactions, monthCount)
    
    ' Clear any existing chart
    DeleteChartIfExists MONTHLY_CHART_NAME
    
    ' Create data for chart
    Dim chartData As Object
    Set chartData = CreateObject("Scripting.Dictionary")
    
    ' Get sorted keys (months)
    Dim months As Collection
    Set months = GetSortedKeys(monthlySummary)
    
    ' Add monthly expenses to chart data
    Dim month As Variant
    For Each month In months
        ' Format month for display
        Dim displayMonth As String
        displayMonth = MonthName(CInt(Mid(month, 6, 2)), True) & " " & Left(month, 4)
        
        ' Get expense value (positive number for chart)
        Dim expense As Double
        expense = monthlySummary(month & "_Expense")
        
        chartData.Add displayMonth, expense
    Next month
    
    ' Create the chart (Use correct method name)
    m_ChartFactory.CreateMonthlyColumnChart m_DashboardSheet, "F11:I24", chartData, _
                                    "Monthly Spending", MONTHLY_CHART_NAME
                                    ' Removed axis titles as they are handled within the factory method now
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "DashboardManager.UpdateMonthlySpending", Err.Number, Err.Description
    End If
End Sub

' Public wrapper for UpdateMonthlySpending
Public Sub UpdateMonthlySpending()
    IDashboardManager_UpdateMonthlySpending
End Sub

' Update day of week spending chart
Private Sub IDashboardManager_UpdateDayOfWeekSpending()
    On Error GoTo ErrorHandler
    
    ' Ensure dashboard is initialized
    If Not m_IsInitialized Then IDashboardManager_Initialize
    
    ' Check for dependencies
    If m_Repository Is Nothing Or m_Calculator Is Nothing Then
        If Not m_Logger Is Nothing Then
            m_Logger.LogWarning "DashboardManager.UpdateDayOfWeekSpending", _
                "Cannot update day of week spending - missing dependencies"
        End If
        Exit Sub
    End If
    
    ' Get transactions
    Dim transactions As Collection
    Set transactions = m_Repository.GetTransactions()
    
    ' Calculate day of week spending
    Dim dayOfWeekSummary As Object
    Set dayOfWeekSummary = m_Calculator.CalculateDayOfWeekSummary(transactions, m_StartDate, m_EndDate)
    
    ' Clear any existing chart
    DeleteChartIfExists DAILY_CHART_NAME
    
    ' Create data for chart
    Dim chartData As Object
    Set chartData = CreateObject("Scripting.Dictionary")
    
    ' Add days of week in correct order
    Dim i As Integer
    For i = 1 To 7
        Dim dayName As String
        dayName = WeekdayName(i, True)
        
        ' Skip adding total or percentage entries
        If dayOfWeekSummary.Exists(dayName) Then
            chartData.Add dayName, dayOfWeekSummary(dayName)
        End If
    Next i
    
    ' Create the chart (Use correct method name)
    m_ChartFactory.CreateDayOfWeekChart m_DashboardSheet, "B26:E38", chartData, _
                                    "Spending by Day of Week", DAILY_CHART_NAME
                                    ' Removed axis titles as they are handled within the factory method now
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "DashboardManager.UpdateDayOfWeekSpending", Err.Number, Err.Description
    End If
End Sub

' Public wrapper for UpdateDayOfWeekSpending
Public Sub UpdateDayOfWeekSpending()
    IDashboardManager_UpdateDayOfWeekSpending
End Sub

'=========================================================================
' Interface Implementation - Event Handlers
'=========================================================================

' Handle transaction data changes
Private Sub IDashboardManager_HandleTransactionsChanged()
    On Error GoTo ErrorHandler
    
    ' Log event handling
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "DashboardManager.HandleTransactionsChanged", _
            "Handling transactions changed event"
    End If
    
    ' Refresh the dashboard when transactions change
    IDashboardManager_RefreshDashboard
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "DashboardManager.HandleTransactionsChanged", Err.Number, Err.Description
    End If
End Sub

' Public wrapper for HandleTransactionsChanged
Public Sub HandleTransactionsChanged()
    IDashboardManager_HandleTransactionsChanged
End Sub

' Handle balance updates
Private Sub IDashboardManager_HandleBalanceUpdated(ByVal balanceData As Variant)
    On Error GoTo ErrorHandler
    
    ' This method is called when the balance is updated
    ' The dashboard already updates the balance display, so this is primarily for other
    ' UI components that might need to react to balance changes
    
    ' For now, just log the event
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "DashboardManager.HandleBalanceUpdated", _
            "Handling balance updated event"
    End If
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "DashboardManager.HandleBalanceUpdated", Err.Number, Err.Description
    End If
End Sub

' Public wrapper for HandleBalanceUpdated
Public Sub HandleBalanceUpdated(ByVal balanceData As Variant)
    IDashboardManager_HandleBalanceUpdated balanceData
End Sub

' Handle category data changes
Private Sub IDashboardManager_HandleCategoryDataChanged(ByVal categoryData As Variant)
    On Error GoTo ErrorHandler
    
    ' This method is called when category data changes
    ' The dashboard already updates the category display, so this is primarily for other
    ' UI components that might need to react to category changes
    
    ' For now, just log the event
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "DashboardManager.HandleCategoryDataChanged", _
            "Handling category data changed event"
    End If
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "DashboardManager.HandleCategoryDataChanged", Err.Number, Err.Description
    End If
End Sub

' Public wrapper for HandleCategoryDataChanged
Public Sub HandleCategoryDataChanged(ByVal categoryData As Variant)
    IDashboardManager_HandleCategoryDataChanged categoryData
End Sub

' Handle settings changes
Private Sub IDashboardManager_HandleSettingsChanged()
    On Error GoTo ErrorHandler
    
    ' Log event handling
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "DashboardManager.HandleSettingsChanged", _
            "Handling settings changed event"
    End If
    
    ' Refresh the dashboard when settings change
    IDashboardManager_RefreshDashboard
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "DashboardManager.HandleSettingsChanged", Err.Number, Err.Description
    End If
End Sub

' Public wrapper for HandleSettingsChanged
Public Sub HandleSettingsChanged()
    IDashboardManager_HandleSettingsChanged
End Sub

'=========================================================================
' Interface Implementation - Filter Methods
'=========================================================================

' Apply date filter
Private Sub IDashboardManager_ApplyDateFilter(ByVal startDate As Date, ByVal endDate As Date)
    On Error GoTo ErrorHandler
    
    ' Set new filter dates
    m_StartDate = startDate
    m_EndDate = endDate
    
    ' Update UI to reflect new dates
    With m_DashboardSheet
        .Range("H2").Value = startDate
        .Range("H3").Value = endDate
    End With
    
    ' Refresh the dashboard with new filter
    IDashboardManager_RefreshDashboard
    
    ' Publish filter applied event
    EventManager.PublishEvent EventType.FilterApplied
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "DashboardManager.ApplyDateFilter", Err.Number, Err.Description
    End If
End Sub

' Public wrapper for ApplyDateFilter
Public Sub ApplyDateFilter(ByVal startDate As Date, ByVal endDate As Date)
    IDashboardManager_ApplyDateFilter startDate, endDate
End Sub

' Get the current filter settings
Private Function IDashboardManager_GetCurrentFilter() As Object
    On Error GoTo ErrorHandler
    
    ' Create dictionary for filter settings
    Dim filter As Object
    Set filter = CreateObject("Scripting.Dictionary")
    
    ' Add date range
    filter.Add "StartDate", m_StartDate
    filter.Add "EndDate", m_EndDate
    
    Set IDashboardManager_GetCurrentFilter = filter
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "DashboardManager.GetCurrentFilter", Err.Number, Err.Description
    End If
    
    Set filter = CreateObject("Scripting.Dictionary")
    Set IDashboardManager_GetCurrentFilter = filter
End Function

' Public wrapper for GetCurrentFilter
Public Function GetCurrentFilter() As Object
    Set GetCurrentFilter = IDashboardManager_GetCurrentFilter()
End Function

'=========================================================================
' Helper Methods
'=========================================================================

' Delete a chart if it exists
Private Sub DeleteChartIfExists(ByVal chartName As String)
    On Error Resume Next
    
    Dim cht As ChartObject
    For Each cht In m_DashboardSheet.ChartObjects
        If cht.Name = chartName Then
            cht.Delete
            Exit Sub
        End If
    Next cht
    
    On Error GoTo 0
End Sub

' Update date filter from UI controls
Private Sub UpdateDateFilterFromUI()
    On Error Resume Next
    
    ' Get dates from UI
    Dim startDateValue As Variant
    Dim endDateValue As Variant
    
    startDateValue = m_DashboardSheet.Range("H2").Value
    endDateValue = m_DashboardSheet.Range("H3").Value
    
    ' Validate and update if valid
    If IsDate(startDateValue) And IsDate(endDateValue) Then
        m_StartDate = CDate(startDateValue)
        m_EndDate = CDate(endDateValue)
    End If
    
    On Error GoTo 0
End Sub

' Get sorted keys from a dictionary (for months)
Private Function GetSortedKeys(ByVal dict As Object) As Collection
    On Error GoTo ErrorHandler
    
    ' Create collection for results
    Dim result As New Collection
    
    ' Get all keys
    Dim keys() As Variant
    keys = dict.keys
    
    ' Sort keys using bubble sort
    Dim i As Long, j As Long
    Dim temp As Variant
    
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(i) > keys(j) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
    
    ' Add sorted keys to collection
    For i = LBound(keys) To UBound(keys)
        ' Only add keys that are month identifiers (yyyy-mm)
        If Len(keys(i)) = 7 And Mid$(keys(i), 5, 1) = "-" Then
            If Not (InStr(1, keys(i), "_") > 0) Then  ' Exclude derived keys
                result.Add keys(i)
            End If
        End If
    Next i
    
    Set GetSortedKeys = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "DashboardManager.GetSortedKeys", Err.Number, Err.Description
    End If
    
    Set result = New Collection
    Set GetSortedKeys = result
End Function

'=========================================================================
' Cleanup
'=========================================================================

Private Sub Class_Terminate()
    ' Unsubscribe from all events
    EventManager.UnsubscribeAll Me
    
    ' Clean up object references
    Set m_Repository = Nothing
    Set m_Calculator = Nothing
    Set m_ChartFactory = Nothing
    Set m_DashboardSheet = Nothing
    Set m_Logger = Nothing
End Sub
