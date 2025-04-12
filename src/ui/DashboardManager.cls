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
Private Const BALANCE_CARD_RANGE As String = "B4:D8" ' Example range
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
    
    ' Prevent re-initialization if already done
    If m_IsInitialized Then Exit Sub

    ' Set dependencies using ServiceLocator if not provided
    If repository Is Nothing Then
        Set m_Repository = ServiceLocator.GetTransactionRepository()
    Else
        Set m_Repository = repository
    End If
    
    If calculator Is Nothing Then
        Set m_Calculator = ServiceLocator.GetBalanceCalculator()
    Else
        Set m_Calculator = calculator
    End If
    
    If logger Is Nothing Then
        Set m_Logger = ServiceLocator.GetErrorLogger()
    Else
        Set m_Logger = logger
    End If
    
    ' Get or create the dashboard sheet
    Set m_DashboardSheet = GetOrCreateDashboardSheet()
    If m_DashboardSheet Is Nothing Then Err.Raise vbObjectError + 517, "IDashboardManager_Initialize", "Failed to get or create Dashboard sheet."

    ' Initialize chart factory (Use ServiceLocator or PredeclaredId instance)
    Set m_ChartFactory = ChartFactory ' Assuming ChartFactory has PredeclaredId=True
    m_ChartFactory.Initialize ' ChartFactory Initialize might take logger if needed

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
    Else
        Debug.Print "ERROR in DashboardManager.Initialize: " & Err.Number & " - " & Err.Description
    End If
    ' Optionally re-raise or handle error appropriately
End Sub

' Public wrapper for Initialize
Public Sub Initialize(Optional ByVal repository As ITransactionRepository = Nothing, _
                      Optional ByVal calculator As IBalanceCalculator = Nothing, _
                      Optional ByVal logger As IErrorLogger = Nothing)
    IDashboardManager_Initialize repository, calculator, logger
End Sub

' ***** ADDED PUBLIC PROPERTY GET *****
' Expose the dashboard sheet object
Public Property Get DashboardSheet() As Worksheet
    ' Ensure manager is initialized before returning sheet
    If Not m_IsInitialized Then Me.Initialize
    Set DashboardSheet = m_DashboardSheet
End Property
' ***** END ADDED PUBLIC PROPERTY GET *****


' Get or create the dashboard sheet
Private Function GetOrCreateDashboardSheet() As Worksheet
    On Error Resume Next ' Use Resume Next carefully for object checking
    
    Dim ws As Worksheet
    
    ' Try to get existing sheet
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET_NAME)
    
    ' If sheet doesn't exist, create it
    If ws Is Nothing Then
        Err.Clear ' Clear any error from trying to get the sheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)) ' Add as last sheet
        If Err.Number <> 0 Then GoTo HandleSheetError ' Check if Add failed
        
        On Error Resume Next ' Handle potential error during rename
        ws.Name = DASHBOARD_SHEET_NAME
        If Err.Number <> 0 Then GoTo HandleSheetError ' Check if rename failed
        On Error GoTo 0 ' Restore default error handling
        
        ' Set up basic dashboard structure
        SetupDashboardLayout ws ' Call layout setup for new sheet
                
        If Not m_Logger Is Nothing Then
             m_Logger.LogInfo "DashboardManager.GetOrCreateDashboardSheet", "Created new sheet: " & DASHBOARD_SHEET_NAME
        End If
    End If
    
    ' Ensure sheet is visible
    If Not ws Is Nothing Then ws.Visible = xlSheetVisible
    
    Set GetOrCreateDashboardSheet = ws
    Exit Function

HandleSheetError:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "DashboardManager.GetOrCreateDashboardSheet", Err.Number, "Failed to create/rename dashboard sheet: " & Err.Description
    Else
        Debug.Print "ERROR in DashboardManager.GetOrCreateDashboardSheet: " & Err.Number & " - " & Err.Description
    End If
    Set GetOrCreateDashboardSheet = Nothing ' Return Nothing on error
    On Error GoTo 0 ' Clear error context
End Function


' Set up the basic dashboard layout
Private Sub SetupDashboardLayout(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Clear the sheet
    ws.Cells.Clear
    Utilities.ClearSheet ws ' Also remove shapes if Utilities function does that

    ' Apply base formatting
    Utilities.ApplyModernFormatting ws
    
    ' Add title
    With ws.Range("B2")
        .Value = AppSettings.AppTitle & " - Financial Dashboard"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = AppSettings.ColorDark
    End With
    
    ' Add date filter controls
    ws.Range("G2").Value = "Start Date:"
    ws.Range("G3").Value = "End Date:"
    ws.Range("G2:G3").Font.Bold = True
    ws.Range("G2:G3").HorizontalAlignment = xlRight
    
    ws.Range("H2").Value = m_StartDate
    ws.Range("H2").NumberFormat = "yyyy-mm-dd" ' Use consistent format
    
    ws.Range("H3").Value = m_EndDate
    ws.Range("H3").NumberFormat = "yyyy-mm-dd" ' Use consistent format
    ws.Range("H2:H3").HorizontalAlignment = xlLeft

    ' Add Buttons using Utility
    Dim btnTop As Double: btnTop = ws.Range("G2").Top
    Dim btnLeft As Double: btnLeft = ws.Range("I2").Left + 20

    Utilities.AddButton ws, btnLeft, btnTop, 100, 25, "Refresh", "BALANCE.RefreshData", AppSettings.ColorPrimary
    btnTop = btnTop + 30 ' Move down for next button
    Utilities.AddButton ws, btnLeft, btnTop, 100, 25, "Import CSV", "BALANCE.ShowImportDialog", AppSettings.ColorSecondary
    btnTop = btnTop + 30
    Utilities.AddButton ws, btnLeft, btnTop, 100, 25, "Transactions", "BALANCE.ShowTransactionList", AppSettings.ColorInfo
    btnTop = btnTop + 30
    Utilities.AddButton ws, btnLeft, btnTop, 100, 25, "Insights", "BALANCE.ShowInsights", AppSettings.ColorAccent
    btnTop = btnTop + 30
    Utilities.AddButton ws, btnLeft, btnTop, 100, 25, "Settings", "BALANCE.ShowSettings", AppSettings.ColorNeutral ' Use Neutral for Settings
    ' Adjust button colors if needed

    ' Add placeholders for cards/charts - actual creation happens during Refresh
    ws.Range("B4").Value = "Balance Summary (Loading...)"
    ws.Range("B4").Font.Bold = True
    ws.Range("B4").Font.Size = 12
    
    ws.Range("B12").Value = "Spending by Category (Loading...)"
    ws.Range("B12").Font.Bold = True
    ws.Range("B12").Font.Size = 12
    
    ws.Range("F12").Value = "Monthly Spending (Loading...)"
    ws.Range("F12").Font.Bold = True
    ws.Range("F12").Font.Size = 12
    
    ws.Range("B27").Value = "Day of Week Spending (Loading...)"
    ws.Range("B27").Font.Bold = True
    ws.Range("B27").Font.Size = 12

    ' Deselect
    ws.Range("A1").Select
    
    Exit Sub
ErrorHandler:
     If Not m_Logger Is Nothing Then
        m_Logger.LogError "DashboardManager.SetupDashboardLayout", Err.Number, Err.Description
    End If
End Sub


' Subscribe to events from data sources
Private Sub SubscribeToEvents()
    On Error GoTo ErrorHandler
    
    ' Subscribe to transaction changes
    EventManager.Subscribe EventType.TransactionsChanged, Me, "HandleTransactionsChanged"
    
    ' Subscribe to balance updates (Maybe not needed if dashboard triggers its own updates)
    ' EventManager.Subscribe EventType.BalanceUpdated, Me, "HandleBalanceUpdated"
    
    ' Subscribe to category data changes (Maybe not needed)
    ' EventManager.Subscribe EventType.CategoryDataChanged, Me, "HandleCategoryDataChanged"
    
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
    If Not m_IsInitialized Then Me.Initialize
    
    ' Prevent refresh loops
    If m_IsRefreshing Then Exit Sub
    m_IsRefreshing = True ' Set flag
    
    ' Log refresh start
    If Not m_Logger Is Nothing Then m_Logger.LogInfo "DashboardManager.RefreshDashboard", "Refreshing dashboard"
    
    ' Ensure sheet object is valid
    If m_DashboardSheet Is Nothing Then
        Set m_DashboardSheet = GetOrCreateDashboardSheet()
        If m_DashboardSheet Is Nothing Then Err.Raise vbObjectError + 517, "RefreshDashboard", "Dashboard sheet is not available."
    End If
    
    ' Get updated date filter from UI
    UpdateDateFilterFromUI
    
    ' Performance optimizations
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Update all dashboard components
    Me.UpdateBalanceSummary ' Use Me to call public wrapper which calls interface impl
    Me.UpdateCategorySummary
    Me.UpdateMonthlySpending
    Me.UpdateDayOfWeekSpending
    ' Add calls to update other potential dashboard elements here
    
    ' Restore Excel settings
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' Trigger dashboard refreshed event
    EventManager.PublishEvent EventType.DashboardRefreshed
    
    ' Reset refreshing flag
    m_IsRefreshing = False
    
    ' Log refresh completion
    If Not m_Logger Is Nothing Then m_Logger.LogInfo "DashboardManager.RefreshDashboard", "Dashboard refresh completed"
    
    Exit Sub
    
ErrorHandler:
    ' Restore Excel settings even on error
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ' Reset refreshing flag
    m_IsRefreshing = False
    If Not m_Logger Is Nothing Then m_Logger.LogError "DashboardManager.RefreshDashboard", Err.Number, Err.Description
End Sub


' Public wrapper for RefreshDashboard
Public Sub RefreshDashboard()
    IDashboardManager_RefreshDashboard
End Sub

' Update the balance summary card
Private Sub IDashboardManager_UpdateBalanceSummary()
    On Error GoTo ErrorHandler
    
    ' Ensure dashboard is initialized and dependencies are met
    If Not m_IsInitialized Then Me.Initialize
    If m_Repository Is Nothing Or m_Calculator Is Nothing Or m_DashboardSheet Is Nothing Then
        If Not m_Logger Is Nothing Then m_Logger.LogWarning "DashboardManager.UpdateBalanceSummary", "Cannot update - missing dependencies or sheet."
        Exit Sub
    End If
    
    ' Get transactions for the current date range
    Dim transactions As Collection
    Set transactions = m_Repository.GetTransactionsByDateRange(m_StartDate, m_EndDate) ' Filter by date
    
    ' Calculate balance summary
    Dim balanceSummary As Object ' Dictionary
    Set balanceSummary = m_Calculator.CalculateBalanceSummary(transactions)
    
    ' Clear previous card content (adjust range if layout changed)
    m_DashboardSheet.Range("B5:D8").ClearContents ' Clear values, keep formatting/borders if any

    ' Update balance card using dictionary keys
    With m_DashboardSheet
        .Range("B4").Value = "Balance Summary (" & Format(m_StartDate, "yyyy-mm-dd") & " to " & Format(m_EndDate, "yyyy-mm-dd") & ")" ' Add date range
        .Range("B4").Font.Bold = True

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

        ' Align text
        .Range("B5:B8").HorizontalAlignment = xlRight
        .Range("C5:C8").HorizontalAlignment = xlLeft
    End With
    
    ' Display who owes whom (Merge cells for better display)
    On Error Resume Next ' Ignore error if already merged
    m_DashboardSheet.Range("D5:D8").Merge
    m_DashboardSheet.Range("D5:D8").ClearContents ' Clear previous merged value
    On Error GoTo ErrorHandler ' Restore default error handling

    With m_DashboardSheet.Range("D5") ' Write to top-left cell of merged area
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .WrapText = True
        If balanceSummary("OwedAmount") > 0.01 Then ' Use small tolerance
            .Value = balanceSummary("OwingUser") & " owes " & _
                     balanceSummary("OwedUser") & vbCrLf & _
                     Format(balanceSummary("OwedAmount"), "$#,##0.00")
            .Interior.Color = AppSettings.ColorWarning ' Yellow for owing
            .Font.Bold = True
        Else
            .Value = "Expenses Balanced"
            .Interior.Color = AppSettings.ColorSuccess ' Green for balanced
            .Font.Bold = True
        End If
    End With
    
    ' Publish balance updated event with the summary data
    EventManager.PublishEventWithData EventType.BalanceUpdated, balanceSummary
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "DashboardManager.UpdateBalanceSummary", Err.Number, Err.Description
End Sub

' Public wrapper for UpdateBalanceSummary
Public Sub UpdateBalanceSummary()
    IDashboardManager_UpdateBalanceSummary
End Sub


' Update category summary and chart
Private Sub IDashboardManager_UpdateCategorySummary()
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Me.Initialize
    If m_Repository Is Nothing Or m_Calculator Is Nothing Or m_ChartFactory Is Nothing Or m_DashboardSheet Is Nothing Then
         If Not m_Logger Is Nothing Then m_Logger.LogWarning "DashboardManager.UpdateCategorySummary", "Cannot update - missing dependencies or sheet."
        Exit Sub
    End If
    
    ' Get transactions for the current date range
    Dim transactions As Collection
    Set transactions = m_Repository.GetTransactionsByDateRange(m_StartDate, m_EndDate) ' Filter by date
    
    ' Calculate category summary
    Dim categorySummaryDict As Object ' Dictionary
    Set categorySummaryDict = m_Calculator.CalculateCategorySummary(transactions, m_StartDate, m_EndDate) ' Pass dates
    
    ' Clear any existing chart by name
    DeleteChartIfExists CATEGORY_CHART_NAME
    
    ' Prepare data for chart (Dictionary: Category -> Amount)
    Dim chartData As Object
    Set chartData = CreateObject("Scripting.Dictionary")
    
    Dim categoryKey As Variant
    For Each categoryKey In categorySummaryDict.Keys
        ' Only include categories, not metadata like percentages and totals
        If Right(categoryKey, 8) <> "_Percent" And categoryKey <> "Total" Then
            If categorySummaryDict(categoryKey) > 0 Then ' Only chart categories with spending
                 chartData.Add categoryKey, categorySummaryDict(categoryKey)
            End If
        End If
    Next categoryKey
    
    ' Create the chart using ChartFactory
    If chartData.Count > 0 Then
        m_ChartFactory.CreateCategoryPieChart m_DashboardSheet, "B13:E26", chartData, "Spending by Category", CATEGORY_CHART_NAME ' Example Range
    Else
         m_DashboardSheet.Range("B13").Value = "No category spending data for selected period." ' Placeholder text
    End If

    ' Publish category data changed event with the summary data
    EventManager.PublishEventWithData EventType.CategoryDataChanged, categorySummaryDict
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "DashboardManager.UpdateCategorySummary", Err.Number, Err.Description
End Sub

' Public wrapper for UpdateCategorySummary
Public Sub UpdateCategorySummary()
    IDashboardManager_UpdateCategorySummary
End Sub

' Update monthly spending chart
Private Sub IDashboardManager_UpdateMonthlySpending()
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Me.Initialize
    If m_Repository Is Nothing Or m_Calculator Is Nothing Or m_ChartFactory Is Nothing Or m_DashboardSheet Is Nothing Then
         If Not m_Logger Is Nothing Then m_Logger.LogWarning "DashboardManager.UpdateMonthlySpending", "Cannot update - missing dependencies or sheet."
        Exit Sub
    End If
    
    ' Get transactions for the current date range (or maybe use all for trends?)
    ' For consistency with date filter, let's filter, but maybe trend needs all data?
    Dim transactions As Collection
    Set transactions = m_Repository.GetTransactionsByDateRange(m_StartDate, m_EndDate) ' Filter by date

    ' Calculate monthly spending based on filtered range
    Dim monthCount As Integer
    monthCount = DateDiff("m", m_StartDate, m_EndDate) + 1
    If monthCount <= 0 Then monthCount = 1

    Dim monthlySummaryDict As Object ' Dictionary
    Set monthlySummaryDict = m_Calculator.CalculateMonthSummary(transactions, monthCount) ' Pass filtered transactions

    ' Clear any existing chart
    DeleteChartIfExists MONTHLY_CHART_NAME
    
    ' Prepare data for chart (Dictionary: Formatted Month -> Expense Amount)
    Dim chartData As Object
    Set chartData = CreateObject("Scripting.Dictionary")
    
    Dim months As Collection
    Set months = Utilities.GetSortedKeys(monthlySummaryDict) ' Use Utility helper
    
    Dim monthKey As Variant
    For Each monthKey In months ' Iterate sorted keys
        ' Format month for display (e.g., "Apr 2025")
        Dim displayMonth As String
        On Error Resume Next ' Handle potential error parsing monthKey
        displayMonth = MonthName(CInt(Mid(monthKey, 6, 2)), True) & " " & Left(monthKey, 4)
        On Error GoTo ErrorHandler ' Restore error handling
        If displayMonth = "" Then displayMonth = monthKey ' Fallback

        ' Get expense value (positive number for chart)
        Dim expense As Double
        expense = monthlySummaryDict(monthKey & "_Expense")
        
        chartData.Add displayMonth, expense ' Use formatted month as key
    Next monthKey
    
    ' Create the chart using ChartFactory
    If chartData.Count > 0 Then
        m_ChartFactory.CreateMonthlyColumnChart m_DashboardSheet, "F13:I26", chartData, "Monthly Expenses", MONTHLY_CHART_NAME ' Example Range
    Else
        m_DashboardSheet.Range("F13").Value = "No monthly spending data for selected period." ' Placeholder text
    End If

    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "DashboardManager.UpdateMonthlySpending", Err.Number, Err.Description
End Sub


' Public wrapper for UpdateMonthlySpending
Public Sub UpdateMonthlySpending()
    IDashboardManager_UpdateMonthlySpending
End Sub


' Update day of week spending chart
Private Sub IDashboardManager_UpdateDayOfWeekSpending()
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Me.Initialize
    If m_Repository Is Nothing Or m_Calculator Is Nothing Or m_ChartFactory Is Nothing Or m_DashboardSheet Is Nothing Then
        If Not m_Logger Is Nothing Then m_Logger.LogWarning "DashboardManager.UpdateDayOfWeekSpending", "Cannot update - missing dependencies or sheet."
        Exit Sub
    End If
    
    ' Get transactions for the current date range
    Dim transactions As Collection
    Set transactions = m_Repository.GetTransactionsByDateRange(m_StartDate, m_EndDate) ' Filter by date
    
    ' Calculate day of week spending
    Dim dayOfWeekSummaryDict As Object ' Dictionary
    Set dayOfWeekSummaryDict = m_Calculator.CalculateDayOfWeekSummary(transactions, m_StartDate, m_EndDate) ' Pass dates
    
    ' Clear any existing chart
    DeleteChartIfExists DAILY_CHART_NAME
    
    ' Prepare data for chart (Dictionary: Day Name -> Amount)
    Dim chartData As Object
    Set chartData = CreateObject("Scripting.Dictionary")
    
    ' Add days of week in correct order
    Dim i As Integer
    For i = vbSunday To vbSaturday ' Use VBA constants for days
        Dim dayName As String
        dayName = WeekdayName(i, True) ' Abbreviated day name
        
        ' Check if day exists in summary and has spending
        If dayOfWeekSummaryDict.Exists(dayName) Then
             If dayOfWeekSummaryDict(dayName) > 0 Then
                chartData.Add dayName, dayOfWeekSummaryDict(dayName)
             End If
        End If
    Next i
    
    ' Create the chart using ChartFactory
    If chartData.Count > 0 Then
       m_ChartFactory.CreateDayOfWeekChart m_DashboardSheet, "B28:E38", chartData, "Spending by Day of Week", DAILY_CHART_NAME ' Example Range
    Else
        m_DashboardSheet.Range("B28").Value = "No day of week spending data for selected period." ' Placeholder text
    End If
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "DashboardManager.UpdateDayOfWeekSpending", Err.Number, Err.Description
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
    
    If Not m_IsInitialized Then Exit Sub ' Don't handle if not ready
    If m_IsRefreshing Then Exit Sub ' Avoid loops if called during a refresh
    
    If Not m_Logger Is Nothing Then m_Logger.LogInfo "DashboardManager.HandleTransactionsChanged", "Handling transactions changed event"
    
    ' Refresh the dashboard when transactions change
    Me.RefreshDashboard ' Use public wrapper
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "DashboardManager.HandleTransactionsChanged", Err.Number, Err.Description
End Sub


' Public wrapper for HandleTransactionsChanged
Public Sub HandleTransactionsChanged()
    IDashboardManager_HandleTransactionsChanged
End Sub

' Handle balance updates (Potentially redundant if RefreshDashboard covers it)
Private Sub IDashboardManager_HandleBalanceUpdated(ByVal balanceData As Variant)
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Exit Sub
    If m_IsRefreshing Then Exit Sub

    If Not m_Logger Is Nothing Then m_Logger.LogInfo "DashboardManager.HandleBalanceUpdated", "Handling balance updated event"
    
    ' Optionally, could perform a more targeted update here instead of full refresh
    ' Me.UpdateBalanceSummary ' Example: Update only the balance card

    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "DashboardManager.HandleBalanceUpdated", Err.Number, Err.Description
End Sub

' Public wrapper for HandleBalanceUpdated
Public Sub HandleBalanceUpdated(ByVal balanceData As Variant)
    IDashboardManager_HandleBalanceUpdated balanceData
End Sub

' Handle category data changes (Potentially redundant)
Private Sub IDashboardManager_HandleCategoryDataChanged(ByVal categoryData As Variant)
    On Error GoTo ErrorHandler

    If Not m_IsInitialized Then Exit Sub
    If m_IsRefreshing Then Exit Sub
    
    If Not m_Logger Is Nothing Then m_Logger.LogInfo "DashboardManager.HandleCategoryDataChanged", "Handling category data changed event"

    ' Optionally, update only category chart
    ' Me.UpdateCategorySummary

    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "DashboardManager.HandleCategoryDataChanged", Err.Number, Err.Description
End Sub

' Public wrapper for HandleCategoryDataChanged
Public Sub HandleCategoryDataChanged(ByVal categoryData As Variant)
    IDashboardManager_HandleCategoryDataChanged categoryData
End Sub

' Handle settings changes
Private Sub IDashboardManager_HandleSettingsChanged()
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Exit Sub
    If m_IsRefreshing Then Exit Sub
    
    If Not m_Logger Is Nothing Then m_Logger.LogInfo "DashboardManager.HandleSettingsChanged", "Handling settings changed event"
    
    ' Refresh the dashboard when settings change (user names might affect balance display)
    Me.RefreshDashboard
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "DashboardManager.HandleSettingsChanged", Err.Number, Err.Description
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
    
    If Not m_IsInitialized Then Me.Initialize ' Ensure initialized

    ' Validate dates
    If startDate > endDate Then
        MsgBox "Start date cannot be after end date.", vbExclamation, AppSettings.AppTitle
        Exit Sub
    End If

    ' Set new filter dates
    m_StartDate = startDate
    m_EndDate = endDate
    
    ' Update UI to reflect new dates (if sheet exists)
    If Not m_DashboardSheet Is Nothing Then
        On Error Resume Next ' Ignore error if sheet/cells don't exist yet
        m_DashboardSheet.Range("H2").Value = startDate
        m_DashboardSheet.Range("H3").Value = endDate
        On Error GoTo ErrorHandler ' Restore error handling
    End If
    
    ' Refresh the dashboard with new filter
    Me.RefreshDashboard
    
    ' Publish filter applied event
    EventManager.PublishEvent EventType.FilterApplied
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "DashboardManager.ApplyDateFilter", Err.Number, Err.Description
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
    If Not m_Logger Is Nothing Then m_Logger.LogError "DashboardManager.GetCurrentFilter", Err.Number, Err.Description
    
    Set filter = CreateObject("Scripting.Dictionary") ' Return empty dictionary on error
    Set IDashboardManager_GetCurrentFilter = filter
End Function

' Public wrapper for GetCurrentFilter
Public Function GetCurrentFilter() As Object
    Set GetCurrentFilter = IDashboardManager_GetCurrentFilter()
End Function

'=========================================================================
' Helper Methods
'=========================================================================

' Delete a chart if it exists by name
Private Sub DeleteChartIfExists(ByVal chartName As String)
    On Error Resume Next ' Ignore error if sheet or chart doesn't exist
    
    If m_DashboardSheet Is Nothing Then Exit Sub

    Dim chtObj As ChartObject ' Use ChartObject
    Set chtObj = m_DashboardSheet.ChartObjects(chartName)
    If Not chtObj Is Nothing Then
        chtObj.Delete
    End If
    Set chtObj = Nothing
    Err.Clear ' Clear any error from accessing non-existent chart
    On Error GoTo 0 ' Restore default error handling
End Sub


' Update date filter from UI controls
Private Sub UpdateDateFilterFromUI()
    On Error GoTo ErrorHandler ' Use GoTo for better control
    
    If m_DashboardSheet Is Nothing Then Exit Sub ' Cannot update if sheet doesn't exist

    ' Get dates from UI
    Dim startDateValue As Variant
    Dim endDateValue As Variant
    
    startDateValue = m_DashboardSheet.Range("H2").Value
    endDateValue = m_DashboardSheet.Range("H3").Value
    
    ' Validate and update if valid and different from current
    If IsDate(startDateValue) And IsDate(endDateValue) Then
         If CDate(startDateValue) <= CDate(endDateValue) Then ' Basic validation
            If CDate(startDateValue) <> m_StartDate Or CDate(endDateValue) <> m_EndDate Then
                m_StartDate = CDate(startDateValue)
                m_EndDate = CDate(endDateValue)
                If Not m_Logger Is Nothing Then m_Logger.LogInfo "UpdateDateFilterFromUI", "Date filter updated from UI: " & m_StartDate & " to " & m_EndDate
            End If
         Else
             ' Restore UI to previous valid dates if input is invalid
             m_DashboardSheet.Range("H2").Value = m_StartDate
             m_DashboardSheet.Range("H3").Value = m_EndDate
             If Not m_Logger Is Nothing Then m_Logger.LogWarning "UpdateDateFilterFromUI", "Invalid date range entered in UI. Start date was after end date."
         End If
    Else
        ' Restore UI to previous valid dates if input is invalid
        m_DashboardSheet.Range("H2").Value = m_StartDate
        m_DashboardSheet.Range("H3").Value = m_EndDate
        If Not m_Logger Is Nothing Then m_Logger.LogWarning "UpdateDateFilterFromUI", "Invalid date value entered in UI."
    End If
    
    Exit Sub
ErrorHandler:
     If Not m_Logger Is Nothing Then m_Logger.LogError "DashboardManager.UpdateDateFilterFromUI", Err.Number, Err.Description
     ' Attempt to restore UI dates on error
     On Error Resume Next
     If Not m_DashboardSheet Is Nothing Then
        m_DashboardSheet.Range("H2").Value = m_StartDate
        m_DashboardSheet.Range("H3").Value = m_EndDate
     End If
     On Error GoTo 0

End Sub

' Get sorted keys from a dictionary (Assumed Utility exists)
' This function is duplicated in BalanceCalculator - consider moving to Utilities
' Private Function GetSortedKeys(ByVal dict As Object) As Collection ...

'=========================================================================
' Cleanup
'=========================================================================

Private Sub Class_Terminate()
    ' Unsubscribe from all events to prevent errors if manager is destroyed
    On Error Resume Next ' Ignore errors during cleanup
    EventManager.UnsubscribeAll Me
    On Error GoTo 0
    
    ' Clean up object references
    Set m_Repository = Nothing
    Set m_Calculator = Nothing
    Set m_ChartFactory = Nothing
    Set m_DashboardSheet = Nothing
    Set m_Logger = Nothing
End Sub