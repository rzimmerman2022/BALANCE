VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChartFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' ChartFactory Class - Creates and manages charts for the dashboard

' Private variables
Private m_IsInitialized As Boolean
Private m_ChartsSheet As Worksheet
Private m_ChartSheet As Worksheet

' ===== Properties =====

' Charts sheet
Public Property Get ChartsSheet() As Worksheet
    If Not m_IsInitialized Then Initialize
    Set ChartsSheet = m_ChartsSheet
End Property

' ===== Methods =====

' Initialize the factory
Public Sub Initialize()
    On Error GoTo ErrorHandler
    
    ' Get or create the charts sheet
    Set m_ChartsSheet = Utilities.GetOrCreateSheet("Charts", False) ' Hidden sheet
    
    m_IsInitialized = True
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "ChartFactory.Initialize", Err.Number, Err.Description
End Sub

' Create a category pie chart
Public Function CreateCategoryPieChart(ws As Worksheet, left As Double, top As Double, _
                                       width As Double, height As Double, _
                                       title As String) As ChartObject
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    ' Clear existing data on the charts sheet
    Dim nextRow As Long
    nextRow = Utilities.GetLastRow(m_ChartsSheet, 1) + 2 ' Leave a row gap
    
    ' Get category data
    Dim balanceCalc As BalanceCalculator
    Set balanceCalc = BalanceCalculator
    
    Dim repo As TransactionRepository
    Set repo = TransactionRepository
    repo.Initialize
    
    Dim catSummary As Collection
    Set catSummary = balanceCalc.GetCategorySummary(repo.Transactions, 7) ' Top 7 categories
    
    ' Write data to the charts sheet
    m_ChartsSheet.Cells(nextRow, 1).Value = "Category"
    m_ChartsSheet.Cells(nextRow, 2).Value = "Amount"
    m_ChartsSheet.Cells(nextRow, 3).Value = "Percentage"
    
    m_ChartsSheet.Range(m_ChartsSheet.Cells(nextRow, 1), m_ChartsSheet.Cells(nextRow, 3)).Font.Bold = True
    
    Dim i As Long
    Dim cat As CategorySummary
    i = 1
    For Each cat In catSummary
        m_ChartsSheet.Cells(nextRow + i, 1).Value = cat.Category
        m_ChartsSheet.Cells(nextRow + i, 2).Value = cat.TotalAmount
        m_ChartsSheet.Cells(nextRow + i, 3).Value = cat.Percentage
        i = i + 1
    Next cat
    
    ' Create the chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(left, top, width, height)
    
    With chartObj.Chart
        ' Set chart type
        .ChartType = xlPie
        
        ' Set data source
        .SetSourceData m_ChartsSheet.Range(m_ChartsSheet.Cells(nextRow + 1, 1), m_ChartsSheet.Cells(nextRow + i - 1, 2))
        
        ' Format chart
        .HasTitle = True
        .ChartTitle.Text = title
        
        ' Format legend
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
        
        ' Format data labels
        .ApplyDataLabels xlDataLabelsShowPercent
        
        ' Apply a colorful theme
        .ApplyLayout 3 ' Colorful layout
        .ChartStyle = 26 ' Modern style with 3D effect
        
        ' Apply color to match BALANCE theme
        Dim j As Long
        For j = 1 To .SeriesCollection(1).Points.Count
            .SeriesCollection(1).Points(j).Format.Fill.ForeColor.RGB = GetChartColor(j)
        Next j
    End With
    
    Set CreateCategoryPieChart = chartObj
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "ChartFactory.CreateCategoryPieChart", Err.Number, Err.Description
    Set CreateCategoryPieChart = Nothing
End Function

' Create a monthly column chart
Public Function CreateMonthlyColumnChart(ws As Worksheet, left As Double, top As Double, _
                                         width As Double, height As Double, _
                                         title As String) As ChartObject
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    ' Clear existing data on the charts sheet
    Dim nextRow As Long
    nextRow = Utilities.GetLastRow(m_ChartsSheet, 1) + 2 ' Leave a row gap
    
    ' Get monthly data
    Dim balanceCalc As BalanceCalculator
    Set balanceCalc = BalanceCalculator
    
    Dim repo As TransactionRepository
    Set repo = TransactionRepository
    repo.Initialize
    
    Dim monthlySummary As Collection
    Set monthlySummary = balanceCalc.GetMonthlySummary(repo.Transactions, 6) ' Last 6 months
    
    ' Write data to the charts sheet
    m_ChartsSheet.Cells(nextRow, 1).Value = "Month"
    m_ChartsSheet.Cells(nextRow, 2).Value = "Net Amount" ' Changed Header
    m_ChartsSheet.Cells(nextRow, 3).Value = "Expenses"   ' Changed Header
    m_ChartsSheet.Cells(nextRow, 4).Value = "Income"     ' Added Header

    m_ChartsSheet.Range(m_ChartsSheet.Cells(nextRow, 1), m_ChartsSheet.Cells(nextRow, 4)).Font.Bold = True ' Adjusted range
    
    Dim i As Long
    Dim month As MonthSummary
    i = 1
    
    ' Process in reverse order to get chronological order
    Dim monthItems() As Variant
    ReDim monthItems(1 To monthlySummary.Count)
    For i = 1 To monthlySummary.Count
        Set monthItems(i) = monthlySummary(i)
    Next i
    
    ' Reverse the array
    For i = 1 To monthlySummary.Count
        Set month = monthItems(monthlySummary.Count - i + 1)
        ' Use MonthKey instead of MonthDate, format might need adjustment if parsing needed
        m_ChartsSheet.Cells(nextRow + i, 1).Value = month.MonthKey ' Use MonthKey
        m_ChartsSheet.Cells(nextRow + i, 2).Value = month.TotalAmount ' Net Amount
        m_ChartsSheet.Cells(nextRow + i, 3).Value = month.ExpenseAmount ' Expenses
        m_ChartsSheet.Cells(nextRow + i, 4).Value = month.IncomeAmount  ' Income
    Next i

    ' Create the chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(left, top, width, height)
    
    With chartObj.Chart
        ' Set chart type
        .ChartType = xlColumnClustered
        
        ' Set data source (Plot Expenses only for this example)
        .SetSourceData m_ChartsSheet.Range(m_ChartsSheet.Cells(nextRow, 1), _
                                          m_ChartsSheet.Cells(nextRow + monthlySummary.Count, 3)) ' Plot Month vs Expenses

        ' Format chart
        .HasTitle = True
        .ChartTitle.Text = title
        
        ' Format axes
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Expense Amount" ' Adjusted Title
        .Axes(xlValue).AxisTitle.Orientation = xlUpward

        ' Format data labels
        .HasDataTable = False
        
        ' Apply a colorful theme
        .ChartStyle = 48 ' Modern style
        
        ' Apply color to match BALANCE theme
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = AppSettings.ColorPrimary
    End With
    
    Set CreateMonthlyColumnChart = chartObj
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "ChartFactory.CreateMonthlyColumnChart", Err.Number, Err.Description
    Set CreateMonthlyColumnChart = Nothing
End Function

' Create a day of week column chart
Public Function CreateDayOfWeekChart(ws As Worksheet, left As Double, top As Double, _
                                    width As Double, height As Double, _
                                    title As String) As ChartObject
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    ' Clear existing data on the charts sheet
    Dim nextRow As Long
    nextRow = Utilities.GetLastRow(m_ChartsSheet, 1) + 2 ' Leave a row gap
    
    ' Get day of week data
    Dim balanceCalc As BalanceCalculator
    Set balanceCalc = BalanceCalculator
    
    Dim repo As TransactionRepository
    Set repo = TransactionRepository
    repo.Initialize
    
    Dim dowSummary As Collection
    Set dowSummary = balanceCalc.GetDayOfWeekSummary(repo.Transactions)
    
    ' Write data to the charts sheet
    m_ChartsSheet.Cells(nextRow, 1).Value = "Day"
    m_ChartsSheet.Cells(nextRow, 2).Value = "Amount"
    m_ChartsSheet.Cells(nextRow, 3).Value = "Percentage" ' Use Percentage instead of TransactionCount

    m_ChartsSheet.Range(m_ChartsSheet.Cells(nextRow, 1), m_ChartsSheet.Cells(nextRow, 3)).Font.Bold = True
    
    Dim i As Long
    Dim day As DayOfWeekSummary
    i = 1
    For Each day In dowSummary
        m_ChartsSheet.Cells(nextRow + i, 1).Value = day.DayName
        m_ChartsSheet.Cells(nextRow + i, 2).Value = day.TotalAmount
        m_ChartsSheet.Cells(nextRow + i, 3).Value = day.Percentage ' Use Percentage
        i = i + 1
    Next day
    
    ' Create the chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(left, top, width, height)
    
    With chartObj.Chart
        ' Set chart type
        .ChartType = xlColumnClustered
        
        ' Set data source
        .SetSourceData m_ChartsSheet.Range(m_ChartsSheet.Cells(nextRow, 1), _
                                          m_ChartsSheet.Cells(nextRow + 7, 2))
        
        ' Format chart
        .HasTitle = True
        .ChartTitle.Text = title
        
        ' Format axes
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Amount"
        .Axes(xlValue).AxisTitle.Orientation = xlUpward
        
        ' Format data labels
        .HasDataTable = False
        
        ' Apply a colorful theme
        .ChartStyle = 52 ' Modern style
        
        ' Apply colors to match BALANCE theme
        Dim j As Long
        For j = 1 To .SeriesCollection(1).Points.Count
            .SeriesCollection(1).Points(j).Format.Fill.ForeColor.RGB = GetChartColor(j)
        Next j
    End With
    
    Set CreateDayOfWeekChart = chartObj
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "ChartFactory.CreateDayOfWeekChart", Err.Number, Err.Description
    Set CreateDayOfWeekChart = Nothing
End Function

' Create a balance chart showing what each user owes
Public Function CreateBalanceChart(ws As Worksheet, left As Double, top As Double, _
                                  width As Double, height As Double, _
                                  title As String) As ChartObject
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    ' Clear existing data on the charts sheet
    Dim nextRow As Long
    nextRow = Utilities.GetLastRow(m_ChartsSheet, 1) + 2 ' Leave a row gap
    
    ' Get balance data
    Dim balanceCalc As BalanceCalculator
    Set balanceCalc = BalanceCalculator
    
    Dim repo As TransactionRepository
    Set repo = TransactionRepository
    repo.Initialize
    
    Dim balance As BalanceSummary
    balance = balanceCalc.CalculateBalance(repo.Transactions)
    
    ' Write data to the charts sheet - Adjusting to use available BalanceSummary data
    m_ChartsSheet.Cells(nextRow, 1).Value = "Metric"
    m_ChartsSheet.Cells(nextRow, 2).Value = "Value"

    m_ChartsSheet.Range(m_ChartsSheet.Cells(nextRow, 1), m_ChartsSheet.Cells(nextRow, 2)).Font.Bold = True

    m_ChartsSheet.Cells(nextRow + 1, 1).Value = "Net Balance"
    m_ChartsSheet.Cells(nextRow + 1, 2).Value = balance.NetBalance

    m_ChartsSheet.Cells(nextRow + 2, 1).Value = "Amount Owed"
    m_ChartsSheet.Cells(nextRow + 2, 2).Value = balance.OwedAmount

    ' Add who owes whom if available
    If Len(balance.WhoOwes) > 0 Then
        m_ChartsSheet.Cells(nextRow + 3, 1).Value = "Summary"
        m_ChartsSheet.Cells(nextRow + 3, 2).Value = balance.WhoOwes
    End If

    ' Create the chart - Needs adjustment as Bar chart needs categories/series
    ' For simplicity, let's just show the Owed Amount for now, maybe vs Net Balance
    ' Or perhaps better not to chart this directly if Paid/Share aren't available

    ' Option: Chart Owed Amount vs. 0 (less useful)
    m_ChartsSheet.Cells(nextRow + 5, 1).Value = "Owed" ' Chart Data Start
    m_ChartsSheet.Cells(nextRow + 5, 2).Value = balance.OwedAmount

    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(left, top, width, height)

    With chartObj.Chart
        .ChartType = xlColumnClustered ' Changed to Column for single value
        .SetSourceData m_ChartsSheet.Range(m_ChartsSheet.Cells(nextRow + 5, 1), m_ChartsSheet.Cells(nextRow + 5, 2))
        .HasLegend = False ' No legend needed for single series

        ' Format chart
        .HasTitle = True
        .ChartTitle.Text = title & " (Amount Owed)" ' Adjusted Title

        ' Format axes
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionNone ' Hide category axis label
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Amount"

        ' Format data labels
        .ApplyDataLabels xlDataLabelsShowValue

        ' Apply a colorful theme
        .ChartStyle = 48 ' Modern style

        ' Apply colors to match BALANCE theme
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = AppSettings.ColorPrimary
    End With

    ' Removed duplicated/commented out code block that was here previously
    
    Set CreateBalanceChart = chartObj
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "ChartFactory.CreateBalanceChart", Err.Number, Err.Description
    Set CreateBalanceChart = Nothing
End Function

' Create all charts on the Dashboard
Public Sub CreateAllDashboardCharts(dashboardSheet As Worksheet)
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    ' Clear existing charts
    Dim chartObj As ChartObject
    For Each chartObj In dashboardSheet.ChartObjects
        chartObj.Delete
    Next chartObj
    
    ' Create Category Pie Chart
    CreateCategoryPieChart dashboardSheet, 20, 100, 300, 200, "Spending by Category"
    
    ' Create Monthly Column Chart
    CreateMonthlyColumnChart dashboardSheet, 340, 100, 300, 200, "Monthly Expenses"
    
    ' Create Day of Week Chart
    CreateDayOfWeekChart dashboardSheet, 20, 320, 300, 200, "Spending by Day of Week"
    
    ' Create Balance Chart
    CreateBalanceChart dashboardSheet, 340, 320, 300, 200, "Balance Overview"
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "ChartFactory.CreateAllDashboardCharts", Err.Number, Err.Description
End Sub

' Get a color for chart elements by index
Private Function GetChartColor(index As Long) As Long
    On Error Resume Next
    
    Dim colors(1 To 10) As Long
    
    ' Define a set of colors for charts
    colors(1) = AppSettings.ColorPrimary
    colors(2) = AppSettings.ColorSecondary
    colors(3) = AppSettings.ColorAccent
    colors(4) = AppSettings.ColorInfo
    colors(5) = AppSettings.ColorSuccess
    colors(6) = AppSettings.ColorWarning
    colors(7) = AppSettings.ColorDanger
    colors(8) = RGB(128, 100, 162) ' Purple
    colors(9) = RGB(80, 172, 196)  ' Turquoise
    colors(10) = RGB(195, 141, 86) ' Brown
    
    ' Return color based on index (with cycling)
    GetChartColor = colors(((index - 1) Mod 10) + 1)
    
    On Error GoTo 0
End Function

' Class initialize
Private Sub Class_Initialize()
    ' Nothing needed here as we're using PredeclaredId = True
    ' and explicit initialization via Initialize method
    m_IsInitialized = False
End Sub

' Class terminate
Private Sub Class_Terminate()
    ' Clean up to prevent memory leaks
    Set m_ChartsSheet = Nothing
    Set m_ChartSheet = Nothing
End Sub
