VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransactionAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' TransactionAnalyzer Class - Provides advanced analytics for transactions

' Analysis result type
Public Type AnalysisInsight
    Category As String
    Title As String
    Description As String
    Value As Variant
    Trend As String ' "Up", "Down", "Neutral"
    Importance As Integer ' 1 (Low) to 5 (High)
End Type

' Private variables
Private m_IsInitialized As Boolean
Private m_InsightsSheet As Worksheet

' ===== Properties =====

' Insights sheet
Public Property Get InsightsSheet() As Worksheet
    If Not m_IsInitialized Then Initialize
    Set InsightsSheet = m_InsightsSheet
End Property

' ===== Methods =====

' Initialize the analyzer
Public Sub Initialize()
    On Error GoTo ErrorHandler
    
    ' Get or create the insights sheet
    Set m_InsightsSheet = Utilities.GetOrCreateSheet("Insights", True)
    
    ' Apply modern formatting
    Utilities.ApplyModernFormatting m_InsightsSheet
    
    m_IsInitialized = True
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "TransactionAnalyzer.Initialize", Err.Number, Err.Description
End Sub

' Generate spending insights
Public Function GenerateSpendingInsights() As Collection
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    Dim insights As New Collection
    
    ' Get repository and calculator
    Dim repo As TransactionRepository
    Set repo = TransactionRepository
    repo.Initialize
    
    Dim balanceCalc As BalanceCalculator
    Set balanceCalc = BalanceCalculator
    
    ' Get all transactions
    Dim transactions As Collection
    Set transactions = repo.Transactions
    
    ' Skip if we don't have enough data
    If transactions.Count < 10 Then
        ' Add placeholder insight
        Dim notEnough As AnalysisInsight
        notEnough.Category = "Overview"
        notEnough.Title = "Not enough data"
        notEnough.Description = "Need at least 10 transactions for insights"
        notEnough.Value = transactions.Count
        notEnough.Trend = "Neutral"
        notEnough.Importance = 1
        
        insights.Add notEnough
        Set GenerateSpendingInsights = insights
        Exit Function
    End If
    
    ' 1. Top spending category
    Dim catSummary As Collection
    Set catSummary = balanceCalc.GetCategorySummary(transactions, 10)
    
    If catSummary.Count > 0 Then
        Dim topCategory As CategorySummary
        Set topCategory = catSummary(1)
        
        Dim catInsight As AnalysisInsight
        catInsight.Category = "Categories"
        catInsight.Title = "Top spending category"
        catInsight.Description = topCategory.Category & " accounts for " & _
                              Format(topCategory.Percentage, "0.0%") & " of total spending"
        catInsight.Value = topCategory.TotalAmount
        catInsight.Trend = "Neutral"
        catInsight.Importance = 3
        
        insights.Add catInsight
    End If
    
    ' 2. Monthly trend
    Dim monthlySummary As Collection
    Set monthlySummary = balanceCalc.GetMonthlySummary(transactions, 3)
    
    If monthlySummary.Count >= 2 Then
        Dim currentMonth As MonthSummary
        Dim previousMonth As MonthSummary
        
        Set currentMonth = monthlySummary(1)
        Set previousMonth = monthlySummary(2)
        
        Dim monthlyChange As Double
        If previousMonth.TotalAmount > 0 Then
            monthlyChange = (currentMonth.TotalAmount - previousMonth.TotalAmount) / previousMonth.TotalAmount
        Else
            monthlyChange = 0
        End If
        
        Dim monthTrend As String
        If monthlyChange > 0.1 Then
            monthTrend = "Up"
        ElseIf monthlyChange < -0.1 Then
            monthTrend = "Down"
        Else
            monthTrend = "Neutral"
        End If
        
        Dim monthInsight As AnalysisInsight
        monthInsight.Category = "Trends"
        monthInsight.Title = "Monthly spending trend"
        
        If monthlyChange > 0 Then
            monthInsight.Description = "Spending increased by " & Format(Abs(monthlyChange), "0.0%") & _
                                     " compared to last month"
        ElseIf monthlyChange < 0 Then
            monthInsight.Description = "Spending decreased by " & Format(Abs(monthlyChange), "0.0%") & _
                                     " compared to last month"
        Else
            monthInsight.Description = "Spending is stable compared to last month"
        End If
        
        monthInsight.Value = monthlyChange
        monthInsight.Trend = monthTrend
        monthInsight.Importance = 4
        
        insights.Add monthInsight
    End If
    
    ' 3. Day of week analysis
    Dim dowSummary As Collection
    Set dowSummary = balanceCalc.GetDayOfWeekSummary(transactions)
    
    Dim highestDay As DayOfWeekSummary
    Dim totalDayAmount As Currency
    Dim highestAmount As Currency
    
    highestAmount = 0
    
    Dim day As DayOfWeekSummary
    For Each day In dowSummary
        totalDayAmount = totalDayAmount + day.TotalAmount
        
        If day.TotalAmount > highestAmount Then
            highestAmount = day.TotalAmount
            Set highestDay = day
        End If
    Next day
    
    If Not highestDay Is Nothing Then
        Dim dayPercent As Double
        If totalDayAmount > 0 Then
            dayPercent = highestDay.TotalAmount / totalDayAmount
        Else
            dayPercent = 0
        End If
        
        Dim dowInsight As AnalysisInsight
        dowInsight.Category = "Patterns"
        dowInsight.Title = "Highest spending day"
        dowInsight.Description = highestDay.DayName & " accounts for " & Format(dayPercent, "0.0%") & _
                               " of weekly spending"
        dowInsight.Value = highestDay.TotalAmount
        dowInsight.Trend = "Neutral"
        dowInsight.Importance = 2
        
        insights.Add dowInsight
    End If
    
    ' 4. Balance insight
    Dim balance As BalanceSummary
    balance = balanceCalc.CalculateBalance(transactions)
    
    Dim balanceInsight As AnalysisInsight
    balanceInsight.Category = "Balance"
    
    If Abs(balance.NetBalance) < 0.01 Then
        balanceInsight.Title = "Perfect balance"
        balanceInsight.Description = "Expenses are perfectly balanced between users"
        balanceInsight.Trend = "Neutral"
        balanceInsight.Importance = 5
    ElseIf Abs(balance.NetBalance) > 100 Then
        balanceInsight.Title = "Significant imbalance"
        balanceInsight.Description = balance.WhoOwes
        balanceInsight.Trend = "Up"
        balanceInsight.Importance = 5
    Else
        balanceInsight.Title = "Minor imbalance"
        balanceInsight.Description = balance.WhoOwes
        balanceInsight.Trend = "Neutral"
        balanceInsight.Importance = 3
    End If
    
    balanceInsight.Value = balance.OwedAmount
    
    insights.Add balanceInsight
    
    ' 5. Add any outlier transactions (unusually large)
    Dim averageAmount As Currency
    Dim totalAmount As Currency
    Dim transCount As Long
    
    Dim trans As Transaction
    For Each trans In transactions
        If trans.IsExpense Then
            totalAmount = totalAmount + Abs(trans.Amount)
            transCount = transCount + 1
        End If
    Next trans
    
    If transCount > 0 Then
        averageAmount = totalAmount / transCount
        
        ' Find outliers (more than 3x average)
        Dim outlierThreshold As Currency
        outlierThreshold = averageAmount * 3
        
        Dim outlierCount As Long
        outlierCount = 0
        
        For Each trans In transactions
            If trans.IsExpense And Abs(trans.Amount) > outlierThreshold Then
                outlierCount = outlierCount + 1
            End If
        Next trans
        
        If outlierCount > 0 Then
            Dim outlierInsight As AnalysisInsight
            outlierInsight.Category = "Outliers"
            outlierInsight.Title = "Large transactions detected"
            outlierInsight.Description = "Found " & outlierCount & " transactions over " & _
                                       Format(outlierThreshold, "$#,##0.00") & " (3x average)"
            outlierInsight.Value = outlierCount
            outlierInsight.Trend = "Up"
            outlierInsight.Importance = 4
            
            insights.Add outlierInsight
        End If
    End If
    
    Set GenerateSpendingInsights = insights
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "TransactionAnalyzer.GenerateSpendingInsights", Err.Number, Err.Description
    Set GenerateSpendingInsights = New Collection ' Return empty collection in case of error
End Function

' Generate and display insights on the insights sheet
Public Sub DisplayInsights()
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    ' Clear existing content
    Utilities.ClearSheet m_InsightsSheet
    
    ' Set up header
    With m_InsightsSheet
        ' Add title
        .Range("A1").Value = "BALANCE Financial Insights"
        .Range("A1").Font.Size = 24
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = AppSettings.ColorPrimary
        
        ' Add subtitle
        .Range("A2").Value = "Automated analysis of your spending patterns"
        .Range("A2").Font.Size = 12
        .Range("A2").Font.Italic = True
        
        ' Add last update info
        .Range("A3").Value = "Generated: " & Format(Now, "mm/dd/yyyy hh:mm:ss")
        
        ' Add divider line
        With .Range("A4:J4").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = AppSettings.ColorPrimary
        End With
    End With
    
    ' Generate insights
    Dim insights As Collection
    Set insights = GenerateSpendingInsights()
    
    ' Display insights
    Dim insightRow As Long
    insightRow = 6
    
    Dim insight As AnalysisInsight
    For Each insight In insights
        ' Display category
        m_InsightsSheet.Cells(insightRow, 1).Value = insight.Category
        m_InsightsSheet.Cells(insightRow, 1).Font.Bold = True
        
        ' Display title
        m_InsightsSheet.Cells(insightRow, 2).Value = insight.Title
        m_InsightsSheet.Cells(insightRow, 2).Font.Bold = True
        
        ' Display description
        m_InsightsSheet.Cells(insightRow, 3).Value = insight.Description
        
        ' Display value (if applicable)
        If Not IsEmpty(insight.Value) Then
            If IsNumeric(insight.Value) Then
                If insight.Category = "Balance" Or CCur(insight.Value) > 100 Then
                    m_InsightsSheet.Cells(insightRow, 4).Value = insight.Value
                    m_InsightsSheet.Cells(insightRow, 4).NumberFormat = "$#,##0.00"
                Else
                    m_InsightsSheet.Cells(insightRow, 4).Value = insight.Value
                End If
            Else
                m_InsightsSheet.Cells(insightRow, 4).Value = insight.Value
            End If
        End If
        
        ' Display trend indicator
        Select Case insight.Trend
            Case "Up"
                m_InsightsSheet.Cells(insightRow, 5).Value = "↑"
                m_InsightsSheet.Cells(insightRow, 5).Font.Color = AppSettings.ColorDanger
            Case "Down"
                m_InsightsSheet.Cells(insightRow, 5).Value = "↓"
                m_InsightsSheet.Cells(insightRow, 5).Font.Color = AppSettings.ColorSuccess
            Case Else ' "Neutral"
                m_InsightsSheet.Cells(insightRow, 5).Value = "→"
                m_InsightsSheet.Cells(insightRow, 5).Font.Color = AppSettings.ColorInfo
        End Select
        
        ' Add importance indicator (1-5 stars)
        Dim starCount As String
        starCount = String(insight.Importance, "★")
        m_InsightsSheet.Cells(insightRow, 6).Value = starCount
        m_InsightsSheet.Cells(insightRow, 6).Font.Color = AppSettings.ColorAccent
        
        ' Add separator line
        With m_InsightsSheet.Range("A" & (insightRow + 1) & ":F" & (insightRow + 1)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(220, 220, 220)
        End With
        
        ' Move to next row
        insightRow = insightRow + 2
    Next insight
    
    ' Format columns
    m_InsightsSheet.Columns("A:F").AutoFit
    
    ' Add recommendation section if we have insights
    If insights.Count > 0 Then
        insightRow = insightRow + 1
        
        m_InsightsSheet.Cells(insightRow, 1).Value = "Recommendations"
        m_InsightsSheet.Cells(insightRow, 1).Font.Size = 14
        m_InsightsSheet.Cells(insightRow, 1).Font.Bold = True
        m_InsightsSheet.Cells(insightRow, 1).Font.Color = AppSettings.ColorPrimary
        
        insightRow = insightRow + 2
        
        ' Generate recommendations based on insights
        m_InsightsSheet.Cells(insightRow, 1).Value = "1."
        m_InsightsSheet.Cells(insightRow, 2).Value = "Consider using the category filter to analyze your top spending areas"
        
        insightRow = insightRow + 1
        m_InsightsSheet.Cells(insightRow, 1).Value = "2."
        m_InsightsSheet.Cells(insightRow, 2).Value = "Review transactions flagged as outliers for potential errors or one-time expenses"
        
        insightRow = insightRow + 1
        m_InsightsSheet.Cells(insightRow, 1).Value = "3."
        m_InsightsSheet.Cells(insightRow, 2).Value = "Try to even out expenses between " & AppSettings.User1Name & " and " & AppSettings.User2Name & " for easier balancing"
    End If
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "TransactionAnalyzer.DisplayInsights", Err.Number, Err.Description
End Sub

' Class initialize
Private Sub Class_Initialize()
    ' Nothing needed here as we're using PredeclaredId = True
    ' and explicit initialization via Initialize method
    m_IsInitialized = False
End Sub

' Class terminate
Private Sub Class_Terminate()
    ' Clean up to prevent memory leaks
    Set m_InsightsSheet = Nothing
End Sub