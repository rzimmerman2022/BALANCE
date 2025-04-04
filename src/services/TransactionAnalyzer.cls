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
' Implements ICategoryAnalyzer, ITimeAnalyzer, ITrendAnalyzer

' Interface Implementations
Implements ICategoryAnalyzer
Implements ITimeAnalyzer
Implements ITrendAnalyzer

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
Private m_Logger As IErrorLogger ' Add logger dependency
Private m_Calculator As IBalanceCalculator ' Dependency for calculations

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
    ' Get dependencies via ServiceLocator if not already set
    If m_Logger Is Nothing Then Set m_Logger = ServiceLocator.GetErrorLogger()
    If m_Calculator Is Nothing Then Set m_Calculator = ServiceLocator.GetBalanceCalculator()

    m_IsInitialized = True

    Exit Sub

ErrorHandler:
    ' Use the member logger if available
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionAnalyzer.Initialize", Err.Number, Err.Description
    Else ' Fallback if logger itself failed
        Debug.Print "ERROR in TransactionAnalyzer.Initialize: " & Err.Description
    End If
End Sub

' ===== Interface Implementations =====

' --- ICategoryAnalyzer ---
Private Function ICategoryAnalyzer_CalculateCategorySummary(ByVal transactions As Collection, _
                                                          Optional ByVal startDate As Date, _
                                                          Optional ByVal endDate As Date) As Object
    ' Ensure initialized and dependencies are met
    If Not m_IsInitialized Then Initialize
    If m_Calculator Is Nothing Then Set m_Calculator = ServiceLocator.GetBalanceCalculator

    ' Delegate to BalanceCalculator (which has the implementation)
    Set ICategoryAnalyzer_CalculateCategorySummary = m_Calculator.CalculateCategorySummary(transactions, startDate, endDate)
End Function

Private Sub ICategoryAnalyzer_Initialize(Optional ByVal logger As IErrorLogger = Nothing)
    ' Set logger if provided, otherwise use existing or get from ServiceLocator
    If Not logger Is Nothing Then
        Set m_Logger = logger
    ElseIf m_Logger Is Nothing Then
        Set m_Logger = ServiceLocator.GetErrorLogger()
    End If
    ' Ensure base initialization runs
    If Not m_IsInitialized Then Initialize
End Sub

' --- ITimeAnalyzer ---
Private Function ITimeAnalyzer_CalculateDayOfWeekSummary(ByVal transactions As Collection, _
                                                        Optional ByVal startDate As Date, _
                                                        Optional ByVal endDate As Date) As Object
    ' Ensure initialized and dependencies are met
    If Not m_IsInitialized Then Initialize
    If m_Calculator Is Nothing Then Set m_Calculator = ServiceLocator.GetBalanceCalculator

    ' Delegate to BalanceCalculator
    Set ITimeAnalyzer_CalculateDayOfWeekSummary = m_Calculator.CalculateDayOfWeekSummary(transactions, startDate, endDate)
End Function

Private Function ITimeAnalyzer_CalculateMonthSummary(ByVal transactions As Collection, _
                                                    Optional ByVal monthCount As Integer = 12) As Object
    ' Ensure initialized and dependencies are met
    If Not m_IsInitialized Then Initialize
    If m_Calculator Is Nothing Then Set m_Calculator = ServiceLocator.GetBalanceCalculator

    ' Delegate to BalanceCalculator
    Set ITimeAnalyzer_CalculateMonthSummary = m_Calculator.CalculateMonthSummary(transactions, monthCount)
End Function

Private Function ITimeAnalyzer_CalculateDailyAverageSpending(ByVal transactions As Collection, _
                                                            Optional ByVal startDate As Date, _
                                                            Optional ByVal endDate As Date) As Double
    ' Ensure initialized and dependencies are met
    If Not m_IsInitialized Then Initialize
    If m_Calculator Is Nothing Then Set m_Calculator = ServiceLocator.GetBalanceCalculator

    ' Delegate to BalanceCalculator
    ITimeAnalyzer_CalculateDailyAverageSpending = m_Calculator.CalculateDailyAverageSpending(transactions, startDate, endDate)
End Function

Private Sub ITimeAnalyzer_Initialize(Optional ByVal logger As IErrorLogger = Nothing)
    ' Set logger if provided, otherwise use existing or get from ServiceLocator
    If Not logger Is Nothing Then
        Set m_Logger = logger
    ElseIf m_Logger Is Nothing Then
        Set m_Logger = ServiceLocator.GetErrorLogger()
    End If
    ' Ensure base initialization runs
    If Not m_IsInitialized Then Initialize
End Sub

' --- ITrendAnalyzer ---
Private Function ITrendAnalyzer_CalculateSpendingTrends(ByVal transactions As Collection, _
                                                       Optional ByVal monthCount As Integer = 12) As Object
    ' Ensure initialized and dependencies are met
    If Not m_IsInitialized Then Initialize
    If m_Calculator Is Nothing Then Set m_Calculator = ServiceLocator.GetBalanceCalculator

    ' Delegate to BalanceCalculator
    Set ITrendAnalyzer_CalculateSpendingTrends = m_Calculator.CalculateSpendingTrends(transactions, monthCount)
End Function

Private Sub ITrendAnalyzer_Initialize(Optional ByVal logger As IErrorLogger = Nothing)
    ' Set logger if provided, otherwise use existing or get from ServiceLocator
    If Not logger Is Nothing Then
        Set m_Logger = logger
    ElseIf m_Logger Is Nothing Then
        Set m_Logger = ServiceLocator.GetErrorLogger()
    End If
    ' Ensure base initialization runs
    If Not m_IsInitialized Then Initialize
End Sub


' ===== Existing Public Methods =====

' Generate spending insights
Public Function GenerateSpendingInsights() As Collection
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    Dim insights As New Collection
    ' Get dependencies
    Dim repo As ITransactionRepository
    Set repo = ServiceLocator.GetTransactionRepository() ' Use ServiceLocator
    ' No need to initialize repo here, ServiceLocator handles it

    ' Use the member calculator instance
    If m_Calculator Is Nothing Then Set m_Calculator = ServiceLocator.GetBalanceCalculator()
    
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
    Dim catSummaryObj As Object ' Dictionary from calculator
    Set catSummaryObj = m_Calculator.CalculateCategorySummary(transactions) ' Use member calculator

    ' Need to adapt insight generation based on Dictionary structure returned by BalanceCalculator
    ' For now, commenting out the parts that rely on specific Collection structure
    ' This section needs review based on actual data returned by CalculateCategorySummary
'    If catSummaryObj.Count > 0 Then
'        Dim topCategory As CategorySummary ' This Type is defined in Module1
'        ' Need logic to find top category from dictionary
'        ' Set topCategory = ... ' Logic to extract from catSummaryObj
'
'        Dim catInsight As AnalysisInsight
'        catInsight.Category = "Categories"
'        catInsight.Title = "Top spending category"
'        catInsight.Description = topCategory.Category & " accounts for " & _
'                              Format(topCategory.Percentage, "0.0%") & " of total spending"
'        catInsight.Value = topCategory.TotalAmount
'        catInsight.Trend = "Neutral"
'        catInsight.Importance = 3
'
'        insights.Add catInsight
'    End If
        
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
    Dim monthlySummaryObj As Object ' Dictionary from calculator
    Set monthlySummaryObj = m_Calculator.CalculateMonthSummary(transactions, 3) ' Use member calculator

    ' Need to adapt insight generation based on Dictionary structure
    ' This section needs review based on actual data returned by CalculateMonthSummary
'    If monthlySummaryObj.Count >= 2 Then ' Check dictionary count
'        Dim currentMonth As MonthSummary ' This Type is defined in Module1
'        Dim previousMonth As MonthSummary ' This Type is defined in Module1
'
'        ' Need logic to get current/previous month data from dictionary
'        ' Set currentMonth = ...
'        ' Set previousMonth = ...
'        Dim monthlyChange As Double
'        If previousMonth.ExpenseAmount > 0 Then ' Use ExpenseAmount from Type
'            monthlyChange = (currentMonth.ExpenseAmount - previousMonth.ExpenseAmount) / previousMonth.ExpenseAmount
'        Else
'            monthlyChange = 0
'        End If
        
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
'        insights.Add monthInsight
'    End If

    ' 3. Day of week analysis
    Dim dowSummaryObj As Object ' Dictionary from calculator
    Set dowSummaryObj = m_Calculator.CalculateDayOfWeekSummary(transactions) ' Use member calculator

    ' Need to adapt insight generation based on Dictionary structure
    ' This section needs review based on actual data returned by CalculateDayOfWeekSummary
'    Dim highestDay As DayOfWeekSummary ' This Type is defined in Module1
'    Dim totalDayAmount As Currency
'    Dim highestAmount As Currency
'
'    highestAmount = 0
'
'    Dim dayKey As Variant
'    For Each dayKey In dowSummaryObj.Keys
'        If Right(dayKey, 8) <> "_Percent" And dayKey <> "Total" Then ' Check it's a day name key
'            Dim dayAmount As Currency
'            dayAmount = dowSummaryObj(dayKey)
'            totalDayAmount = totalDayAmount + dayAmount
'
'            If dayAmount > highestAmount Then
'                highestAmount = dayAmount
'                ' Need to reconstruct DayOfWeekSummary Type if needed, or just use values
'                ' Set highestDay = ... ' Or just store highestDayName and highestAmount
'            End If
'        End If
'    Next dayKey
'
'    If highestAmount > 0 Then ' Check if any highest day was found
'        Dim dayPercent As Double
'        If totalDayAmount > 0 Then
'            dayPercent = highestAmount / totalDayAmount ' Use highestAmount found
'        Else
'            dayPercent = 0
'        End If
        
        Dim dowInsight As AnalysisInsight
        dowInsight.Category = "Patterns"
        dowInsight.Title = "Highest spending day"
        dowInsight.Description = highestDay.DayName & " accounts for " & Format(dayPercent, "0.0%") & _
                               " of weekly spending"
        dowInsight.Value = highestDay.TotalAmount
        dowInsight.Trend = "Neutral"
        dowInsight.Importance = 2
'        insights.Add dowInsight
'    End If

    ' 4. Balance insight
    Dim balanceSummaryObj As Object ' Dictionary from calculator
    Set balanceSummaryObj = m_Calculator.CalculateBalanceSummary(transactions) ' Use member calculator

    Dim balanceInsight As AnalysisInsight
    balanceInsight.Category = "Balance"

    Dim netBalance As Currency
    Dim owedAmount As Currency
    Dim owingUser As String
    Dim owedUser As String

    netBalance = balanceSummaryObj("Balance")
    owedAmount = balanceSummaryObj("OwedAmount")
    owingUser = balanceSummaryObj("OwingUser")
    owedUser = balanceSummaryObj("OwedUser")

    If Abs(netBalance) < 0.01 Then
        balanceInsight.Title = "Perfect balance"
        balanceInsight.Description = "Expenses are perfectly balanced between users"
        balanceInsight.Trend = "Neutral"
        balanceInsight.Importance = 5
    ElseIf Abs(netBalance) > 100 Then
        balanceInsight.Title = "Significant imbalance"
        balanceInsight.Description = owingUser & " owes " & owedUser & " " & Format(owedAmount, "$#,##0.00")
        balanceInsight.Trend = "Up" ' Assuming imbalance is generally 'Up' in importance
        balanceInsight.Importance = 5
    Else
        balanceInsight.Title = "Minor imbalance"
        balanceInsight.Description = owingUser & " owes " & owedUser & " " & Format(owedAmount, "$#,##0.00")
        balanceInsight.Trend = "Neutral"
        balanceInsight.Importance = 3
    End If

    balanceInsight.Value = owedAmount
    
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
    If Not m_Logger Is Nothing Then m_Logger.LogError "TransactionAnalyzer.GenerateSpendingInsights", Err.Number, Err.Description
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
    If Not m_Logger Is Nothing Then m_Logger.LogError "TransactionAnalyzer.DisplayInsights", Err.Number, Err.Description
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
    Set m_Logger = Nothing
    Set m_Calculator = Nothing
End Sub
