VERSION 1.0 CLASS
Attribute VB_Name = "BalanceCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' BalanceCalculator Class - Implements IBalanceCalculator, ICategoryAnalyzer, ITimeAnalyzer, ITrendAnalyzer
'
' Purpose: Performs financial calculations and analysis on transaction data.
' This class handles balance calculations, spending summaries by category,
' time-based analysis, and other financial metrics.
'
' Design Decisions:
' - Uses PredeclaredId = True to enable Singleton pattern
' - Implements calculation/analysis interfaces for loose coupling & dependency injection via ServiceLocator
' - Uses array-based operations for performance with large datasets
' - Implements dependency injection for the logger component
' - Includes comprehensive error handling

' Interface implementation
Implements IBalanceCalculator
Implements ICategoryAnalyzer
Implements ITimeAnalyzer
Implements ITrendAnalyzer


' Private member variables
Private m_Logger As IErrorLogger
Private m_IsInitialized As Boolean
Private m_User1 As String
Private m_User2 As String

'=========================================================================
' Initialization
'=========================================================================

Private Sub Class_Initialize()
    ' Set default values
    m_IsInitialized = False
    
    ' Get user names from settings (handle potential errors)
    On Error Resume Next ' Defer error handling until after attempt
    m_User1 = AppSettings.User1Name ' Assumes AppSettings is initialized
    If Err.Number <> 0 Or Len(Trim$(m_User1)) = 0 Then m_User1 = "User 1": Err.Clear
    m_User2 = AppSettings.User2Name
    If Err.Number <> 0 Or Len(Trim$(m_User2)) = 0 Then m_User2 = "User 2": Err.Clear
    On Error GoTo 0 ' Restore error handling
End Sub

' Initialize the calculator with dependencies (Called by interfaces & public wrapper)
Private Sub InitializeCalculator(Optional ByVal logger As IErrorLogger = Nothing)
    ' Prevent re-initialization
    If m_IsInitialized Then Exit Sub
    On Error GoTo ErrorHandler
    
    ' Set logger if provided, otherwise get from ServiceLocator
    If Not logger Is Nothing Then
        Set m_Logger = logger
    ElseIf m_Logger Is Nothing Then ' Only get from locator if not already set
        Set m_Logger = ServiceLocator.GetErrorLogger()
    End If
    
    ' Refresh user names in case settings changed after Class_Initialize
    On Error Resume Next
    m_User1 = AppSettings.User1Name
    If Err.Number <> 0 Or Len(Trim$(m_User1)) = 0 Then m_User1 = "User 1": Err.Clear
    m_User2 = AppSettings.User2Name
    If Err.Number <> 0 Or Len(Trim$(m_User2)) = 0 Then m_User2 = "User 2": Err.Clear
    On Error GoTo ErrorHandler
    
    ' Log initialization
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "BalanceCalculator.Initialize", "Balance Calculator initialized/re-initialized"
    End If
    
    m_IsInitialized = True
    
    Exit Sub
    
ErrorHandler:
    ' Can't reliably use logger here if that's what caused the error or wasn't set yet
    Debug.Print "ERROR in BalanceCalculator.InitializeCalculator: (" & Err.Number & ") " & Err.Description
End Sub

' Public wrapper for Initialize (Can be called directly if needed)
Public Sub Initialize(Optional ByVal logger As IErrorLogger = Nothing)
    InitializeCalculator logger
End Sub

' Interface Initialize methods - delegate to common private initializer
Private Sub IBalanceCalculator_Initialize(Optional ByVal logger As IErrorLogger = Nothing)
    InitializeCalculator logger
End Sub
Private Sub ICategoryAnalyzer_Initialize(Optional ByVal logger As IErrorLogger = Nothing)
    InitializeCalculator logger
End Sub
Private Sub ITimeAnalyzer_Initialize(Optional ByVal logger As IErrorLogger = Nothing)
    InitializeCalculator logger
End Sub
Private Sub ITrendAnalyzer_Initialize(Optional ByVal logger As IErrorLogger = Nothing)
    InitializeCalculator logger
End Sub

'=========================================================================
' Interface Implementation - Balance Calculations (IBalanceCalculator)
'=========================================================================

' Calculate the overall balance between users
Private Function IBalanceCalculator_CalculateBalance(ByVal transactions As Collection) As Double
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Me.Initialize ' Call public Initialize wrapper
    
    Dim transArray() As Variant: transArray = TransactionsToArray(transactions)
    Dim balance As Double: balance = CalculateBalanceFromArray(transArray)
    IBalanceCalculator_CalculateBalance = balance
    
    Exit Function
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "BalanceCalculator.IBalanceCalculator_CalculateBalance", Err.Number, Err.Description
    IBalanceCalculator_CalculateBalance = 0
End Function

' Public wrapper for CalculateBalance
Public Function CalculateBalance(ByVal transactions As Collection) As Double
    CalculateBalance = IBalanceCalculator_CalculateBalance(transactions)
End Function

' Calculate who owes whom and how much
Private Function IBalanceCalculator_CalculateBalanceSummary(ByVal transactions As Collection) As Object
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Me.Initialize
    
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    Dim transArray() As Variant: transArray = TransactionsToArray(transactions)
    Dim balance As Double: balance = CalculateBalanceFromArray(transArray)
    
    Dim owes As String, owed As String
    Dim amount As Double: amount = Abs(balance)
    
    If balance > 0.001 Then ' User1 net positive -> User1 owes User2
        owes = m_User1: owed = m_User2
    ElseIf balance < -0.001 Then ' User1 net negative -> User2 owes User1
        owes = m_User2: owed = m_User1
    Else ' Balanced
        owes = "Nobody": owed = "Nobody": amount = 0
    End If
    
    result.Add "Balance", balance
    result.Add "OwedAmount", amount
    result.Add "OwingUser", owes
    result.Add "OwedUser", owed
    result.Add "User1", m_User1
    result.Add "User2", m_User2
    result.Add "User1Spent", GetUserSpentAmount(transArray, m_User1)
    result.Add "User2Spent", GetUserSpentAmount(transArray, m_User2)
    result.Add "TotalShared", GetTotalSharedAmount(transArray)
    result.Add "TotalSpent", GetTotalSpentAmount(transArray)
    
    Set IBalanceCalculator_CalculateBalanceSummary = result
    Exit Function
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "BalanceCalculator.CalculateBalanceSummary", Err.Number, Err.Description
    Set result = CreateObject("Scripting.Dictionary") ' Return empty
    Set IBalanceCalculator_CalculateBalanceSummary = result
End Function

' Public wrapper for CalculateBalanceSummary
Public Function CalculateBalanceSummary(ByVal transactions As Collection) As Object
    Set CalculateBalanceSummary = IBalanceCalculator_CalculateBalanceSummary(transactions)
End Function

'=========================================================================
' Interface Implementation - Category Analysis (ICategoryAnalyzer)
'=========================================================================
' Calculate spending by category
' ***** CORRECTED Signature: Optional parameters are Date type, no default values *****
Private Function ICategoryAnalyzer_CalculateCategorySummary(ByVal transactions As Collection, _
                                                            Optional ByVal startDate As Date, _
                                                            Optional ByVal endDate As Date) As Object
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Me.Initialize

    Dim useStartDate As Date
    Dim useEndDate As Date
    Dim applyDateFilter As Boolean: applyDateFilter = True

    ' Handle optional parameters *inside* the function
    If startDate = 0 And endDate = 0 Then ' Check if both were omitted (default to 0)
        applyDateFilter = False
    Else
        If startDate = 0 Then useStartDate = DateSerial(1900, 1, 1) Else useStartDate = startDate
        If endDate = 0 Then useEndDate = Date Else useEndDate = endDate
        If useStartDate > useEndDate Then useStartDate = useEndDate ' Basic validation
    End If
        
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    Dim transArray() As Variant: transArray = TransactionsToArray(transactions)
    
    Dim dataToAnalyze() As Variant
    If applyDateFilter Then dataToAnalyze = FilterTransactionsByDate(transArray, useStartDate, useEndDate) Else dataToAnalyze = transArray
    
    Dim rowCount As Long
    On Error Resume Next: rowCount = UBound(dataToAnalyze, 1): On Error GoTo ErrorHandler
    If rowCount = 0 Then GoTo NoData

    result = CalculateCategoryTotalsFromArray(dataToAnalyze) ' Get category totals
    Dim totalSpent As Double: totalSpent = GetTotalSpentAmount(dataToAnalyze) ' Get total for percentage calc

    If result.Count > 0 Then ' Calculate percentages
        Dim category As Variant
        For Each category In result.Keys
             If Right(category, 8) <> "_Percent" And category <> "Total" Then ' Avoid metadata keys
                Dim categoryAmount As Double: categoryAmount = result(category)
                If totalSpent <> 0 Then result.Add category & "_Percent", Round(categoryAmount / totalSpent * 100, 2) Else result.Add category & "_Percent", 0
             End If
        Next category
    End If
    If Not result.Exists("Total") Then result.Add "Total", totalSpent Else result("Total") = totalSpent ' Add/Update total

NoData:
    Set ICategoryAnalyzer_CalculateCategorySummary = result
    Exit Function
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "BalanceCalculator.ICategoryAnalyzer_CalculateCategorySummary", Err.Number, Err.Description
    Set result = CreateObject("Scripting.Dictionary"): Set ICategoryAnalyzer_CalculateCategorySummary = result ' Return empty
End Function

' Public wrapper for CalculateCategorySummary
Public Function CalculateCategorySummary(ByVal transactions As Collection, _
                                       Optional ByVal startDate As Date = 0, _
                                       Optional ByVal endDate As Date = 0) As Object
    ' Public function handles default values if needed before calling private implementation
    Dim callStartDate As Date, callEndDate As Date
    If startDate = 0 And endDate = 0 Then
        ' If called publicly with no dates, maybe default to last 6 months? Or pass 0 to impl?
        ' Passing 0 to implementation will cause it to analyze all data based on current logic.
        callStartDate = 0
        callEndDate = 0
    Else
        If startDate = 0 Then callStartDate = DateSerial(1900, 1, 1) Else callStartDate = startDate
        If endDate = 0 Then callEndDate = Date Else callEndDate = endDate
    End If
    Set CalculateCategorySummary = ICategoryAnalyzer_CalculateCategorySummary(transactions, callStartDate, callEndDate)
End Function


'=========================================================================
' Interface Implementation - Time Analysis (ITimeAnalyzer)
'=========================================================================
' Calculate spending by day of week
' ***** CORRECTED Signature: Optional parameters are Date type, no default values *****
Private Function ITimeAnalyzer_CalculateDayOfWeekSummary(ByVal transactions As Collection, _
                                                         Optional ByVal startDate As Date, _
                                                         Optional ByVal endDate As Date) As Object
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Me.Initialize

    Dim useStartDate As Date
    Dim useEndDate As Date
    Dim applyDateFilter As Boolean: applyDateFilter = True

    ' Handle optional parameters *inside* the function
    If startDate = 0 And endDate = 0 Then
        applyDateFilter = False
    Else
        If startDate = 0 Then useStartDate = DateSerial(1900, 1, 1) Else useStartDate = startDate
        If endDate = 0 Then useEndDate = Date Else useEndDate = endDate
        If useStartDate > useEndDate Then useStartDate = useEndDate
    End If

    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    Dim i As Integer
    For i = vbSunday To vbSaturday: result.Add WeekdayName(i, True), 0: Next i ' Init days
    
    Dim transArray() As Variant: transArray = TransactionsToArray(transactions)
    Dim dataToAnalyze() As Variant
    If applyDateFilter Then dataToAnalyze = FilterTransactionsByDate(transArray, useStartDate, useEndDate) Else dataToAnalyze = transArray

    Dim rowCount As Long
    On Error Resume Next: rowCount = UBound(dataToAnalyze, 1): On Error GoTo ErrorHandler
    If rowCount = 0 Then GoTo NoTimeData

    ' Calculate totals
    Dim transDate As Date, transAmount As Double, dayName As String
    For i = 1 To rowCount
        transDate = dataToAnalyze(i, 2): transAmount = dataToAnalyze(i, 5)
        If transAmount < 0 Then ' Only expenses
            dayName = WeekdayName(Weekday(transDate), True)
            If result.Exists(dayName) Then result(dayName) = result(dayName) + Abs(transAmount)
        End If
    Next i
    
    ' Calculate percentages
    Dim totalSpent As Double: totalSpent = 0
    For i = vbSunday To vbSaturday: dayName = WeekdayName(i, True): If result.Exists(dayName) Then totalSpent = totalSpent + result(dayName): Next i
    For i = vbSunday To vbSaturday
        dayName = WeekdayName(i, True)
        If result.Exists(dayName) Then
            If totalSpent <> 0 Then result.Add dayName & "_Percent", Round(result(dayName) / totalSpent * 100, 2) Else result.Add dayName & "_Percent", 0
        End If
    Next i
    If Not result.Exists("Total") Then result.Add "Total", totalSpent Else result("Total") = totalSpent

NoTimeData:
    Set ITimeAnalyzer_CalculateDayOfWeekSummary = result
    Exit Function
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "BalanceCalculator.ITimeAnalyzer_CalculateDayOfWeekSummary", Err.Number, Err.Description
    Set result = CreateObject("Scripting.Dictionary"): Set ITimeAnalyzer_CalculateDayOfWeekSummary = result ' Return empty
End Function

' Public wrapper for CalculateDayOfWeekSummary
Public Function CalculateDayOfWeekSummary(ByVal transactions As Collection, _
                                        Optional ByVal startDate As Date = 0, _
                                        Optional ByVal endDate As Date = 0) As Object
    ' Public wrapper handles default value logic if needed before calling private implementation
     Dim callStartDate As Date, callEndDate As Date
     If startDate = 0 And endDate = 0 Then callStartDate = 0: callEndDate = 0 Else ' Pass 0 if both missing
         If startDate = 0 Then callStartDate = DateSerial(1900, 1, 1) Else callStartDate = startDate
         If endDate = 0 Then callEndDate = Date Else callEndDate = endDate
     End If
    Set CalculateDayOfWeekSummary = ITimeAnalyzer_CalculateDayOfWeekSummary(transactions, callStartDate, callEndDate)
End Function


' Calculate spending by month (Signature matches interface, monthCount has default)
Private Function ITimeAnalyzer_CalculateMonthSummary(ByVal transactions As Collection, _
                                                     Optional ByVal monthCount As Integer = 12) As Object
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Me.Initialize
    If monthCount <= 0 Then monthCount = 1 ' Ensure at least one month
    
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    
    Dim endDate As Date: endDate = DateSerial(Year(Date), Month(Date) + 1, 0)
    Dim startDate As Date: startDate = DateAdd("m", -(monthCount - 1), DateSerial(Year(endDate), Month(endDate), 1))

    Dim tempDate As Date: tempDate = startDate
    Do While tempDate <= endDate ' Initialize dictionary keys
        Dim monthKey As String: monthKey = Format(tempDate, "yyyy-mm")
        If Not result.Exists(monthKey) Then
             result.Add monthKey, CDbl(0): result.Add monthKey & "_Income", CDbl(0): result.Add monthKey & "_Expense", CDbl(0)
        End If
        tempDate = DateAdd("m", 1, tempDate)
    Loop
    
    Dim transArray() As Variant: transArray = TransactionsToArray(transactions)
    Dim filteredArray() As Variant: filteredArray = FilterTransactionsByDate(transArray, startDate, endDate)

    Dim rowCount As Long
    On Error Resume Next: rowCount = UBound(filteredArray, 1): On Error GoTo ErrorHandler
    If rowCount = 0 Then GoTo NoMonthData

    Dim i As Long, transDate As Date, transAmount As Double, currentMonthKey As String
    For i = 1 To rowCount
        transDate = filteredArray(i, 2): transAmount = filteredArray(i, 5)
        currentMonthKey = Format(transDate, "yyyy-mm")
        If result.Exists(currentMonthKey) Then
            result(currentMonthKey) = result(currentMonthKey) + transAmount
            If transAmount > 0 Then result(currentMonthKey & "_Income") = result(currentMonthKey & "_Income") + transAmount Else result(currentMonthKey & "_Expense") = result(currentMonthKey & "_Expense") + Abs(transAmount)
        End If
    Next i
    
NoMonthData:
    Set ITimeAnalyzer_CalculateMonthSummary = result
    Exit Function
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "BalanceCalculator.ITimeAnalyzer_CalculateMonthSummary", Err.Number, Err.Description
    Set result = CreateObject("Scripting.Dictionary"): Set ITimeAnalyzer_CalculateMonthSummary = result
End Function

' Public wrapper for CalculateMonthSummary
Public Function CalculateMonthSummary(ByVal transactions As Collection, _
                                    Optional ByVal monthCount As Integer = 12) As Object
    Set CalculateMonthSummary = ITimeAnalyzer_CalculateMonthSummary(transactions, monthCount)
End Function


' Calculate daily average spending
' ***** CORRECTED Signature: Optional parameters are Date type, no default values *****
Private Function ITimeAnalyzer_CalculateDailyAverageSpending(ByVal transactions As Collection, _
                                                             Optional ByVal startDate As Date, _
                                                             Optional ByVal endDate As Date) As Double
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Me.Initialize
    
    Dim useStartDate As Date
    Dim useEndDate As Date
    Dim applyDateFilter As Boolean: applyDateFilter = True

    ' Handle optional parameters *inside* the function
    If startDate = 0 And endDate = 0 Then ' Default to last 30 days if no range
        useEndDate = Date
        useStartDate = DateAdd("d", -29, useEndDate)
    Else
        If startDate = 0 Then useStartDate = DateSerial(1900, 1, 1) Else useStartDate = startDate
        If endDate = 0 Then useEndDate = Date Else useEndDate = endDate
        If useStartDate > useEndDate Then useStartDate = useEndDate
    End If

    Dim transArray() As Variant: transArray = TransactionsToArray(transactions)
    Dim filteredArray() As Variant: filteredArray = FilterTransactionsByDate(transArray, useStartDate, useEndDate)
    
    Dim rowCount As Long
    On Error Resume Next: rowCount = UBound(filteredArray, 1): On Error GoTo ErrorHandler
    If rowCount = 0 Then GoTo NoSpendingData

    Dim totalExpenses As Double: totalExpenses = 0
    Dim i As Long, transAmount As Double
    For i = 1 To rowCount
        transAmount = filteredArray(i, 5)
        If transAmount < 0 Then totalExpenses = totalExpenses + Abs(transAmount)
    Next i
    
    Dim dayCount As Long: dayCount = DateDiff("d", useStartDate, useEndDate) + 1
    Dim dailyAverage As Double
    If dayCount > 0 Then dailyAverage = totalExpenses / dayCount Else dailyAverage = 0
    
NoSpendingData:
    ITimeAnalyzer_CalculateDailyAverageSpending = dailyAverage
    Exit Function
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "BalanceCalculator.ITimeAnalyzer_CalculateDailyAverageSpending", Err.Number, Err.Description
    ITimeAnalyzer_CalculateDailyAverageSpending = 0
End Function

' Public wrapper for CalculateDailyAverageSpending
Public Function CalculateDailyAverageSpending(ByVal transactions As Collection, _
                                            Optional ByVal startDate As Date = 0, _
                                            Optional ByVal endDate As Date = 0) As Double
     Dim callStartDate As Date, callEndDate As Date
     If startDate = 0 And endDate = 0 Then callStartDate = 0: callEndDate = 0 Else ' Pass 0 if both missing to trigger default logic
         If startDate = 0 Then callStartDate = DateSerial(1900, 1, 1) Else callStartDate = startDate
         If endDate = 0 Then callEndDate = Date Else callEndDate = endDate
     End If
    CalculateDailyAverageSpending = ITimeAnalyzer_CalculateDailyAverageSpending(transactions, callStartDate, callEndDate)
End Function


'=========================================================================
' Interface Implementation - Trend Analysis (ITrendAnalyzer)
'=========================================================================
' Calculate spending trends (month-over-month changes) (Signature matches interface)
Private Function ITrendAnalyzer_CalculateSpendingTrends(ByVal transactions As Collection, _
                                                        Optional ByVal monthCount As Integer = 12) As Object
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Me.Initialize
    If monthCount <= 1 Then monthCount = 2 ' Need at least 2 months for trends
    
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    Dim monthSummary As Object: Set monthSummary = ITimeAnalyzer_CalculateMonthSummary(transactions, monthCount + 1) ' Get extra month
    Dim months As Collection: Set months = Utilities.GetSortedKeys(monthSummary)
    
    Dim i As Long, monthsAvailable As Long: monthsAvailable = months.Count
    If monthsAvailable < 2 Then GoTo NoTrendData
    
    Dim currentMonthKey As String, previousMonthKey As String
    Dim currentExpense As Double, previousExpense As Double
    Dim changeAmount As Double, changePercent As Double
    
    For i = 2 To monthsAvailable
        currentMonthKey = months(i): previousMonthKey = months(i - 1)
        If monthSummary.Exists(currentMonthKey & "_Expense") And monthSummary.Exists(previousMonthKey & "_Expense") Then
            currentExpense = monthSummary(currentMonthKey & "_Expense")
            previousExpense = monthSummary(previousMonthKey & "_Expense")
            changeAmount = currentExpense - previousExpense
            If previousExpense <> 0 Then changePercent = Round(changeAmount / previousExpense * 100, 2) Else changePercent = 0
            result.Add currentMonthKey & "_Change", changeAmount
            result.Add currentMonthKey & "_ChangePercent", changePercent
        End If
    Next i
    
NoTrendData:
    Set ITrendAnalyzer_CalculateSpendingTrends = result
    Exit Function
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "BalanceCalculator.ITrendAnalyzer_CalculateSpendingTrends", Err.Number, Err.Description
    Set result = CreateObject("Scripting.Dictionary"): Set ITrendAnalyzer_CalculateSpendingTrends = result
End Function

' Public wrapper for CalculateSpendingTrends
Public Function CalculateSpendingTrends(ByVal transactions As Collection, _
                                      Optional ByVal monthCount As Integer = 12) As Object
    Set CalculateSpendingTrends = ITrendAnalyzer_CalculateSpendingTrends(transactions, monthCount)
End Function


'=========================================================================
' Helper Methods - Array Operations
'=========================================================================

' Convert transactions collection to array for faster processing
Private Function TransactionsToArray(ByVal transactions As Collection) As Variant
    On Error GoTo ErrorHandler
    TransactionsToArray = Array() ' Default empty
    Dim transCount As Long: transCount = transactions.Count
    If transCount = 0 Then Exit Function
    
    Dim result() As Variant: ReDim result(1 To transCount, 1 To 10)
    Dim i As Long: i = 1
    Dim trans As ITransaction
    
    For Each trans In transactions
        If TypeName(trans) = "Transaction" Or TypeOf trans Is ITransaction Then ' Defensive check
            result(i, 1) = trans.ID: result(i, 2) = trans.TransactionDate: result(i, 3) = trans.Merchant
            result(i, 4) = trans.Category: result(i, 5) = trans.Amount: result(i, 6) = trans.Account
            result(i, 7) = trans.Owner: result(i, 8) = trans.IsShared: result(i, 9) = trans.Notes
            result(i, 10) = trans.SourceFile
            i = i + 1
        Else
             If Not m_Logger Is Nothing Then m_Logger.LogWarning "TransactionsToArray", 0, "Skipping item in collection - not an ITransaction object. Type: " & TypeName(trans)
        End If
    Next trans
    
    If i <= transCount Then ReDim Preserve result(1 To i - 1, 1 To 10) ' Resize if items were skipped
    TransactionsToArray = result
    Exit Function
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "BalanceCalculator.TransactionsToArray", Err.Number, Err.Description
    TransactionsToArray = Array()
End Function

' Calculate balance from transaction array
Private Function CalculateBalanceFromArray(ByVal transArray As Variant) As Double
    On Error GoTo ErrorHandler
    CalculateBalanceFromArray = 0: If Not IsArray(transArray) Then Exit Function
    Dim rowCount As Long
    On Error Resume Next: rowCount = UBound(transArray, 1): On Error GoTo ErrorHandler
    If rowCount = 0 Then Exit Function

    Dim user1Net As Double: user1Net = 0
    Dim i As Long, transAmount As Double, transOwner As String, isShared As Boolean, user1IsOwner As Boolean

    For i = 1 To rowCount
        transAmount = transArray(i, 5): transOwner = transArray(i, 7): isShared = CBool(transArray(i, 8))
        user1IsOwner = (StrComp(transOwner, m_User1, vbTextCompare) = 0)

        If isShared Then
            If transAmount < 0 Then ' Shared Expense
                 If user1IsOwner Then user1Net = user1Net + (transAmount / 2) Else user1Net = user1Net + (-transAmount / 2)
            Else ' Shared Income (Assuming 50/50 split)
                 user1Net = user1Net + (transAmount / 2)
            End If
        Else ' Individual Transaction
            If user1IsOwner Then user1Net = user1Net + transAmount
        End If
    Next i
    CalculateBalanceFromArray = -user1Net ' Positive means User1 owes
    Exit Function
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "BalanceCalculator.CalculateBalanceFromArray", Err.Number, Err.Description
    CalculateBalanceFromArray = 0
End Function


' Filter transactions by date range
Private Function FilterTransactionsByDate(ByVal transArray As Variant, _
                                          ByVal startDate As Date, _
                                          ByVal endDate As Date) As Variant
    On Error GoTo ErrorHandler
    FilterTransactionsByDate = Array(): If Not IsArray(transArray) Then Exit Function
    Dim lBound1 As Long, uBound1 As Long, lBound2 As Long, uBound2 As Long
    On Error Resume Next: lBound1 = LBound(transArray, 1): uBound1 = UBound(transArray, 1)
    lBound2 = LBound(transArray, 2): uBound2 = UBound(transArray, 2): On Error GoTo ErrorHandler
    If uBound1 = 0 Or uBound2 = 0 Then Exit Function

    Dim matchingRows As New Collection, i As Long, transDate As Date
    For i = lBound1 To uBound1
        On Error Resume Next: transDate = CDate(transArray(i, 2)): If Err.Number = 0 Then
            If transDate >= startDate And transDate <= endDate Then matchingRows.Add i
        Else: Err.Clear ' Ignore rows with bad dates for filtering
        End If: On Error GoTo ErrorHandler
    Next i
    If matchingRows.Count = 0 Then Exit Function ' Return empty array if no matches

    Dim result() As Variant: ReDim result(1 To matchingRows.Count, lBound2 To uBound2)
    Dim resultIndex As Long: resultIndex = 1
    Dim originalRowIndex As Variant, j As Long
    For Each originalRowIndex In matchingRows
        For j = lBound2 To uBound2: result(resultIndex, j) = transArray(originalRowIndex, j): Next j
        resultIndex = resultIndex + 1
    Next originalRowIndex
    FilterTransactionsByDate = result
    Exit Function
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "BalanceCalculator.FilterTransactionsByDate", Err.Number, Err.Description
    FilterTransactionsByDate = Array()
End Function

' Calculate category totals from transaction array
Private Function CalculateCategoryTotalsFromArray(ByVal dataToAnalyze As Variant) As Object
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    On Error GoTo ErrorHandler
    If Not IsArray(dataToAnalyze) Then GoTo NoCatData
    Dim rowCount As Long: On Error Resume Next: rowCount = UBound(dataToAnalyze, 1): On Error GoTo ErrorHandler
    If rowCount = 0 Then GoTo NoCatData

    Dim i As Long, transCategory As String, transAmount As Double
    For i = 1 To rowCount
        transCategory = Utilities.SafeString(dataToAnalyze(i, 4), "Uncategorized")
        transAmount = CDbl(dataToAnalyze(i, 5)) ' Assumes numeric, might need error check
        If transAmount < 0 Then result(transCategory) = result(transCategory) + Abs(transAmount)
    Next i
NoCatData:
    Set CalculateCategoryTotalsFromArray = result
    Exit Function
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "BalanceCalculator.CalculateCategoryTotalsFromArray", Err.Number, Err.Description
    Set result = CreateObject("Scripting.Dictionary"): Set CalculateCategoryTotalsFromArray = result
End Function


' Get amount spent by a specific user (expenses only)
Private Function GetUserSpentAmount(ByVal transArray As Variant, ByVal userName As String) As Double
    On Error GoTo ErrorHandler
    GetUserSpentAmount = 0: If Not IsArray(transArray) Then Exit Function
    Dim rowCount As Long: On Error Resume Next: rowCount = UBound(transArray, 1): On Error GoTo ErrorHandler
    If rowCount = 0 Then Exit Function

    Dim totalAmount As Double: totalAmount = 0
    Dim i As Long, transOwner As String, transAmount As Double
    For i = 1 To rowCount
        transOwner = transArray(i, 7): transAmount = transArray(i, 5)
        If StrComp(transOwner, userName, vbTextCompare) = 0 And transAmount < 0 Then totalAmount = totalAmount + Abs(transAmount)
    Next i
    GetUserSpentAmount = totalAmount
    Exit Function
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "BalanceCalculator.GetUserSpentAmount", Err.Number, Err.Description
    GetUserSpentAmount = 0
End Function

' Get total amount of shared expenses
Private Function GetTotalSharedAmount(ByVal transArray As Variant) As Double
    On Error GoTo ErrorHandler
    GetTotalSharedAmount = 0: If Not IsArray(transArray) Then Exit Function
    Dim rowCount As Long: On Error Resume Next: rowCount = UBound(transArray, 1): On Error GoTo ErrorHandler
    If rowCount = 0 Then Exit Function

    Dim totalAmount As Double: totalAmount = 0
    Dim i As Long, isShared As Boolean, transAmount As Double
    For i = 1 To rowCount
        isShared = CBool(transArray(i, 8)): transAmount = transArray(i, 5)
        If isShared And transAmount < 0 Then totalAmount = totalAmount + Abs(transAmount)
    Next i
    GetTotalSharedAmount = totalAmount
    Exit Function
ErrorHandler:
     If Not m_Logger Is Nothing Then m_Logger.LogError "BalanceCalculator.GetTotalSharedAmount", Err.Number, Err.Description
    GetTotalSharedAmount = 0
End Function

' Get total amount spent (all expenses)
Private Function GetTotalSpentAmount(ByVal transArray As Variant) As Double
     On Error GoTo ErrorHandler
    GetTotalSpentAmount = 0: If Not IsArray(transArray) Then Exit Function
    Dim rowCount As Long: On Error Resume Next: rowCount = UBound(transArray, 1): On Error GoTo ErrorHandler
    If rowCount = 0 Then Exit Function

    Dim totalAmount As Double: totalAmount = 0
    Dim i As Long, transAmount As Double
    For i = 1 To rowCount
        transAmount = transArray(i, 5)
        If transAmount < 0 Then totalAmount = totalAmount + Abs(transAmount)
    Next i
    GetTotalSpentAmount = totalAmount
    Exit Function
ErrorHandler:
     If Not m_Logger Is Nothing Then m_Logger.LogError "BalanceCalculator.GetTotalSpentAmount", Err.Number, Err.Description
    GetTotalSpentAmount = 0
End Function


'=========================================================================
' Properties
'=========================================================================

' Get/set User1 name
Public Property Get User1() As String
    If Not m_IsInitialized Then Me.Initialize
    User1 = m_User1
End Property

Public Property Let User1(ByVal value As String)
    m_User1 = value
End Property

' Get/set User2 name
Public Property Get User2() As String
    If Not m_IsInitialized Then Me.Initialize
    User2 = m_User2
End Property

Public Property Let User2(ByVal value As String)
    m_User2 = value
End Property

'=========================================================================
' Cleanup
'=========================================================================

Private Sub Class_Terminate()
    ' Clean up object references
    Set m_Logger = Nothing
End Sub