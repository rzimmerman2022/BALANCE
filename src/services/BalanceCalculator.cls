' File: src/services/BalanceCalculator.cls
'---------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BalanceCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' BalanceCalculator Class - Implements IBalanceCalculator
'
' Purpose: Performs financial calculations and analysis on transaction data.
' This class handles balance calculations, spending summaries by category,
' time-based analysis, and other financial metrics.
'
' Design Decisions:
' - Uses PredeclaredId = True to enable Singleton pattern
' - Implements IBalanceCalculator for loose coupling
' - Uses array-based operations for performance with large datasets
' - Implements dependency injection for the logger component
' - Includes comprehensive error handling

' Interface implementation
Implements IBalanceCalculator

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
    
    ' Get user names from settings (with defaults)
    m_User1 = GetSetting("BALANCE", "Settings", "User1", "User 1")
    m_User2 = GetSetting("BALANCE", "Settings", "User2", "User 2")
End Sub

' Initialize the calculator with dependencies
Private Sub IBalanceCalculator_Initialize(Optional ByVal logger As IErrorLogger = Nothing)
    On Error GoTo ErrorHandler
    
    ' Set logger if provided
    If Not logger Is Nothing Then
        Set m_Logger = logger
    End If
    
    ' Log initialization
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "BalanceCalculator.Initialize", "Balance Calculator initialized"
    End If
    
    m_IsInitialized = True
    
    Exit Sub
    
ErrorHandler:
    ' Can't use logger here if that's what caused the error
    If Not (m_Logger Is Nothing) And Err.Source <> "m_Logger" Then
        m_Logger.LogError "BalanceCalculator.Initialize", Err.Number, Err.Description
    End If
End Sub

' Public wrapper for Initialize
Public Sub Initialize(Optional ByVal logger As IErrorLogger = Nothing)
    IBalanceCalculator_Initialize logger
End Sub

'=========================================================================
' Interface Implementation - Balance Calculations
'=========================================================================

' Calculate the overall balance between users
Private Function IBalanceCalculator_CalculateBalance(ByVal transactions As Collection) As Double
    On Error GoTo ErrorHandler
    
    ' Ensure calculator is initialized
    If Not m_IsInitialized Then IBalanceCalculator_Initialize
    
    ' Convert transactions to array for faster processing
    Dim transArray() As Variant
    transArray = TransactionsToArray(transactions)
    
    ' Calculate balance using array operations
    Dim balance As Double
    balance = CalculateBalanceFromArray(transArray)
    
    IBalanceCalculator_CalculateBalance = balance
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BalanceCalculator.CalculateBalance", Err.Number, Err.Description
    End If
    IBalanceCalculator_CalculateBalance = 0
End Function

' Public wrapper for CalculateBalance
Public Function CalculateBalance(ByVal transactions As Collection) As Double
    CalculateBalance = IBalanceCalculator_CalculateBalance(transactions)
End Function

' Calculate who owes whom and how much
Private Function IBalanceCalculator_CalculateBalanceSummary(ByVal transactions As Collection) As Object
    On Error GoTo ErrorHandler
    
    ' Ensure calculator is initialized
    If Not m_IsInitialized Then IBalanceCalculator_Initialize
    
    ' Create dictionary for results
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    ' Convert transactions to array for faster processing
    Dim transArray() As Variant
    transArray = TransactionsToArray(transactions)
    
    ' Calculate balance
    Dim balance As Double
    balance = CalculateBalanceFromArray(transArray)
    
    ' Determine who owes whom
    Dim owes As String, owed As String
    Dim amount As Double
    
    amount = Abs(balance)
    
    If balance > 0 Then
        ' User1 owes User2
        owes = m_User1
        owed = m_User2
    ElseIf balance < 0 Then
        ' User2 owes User1
        owes = m_User2
        owed = m_User1
    Else
        ' Balance is even
        owes = ""
        owed = ""
    End If
    
    ' Build result dictionary
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
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BalanceCalculator.CalculateBalanceSummary", Err.Number, Err.Description
    End If
    
    ' Return empty dictionary on error
    Set result = CreateObject("Scripting.Dictionary")
    Set IBalanceCalculator_CalculateBalanceSummary = result
End Function

' Public wrapper for CalculateBalanceSummary
Public Function CalculateBalanceSummary(ByVal transactions As Collection) As Object
    Set CalculateBalanceSummary = IBalanceCalculator_CalculateBalanceSummary(transactions)
End Function

' Calculate spending by category
Private Function IBalanceCalculator_CalculateCategorySummary(ByVal transactions As Collection, _
                                                           Optional ByVal startDate As Date, _
                                                           Optional ByVal endDate As Date) As Object
    On Error GoTo ErrorHandler
    
    ' Ensure calculator is initialized
    If Not m_IsInitialized Then IBalanceCalculator_Initialize
    
    ' Use current date range if not specified
    If startDate = 0 Then startDate = DateSerial(Year(Date) - 1, Month(Date), 1)
    If endDate = 0 Then endDate = Date
    
    ' Create dictionary for results
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    ' Convert transactions to array for faster processing
    Dim transArray() As Variant
    transArray = TransactionsToArray(transactions)
    
    ' Filter transactions by date range
    Dim filteredArray() As Variant
    filteredArray = FilterTransactionsByDate(transArray, startDate, endDate)
    
    ' Calculate category totals
    result = CalculateCategoryTotalsFromArray(filteredArray)
    
    ' Calculate percentages
    Dim totalSpent As Double
    totalSpent = GetTotalSpentAmount(filteredArray)
    
    Dim category As Variant
    For Each category In result.Keys
        Dim categoryAmount As Double
        categoryAmount = result(category)
        
        ' Add percentage to the dictionary with a different key
        If totalSpent > 0 Then
            result.Add category & "_Percent", Round(categoryAmount / totalSpent * 100, 2)
        Else
            result.Add category & "_Percent", 0
        End If
    Next category
    
    ' Add total to the results
    result.Add "Total", totalSpent
    
    Set IBalanceCalculator_CalculateCategorySummary = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BalanceCalculator.CalculateCategorySummary", Err.Number, Err.Description
    End If
    
    ' Return empty dictionary on error
    Set result = CreateObject("Scripting.Dictionary")
    Set IBalanceCalculator_CalculateCategorySummary = result
End Function

' Public wrapper for CalculateCategorySummary
Public Function CalculateCategorySummary(ByVal transactions As Collection, _
                                       Optional ByVal startDate As Date, _
                                       Optional ByVal endDate As Date) As Object
    Set CalculateCategorySummary = IBalanceCalculator_CalculateCategorySummary(transactions, startDate, endDate)
End Function

' Calculate spending by day of week
Private Function IBalanceCalculator_CalculateDayOfWeekSummary(ByVal transactions As Collection, _
                                                             Optional ByVal startDate As Date, _
                                                             Optional ByVal endDate As Date) As Object
    On Error GoTo ErrorHandler
    
    ' Ensure calculator is initialized
    If Not m_IsInitialized Then IBalanceCalculator_Initialize
    
    ' Use current date range if not specified
    If startDate = 0 Then startDate = DateSerial(Year(Date) - 1, Month(Date), 1)
    If endDate = 0 Then endDate = Date
    
    ' Create dictionary for results
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    ' Initialize day of week counters
    Dim i As Integer
    For i = 1 To 7
        result.Add WeekdayName(i, True), 0
    Next i
    
    ' Convert transactions to array for faster processing
    Dim transArray() As Variant
    transArray = TransactionsToArray(transactions)
    
    ' Filter transactions by date range
    Dim filteredArray() As Variant
    filteredArray = FilterTransactionsByDate(transArray, startDate, endDate)
    
    ' Calculate day of week totals
    Dim rowCount As Long
    rowCount = UBound(filteredArray, 1)
    
    For i = 1 To rowCount
        Dim transDate As Date
        transDate = filteredArray(i, 2) ' Date is in column 2
        
        Dim transAmount As Double
        transAmount = filteredArray(i, 5) ' Amount is in column 5
        
        Dim dayName As String
        dayName = WeekdayName(Weekday(transDate), True)
        
        ' Only count expenses (negative amounts)
        If transAmount < 0 Then
            result(dayName) = result(dayName) + Abs(transAmount)
        End If
    Next i
    
    ' Calculate total spending
    Dim totalSpent As Double
    totalSpent = 0
    
    For i = 1 To 7
        Dim dayOfWeek As String
        dayOfWeek = WeekdayName(i, True)
        
        totalSpent = totalSpent + result(dayOfWeek)
    Next i
    
    ' Calculate percentages
    For i = 1 To 7
        dayOfWeek = WeekdayName(i, True)
        
        ' Add percentage to the dictionary with a different key
        If totalSpent > 0 Then
            result.Add dayOfWeek & "_Percent", Round(result(dayOfWeek) / totalSpent * 100, 2)
        Else
            result.Add dayOfWeek & "_Percent", 0
        End If
    Next i
    
    ' Add total to the results
    result.Add "Total", totalSpent
    
    Set IBalanceCalculator_CalculateDayOfWeekSummary = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BalanceCalculator.CalculateDayOfWeekSummary", Err.Number, Err.Description
    End If
    
    ' Return empty dictionary on error
    Set result = CreateObject("Scripting.Dictionary")
    Set IBalanceCalculator_CalculateDayOfWeekSummary = result
End Function

' Public wrapper for CalculateDayOfWeekSummary
Public Function CalculateDayOfWeekSummary(ByVal transactions As Collection, _
                                        Optional ByVal startDate As Date, _
                                        Optional ByVal endDate As Date) As Object
    Set CalculateDayOfWeekSummary = IBalanceCalculator_CalculateDayOfWeekSummary(transactions, startDate, endDate)
End Function

' Calculate spending by month
Private Function IBalanceCalculator_CalculateMonthSummary(ByVal transactions As Collection, _
                                                        Optional ByVal monthCount As Integer = 12) As Object
    On Error GoTo ErrorHandler
    
    ' Ensure calculator is initialized
    If Not m_IsInitialized Then IBalanceCalculator_Initialize
    
    ' Create dictionary for results
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    ' Determine date range
    Dim endDate As Date
    endDate = DateSerial(Year(Date), Month(Date) + 1, 0) ' Last day of current month
    
    Dim startDate As Date
    startDate = DateSerial(Year(endDate) - Int(monthCount / 12), Month(endDate) - (monthCount Mod 12), 1)
    If Month(endDate) - (monthCount Mod 12) <= 0 Then
        startDate = DateSerial(Year(endDate) - Int(monthCount / 12) - 1, Month(endDate) - (monthCount Mod 12) + 12, 1)
    End If
    
    ' Initialize month counters
    Dim currentDate As Date
    currentDate = startDate
    
    Do While DateDiff("m", currentDate, endDate) >= 0
        Dim monthKey As String
        monthKey = Format(currentDate, "yyyy-mm")
        
        result.Add monthKey, 0
        result.Add monthKey & "_Income", 0
        result.Add monthKey & "_Expense", 0
        
        currentDate = DateAdd("m", 1, currentDate)
    Loop
    
    ' Convert transactions to array for faster processing
    Dim transArray() As Variant
    transArray = TransactionsToArray(transactions)
    
    ' Filter transactions by date range
    Dim filteredArray() As Variant
    filteredArray = FilterTransactionsByDate(transArray, startDate, endDate)
    
    ' Calculate monthly totals
    Dim i As Long
    Dim rowCount As Long
    rowCount = UBound(filteredArray, 1)
    
    For i = 1 To rowCount
        Dim transDate As Date
        transDate = filteredArray(i, 2) ' Date is in column 2
        
        Dim transAmount As Double
        transAmount = filteredArray(i, 5) ' Amount is in column 5
        
        monthKey = Format(transDate, "yyyy-mm")
        
        ' Only process transactions in our result months
        If result.Exists(monthKey) Then
            ' Update net amount
            result(monthKey) = result(monthKey) + transAmount
            
            ' Update income or expense
            If transAmount > 0 Then
                result(monthKey & "_Income") = result(monthKey & "_Income") + transAmount
            Else
                result(monthKey & "_Expense") = result(monthKey & "_Expense") + Abs(transAmount)
            End If
        End If
    Next i
    
    Set IBalanceCalculator_CalculateMonthSummary = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BalanceCalculator.CalculateMonthSummary", Err.Number, Err.Description
    End If
    
    ' Return empty dictionary on error
    Set result = CreateObject("Scripting.Dictionary")
    Set IBalanceCalculator_CalculateMonthSummary = result
End Function

' Public wrapper for CalculateMonthSummary
Public Function CalculateMonthSummary(ByVal transactions As Collection, _
                                    Optional ByVal monthCount As Integer = 12) As Object
    Set CalculateMonthSummary = IBalanceCalculator_CalculateMonthSummary(transactions, monthCount)
End Function

' Calculate daily average spending
Private Function IBalanceCalculator_CalculateDailyAverageSpending(ByVal transactions As Collection, _
                                                                 Optional ByVal startDate As Date, _
                                                                 Optional ByVal endDate As Date) As Double
    On Error GoTo ErrorHandler
    
    ' Ensure calculator is initialized
    If Not m_IsInitialized Then IBalanceCalculator_Initialize
    
    ' Use current date range if not specified
    If startDate = 0 Then startDate = DateSerial(Year(Date) - 1, Month(Date), 1)
    If endDate = 0 Then endDate = Date
    
    ' Convert transactions to array for faster processing
    Dim transArray() As Variant
    transArray = TransactionsToArray(transactions)
    
    ' Filter transactions by date range
    Dim filteredArray() As Variant
    filteredArray = FilterTransactionsByDate(transArray, startDate, endDate)
    
    ' Calculate total expenses (negative amounts)
    Dim totalExpenses As Double
    totalExpenses = 0
    
    Dim i As Long
    Dim rowCount As Long
    rowCount = UBound(filteredArray, 1)
    
    For i = 1 To rowCount
        Dim transAmount As Double
        transAmount = filteredArray(i, 5) ' Amount is in column 5
        
        ' Only count expenses (negative amounts)
        If transAmount < 0 Then
            totalExpenses = totalExpenses + Abs(transAmount)
        End If
    Next i
    
    ' Calculate number of days in date range
    Dim dayCount As Integer
    dayCount = DateDiff("d", startDate, endDate) + 1
    
    ' Calculate daily average
    Dim dailyAverage As Double
    If dayCount > 0 Then
        dailyAverage = totalExpenses / dayCount
    Else
        dailyAverage = 0
    End If
    
    IBalanceCalculator_CalculateDailyAverageSpending = dailyAverage
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BalanceCalculator.CalculateDailyAverageSpending", Err.Number, Err.Description
    End If
    IBalanceCalculator_CalculateDailyAverageSpending = 0
End Function

' Public wrapper for CalculateDailyAverageSpending
Public Function CalculateDailyAverageSpending(ByVal transactions As Collection, _
                                             Optional ByVal startDate As Date, _
                                             Optional ByVal endDate As Date) As Double
    CalculateDailyAverageSpending = IBalanceCalculator_CalculateDailyAverageSpending(transactions, startDate, endDate)
End Function

' Calculate spending trends (month-over-month changes)
Private Function IBalanceCalculator_CalculateSpendingTrends(ByVal transactions As Collection, _
                                                           Optional ByVal monthCount As Integer = 12) As Object
    On Error GoTo ErrorHandler
    
    ' Ensure calculator is initialized
    If Not m_IsInitialized Then IBalanceCalculator_Initialize
    
    ' Create dictionary for results
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    ' Get monthly summary
    Dim monthSummary As Object
    Set monthSummary = IBalanceCalculator_CalculateMonthSummary(transactions, monthCount)
    
    ' Calculate month-over-month changes
    Dim months As Collection
    Set months = GetSortedKeys(monthSummary)
    
    Dim i As Long
    Dim monthCount As Long
    monthCount = months.Count
    
    ' Need at least 2 months to calculate trends
    If monthCount < 2 Then
        Set IBalanceCalculator_CalculateSpendingTrends = result
        Exit Function
    End If
    
    ' Calculate change for each month after the first
    For i = 2 To monthCount
        Dim currentMonth As String
        currentMonth = months(i)
        
        Dim previousMonth As String
        previousMonth = months(i - 1)
        
        ' Check if keys exist
        If monthSummary.Exists(currentMonth & "_Expense") And monthSummary.Exists(previousMonth & "_Expense") Then
            Dim currentExpense As Double
            currentExpense = monthSummary(currentMonth & "_Expense")
            
            Dim previousExpense As Double
            previousExpense = monthSummary(previousMonth & "_Expense")
            
            Dim changeAmount As Double
            changeAmount = currentExpense - previousExpense
            
            Dim changePercent As Double
            If previousExpense > 0 Then
                changePercent = Round(changeAmount / previousExpense * 100, 2)
            Else
                changePercent = 0
            End If
            
            ' Add to results
            result.Add currentMonth & "_Change", changeAmount
            result.Add currentMonth & "_ChangePercent", changePercent
        End If
    Next i
    
    Set IBalanceCalculator_CalculateSpendingTrends = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BalanceCalculator.CalculateSpendingTrends", Err.Number, Err.Description
    End If
    
    ' Return empty dictionary on error
    Set result = CreateObject("Scripting.Dictionary")
    Set IBalanceCalculator_CalculateSpendingTrends = result
End Function

' Public wrapper for CalculateSpendingTrends
Public Function CalculateSpendingTrends(ByVal transactions As Collection, _
                                      Optional ByVal monthCount As Integer = 12) As Object
    Set CalculateSpendingTrends = IBalanceCalculator_CalculateSpendingTrends(transactions, monthCount)
End Function

'=========================================================================
' Helper Methods - Array Operations
'=========================================================================

' Convert transactions collection to array for faster processing
Private Function TransactionsToArray(ByVal transactions As Collection) As Variant
    On Error GoTo ErrorHandler
    
    Dim transCount As Long
    transCount = transactions.Count
    
    ' If no transactions, return empty array
    If transCount = 0 Then
        TransactionsToArray = Array()
        Exit Function
    End If
    
    ' Create array with specific structure:
    ' Column 1: ID
    ' Column 2: Date
    ' Column 3: Merchant
    ' Column 4: Category
    ' Column 5: Amount
    ' Column 6: Account
    ' Column 7: Owner
    ' Column 8: IsShared
    ' Column 9: Notes
    ' Column 10: SourceFile
    Dim result() As Variant
    ReDim result(1 To transCount, 1 To 10)
    
    Dim i As Long
    i = 1
    
    Dim trans As Transaction
    For Each trans In transactions
        result(i, 1) = trans.ID
        result(i, 2) = trans.TransactionDate
        result(i, 3) = trans.Merchant
        result(i, 4) = trans.Category
        result(i, 5) = trans.Amount
        result(i, 6) = trans.Account
        result(i, 7) = trans.Owner
        result(i, 8) = trans.IsShared
        result(i, 9) = trans.Notes
        result(i, 10) = trans.SourceFile
        
        i = i + 1
    Next trans
    
    TransactionsToArray = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BalanceCalculator.TransactionsToArray", Err.Number, Err.Description
    End If
    TransactionsToArray = Array()
End Function

' Calculate balance from transaction array
Private Function CalculateBalanceFromArray(ByVal transArray As Variant) As Double
    On Error GoTo ErrorHandler
    
    ' Check if array is empty
    If Not IsArray(transArray) Then
        CalculateBalanceFromArray = 0
        Exit Function
    End If
    
    Dim rowCount As Long
    On Error Resume Next
    rowCount = UBound(transArray, 1)
    On Error GoTo ErrorHandler
    
    If rowCount = 0 Then
        CalculateBalanceFromArray = 0
        Exit Function
    End If
    
    ' Track share of each user
    Dim user1Amount As Double
    Dim user2Amount As Double
    
    user1Amount = 0
    user2Amount = 0
    
    ' Process each transaction
    Dim i As Long
    For i = 1 To rowCount
        Dim transAmount As Double
        transAmount = transArray(i, 5) ' Amount is in column 5
        
        Dim transOwner As String
        transOwner = transArray(i, 7) ' Owner is in column 7
        
        Dim isShared As Boolean
        isShared = transArray(i, 8) ' IsShared is in column 8
        
        ' Handle shared expenses
        If isShared Then
            ' If shared, each user bears half the expense regardless of who paid
            If transAmount < 0 Then ' Expense
                ' The owner pays the full amount but should only bear half
                If transOwner = m_User1 Then
                    user1Amount = user1Amount + (transAmount / 2) ' Half of expense
                    user2Amount = user2Amount + (transAmount / 2) ' Half of expense
                ElseIf transOwner = m_User2 Then
                    user1Amount = user1Amount + (transAmount / 2) ' Half of expense
                    user2Amount = user2Amount + (transAmount / 2) ' Half of expense
                End If
            Else ' Income
                ' Shared income is split evenly
                If transOwner = m_User1 Then
                    user1Amount = user1Amount + (transAmount / 2) ' Half of income
                    user2Amount = user2Amount + (transAmount / 2) ' Half of income
                ElseIf transOwner = m_User2 Then
                    user1Amount = user1Amount + (transAmount / 2) ' Half of income
                    user2Amount = user2Amount + (transAmount / 2) ' Half of income
                End If
            End If
        Else
            ' If not shared, only the owner is affected
            If transOwner = m_User1 Then
                user1Amount = user1Amount + transAmount
            ElseIf transOwner = m_User2 Then
                user2Amount = user2Amount + transAmount
            End If
        End If
    Next i
    
    ' Calculate final balance
    ' A positive balance means User1 owes User2
    ' A negative balance means User2 owes User1
    CalculateBalanceFromArray = user1Amount - user2Amount
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BalanceCalculator.CalculateBalanceFromArray", Err.Number, Err.Description
    End If
    CalculateBalanceFromArray = 0
End Function

' Filter transactions by date range
Private Function FilterTransactionsByDate(ByVal transArray As Variant, _
                                         ByVal startDate As Date, _
                                         ByVal endDate As Date) As Variant
    On Error GoTo ErrorHandler
    
    ' Check if array is empty
    If Not IsArray(transArray) Then
        FilterTransactionsByDate = Array()
        Exit Function
    End If
    
    Dim rowCount As Long
    On Error Resume Next
    rowCount = UBound(transArray, 1)
    On Error GoTo ErrorHandler
    
    If rowCount = 0 Then
        FilterTransactionsByDate = Array()
        Exit Function
    End If
    
    ' Count transactions in date range
    Dim filterCount As Long
    filterCount = 0
    
    Dim i As Long
    For i = 1 To rowCount
        Dim transDate As Date
        transDate = transArray(i, 2) ' Date is in column 2
        
        If transDate >= startDate And transDate <= endDate Then
            filterCount = filterCount + 1
        End If
    Next i
    
    ' If no transactions in range, return empty array
    If filterCount = 0 Then
        FilterTransactionsByDate = Array()
        Exit Function
    End If
    
    ' Create filtered array
    Dim result() As Variant
    ReDim result(1 To filterCount, 1 To UBound(transArray, 2))
    
    Dim resultIndex As Long
    resultIndex = 1
    
    ' Fill filtered array
    For i = 1 To rowCount
        transDate = transArray(i, 2) ' Date is in column 2
        
        If transDate >= startDate And transDate <= endDate Then
            ' Copy all columns
            Dim j As Long
            For j = 1 To UBound(transArray, 2)
                result(resultIndex, j) = transArray(i, j)
            Next j
            
            resultIndex = resultIndex + 1
        End If
    Next i
    
    FilterTransactionsByDate = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BalanceCalculator.FilterTransactionsByDate", Err.Number, Err.Description
    End If
    FilterTransactionsByDate = Array()
End Function

' Calculate category totals from transaction array
Private Function CalculateCategoryTotalsFromArray(ByVal transArray As Variant) As Object
    On Error GoTo ErrorHandler
    
    ' Create dictionary for results
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    ' Check if array is empty
    If Not IsArray(transArray) Then
        Set CalculateCategoryTotalsFromArray = result
        Exit Function
    End If
    
    Dim rowCount As Long
    On Error Resume Next
    rowCount = UBound(transArray, 1)
    On Error GoTo ErrorHandler
    
    If rowCount = 0 Then
        Set CalculateCategoryTotalsFromArray = result
        Exit Function
    End If
    
    ' Process each transaction
    Dim i As Long
    For i = 1 To rowCount
        Dim transCategory As String
        transCategory = transArray(i, 4) ' Category is in column 4
        
        Dim transAmount As Double
        transAmount = transArray(i, 5) ' Amount is in column 5
        
        ' Only count expenses (negative amounts)
        If transAmount < 0 Then
            ' Ensure category is not empty
            If Len(Trim$(transCategory)) = 0 Then
                transCategory = "Uncategorized"
            End If
            
            ' Update category total
            If Not result.Exists(transCategory) Then
                result.Add transCategory, 0
            End If
            
            result(transCategory) = result(transCategory) + Abs(transAmount)
        End If
    Next i
    
    Set CalculateCategoryTotalsFromArray = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BalanceCalculator.CalculateCategoryTotalsFromArray", Err.Number, Err.Description
    End If
    
    Set result = CreateObject("Scripting.Dictionary")
    Set CalculateCategoryTotalsFromArray = result
End Function

' Get amount spent by a specific user
Private Function GetUserSpentAmount(ByVal transArray As Variant, ByVal userName As String) As Double
    On Error GoTo ErrorHandler
    
    ' Check if array is empty
    If Not IsArray(transArray) Then
        GetUserSpentAmount = 0
        Exit Function
    End If
    
    Dim rowCount As Long
    On Error Resume Next
    rowCount = UBound(transArray, 1)
    On Error GoTo ErrorHandler
    
    If rowCount = 0 Then
        GetUserSpentAmount = 0
        Exit Function
    End If
    
    ' Calculate amount
    Dim totalAmount As Double
    totalAmount = 0
    
    ' Process each transaction
    Dim i As Long
    For i = 1 To rowCount
        Dim transOwner As String
        transOwner = transArray(i, 7) ' Owner is in column 7
        
        Dim transAmount As Double
        transAmount = transArray(i, 5) ' Amount is in column 5
        
        ' Only count expenses (negative amounts) for this user
        If transOwner = userName And transAmount < 0 Then
            totalAmount = totalAmount + Abs(transAmount)
        End If
    Next i
    
    GetUserSpentAmount = totalAmount
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BalanceCalculator.GetUserSpentAmount", Err.Number, Err.Description
    End If
    GetUserSpentAmount = 0
End Function

' Get total amount of shared expenses
Private Function GetTotalSharedAmount(ByVal transArray As Variant) As Double
    On Error GoTo ErrorHandler
    
    ' Check if array is empty
    If Not IsArray(transArray) Then
        GetTotalSharedAmount = 0
        Exit Function
    End If
    
    Dim rowCount As Long
    On Error Resume Next
    rowCount = UBound(transArray, 1)
    On Error GoTo ErrorHandler
    
    If rowCount = 0 Then
        GetTotalSharedAmount = 0
        Exit Function
    End If
    
    ' Calculate amount
    Dim totalAmount As Double
    totalAmount = 0
    
    ' Process each transaction
    Dim i As Long
    For i = 1 To rowCount
        Dim isShared As Boolean
        isShared = transArray(i, 8) ' IsShared is in column 8
        
        Dim transAmount As Double
        transAmount = transArray(i, 5) ' Amount is in column 5
        
        ' Only count shared expenses (negative amounts)
        If isShared And transAmount < 0 Then
            totalAmount = totalAmount + Abs(transAmount)
        End If
    Next i
    
    GetTotalSharedAmount = totalAmount
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BalanceCalculator.GetTotalSharedAmount", Err.Number, Err.Description
    End If
    GetTotalSharedAmount = 0
End Function

' Get total amount spent
Private Function GetTotalSpentAmount(ByVal transArray As Variant) As Double
    On Error GoTo ErrorHandler
    
    ' Check if array is empty
    If Not IsArray(transArray) Then
        GetTotalSpentAmount = 0
        Exit Function
    End If
    
    Dim rowCount As Long
    On Error Resume Next
    rowCount = UBound(transArray, 1)
    On Error GoTo ErrorHandler
    
    If rowCount = 0 Then
        GetTotalSpentAmount = 0
        Exit Function
    End If
    
    ' Calculate amount
    Dim totalAmount As Double
    totalAmount = 0
    
    ' Process each transaction
    Dim i As Long
    For i = 1 To rowCount
        Dim transAmount As Double
        transAmount = transArray(i, 5) ' Amount is in column 5
        
        ' Only count expenses (negative amounts)
        If transAmount < 0 Then
            totalAmount = totalAmount + Abs(transAmount)
        End If
    Next i
    
    GetTotalSpentAmount = totalAmount
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BalanceCalculator.GetTotalSpentAmount", Err.Number, Err.Description
    End If
    GetTotalSpentAmount = 0
End Function

' Get keys from a dictionary, sorted alphabetically
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
        m_Logger.LogError "BalanceCalculator.GetSortedKeys", Err.Number, Err.Description
    End If
    
    Set result = New Collection
    Set GetSortedKeys = result
End Function

'=========================================================================
' Properties
'=========================================================================

' Get/set User1 name
Public Property Get User1() As String
    User1 = m_User1
End Property

Public Property Let User1(ByVal value As String)
    m_User1 = value
End Property

' Get/set User2 name
Public Property Get User2() As String
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