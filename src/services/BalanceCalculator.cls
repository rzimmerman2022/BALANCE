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
' BalanceCalculator Class - Handles calculating balances between users

' Define result types for return values
Public Type BalanceSummary
    TotalAmount As Currency
    User1Share As Currency
    User2Share As Currency
    User1Paid As Currency
    User2Paid As Currency
    NetBalance As Currency  ' Positive: User1 owes User2, Negative: User2 owes User1
    WhoOwes As String
    OwedAmount As Currency
End Type

Public Type CategorySummary
    Category As String
    TotalAmount As Currency
    Percentage As Double
End Type

Public Type MonthSummary
    MonthDate As Date
    TotalAmount As Currency
    TransactionCount As Long
End Type

Public Type DayOfWeekSummary
    DayName As String
    DayNumber As Integer
    TotalAmount As Currency
    TransactionCount As Long
End Type

' ===== Balance Calculation =====

' Calculate balance summary from a collection of transactions
Public Function CalculateBalance(transactions As Collection) As BalanceSummary
    On Error GoTo ErrorHandler
    
    Dim result As BalanceSummary
    Dim trans As Transaction
    Dim user1Amount As Currency
    Dim user2Amount As Currency
    
    ' Initialize result
    result.TotalAmount = 0
    result.User1Share = 0
    result.User2Share = 0
    result.User1Paid = 0
    result.User2Paid = 0
    result.NetBalance = 0
    result.WhoOwes = ""
    result.OwedAmount = 0
    
    ' Process each transaction
    For Each trans In transactions
        ' Update total amount (use negative for expenses, positive for income)
        result.TotalAmount = result.TotalAmount + trans.Amount
        
        ' Calculate split
        trans.CalculateSplit user1Amount, user2Amount
        
        ' Update shares
        result.User1Share = result.User1Share + user1Amount
        result.User2Share = result.User2Share + user2Amount
        
        ' Update paid amounts
        If LCase(trans.Owner) = LCase(AppSettings.User1Name) Then
            result.User1Paid = result.User1Paid + Abs(trans.Amount)
        Else
            result.User2Paid = result.User2Paid + Abs(trans.Amount)
        End If
    Next trans
    
    ' Calculate net balance (positive: User1 owes User2, negative: User2 owes User1)
    result.NetBalance = result.User1Share - result.User2Share
    
    ' Determine who owes whom
    If Abs(result.NetBalance) < 0.01 Then
        result.WhoOwes = "No one owes anything"
        result.OwedAmount = 0
    ElseIf result.NetBalance > 0 Then
        result.WhoOwes = AppSettings.User1Name & " owes " & AppSettings.User2Name
        result.OwedAmount = result.NetBalance
    Else
        result.WhoOwes = AppSettings.User2Name & " owes " & AppSettings.User1Name
        result.OwedAmount = Abs(result.NetBalance)
    End If
    
    CalculateBalance = result
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "BalanceCalculator.CalculateBalance", Err.Number, Err.Description
    ' Return the initialized result structure in case of error
    CalculateBalance = result
End Function

' ===== Category Analysis =====

' Get spending by category
Public Function GetCategorySummary(transactions As Collection, Optional topCount As Long = 8) As Collection
    On Error GoTo ErrorHandler
    
    Dim result As New Collection
    Dim categoryDict As Object
    Set categoryDict = CreateObject("Scripting.Dictionary")
    Dim totalExpenses As Currency
    
    totalExpenses = 0
    
    ' Group by category
    Dim trans As Transaction
    For Each trans In transactions
        ' Only include expenses
        If trans.IsExpense() Then
            Dim category As String
            Dim amount As Currency
            
            ' Get category name (use "Uncategorized" if empty)
            category = IIf(Len(trans.Category) > 0, trans.Category, "Uncategorized")
            
            ' Get absolute amount (positive value)
            amount = trans.GetAbsoluteAmount()
            
            ' Add to total
            totalExpenses = totalExpenses + amount
            
            ' Add or update category
            If categoryDict.Exists(category) Then
                categoryDict(category) = categoryDict(category) + amount
            Else
                categoryDict.Add category, amount
            End If
        End If
    Next trans
    
    ' Convert to array for sorting
    Dim categories() As String
    Dim amounts() As Currency
    Dim i As Long
    
    ReDim categories(1 To categoryDict.Count)
    ReDim amounts(1 To categoryDict.Count)
    
    i = 1
    Dim key As Variant
    For Each key In categoryDict.Keys
        categories(i) = key
        amounts(i) = categoryDict(key)
        i = i + 1
    Next key
    
    ' Sort by amount (using QuickSort instead of bubble sort for better performance)
    QuickSortCategories categories, amounts, 1, UBound(categories)
    
    ' Create category summaries
    Dim displayCount As Long
    Dim otherAmount As Currency
    
    ' Determine how many categories to display
    displayCount = WorksheetFunction.Min(topCount, UBound(categories))
    
    ' Add top categories
    For i = 1 To displayCount
        Dim catSummary As CategorySummary
        
        catSummary.Category = categories(i)
        catSummary.TotalAmount = amounts(i)
        
        ' Calculate percentage (handle division by zero)
        If totalExpenses > 0 Then
            catSummary.Percentage = amounts(i) / totalExpenses
        Else
            catSummary.Percentage = 0
        End If
        
        ' Add to result
        result.Add catSummary
    Next i
    
    ' Add "Other" category if needed
    If UBound(categories) > displayCount Then
        Dim otherSummary As CategorySummary
        
        otherSummary.Category = "Other Categories"
        otherSummary.TotalAmount = 0
        
        ' Sum the remaining categories
        For i = displayCount + 1 To UBound(categories)
            otherSummary.TotalAmount = otherSummary.TotalAmount + amounts(i)
        Next i
        
        ' Calculate percentage
        If totalExpenses > 0 Then
            otherSummary.Percentage = otherSummary.TotalAmount / totalExpenses
        Else
            otherSummary.Percentage = 0
        End If
        
        ' Add to result
        result.Add otherSummary
    End If
    
    Set GetCategorySummary = result
    
    ' Clean up
    Set categoryDict = Nothing
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "BalanceCalculator.GetCategorySummary", Err.Number, Err.Description
    Set GetCategorySummary = New Collection ' Return empty collection in case of error
End Function

' QuickSort algorithm for categories (much faster than bubble sort)
Private Sub QuickSortCategories(ByRef categories() As String, ByRef amounts() As Currency, ByVal low As Long, ByVal high As Long)
    On Error GoTo ErrorHandler
    
    Dim pivot As Currency
    Dim tempAmount As Currency
    Dim tempCategory As String
    Dim i As Long, j As Long
    
    If low < high Then
        pivot = amounts((low + high) \ 2)
        i = low
        j = high
        
        Do
            Do While amounts(i) > pivot
                i = i + 1
            Loop
            
            Do While amounts(j) < pivot
                j = j - 1
            Loop
            
            If i <= j Then
                ' Swap amounts
                tempAmount = amounts(i)
                amounts(i) = amounts(j)
                amounts(j) = tempAmount
                
                ' Swap categories
                tempCategory = categories(i)
                categories(i) = categories(j)
                categories(j) = tempCategory
                
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j
        
        If low < j Then QuickSortCategories categories, amounts, low, j
        If i < high Then QuickSortCategories categories, amounts, i, high
    End If
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BalanceCalculator.QuickSortCategories", Err.Number, Err.Description
End Sub

' ===== Monthly Analysis =====

' Get monthly spending summary
Public Function GetMonthlySummary(transactions As Collection, Optional months As Long = 6) As Collection
    On Error GoTo ErrorHandler
    
    Dim result As New Collection
    Dim monthDict As Object
    Set monthDict = CreateObject("Scripting.Dictionary")
    
    ' Group by month
    Dim trans As Transaction
    For Each trans In transactions
        ' Only include expenses
        If trans.IsExpense() Then
            Dim monthKey As String
            Dim monthDate As Date
            Dim amount As Currency
            
            ' Get first day of month
            monthDate = DateSerial(Year(trans.TransactionDate), Month(trans.TransactionDate), 1)
            monthKey = Format(monthDate, "yyyy-mm")
            
            ' Get absolute amount
            amount = trans.GetAbsoluteAmount()
            
            ' Add or update month
            If monthDict.Exists(monthKey) Then
                monthDict(monthKey)(0) = monthDict(monthKey)(0) + amount  ' Sum
                monthDict(monthKey)(1) = monthDict(monthKey)(1) + 1       ' Count
                monthDict(monthKey)(2) = monthDate                        ' Date
            Else
                Dim newMonth(0 To 2) As Variant
                newMonth(0) = amount  ' Sum
                newMonth(1) = 1       ' Count
                newMonth(2) = monthDate  ' Date
                monthDict.Add monthKey, newMonth
            End If
        End If
    Next trans
    
    ' Sort months chronologically (newest first)
    Dim sortedMonths() As String
    ReDim sortedMonths(1 To monthDict.Count)
    
    Dim i As Long
    i = 1
    Dim key As Variant
    For Each key In monthDict.Keys
        sortedMonths(i) = key
        i = i + 1
    Next key
    
    ' Sort (using QuickSort for better performance)
    QuickSortMonths sortedMonths, 1, UBound(sortedMonths)
    
    ' Create month summaries (most recent first)
    Dim displayMonths As Long
    displayMonths = WorksheetFunction.Min(months, UBound(sortedMonths))
    
    For i = 1 To displayMonths
        Dim monthSummary As MonthSummary
        Dim monthKey As String
        
        monthKey = sortedMonths(i)
        
        monthSummary.MonthDate = monthDict(monthKey)(2)
        monthSummary.TotalAmount = monthDict(monthKey)(0)
        monthSummary.TransactionCount = monthDict(monthKey)(1)
        
        ' Add to result
        result.Add monthSummary
    Next i
    
    Set GetMonthlySummary = result
    
    ' Clean up
    Set monthDict = Nothing
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "BalanceCalculator.GetMonthlySummary", Err.Number, Err.Description
    Set GetMonthlySummary = New Collection ' Return empty collection in case of error
End Function

' QuickSort algorithm for months (descending order - newest first)
Private Sub QuickSortMonths(ByRef months() As String, ByVal low As Long, ByVal high As Long)
    On Error GoTo ErrorHandler
    
    Dim pivot As String
    Dim temp As String
    Dim i As Long, j As Long
    
    If low < high Then
        pivot = months((low + high) \ 2)
        i = low
        j = high
        
        Do
            Do While months(i) > pivot
                i = i + 1
            Loop
            
            Do While months(j) < pivot
                j = j - 1
            Loop
            
            If i <= j Then
                ' Swap months
                temp = months(i)
                months(i) = months(j)
                months(j) = temp
                
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j
        
        If low < j Then QuickSortMonths months, low, j
        If i < high Then QuickSortMonths months, i, high
    End If
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BalanceCalculator.QuickSortMonths", Err.Number, Err.Description
End Sub

' ===== Day of Week Analysis =====

' Get spending by day of week
Public Function GetDayOfWeekSummary(transactions As Collection) As Collection
    On Error GoTo ErrorHandler
    
    Dim result As New Collection
    Dim dayDict As Object
    Set dayDict = CreateObject("Scripting.Dictionary")
    
    ' Initialize all days to zero
    Dim days As Variant
    days = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
    Dim i As Long
    
    For i = 0 To 6
        Dim newDay(0 To 1) As Currency
        newDay(0) = 0  ' Sum
        newDay(1) = 0  ' Count
        dayDict.Add days(i), newDay
    Next i
    
    ' Group by day of week
    Dim trans As Transaction
    For Each trans In transactions
        ' Only include expenses
        If trans.IsExpense() Then
            Dim dayName As String
            Dim amount As Currency
            
            ' Get day of week name
            dayName = Format(trans.TransactionDate, "dddd")
            
            ' Get absolute amount
            amount = trans.GetAbsoluteAmount()
            
            ' Update day data
            dayDict(dayName)(0) = dayDict(dayName)(0) + amount  ' Sum
            dayDict(dayName)(1) = dayDict(dayName)(1) + 1       ' Count
        End If
    Next trans
    
    ' Create day summaries (in order of day of week)
    For i = 0 To 6
        Dim daySummary As DayOfWeekSummary
        Dim dayName As String
        
        dayName = days(i)
        
        daySummary.DayName = dayName
        daySummary.DayNumber = i
        daySummary.TotalAmount = dayDict(dayName)(0)
        daySummary.TransactionCount = dayDict(dayName)(1)
        
        ' Add to result
        result.Add daySummary
    Next i
    
    Set GetDayOfWeekSummary = result
    
    ' Clean up
    Set dayDict = Nothing
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "BalanceCalculator.GetDayOfWeekSummary", Err.Number, Err.Description
    Set GetDayOfWeekSummary = New Collection ' Return empty collection in case of error
End Function

' Class initialize
Private Sub Class_Initialize()
    ' Nothing needed here
End Sub

' Class terminate
Private Sub Class_Terminate()
    ' Nothing to clean up
End Sub