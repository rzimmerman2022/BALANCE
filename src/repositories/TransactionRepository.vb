' Get transactions filtered by date range
Public Function GetTransactionsByDateRange(startDate As Date, endDate As Date) As Collection
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    Dim result As New Collection
    Dim trans As Transaction
    
    For Each trans In m_Transactions
        If trans.TransactionDate >= startDate And trans.TransactionDate <= endDate Then
            result.Add trans
        End If
    Next trans
    
    Set GetTransactionsByDateRange = result
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "TransactionRepository.GetTransactionsByDateRange", Err.Number, Err.Description
    Set GetTransactionsByDateRange = New Collection
End Function

' Get transactions filtered by category
Public Function GetTransactionsByCategory(category As String) As Collection
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    Dim result As New Collection
    Dim trans As Transaction
    
    For Each trans In m_Transactions
        If LCase(trans.Category) = LCase(category) Then
            result.Add trans
        End If
    Next trans
    
    Set GetTransactionsByCategory = result
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "TransactionRepository.GetTransactionsByCategory", Err.Number, Err.Description
    Set GetTransactionsByCategory = New Collection
End Function

' Get transactions filtered by owner
Public Function GetTransactionsByOwner(owner As String) As Collection
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    Dim result As New Collection
    Dim trans As Transaction
    
    For Each trans In m_Transactions
        If LCase(trans.Owner) = LCase(owner) Then
            result.Add trans
        End If
    Next trans
    
    Set GetTransactionsByOwner = result
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "TransactionRepository.GetTransactionsByOwner", Err.Number, Err.Description
    Set GetTransactionsByOwner = New Collection
End Function

' Get transactions filtered by source file
Public Function GetTransactionsBySourceFile(sourceFile As String) As Collection
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    Dim result As New Collection
    Dim trans As Transaction
    
    For Each trans In m_Transactions
        If LCase(trans.SourceFile) = LCase(sourceFile) Then
            result.Add trans
        End If
    Next trans
    
    Set GetTransactionsBySourceFile = result
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "TransactionRepository.GetTransactionsBySourceFile", Err.Number, Err.Description
    Set GetTransactionsBySourceFile = New Collection
End Function

' Save changes to the data sheet
Public Sub SaveChanges()
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    ' Only save if there are unsaved changes
    If Not m_IsDirty Then Exit Sub
    
    ' Clear existing data (keeping header)
    Dim lastRow As Long
    lastRow = Utilities.GetLastRow(m_DataSheet, 1)
    
    If lastRow > HEADER_ROW Then
        m_DataSheet.Range(m_DataSheet.Cells(HEADER_ROW + 1, 1), m_DataSheet.Cells(lastRow, COL_SOURCE_FILE)).Clear
    End If
    
    ' Write transactions to sheet
    Dim trans As Transaction
    Dim row As Long
    
    row = HEADER_ROW + 1
    
    For Each trans In m_Transactions
        ' ID
        m_DataSheet.Cells(row, COL_ID).Value = trans.ID
        
        ' Date
        m_DataSheet.Cells(row, COL_DATE).Value = trans.TransactionDate
        m_DataSheet.Cells(row, COL_DATE).NumberFormat = "yyyy-mm-dd"
        
        ' Merchant
        m_DataSheet.Cells(row, COL_MERCHANT).Value = trans.Merchant
        
        ' Category
        m_DataSheet.Cells(row, COL_CATEGORY).Value = trans.Category
        
        ' Amount
        m_DataSheet.Cells(row, COL_AMOUNT).Value = trans.Amount
        m_DataSheet.Cells(row, COL_AMOUNT).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        
        ' Account
        m_DataSheet.Cells(row, COL_ACCOUNT).Value = trans.Account
        
        ' Owner
        m_DataSheet.Cells(row, COL_OWNER).Value = trans.Owner
        
        ' IsShared
        m_DataSheet.Cells(row, COL_IS_SHARED).Value = trans.IsShared
        
        ' Notes
        m_DataSheet.Cells(row, COL_NOTES).Value = trans.Notes
        
        ' SourceFile
        m_DataSheet.Cells(row, COL_SOURCE_FILE).Value = trans.SourceFile
        
        row = row + 1
    Next trans
    
    ' Mark as clean (all changes saved)
    m_IsDirty = False
    
    ' Update last update date
    AppSettings.LastUpdateDate = Now
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "TransactionRepository.SaveChanges", Err.Number, Err.Description
End Sub

' Clear all transactions
Public Sub ClearAll()
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    ' Create a new empty collection
    Set m_Transactions = New Collection
    
    ' Mark as dirty
    m_IsDirty = True
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "TransactionRepository.ClearAll", Err.Number, Err.Description
End Sub

' Load transactions from sheet
Private Sub LoadTransactions()
    On Error GoTo ErrorHandler
    
    ' Clear existing collection
    Set m_Transactions = New Collection
    
    ' Get last row
    Dim lastRow As Long
    lastRow = Utilities.GetLastRow(m_DataSheet, 1)
    
    ' If no data, just return
    If lastRow <= HEADER_ROW Then Exit Sub
    
    ' Load each row
    Dim row As Long
    Dim trans As Transaction
    
    For row = HEADER_ROW + 1 To lastRow
        ' Create new transaction
        Set trans = New Transaction
        
        ' Set ID
        trans.ID = m_DataSheet.Cells(row, COL_ID).Value
        
        ' Initialize from row data
        trans.InitFromRow _
            m_DataSheet.Cells(row, COL_DATE).Value, _
            m_DataSheet.Cells(row, COL_MERCHANT).Value, _
            m_DataSheet.Cells(row, COL_CATEGORY).Value, _
            m_DataSheet.Cells(row, COL_AMOUNT).Value, _
            m_DataSheet.Cells(row, COL_ACCOUNT).Value, _
            m_DataSheet.Cells(row, COL_OWNER).Value, _
            m_DataSheet.Cells(row, COL_IS_SHARED).Value, _
            m_DataSheet.Cells(row, COL_NOTES).Value, _
            m_DataSheet.Cells(row, COL_SOURCE_FILE).Value
        
        ' Add to collection
        On Error Resume Next
        m_Transactions.Add trans, trans.ID
        
        If Err.Number = 457 Then  ' Already exists - key is duplicated
            ' Skip this transaction (should not happen)
            Debug.Print "Duplicate transaction ID found: " & trans.ID
        End If
        On Error GoTo ErrorHandler
    Next row
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "TransactionRepository.LoadTransactions", Err.Number, Err.Description
End Sub

' Set up data sheet with headers
Private Sub SetupDataSheet()
    On Error GoTo ErrorHandler
    
    ' Set headers
    m_DataSheet.Cells(HEADER_ROW, COL_ID).Value = "ID"
    m_DataSheet.Cells(HEADER_ROW, COL_DATE).Value = "Date"
    m_DataSheet.Cells(HEADER_ROW, COL_MERCHANT).Value = "Merchant"
    m_DataSheet.Cells(HEADER_ROW, COL_CATEGORY).Value = "Category"
    m_DataSheet.Cells(HEADER_ROW, COL_AMOUNT).Value = "Amount"
    m_DataSheet.Cells(HEADER_ROW, COL_ACCOUNT).Value = "Account"
    m_DataSheet.Cells(HEADER_ROW, COL_OWNER).Value = "Owner"
    m_DataSheet.Cells(HEADER_ROW, COL_IS_SHARED).Value = "IsShared"
    m_DataSheet.Cells(HEADER_ROW, COL_NOTES).Value = "Notes"
    m_DataSheet.Cells(HEADER_ROW, COL_SOURCE_FILE).Value = "SourceFile"
    
    ' Format headers
    With m_DataSheet.Range(m_DataSheet.Cells(HEADER_ROW, 1), m_DataSheet.Cells(HEADER_ROW, COL_SOURCE_FILE))
        .Font.Bold = True
        .Interior.Color = AppSettings.ColorPrimary
        .Font.Color = AppSettings.ColorLightText
    End With
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "TransactionRepository.SetupDataSheet", Err.Number, Err.Description
End Sub

' Add a collection of transactions and save
Public Sub AddTransactions(transactions As Collection)
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    ' Add each transaction
    Dim trans As Transaction
    
    For Each trans In transactions
        AddTransaction trans
    Next trans
    
    ' Save changes
    SaveChanges
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "TransactionRepository.AddTransactions", Err.Number, Err.Description
End Sub

' Check for duplicate transactions and return the newer one
Public Function ResolveDuplicate(newTrans As Transaction) As Transaction
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    Dim existingTrans As Transaction
    Dim dateStr As String
    Dim key As String
    
    ' Create a key to detect potential duplicates
    dateStr = Format(newTrans.TransactionDate, "yyyy-mm-dd")
    key = dateStr & "|" & newTrans.Merchant & "|" & newTrans.Amount
    
    ' Look for matching transactions
    For Each existingTrans In m_Transactions
        Dim existingDateStr As String
        Dim existingKey As String
        
        existingDateStr = Format(existingTrans.TransactionDate, "yyyy-mm-dd")
        existingKey = existingDateStr & "|" & existingTrans.Merchant & "|" & existingTrans.Amount
        
        If existingKey = key Then
            ' Found a duplicate - need to decide which to keep
            
            ' Compare source files for dates
            Dim existingFileDate As String
            Dim newFileDate As String
            
            existingFileDate = Utilities.ExtractDateFromFilename(existingTrans.SourceFile)
            newFileDate = Utilities.ExtractDateFromFilename(newTrans.SourceFile)
            
            ' If we can compare dates, return the newer one
            If Len(existingFileDate) > 0 And Len(newFileDate) > 0 Then
                If newFileDate > existingFileDate Then
                    Set ResolveDuplicate = newTrans
                Else
                    Set ResolveDuplicate = existingTrans
                End If
                Exit Function
            End If
        End If
    Next existingTrans
    
    ' No duplicate found, return the new transaction
    Set ResolveDuplicate = newTrans
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "TransactionRepository.ResolveDuplicate", Err.Number, Err.Description
    Set ResolveDuplicate = newTrans ' Default to new transaction in case of error
End Function

' Check for duplicates by date range
Public Function GetDuplicatesInDateRange(startDate As Date, endDate As Date) As Collection
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    Dim result As New Collection
    Dim transDict As Object
    Set transDict = CreateObject("Scripting.Dictionary")
    
    ' Get transactions in date range
    Dim trans As Transaction
    Dim transactionList As Collection
    Set transactionList = GetTransactionsByDateRange(startDate, endDate)
    
    ' Group by transaction signature
    For Each trans In transactionList
        Dim key As String
        key = Format(trans.TransactionDate, "yyyy-mm-dd") & "|" & trans.Merchant & "|" & trans.Amount
        
        If transDict.Exists(key) Then
            ' Found a duplicate
            If Not ExistsInCollection(result, key) Then
                result.Add key
            End If
        Else
            transDict.Add key, trans
        End If
    Next trans
    
    Set GetDuplicatesInDateRange = result
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "TransactionRepository.GetDuplicatesInDateRange", Err.Number, Err.Description
    Set GetDuplicatesInDateRange = New Collection
End Function

' Helper function to check if a value exists in a collection
Private Function ExistsInCollection(col As Collection, value As Variant) As Boolean
    On Error Resume Next
    
    Dim item As Variant
    For Each item In col
        If item = value Then
            ExistsInCollection = True
            Exit Function
        End If
    Next item
    
    ExistsInCollection = False
    
    On Error GoTo 0
End Function

' Class initialize
Private Sub Class_Initialize()
    ' Nothing needed here as we're using PredeclaredId = True
    ' and explicit initialization via Initialize method
    m_IsInitialized = False
    m_IsDirty = False
End Sub

' Class terminate
Private Sub Class_Terminate()
    ' Clean up to prevent memory leaks
    Set m_Transactions = Nothing
    Set m_DataSheet = Nothing
End Sub