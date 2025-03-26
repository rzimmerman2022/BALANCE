' File: src/repositories/TransactionRepository.cls
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransactionRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' TransactionRepository Class - Implements ITransactionRepository interface

Implements ITransactionRepository

' Constants for column indices
Private Const COL_ID As Long = 1
Private Const COL_DATE As Long = 2
Private Const COL_MERCHANT As Long = 3
Private Const COL_CATEGORY As Long = 4
Private Const COL_AMOUNT As Long = 5
Private Const COL_ACCOUNT As Long = 6
Private Const COL_OWNER As Long = 7
Private Const COL_IS_SHARED As Long = 8
Private Const COL_NOTES As Long = 9
Private Const COL_SOURCE_FILE As Long = 10

Private Const DATA_SHEET_NAME As String = "TransactionData"
Private Const HEADER_ROW As Long = 1
Private Const TABLE_NAME As String = "TransactionsTable"

' Private variables
Private m_Transactions As Collection
Private m_DataSheet As Worksheet
Private m_DataTable As ListObject
Private m_IsInitialized As Boolean
Private m_IsDirty As Boolean
Private m_Logger As ErrorLogger

' ===== Constructor & Initialization =====

Private Sub Class_Initialize()
    ' Nothing needed here as we're using PredeclaredId = True
    m_IsInitialized = False
    m_IsDirty = False
End Sub

' Initialize with dependency injection for logger
Public Sub Initialize(Optional logger As ErrorLogger = Nothing)
    On Error GoTo ErrorHandler
    
    ' Set logger if provided, otherwise get singleton
    If logger Is Nothing Then
        Set m_Logger = ErrorLogger
    Else
        Set m_Logger = logger
    End If
    
    ' Get or create data sheet
    Set m_DataSheet = GetOrCreateDataSheet()
    
    ' Get or create data table
    Set m_DataTable = GetOrCreateDataTable()
    
    ' Load transactions from table
    LoadTransactions
    
    m_IsInitialized = True
    m_IsDirty = False
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.Initialize", Err.Number, Err.Description
    End If
End Sub

' Get or create data sheet
Private Function GetOrCreateDataSheet() As Worksheet
    On Error Resume Next
    
    Dim ws As Worksheet
    
    ' Check if sheet exists
    If Utilities.SheetExists(DATA_SHEET_NAME) Then
        Set ws = ThisWorkbook.Worksheets(DATA_SHEET_NAME)
    Else
        ' Create new hidden sheet
        Set ws = ThisWorkbook.Worksheets.Add
        If Not ws Is Nothing Then
            ws.Name = DATA_SHEET_NAME
            ws.Visible = xlSheetVeryHidden
        End If
    End If
    
    Set GetOrCreateDataSheet = ws
    
    On Error GoTo 0
End Function

' Get or create data table
Private Function GetOrCreateDataTable() As ListObject
    On Error Resume Next
    
    Dim tbl As ListObject
    
    ' Check if table exists
    If TableExists(TABLE_NAME) Then
        Set tbl = m_DataSheet.ListObjects(TABLE_NAME)
    Else
        ' Set up headers first
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
        
        ' Create table
        Set tbl = m_DataSheet.ListObjects.Add(xlSrcRange, _
                                   m_DataSheet.Range(m_DataSheet.Cells(HEADER_ROW, COL_ID), _
                                                   m_DataSheet.Cells(HEADER_ROW, COL_SOURCE_FILE)), , xlYes)
        
        ' Set table name
        If Not tbl Is Nothing Then
            tbl.Name = TABLE_NAME
            
            ' Format headers
            With tbl.HeaderRowRange
                .Font.Bold = True
                .Interior.Color = AppSettings.ColorPrimary
                .Font.Color = AppSettings.ColorLightText
            End With
        End If
    End If
    
    Set GetOrCreateDataTable = tbl
    
    On Error GoTo 0
End Function

' Check if table exists
Private Function TableExists(tableName As String) As Boolean
    On Error Resume Next
    
    Dim tbl As ListObject
    Set tbl = m_DataSheet.ListObjects(tableName)
    
    TableExists = Not tbl Is Nothing
    
    On Error GoTo 0
End Function

' ===== Interface Implementation =====

' Get all transactions
Private Function ITransactionRepository_GetTransactions() As Collection
    If Not m_IsInitialized Then Initialize
    
    ' Return a copy of the collection to prevent external modification
    Dim result As New Collection
    Dim trans As Transaction
    
    For Each trans In m_Transactions
        result.Add trans, trans.ID
    Next trans
    
    Set ITransactionRepository_GetTransactions = result
End Function

' Public accessor for transactions collection
Public Property Get Transactions() As Collection
    Set Transactions = ITransactionRepository_GetTransactions()
End Property

' Get transaction count
Private Property Get ITransactionRepository_Count() As Long
    If Not m_IsInitialized Then Initialize
    ITransactionRepository_Count = m_Transactions.Count
End Property

' Get transactions by date range
Private Function ITransactionRepository_GetTransactionsByDateRange(startDate As Date, endDate As Date) As Collection
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    Dim result As New Collection
    Dim trans As Transaction
    
    For Each trans In m_Transactions
        If trans.TransactionDate >= startDate And trans.TransactionDate <= endDate Then
            result.Add trans
        End If
    Next trans
    
    Set ITransactionRepository_GetTransactionsByDateRange = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.GetTransactionsByDateRange", Err.Number, Err.Description
    End If
    Set ITransactionRepository_GetTransactionsByDateRange = New Collection
End Function

' Public wrapper for GetTransactionsByDateRange
Public Function GetTransactionsByDateRange(startDate As Date, endDate As Date) As Collection
    Set GetTransactionsByDateRange = ITransactionRepository_GetTransactionsByDateRange(startDate, endDate)
End Function

' Get transactions by category
Private Function ITransactionRepository_GetTransactionsByCategory(category As String) As Collection
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    Dim result As New Collection
    Dim trans As Transaction
    
    For Each trans In m_Transactions
        If LCase(trans.Category) = LCase(category) Then
            result.Add trans
        End If
    Next trans
    
    Set ITransactionRepository_GetTransactionsByCategory = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.GetTransactionsByCategory", Err.Number, Err.Description
    End If
    Set ITransactionRepository_GetTransactionsByCategory = New Collection
End Function

' Public wrapper for GetTransactionsByCategory
Public Function GetTransactionsByCategory(category As String) As Collection
    Set GetTransactionsByCategory = ITransactionRepository_GetTransactionsByCategory(category)
End Function

' Get transactions by owner
Private Function ITransactionRepository_GetTransactionsByOwner(owner As String) As Collection
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    Dim result As New Collection
    Dim trans As Transaction
    
    For Each trans In m_Transactions
        If LCase(trans.Owner) = LCase(owner) Then
            result.Add trans
        End If
    Next trans
    
    Set ITransactionRepository_GetTransactionsByOwner = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.GetTransactionsByOwner", Err.Number, Err.Description
    End If
    Set ITransactionRepository_GetTransactionsByOwner = New Collection
End Function

' Public wrapper for GetTransactionsByOwner
Public Function GetTransactionsByOwner(owner As String) As Collection
    Set GetTransactionsByOwner = ITransactionRepository_GetTransactionsByOwner(owner)
End Function

' Add a single transaction
Private Function ITransactionRepository_AddTransaction(transaction As Transaction) As Boolean
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    ' Check for duplicates and resolve if needed
    Dim transToAdd As Transaction
    Set transToAdd = ResolveDuplicate(transaction)
    
    ' Add to collection (replace if exists with same ID)
    On Error Resume Next
    
    If Len(transToAdd.ID) = 0 Then
        ' Generate ID if not set
        transToAdd.ID = GenerateUniqueID
    End If
    
    m_Transactions.Remove transToAdd.ID
    m_Transactions.Add transToAdd, transToAdd.ID
    
    On Error GoTo ErrorHandler
    
    ' Mark as dirty
    m_IsDirty = True
    
    ITransactionRepository_AddTransaction = True
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.AddTransaction", Err.Number, Err.Description
    End If
    ITransactionRepository_AddTransaction = False
End Function

' Public wrapper for AddTransaction
Public Function AddTransaction(transaction As Transaction) As Boolean
    AddTransaction = ITransactionRepository_AddTransaction(transaction)
End Function

' Generate unique ID
Private Function GenerateUniqueID() As String
    ' Create a GUID-like ID
    Dim guidPart1 As String: guidPart1 = Hex(Int(Rnd * &HFFFF&))
    Dim guidPart2 As String: guidPart2 = Hex(Int(Rnd * &HFFFF&))
    Dim guidPart3 As String: guidPart3 = Hex(Int(Rnd * &HFFFF&))
    Dim guidPart4 As String: guidPart4 = Hex(Int(Rnd * &HFFFF&))
    Dim timeStamp As String: timeStamp = Hex(CLng(Now * 86400))
    
    ' Format parts to ensure consistent length
    While Len(guidPart1) < 4: guidPart1 = "0" & guidPart1: Wend
    While Len(guidPart2) < 4: guidPart2 = "0" & guidPart2: Wend
    While Len(guidPart3) < 4: guidPart3 = "0" & guidPart3: Wend
    While Len(guidPart4) < 4: guidPart4 = "0" & guidPart4: Wend
    
    GenerateUniqueID = "TXID-" & guidPart1 & "-" & guidPart2 & "-" & guidPart3 & "-" & guidPart4 & "-" & timeStamp
End Function

' Add multiple transactions
Private Function ITransactionRepository_AddTransactions(transactions As Collection) As Long
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    Dim addedCount As Long
    addedCount = 0
    
    ' Add each transaction
    Dim trans As Transaction
    For Each trans In transactions
        If ITransactionRepository_AddTransaction(trans) Then
            addedCount = addedCount + 1
        End If
    Next trans
    
    ' Save changes
    If addedCount > 0 Then
        ITransactionRepository_SaveChanges
    End If
    
    ITransactionRepository_AddTransactions = addedCount
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.AddTransactions", Err.Number, Err.Description
    End If
    ITransactionRepository_AddTransactions = addedCount
End Function

' Public wrapper for AddTransactions
Public Function AddTransactions(transactions As Collection) As Long
    AddTransactions = ITransactionRepository_AddTransactions(transactions)
End Function

' Remove a transaction by ID
Private Function ITransactionRepository_RemoveTransaction(transactionId As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    ' Remove from collection
    On Error Resume Next
    m_Transactions.Remove transactionId
    
    If Err.Number = 0 Then
        ' Mark as dirty only if we actually removed something
        m_IsDirty = True
        ITransactionRepository_RemoveTransaction = True
    Else
        ITransactionRepository_RemoveTransaction = False
    End If
    
    On Error GoTo ErrorHandler
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.RemoveTransaction", Err.Number, Err.Description
    End If
    ITransactionRepository_RemoveTransaction = False
End Function

' Public wrapper for RemoveTransaction
Public Function RemoveTransaction(transactionId As String) As Boolean
    RemoveTransaction = ITransactionRepository_RemoveTransaction(transactionId)
End Function

' Save changes to storage
Private Function ITransactionRepository_SaveChanges() As Boolean
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    ' Only save if there are unsaved changes
    If Not m_IsDirty Then
        ITransactionRepository_SaveChanges = True
        Exit Function
    End If
    
    ' Performance optimization: use arrays for batch operation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Clear existing data (except header)
    If m_DataTable.ListRows.Count > 0 Then
        m_DataTable.DataBodyRange.Delete
    End If
    
    ' Prepare data array for bulk insertion
    Dim transCount As Long
    transCount = m_Transactions.Count
    
    If transCount > 0 Then
        ' Create an array to hold all transaction data
        Dim dataArray() As Variant
        ReDim dataArray(1 To transCount, 1 To 10)
        
        ' Fill array with transaction data
        Dim i As Long
        Dim trans As Transaction
        i = 1
        
        For Each trans In m_Transactions
            dataArray(i, COL_ID) = trans.ID
            dataArray(i, COL_DATE) = trans.TransactionDate
            dataArray(i, COL_MERCHANT) = trans.Merchant
            dataArray(i, COL_CATEGORY) = trans.Category
            dataArray(i, COL_AMOUNT) = trans.Amount
            dataArray(i, COL_ACCOUNT) = trans.Account
            dataArray(i, COL_OWNER) = trans.Owner
            dataArray(i, COL_IS_SHARED) = trans.IsShared
            dataArray(i, COL_NOTES) = trans.Notes
            dataArray(i, COL_SOURCE_FILE) = trans.SourceFile
            
            i = i + 1
        Next trans
        
        ' Add all data in one operation
        For i = 1 To transCount
            Dim newRow As ListRow
            Set newRow = m_DataTable.ListRows.Add
            
            For j = 1 To 10
                newRow.Range(1, j).Value = dataArray(i, j)
            Next j
        Next i
        
        ' Format date column
        m_DataTable.ListColumns(COL_DATE).DataBodyRange.NumberFormat = "yyyy-mm-dd"
        
        ' Format amount column
        m_DataTable.ListColumns(COL_AMOUNT).DataBodyRange.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    End If
    
    ' Mark as clean (all changes saved)
    m_IsDirty = False
    
    ' Update last update date
    AppSettings.LastUpdateDate = Now
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ITransactionRepository_SaveChanges = True
    
    Exit Function
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.SaveChanges", Err.Number, Err.Description
    End If
    ITransactionRepository_SaveChanges = False
End Function

' Public wrapper for SaveChanges
Public Function SaveChanges() As Boolean
    SaveChanges = ITransactionRepository_SaveChanges()
End Function

' Clear all transactions
Private Sub ITransactionRepository_ClearAll()
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    ' Create a new empty collection
    Set m_Transactions = New Collection
    
    ' Mark as dirty
    m_IsDirty = True
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.ClearAll", Err.Number, Err.Description
    End If
End Sub

' Public wrapper for ClearAll
Public Sub ClearAll()
    ITransactionRepository_ClearAll
End Sub

' Initialize repository
Private Sub ITransactionRepository_Initialize()
    Initialize
End Sub

' ===== Private Helper Methods =====

' Load transactions from table
Private Sub LoadTransactions()
    On Error GoTo ErrorHandler
    
    ' Clear existing collection
    Set m_Transactions = New Collection
    
    ' If no data, just return
    If m_DataTable.ListRows.Count = 0 Then Exit Sub
    
    ' Load each row using array operations for better performance
    Dim dataRange As Range
    Set dataRange = m_DataTable.DataBodyRange
    
    Dim dataArray As Variant
    dataArray = dataRange.Value
    
    Dim i As Long, j As Long
    Dim rowCount As Long, colCount As Long
    
    rowCount = UBound(dataArray, 1)
    colCount = UBound(dataArray, 2)
    
    For i = 1 To rowCount
        ' Create new transaction
        Dim trans As Transaction
        Set trans = New Transaction
        
        ' Set ID
        trans.ID = dataArray(i, COL_ID)
        
        ' Initialize from row data
        trans.InitFromRow _
            dataArray(i, COL_DATE), _
            dataArray(i, COL_MERCHANT), _
            dataArray(i, COL_CATEGORY), _
            dataArray(i, COL_AMOUNT), _
            dataArray(i, COL_ACCOUNT), _
            dataArray(i, COL_OWNER), _
            dataArray(i, COL_IS_SHARED), _
            dataArray(i, COL_NOTES), _
            dataArray(i, COL_SOURCE_FILE)
        
        ' Add to collection
        On Error Resume Next
        m_Transactions.Add trans, trans.ID
        
        If Err.Number = 457 Then  ' Already exists - key is duplicated
            ' Skip this transaction (should not happen)
            If Not m_Logger Is Nothing Then
                m_Logger.LogWarning "TransactionRepository.LoadTransactions", _
                    "Duplicate transaction ID found: " & trans.ID
            End If
        End If
        On Error GoTo ErrorHandler
    Next i
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.LoadTransactions", Err.Number, Err.Description
    End If
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
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.ResolveDuplicate", Err.Number, Err.Description
    End If
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
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.GetDuplicatesInDateRange", Err.Number, Err.Description
    End If
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

' ===== Cleanup =====

Private Sub Class_Terminate()
    ' Clean up to prevent memory leaks
    Set m_Transactions = Nothing
    Set m_DataSheet = Nothing
    Set m_DataTable = Nothing
    Set m_Logger = Nothing
End Sub