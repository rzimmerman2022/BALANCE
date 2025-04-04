' File: src/repositories/TransactionRepository.cls
'------------------------------------------------
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
' TransactionRepository Class - Implements ITransactionRepository
'
' Purpose: Provides data storage and retrieval for Transaction objects
' using Excel Tables (ListObjects) for efficient data operations.
'
' Design Decisions:
' - Uses PredeclaredId = True to enable Singleton pattern
' - Implements ITransactionRepository for loose coupling
' - Uses Excel Tables (ListObjects) for structured data storage
' - Performs batch operations using arrays for performance
' - Implements dependency injection for the logger component
' - Uses proper error handling with centralized logging
' - Publishes events when data changes to support event-driven architecture

' Interface implementation
Implements ITransactionRepository

' Constants for table and column definitions
Private Const DATA_SHEET_NAME As String = "TransactionData"
Private Const TABLE_NAME As String = "TransactionsTable"
Private Const HEADER_ROW As Long = 1

' Column indices for table
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

' Private member variables
Private m_Transactions As Collection  ' Collection of Transaction objects
Private m_DataSheet As Worksheet      ' Worksheet containing the data
Private m_DataTable As ListObject     ' Excel Table for structured data storage
Private m_IsInitialized As Boolean    ' Flag indicating if repository is initialized
Private m_IsDirty As Boolean          ' Flag indicating unsaved changes
Private m_Logger As IErrorLogger       ' Error logging component (injected)

'=========================================================================
' Initialization and Setup
'=========================================================================

Private Sub Class_Initialize()
    ' Initialize member variables
    Set m_Transactions = New Collection
    m_IsInitialized = False
    m_IsDirty = False
End Sub

' Initialize the repository and set up dependencies
Private Sub ITransactionRepository_Initialize(Optional ByVal logger As IErrorLogger = Nothing)
    On Error GoTo ErrorHandler
    
    ' Set logger if provided, otherwise get singleton instance
    If logger Is Nothing Then
        Set m_Logger = ErrorLogger ' Assume ErrorLogger has a default instance
    Else
        Set m_Logger = logger
    End If
    
    ' Log initialization start
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "TransactionRepository.Initialize", "Initializing transaction repository"
    End If
    
    ' Get or create data sheet
    Set m_DataSheet = GetOrCreateDataSheet()
    
    ' Get or create data table
    Set m_DataTable = GetOrCreateDataTable()
    
    ' Load transactions from table
    LoadTransactionsFromTable
    
    m_IsInitialized = True
    m_IsDirty = False
    
    ' Log initialization complete
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "TransactionRepository.Initialize", _
            "Repository initialized with " & m_Transactions.Count & " transactions"
    End If
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.Initialize", Err.Number, Err.Description
    End If
End Sub

' Public wrapper for initialization
Public Sub Initialize(Optional ByVal logger As IErrorLogger = Nothing)
    ITransactionRepository_Initialize logger
End Sub

' Get or create the data sheet
Private Function GetOrCreateDataSheet() As Worksheet
    On Error Resume Next
    
    Dim ws As Worksheet
    
    ' Try to get existing sheet
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET_NAME)
    
    ' If sheet doesn't exist, create it
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = DATA_SHEET_NAME
        
        ' Hide the sheet (very hidden so it can't be unhidden through the UI)
        ws.Visible = xlSheetVeryHidden
        
        If Not m_Logger Is Nothing Then
            m_Logger.LogInfo "TransactionRepository.GetOrCreateDataSheet", _
                "Created new data sheet: " & DATA_SHEET_NAME
        End If
    End If
    
    Set GetOrCreateDataSheet = ws
    
    On Error GoTo 0
End Function

' Get or create the data table
Private Function GetOrCreateDataTable() As ListObject
    On Error Resume Next
    
    Dim tbl As ListObject
    
    ' Check if table exists
    For Each tbl In m_DataSheet.ListObjects
        If tbl.Name = TABLE_NAME Then
            Set GetOrCreateDataTable = tbl
            Exit Function
        End If
    Next tbl
    
    ' Create table if it doesn't exist
    ' First set up headers
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
    
    ' Create table with the headers
    Set tbl = m_DataSheet.ListObjects.Add(xlSrcRange, _
                                   m_DataSheet.Range(m_DataSheet.Cells(HEADER_ROW, COL_ID), _
                                                   m_DataSheet.Cells(HEADER_ROW, COL_SOURCE_FILE)), , xlYes)
    
    ' Set table name
    tbl.Name = TABLE_NAME
    
    ' Format header row
    With tbl.HeaderRowRange
        .Font.Bold = True
        .Interior.Color = RGB(0, 112, 192)  ' Blue header
        .Font.Color = RGB(255, 255, 255)    ' White text
    End With
    
    ' Format data
    If Not tbl.DataBodyRange Is Nothing Then  ' Only if table has data rows
        ' Format date column
        tbl.ListColumns(COL_DATE).DataBodyRange.NumberFormat = "yyyy-mm-dd"
        
        ' Format amount column
        tbl.ListColumns(COL_AMOUNT).DataBodyRange.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    End If
    
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "TransactionRepository.GetOrCreateDataTable", _
            "Created new transactions table: " & TABLE_NAME
    End If
    
    Set GetOrCreateDataTable = tbl
    
    On Error GoTo 0
End Function

'=========================================================================
' Interface Implementation - Get Methods
'=========================================================================

' Returns all transactions in the repository
Private Function ITransactionRepository_GetTransactions() As Collection
    On Error GoTo ErrorHandler
    
    ' Ensure repository is initialized
    If Not m_IsInitialized Then ITransactionRepository_Initialize
    
    ' Create a new collection to return (don't return direct reference to internal collection)
    Dim result As Collection
    Set result = New Collection
    
    ' Copy all transactions to the new collection
    Dim trans As Transaction
    For Each trans In m_Transactions
        ' Use the transaction ID as the key for fast lookups
        On Error Resume Next
        result.Add trans, trans.ID
        On Error GoTo ErrorHandler
    Next trans
    
    ' Return the collection
    Set ITransactionRepository_GetTransactions = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.GetTransactions", Err.Number, Err.Description
    End If
    ' Return empty collection on error
    Set ITransactionRepository_GetTransactions = New Collection
End Function

' Public wrapper for GetTransactions
Public Function GetTransactions() As Collection
    Set GetTransactions = ITransactionRepository_GetTransactions()
End Function

' Returns transactions within the specified date range
Private Function ITransactionRepository_GetTransactionsByDateRange(ByVal startDate As Date, ByVal endDate As Date) As Collection
    On Error GoTo ErrorHandler
    
    ' Ensure repository is initialized
    If Not m_IsInitialized Then ITransactionRepository_Initialize
    
    ' Create a new collection for the results
    Dim result As Collection
    Set result = New Collection
    
    ' Filter transactions by date range
    Dim trans As Transaction
    For Each trans In m_Transactions
        If trans.TransactionDate >= startDate And trans.TransactionDate <= endDate Then
            result.Add trans
        End If
    Next trans
    
    ' Return the filtered collection
    Set ITransactionRepository_GetTransactionsByDateRange = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.GetTransactionsByDateRange", Err.Number, Err.Description
    End If
    ' Return empty collection on error
    Set ITransactionRepository_GetTransactionsByDateRange = New Collection
End Function

' Public wrapper for GetTransactionsByDateRange
Public Function GetTransactionsByDateRange(ByVal startDate As Date, ByVal endDate As Date) As Collection
    Set GetTransactionsByDateRange = ITransactionRepository_GetTransactionsByDateRange(startDate, endDate)
End Function

' Returns transactions with the specified category
Private Function ITransactionRepository_GetTransactionsByCategory(ByVal category As String) As Collection
    On Error GoTo ErrorHandler
    
    ' Ensure repository is initialized
    If Not m_IsInitialized Then ITransactionRepository_Initialize
    
    ' Create a new collection for the results
    Dim result As Collection
    Set result = New Collection
    
    ' Filter transactions by category (case-insensitive)
    Dim trans As Transaction
    For Each trans In m_Transactions
        If LCase(trans.Category) = LCase(category) Then
            result.Add trans
        End If
    Next trans
    
    ' Return the filtered collection
    Set ITransactionRepository_GetTransactionsByCategory = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.GetTransactionsByCategory", Err.Number, Err.Description
    End If
    ' Return empty collection on error
    Set ITransactionRepository_GetTransactionsByCategory = New Collection
End Function

' Public wrapper for GetTransactionsByCategory
Public Function GetTransactionsByCategory(ByVal category As String) As Collection
    Set GetTransactionsByCategory = ITransactionRepository_GetTransactionsByCategory(category)
End Function

' Returns transactions with the specified owner
Private Function ITransactionRepository_GetTransactionsByOwner(ByVal owner As String) As Collection
    On Error GoTo ErrorHandler
    
    ' Ensure repository is initialized
    If Not m_IsInitialized Then ITransactionRepository_Initialize
    
    ' Create a new collection for the results
    Dim result As Collection
    Set result = New Collection
    
    ' Filter transactions by owner (case-insensitive)
    Dim trans As Transaction
    For Each trans In m_Transactions
        If LCase(trans.Owner) = LCase(owner) Then
            result.Add trans
        End If
    Next trans
    
    ' Return the filtered collection
    Set ITransactionRepository_GetTransactionsByOwner = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.GetTransactionsByOwner", Err.Number, Err.Description
    End If
    ' Return empty collection on error
    Set ITransactionRepository_GetTransactionsByOwner = New Collection
End Function

' Public wrapper for GetTransactionsByOwner
Public Function GetTransactionsByOwner(ByVal owner As String) As Collection
    Set GetTransactionsByOwner = ITransactionRepository_GetTransactionsByOwner(owner)
End Function

' Returns the number of transactions in the repository
Private Property Get ITransactionRepository_Count() As Long
    ' Ensure repository is initialized
    If Not m_IsInitialized Then ITransactionRepository_Initialize
    
    ' Return the count
    ITransactionRepository_Count = m_Transactions.Count
End Property

' Public wrapper for Count
Public Property Get Count() As Long
    Count = ITransactionRepository_Count
End Property

'=========================================================================
' Interface Implementation - Add and Remove Methods
'=========================================================================

' Adds a single transaction to the repository
Private Function ITransactionRepository_AddTransaction(ByVal transaction As ITransaction) As Boolean ' Changed parameter type to ITransaction
    On Error GoTo ErrorHandler
    
    ' Ensure repository is initialized
    If Not m_IsInitialized Then ITransactionRepository_Initialize
    
    ' Check for duplicates and resolve if needed
    ' Note: ResolveDuplicate expects a concrete Transaction. We might need to adjust it
    ' or handle the type check/cast here. For now, assume transaction IS a Transaction object.
    ' This might need refinement if other ITransaction implementers exist.
    Dim concreteTransaction As Transaction
    If TypeOf transaction Is Transaction Then
        Set concreteTransaction = transaction
    Else
        ' Handle error or attempt conversion if possible
        If Not m_Logger Is Nothing Then
            m_Logger.LogError "TransactionRepository.AddTransaction", 0, "AddTransaction received non-Transaction object implementing ITransaction. Cannot resolve duplicates."
        End If
        ' Decide how to proceed: maybe add without duplicate check? Or fail?
        ' For now, let's try adding without duplicate check if it's not a concrete Transaction
        ' This part might need more robust handling depending on requirements.
        Set concreteTransaction = transaction ' This might fail if ResolveDuplicate strictly needs Transaction
        ' A safer approach might be to skip ResolveDuplicate if not TypeOf Transaction
        ' Or modify ResolveDuplicate to accept ITransaction
    End If

    Dim transToAdd As Transaction ' ResolveDuplicate returns Transaction
    Set transToAdd = ResolveDuplicate(concreteTransaction) ' Pass concrete type

    ' Generate ID if not set
    If Len(transToAdd.ID) = 0 Then
        transToAdd.ID = GenerateUniqueID
    End If
    
    ' Add to collection (remove first if it already exists)
    On Error Resume Next
    m_Transactions.Remove transToAdd.ID
    On Error GoTo ErrorHandler
    
    m_Transactions.Add transToAdd, transToAdd.ID
    
    ' Mark as dirty (needs saving)
    m_IsDirty = True
    
    ' Publish event to notify that transactions have changed
    EventManager.PublishEvent EventType.TransactionsChanged
    
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "TransactionRepository.AddTransaction", _
            "Added transaction ID: " & transToAdd.ID
    End If
    
    ITransactionRepository_AddTransaction = True
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.AddTransaction", Err.Number, Err.Description
    End If
    ITransactionRepository_AddTransaction = False
End Function

' Public wrapper for AddTransaction (Keep only the one accepting ITransaction)
Public Function AddTransaction(ByVal transaction As ITransaction) As Boolean ' Changed parameter type
    AddTransaction = ITransactionRepository_AddTransaction(transaction)
End Function

' Adds multiple transactions to the repository
Private Function ITransactionRepository_AddTransactions(ByVal transactions As Collection) As Long
    On Error GoTo ErrorHandler
    
    ' Ensure repository is initialized
    If Not m_IsInitialized Then ITransactionRepository_Initialize
    
    Dim addedCount As Long
    addedCount = 0
    
    ' Process each transaction - Use Variant loop and check type
    Dim transVariant As Variant
    Dim trans As ITransaction ' Use interface type
    For Each transVariant In transactions
        Set trans = transVariant ' Assumes collection contains objects implementing ITransaction
        If ITransactionRepository_AddTransaction(trans) Then ' Pass ITransaction
            addedCount = addedCount + 1
        End If
    Next trans
    
    ' Publish event to notify that transactions have changed
    EventManager.PublishEvent EventType.TransactionsChanged
    
    ' Return the number of transactions added
    ITransactionRepository_AddTransactions = addedCount
    
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "TransactionRepository.AddTransactions", _
            "Added " & addedCount & " transactions"
    End If
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.AddTransactions", Err.Number, Err.Description
    End If
    ITransactionRepository_AddTransactions = addedCount
End Function

' Public wrapper for AddTransactions
Public Function AddTransactions(ByVal transactions As Collection) As Long
    AddTransactions = ITransactionRepository_AddTransactions(transactions)
End Function

' Removes a transaction from the repository by ID
Private Function ITransactionRepository_RemoveTransaction(ByVal transactionId As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Ensure repository is initialized
    If Not m_IsInitialized Then ITransactionRepository_Initialize
    
    ' Try to remove the transaction
    On Error Resume Next
    m_Transactions.Remove transactionId
    
    ' Check if removal was successful
    If Err.Number = 0 Then
        m_IsDirty = True
        ITransactionRepository_RemoveTransaction = True
        
        ' Publish event to notify that transactions have changed
        EventManager.PublishEvent EventType.TransactionsChanged
        
        If Not m_Logger Is Nothing Then
            m_Logger.LogInfo "TransactionRepository.RemoveTransaction", _
                "Removed transaction ID: " & transactionId
        End If
    Else
        ITransactionRepository_RemoveTransaction = False
        
        If Not m_Logger Is Nothing Then
            m_Logger.LogWarning "TransactionRepository.RemoveTransaction", _
                "Transaction not found for removal, ID: " & transactionId
        End If
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
Public Function RemoveTransaction(ByVal transactionId As String) As Boolean
    RemoveTransaction = ITransactionRepository_RemoveTransaction(transactionId)
End Function

' Clears all transactions from the repository
Private Sub ITransactionRepository_ClearAll()
    On Error GoTo ErrorHandler
    
    ' Ensure repository is initialized
    If Not m_IsInitialized Then ITransactionRepository_Initialize
    
    ' Create a new empty collection
    Set m_Transactions = New Collection
    
    ' Mark as dirty
    m_IsDirty = True
    
    ' Publish event to notify that transactions have changed
    EventManager.PublishEvent EventType.TransactionsChanged
    
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "TransactionRepository.ClearAll", "Cleared all transactions"
    End If
    
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

'=========================================================================
' Interface Implementation - Save Changes
'=========================================================================

' Saves all changes to the underlying storage (Excel Table)
Private Function ITransactionRepository_SaveChanges() As Boolean
    On Error GoTo ErrorHandler
    
    ' Ensure repository is initialized
    If Not m_IsInitialized Then ITransactionRepository_Initialize
    
    ' If no changes to save, return success
    If Not m_IsDirty Then
        ITransactionRepository_SaveChanges = True
        Exit Function
    End If
    
    ' Log the start of save operation
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "TransactionRepository.SaveChanges", _
            "Saving " & m_Transactions.Count & " transactions to table"
    End If
    
    ' Performance optimization for Excel
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Clear existing data (except header)
    If m_DataTable.ListRows.Count > 0 Then
        m_DataTable.DataBodyRange.Delete
    End If
    
    ' Check if there are transactions to save
    Dim transCount As Long
    transCount = m_Transactions.Count
    
    If transCount > 0 Then
        ' Prepare data array for bulk insertion
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
        
        ' Add all rows to the table in one operation for better performance
        ' First add the rows
        For i = 1 To transCount
            m_DataTable.ListRows.Add
        Next i
        
        ' Then set the values all at once (if table has data body range)
        If Not m_DataTable.DataBodyRange Is Nothing Then
            m_DataTable.DataBodyRange.Value = dataArray
            
            ' Format date column
            m_DataTable.ListColumns(COL_DATE).DataBodyRange.NumberFormat = "yyyy-mm-dd"
            
            ' Format amount column
            m_DataTable.ListColumns(COL_AMOUNT).DataBodyRange.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        End If
    End If
    
    ' Mark as clean (all changes saved)
    m_IsDirty = False
    
    ' Publish event to notify that transactions have been saved
    EventManager.PublishEvent EventType.TransactionsChanged
    
    ' Restore Excel settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "TransactionRepository.SaveChanges", "Save operation completed successfully"
    End If
    
    ITransactionRepository_SaveChanges = True
    
    Exit Function
    
ErrorHandler:
    ' Restore Excel settings even on error
    Application.EnableEvents = True
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

'=========================================================================
' Helper Methods
'=========================================================================

' Loads transactions from the Excel Table into memory
Private Sub LoadTransactionsFromTable()
    On Error GoTo ErrorHandler
    
    ' Start with a new collection
    Set m_Transactions = New Collection
    
    ' If table has no data, exit
    If m_DataTable.ListRows.Count = 0 Then
        If Not m_Logger Is Nothing Then
            m_Logger.LogInfo "TransactionRepository.LoadTransactionsFromTable", _
                "No transactions found in table"
        End If
        Exit Sub
    End If
    
    ' Get all data at once for better performance
    Dim dataRange As Range
    Set dataRange = m_DataTable.DataBodyRange
    
    Dim dataArray As Variant
    dataArray = dataRange.Value
    
    ' Process data array
    Dim i As Long
    Dim rowCount As Long
    
    ' Use total rows from the array dimension
    rowCount = UBound(dataArray, 1)
    
    For i = 1 To rowCount
        ' Create and initialize new transaction
        Dim trans As New Transaction
        
        ' Set transaction ID
        trans.ID = CStr(dataArray(i, COL_ID))
        
        ' Initialize transaction from the data
        trans.TransactionDate = dataArray(i, COL_DATE)
        trans.Merchant = CStr(dataArray(i, COL_MERCHANT))
        trans.Category = CStr(dataArray(i, COL_CATEGORY))
        trans.Amount = CDbl(dataArray(i, COL_AMOUNT))
        trans.Account = CStr(dataArray(i, COL_ACCOUNT))
        trans.Owner = CStr(dataArray(i, COL_OWNER))
        trans.IsShared = CBool(dataArray(i, COL_IS_SHARED))
        trans.Notes = CStr(dataArray(i, COL_NOTES))
        trans.SourceFile = CStr(dataArray(i, COL_SOURCE_FILE))
        
        ' Add to collection with ID as key for fast lookups
        On Error Resume Next
        m_Transactions.Add trans, trans.ID
        
        If Err.Number = 457 Then  ' Collection key already exists
            If Not m_Logger Is Nothing Then
                m_Logger.LogWarning "TransactionRepository.LoadTransactionsFromTable", _
                    "Duplicate transaction ID found: " & trans.ID
            End If
        End If
        On Error GoTo ErrorHandler
    Next i
    
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "TransactionRepository.LoadTransactionsFromTable", _
            "Loaded " & m_Transactions.Count & " transactions from table"
    End If
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.LoadTransactionsFromTable", Err.Number, Err.Description
    End If
End Sub

' Generates a unique ID for a transaction
Private Function GenerateUniqueID() As String
    ' Create a GUID-like ID with timestamp to ensure uniqueness
    Dim timeStamp As String
    timeStamp = Format(Now, "yyyymmddhhnnss")
    
    Dim randomPart As String
    randomPart = Format(Rnd(), "0000000000")
    
    GenerateUniqueID = "TX-" & timeStamp & "-" & randomPart
End Function

' Checks for duplicate transactions and resolves conflicts
' Consider changing parameter to ITransaction if needed, but requires careful handling
Private Function ResolveDuplicate(ByVal newTrans As Transaction) As Transaction
    On Error GoTo ErrorHandler
    
    ' If transaction has an ID, check if it already exists
    If Len(newTrans.ID) > 0 Then
        On Error Resume Next
        Dim existingTrans As Transaction
        Set existingTrans = m_Transactions(newTrans.ID)
        On Error GoTo ErrorHandler
        
        ' If found, this is an update to an existing transaction
        If Not existingTrans Is Nothing Then
            ' For now, just use the new transaction
            Set ResolveDuplicate = newTrans
            Exit Function
        End If
    End If
    
    ' Check for potential duplicates based on date, merchant, and amount
    Dim potentialDuplicate As Boolean
    potentialDuplicate = False
    
    ' Create a key to detect potential duplicates
    Dim dateStr As String
    dateStr = Format(newTrans.TransactionDate, "yyyy-mm-dd")
    
    Dim dupeKey As String
    dupeKey = dateStr & "|" & newTrans.Merchant & "|" & newTrans.Amount
    
    ' Look for matching transactions
    Dim trans As Transaction
    For Each trans In m_Transactions
        Dim transDateStr As String
        transDateStr = Format(trans.TransactionDate, "yyyy-mm-dd")
        
        Dim transKey As String
        transKey = transDateStr & "|" & trans.Merchant & "|" & trans.Amount
        
        If transKey = dupeKey Then
            ' Found a potential duplicate
            potentialDuplicate = True
            
            ' Log the potential duplicate
            If Not m_Logger Is Nothing Then
                m_Logger.LogWarning "TransactionRepository.ResolveDuplicate", _
                    "Potential duplicate found for transaction on " & dateStr & _
                    " at " & newTrans.Merchant & " for " & Format(newTrans.Amount, "$#,##0.00")
            End If
            
            ' Simple resolution strategy - use the newest file as source of truth
            ' This could be enhanced with more sophisticated rules
            If Len(trans.SourceFile) > 0 And Len(newTrans.SourceFile) > 0 Then
                ' If new transaction has a more recent file date, use it
                If newTrans.SourceFile > trans.SourceFile Then
                    Set ResolveDuplicate = newTrans
                Else
                    Set ResolveDuplicate = trans
                End If
                
                Exit Function
            End If
        End If
    Next trans
    
    ' If no duplicate found, use the new transaction
    Set ResolveDuplicate = newTrans
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.ResolveDuplicate", Err.Number, Err.Description
    End If
    
    ' On error, default to using the new transaction
    Set ResolveDuplicate = newTrans
End Function

'=========================================================================
' Cleanup
'=========================================================================

Private Sub Class_Terminate()
    ' Save any pending changes
    If m_IsDirty Then
        On Error Resume Next
        ITransactionRepository_SaveChanges
        On Error GoTo 0
    End If
    
    ' Clean up object references to prevent memory leaks
    Set m_Transactions = Nothing
    Set m_DataSheet = Nothing
    Set m_DataTable = Nothing
    Set m_Logger = Nothing
End Sub
