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
Private m_Logger As IErrorLogger      ' Error logging component (injected)

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
        ' Use ServiceLocator to get the logger instance
        Set m_Logger = ServiceLocator.GetErrorLogger()
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
    Else
        Debug.Print "ERROR in TransactionRepository.Initialize: " & Err.Number & " - " & Err.Description
    End If
End Sub

' Public wrapper for initialization
Public Sub Initialize(Optional ByVal logger As IErrorLogger = Nothing)
    ITransactionRepository_Initialize logger
End Sub

' Get or create the data sheet
Private Function GetOrCreateDataSheet() As Worksheet
    On Error Resume Next ' Use Resume Next carefully for object checking
    
    Dim ws As Worksheet
    
    ' Try to get existing sheet
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET_NAME)
    
    ' If sheet doesn't exist, create it
    If ws Is Nothing Then
        Err.Clear ' Clear any error from trying to get the sheet
        Set ws = ThisWorkbook.Worksheets.Add ' Add new sheet
        If Err.Number <> 0 Then GoTo HandleSheetError ' Check if Add failed
        
        On Error Resume Next ' Handle potential error during rename
        ws.Name = DATA_SHEET_NAME
        If Err.Number <> 0 Then GoTo HandleSheetError ' Check if rename failed
        On Error GoTo 0 ' Restore default error handling
        
        ' Hide the sheet (very hidden so it can't be unhidden through the UI)
        ws.Visible = xlSheetVeryHidden
        
        If Not m_Logger Is Nothing Then
            m_Logger.LogInfo "TransactionRepository.GetOrCreateDataSheet", _
                "Created new data sheet: " & DATA_SHEET_NAME
        End If
    End If
    
    Set GetOrCreateDataSheet = ws
    Exit Function

HandleSheetError:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.GetOrCreateDataSheet", Err.Number, "Failed to create/rename data sheet: " & Err.Description
    Else
        Debug.Print "ERROR in TransactionRepository.GetOrCreateDataSheet: " & Err.Number & " - " & Err.Description
    End If
    Set GetOrCreateDataSheet = Nothing ' Return Nothing on error
    On Error GoTo 0 ' Clear error context
End Function


' Get or create the data table
Private Function GetOrCreateDataTable() As ListObject
    On Error Resume Next ' Use Resume Next carefully
    
    Dim tbl As ListObject
    
    ' Ensure data sheet exists
    If m_DataSheet Is Nothing Then Exit Function
    
    ' Check if table exists
    Set tbl = m_DataSheet.ListObjects(TABLE_NAME)
    
    ' Create table if it doesn't exist
    If tbl Is Nothing Then
        Err.Clear ' Clear error from trying to get the table
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
        
        ' Define the range for the table (header row only initially)
        Dim headerRange As Range
        Set headerRange = m_DataSheet.Range(m_DataSheet.Cells(HEADER_ROW, COL_ID), _
                                            m_DataSheet.Cells(HEADER_ROW, COL_SOURCE_FILE))
        
        ' Create table with the headers
        Set tbl = m_DataSheet.ListObjects.Add(xlSrcRange, headerRange, , xlYes)
        If Err.Number <> 0 Then GoTo HandleTableError ' Check if Add failed
        
        On Error Resume Next ' Handle potential error during rename
        tbl.Name = TABLE_NAME
        If Err.Number <> 0 Then GoTo HandleTableError ' Check if rename failed
        On Error GoTo 0 ' Restore default error handling
        
        ' Format header row
        With tbl.HeaderRowRange
            .Font.Bold = True
            .Interior.Color = AppSettings.ColorPrimary ' Use color from settings
            .Font.Color = AppSettings.ColorLightText ' Use color from settings
        End With
                
        If Not m_Logger Is Nothing Then
            m_Logger.LogInfo "TransactionRepository.GetOrCreateDataTable", _
                "Created new transactions table: " & TABLE_NAME
        End If
    End If
    
    ' Apply formatting (even if table already existed)
    On Error Resume Next ' Ignore errors if table has no data body range yet
    ' Format date column
    tbl.ListColumns("Date").DataBodyRange.NumberFormat = "yyyy-mm-dd"
    ' Format amount column
    tbl.ListColumns("Amount").DataBodyRange.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    On Error GoTo 0 ' Restore default error handling

    Set GetOrCreateDataTable = tbl
    Exit Function

HandleTableError:
     If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.GetOrCreateDataTable", Err.Number, "Failed to create/rename table: " & Err.Description
    Else
        Debug.Print "ERROR in TransactionRepository.GetOrCreateDataTable: " & Err.Number & " - " & Err.Description
    End If
    Set GetOrCreateDataTable = Nothing ' Return Nothing on error
    On Error GoTo 0 ' Clear error context
End Function

'=========================================================================
' Interface Implementation - Get Methods
'=========================================================================

' Returns all transactions in the repository
Private Function ITransactionRepository_GetTransactions() As Collection
    On Error GoTo ErrorHandler
    
    ' Ensure repository is initialized
    If Not m_IsInitialized Then Me.Initialize ' Use Me.Initialize to call public wrapper
    
    ' Create a new collection to return (don't return direct reference to internal collection)
    Dim result As Collection
    Set result = New Collection
    
    ' Copy all transactions to the new collection
    Dim trans As ITransaction ' Use Interface type
    For Each trans In m_Transactions
        ' Use the transaction ID as the key for fast lookups
        On Error Resume Next ' Handle potential duplicate keys if ID generation had issues
        result.Add trans, trans.ID
        If Err.Number <> 0 Then
             If Not m_Logger Is Nothing Then
                 m_Logger.LogWarning "TransactionRepository.GetTransactions", Err.Number, "Could not add transaction with potentially duplicate key: " & trans.ID
             End If
             Err.Clear
        End If
        On Error GoTo ErrorHandler ' Restore error handling
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
    If Not m_IsInitialized Then Me.Initialize
    
    ' Create a new collection for the results
    Dim result As Collection
    Set result = New Collection
    
    ' Filter transactions by date range
    Dim trans As ITransaction ' Use Interface type
    For Each trans In m_Transactions
        If trans.TransactionDate >= startDate And trans.TransactionDate <= endDate Then
            result.Add trans ' Cannot use key here as it's a subset
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
    If Not m_IsInitialized Then Me.Initialize
    
    ' Create a new collection for the results
    Dim result As Collection
    Set result = New Collection
    
    ' Filter transactions by category (case-insensitive)
    Dim trans As ITransaction ' Use Interface type
    For Each trans In m_Transactions
        If StrComp(trans.Category, category, vbTextCompare) = 0 Then ' Use StrComp for case-insensitivity
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
    If Not m_IsInitialized Then Me.Initialize
    
    ' Create a new collection for the results
    Dim result As Collection
    Set result = New Collection
    
    ' Filter transactions by owner (case-insensitive)
    Dim trans As ITransaction ' Use Interface type
    For Each trans In m_Transactions
        If StrComp(trans.Owner, owner, vbTextCompare) = 0 Then ' Use StrComp
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
    If Not m_IsInitialized Then Me.Initialize
    
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
Private Function ITransactionRepository_AddTransaction(ByVal transaction As ITransaction) As Boolean
    On Error GoTo ErrorHandler
    
    ' Ensure repository is initialized
    If Not m_IsInitialized Then Me.Initialize
    
    ' Check for duplicates and resolve if needed
    Dim transToAdd As ITransaction
    If TypeOf transaction Is Transaction Then
        ' If it's a concrete Transaction, try resolving duplicates
        Set transToAdd = ResolveDuplicate(transaction) ' ResolveDuplicate expects Transaction, returns Transaction (which implements ITransaction)
    Else
        ' If it's some other ITransaction implementation, add without duplicate check for now
        Set transToAdd = transaction
        If Not m_Logger Is Nothing Then
            m_Logger.LogWarning "TransactionRepository.AddTransaction", 0, "Adding ITransaction object without duplicate check (Type: " & TypeName(transaction) & ")"
        End If
    End If

    ' Generate ID if not set (ensure ID is not empty)
    If Len(transToAdd.ID) = 0 Then
        transToAdd.ID = GenerateUniqueID()
        If Not m_Logger Is Nothing Then m_Logger.LogInfo "TransactionRepository.AddTransaction", "Generated new ID: " & transToAdd.ID
    End If
    
    ' Add to collection (remove first if it already exists by ID)
    On Error Resume Next ' To handle case where key doesn't exist
    m_Transactions.Remove transToAdd.ID
    On Error GoTo ErrorHandler ' Restore default error handling
    
    m_Transactions.Add transToAdd, transToAdd.ID ' Add using ID as key
    
    ' Mark as dirty (needs saving)
    m_IsDirty = True
    
    ' Publish event to notify that transactions have changed
    EventManager.PublishEvent EventType.TransactionsChanged
    
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "TransactionRepository.AddTransaction", _
            "Added/Updated transaction ID: " & transToAdd.ID
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
Public Function AddTransaction(ByVal transaction As ITransaction) As Boolean
    AddTransaction = ITransactionRepository_AddTransaction(transaction)
End Function

' Adds multiple transactions to the repository
Private Function ITransactionRepository_AddTransactions(ByVal transactions As Collection) As Long
    On Error GoTo ErrorHandler
    
    ' Ensure repository is initialized
    If Not m_IsInitialized Then Me.Initialize
    
    Dim addedCount As Long
    addedCount = 0
    Dim initialDirtyState As Boolean
    initialDirtyState = m_IsDirty ' Track if it was already dirty
    
    ' Process each transaction - Use Variant loop and check type
    Dim transVariant As Variant
    Dim trans As ITransaction ' Use interface type
    For Each transVariant In transactions
        Set trans = transVariant ' Assumes collection contains objects implementing ITransaction
        If ITransactionRepository_AddTransaction(trans) Then ' Call the single add method which handles duplicates and dirty flag
            addedCount = addedCount + 1
        End If
    Next transVariant ' ***** CORRECTED LINE *****
    
    ' Only publish event if something was actually added OR if state changed from clean to dirty
    If addedCount > 0 Or (Not initialDirtyState And m_IsDirty) Then
       EventManager.PublishEvent EventType.TransactionsChanged
    End If
    
    ' Return the number of transactions successfully processed (added or updated)
    ITransactionRepository_AddTransactions = addedCount
    
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "TransactionRepository.AddTransactions", _
            "Processed " & transactions.Count & " transactions, Added/Updated: " & addedCount
    End If
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "TransactionRepository.AddTransactions", Err.Number, Err.Description
    End If
    ITransactionRepository_AddTransactions = addedCount ' Return count added before error
End Function


' Public wrapper for AddTransactions
Public Function AddTransactions(ByVal transactions As Collection) As Long
    AddTransactions = ITransactionRepository_AddTransactions(transactions)
End Function

' Removes a transaction from the repository by ID
Private Function ITransactionRepository_RemoveTransaction(ByVal transactionId As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Ensure repository is initialized
    If Not m_IsInitialized Then Me.Initialize
    
    ' Try to remove the transaction
    On Error Resume Next ' To check if key exists before removing
    Dim objCheck As Object
    Set objCheck = m_Transactions(transactionId) ' Check if key exists
    
    If Err.Number = 0 Then ' Key exists
        Set objCheck = Nothing
        m_Transactions.Remove transactionId
        If Err.Number = 0 Then ' Removal successful
            m_IsDirty = True
            ITransactionRepository_RemoveTransaction = True
            
            ' Publish event to notify that transactions have changed
            EventManager.PublishEvent EventType.TransactionsChanged
            
            If Not m_Logger Is Nothing Then
                m_Logger.LogInfo "TransactionRepository.RemoveTransaction", _
                    "Removed transaction ID: " & transactionId
            End If
        Else ' Error during removal (unexpected)
             ITransactionRepository_RemoveTransaction = False
             If Not m_Logger Is Nothing Then
                 m_Logger.LogError "TransactionRepository.RemoveTransaction", Err.Number, "Error removing transaction ID: " & transactionId & " - " & Err.Description
             End If
        End If
    Else ' Key does not exist
        ITransactionRepository_RemoveTransaction = False
        If Not m_Logger Is Nothing Then
            m_Logger.LogWarning "TransactionRepository.RemoveTransaction", _
                "Transaction not found for removal, ID: " & transactionId
        End If
        Err.Clear
    End If
    
    On Error GoTo ErrorHandler ' Restore default error handling
    
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
    If Not m_IsInitialized Then Me.Initialize
    
    If m_Transactions.Count > 0 Then ' Only mark dirty and publish if something was cleared
        ' Create a new empty collection
        Set m_Transactions = New Collection
        ' Mark as dirty
        m_IsDirty = True
        ' Publish event to notify that transactions have changed
        EventManager.PublishEvent EventType.TransactionsChanged
        
        If Not m_Logger Is Nothing Then
            m_Logger.LogInfo "TransactionRepository.ClearAll", "Cleared all transactions"
        End If
    Else
         If Not m_Logger Is Nothing Then
            m_Logger.LogInfo "TransactionRepository.ClearAll", "No transactions to clear"
        End If
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
    If Not m_IsInitialized Then Me.Initialize
    
    ' If no changes to save, return success
    If Not m_IsDirty Then
        ITransactionRepository_SaveChanges = True
        If Not m_Logger Is Nothing Then m_Logger.LogInfo "TransactionRepository.SaveChanges", "No changes to save."
        Exit Function
    End If
    
    ' Ensure table object is valid
    If m_DataTable Is Nothing Then
        If Not m_Logger Is Nothing Then m_Logger.LogError "TransactionRepository.SaveChanges", 0, "Data table object is not initialized. Cannot save."
        ITransactionRepository_SaveChanges = False
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
    On Error Resume Next ' Handle error if DataBodyRange is Nothing
    m_DataTable.DataBodyRange.Delete
    On Error GoTo ErrorHandler ' Restore error handling
    
    ' Check if there are transactions to save
    Dim transCount As Long
    transCount = m_Transactions.Count
    
    If transCount > 0 Then
        ' Prepare data array for bulk insertion
        Dim dataArray() As Variant
        ReDim dataArray(1 To transCount, 1 To COL_SOURCE_FILE) ' Match number of columns
        
        ' Fill array with transaction data
        Dim i As Long
        Dim trans As ITransaction ' Use interface
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
        
        ' Resize table to accommodate new data (if needed) - safer than adding rows one by one
        ' Note: Adding rows before setting value might be slightly more performant for large sets,
        ' but resizing is often simpler and handles the empty table case cleanly.
        m_DataTable.Resize m_DataTable.HeaderRowRange.Resize(transCount + 1) ' Resize including header

        ' Add data to the table body range
        m_DataTable.DataBodyRange.Value = dataArray
        
        ' Re-apply formatting after adding data
        On Error Resume Next ' Ignore errors if table has no data body range yet
        m_DataTable.ListColumns("Date").DataBodyRange.NumberFormat = "yyyy-mm-dd"
        m_DataTable.ListColumns("Amount").DataBodyRange.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        On Error GoTo ErrorHandler ' Restore error handling
    End If
    
    ' Mark as clean (all changes saved)
    m_IsDirty = False
    
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
    
    ' Ensure table object is valid
    If m_DataTable Is Nothing Then
       If Not m_Logger Is Nothing Then m_Logger.LogError "TransactionRepository.LoadTransactionsFromTable", 0, "Data table object is not initialized. Cannot load."
       Exit Sub
    End If
    
    ' If table has no data rows, exit
    If m_DataTable.ListRows.Count = 0 Then
        If Not m_Logger Is Nothing Then
            m_Logger.LogInfo "TransactionRepository.LoadTransactionsFromTable", _
                "No transactions found in table"
        End If
        Exit Sub
    End If
    
    ' Get all data at once for better performance
    Dim dataArray As Variant
    dataArray = m_DataTable.DataBodyRange.Value
    
    ' Check if dataArray is actually an array (it won't be if there's only one row)
    If Not IsArray(dataArray) Then
        ' Handle single row case
        Dim singleRowArray(1 To 1, 1 To UBound(m_DataTable.DataBodyRange.Value, 2)) As Variant
        Dim k As Long
        For k = 1 To UBound(m_DataTable.DataBodyRange.Value, 2)
            singleRowArray(1, k) = m_DataTable.DataBodyRange.Cells(1, k).Value
        Next k
        dataArray = singleRowArray
    End If

    ' Process data array
    Dim i As Long
    Dim rowCount As Long
    rowCount = UBound(dataArray, 1)
    
    For i = 1 To rowCount
        ' Create and initialize new transaction
        Dim trans As New Transaction ' Use concrete class here
        
        ' Use helper to populate transaction, handles potential errors per row
        PopulateTransactionFromRow trans, dataArray, i
        
        ' Add to collection with ID as key for fast lookups
        If Len(trans.ID) > 0 Then ' Only add if ID is valid
            On Error Resume Next ' Handle potential duplicate keys if ID generation had issues
            m_Transactions.Add trans, trans.ID
            If Err.Number = 457 Then ' Collection key already exists
                If Not m_Logger Is Nothing Then
                    m_Logger.LogWarning "TransactionRepository.LoadTransactionsFromTable", _
                        "Duplicate transaction ID found during load: " & trans.ID & ". Skipping row " & i + 1 & "."
                End If
                Err.Clear
            ElseIf Err.Number <> 0 Then
                 If Not m_Logger Is Nothing Then
                    m_Logger.LogError "TransactionRepository.LoadTransactionsFromTable", Err.Number, "Error adding transaction ID " & trans.ID & " to collection: " & Err.Description
                End If
                Err.Clear
            End If
            On Error GoTo ErrorHandler ' Restore error handling
        Else
            If Not m_Logger Is Nothing Then
                 m_Logger.LogWarning "TransactionRepository.LoadTransactionsFromTable", 0, "Transaction in row " & i + 1 & " has invalid/empty ID. Skipping."
            End If
        End If
        
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

' Helper to populate a Transaction object from an array row
Private Sub PopulateTransactionFromRow(trans As Transaction, dataArray As Variant, rowIndex As Long)
    On Error Resume Next ' Handle errors for individual fields gracefully

    trans.ID = Utilities.SafeString(dataArray(rowIndex, COL_ID))
    trans.TransactionDate = CDate(dataArray(rowIndex, COL_DATE))
    If Err.Number <> 0 Then trans.TransactionDate = Date: Err.Clear ' Default if conversion fails

    trans.Merchant = Utilities.SafeString(dataArray(rowIndex, COL_MERCHANT), "Unknown")
    trans.Category = Utilities.SafeString(dataArray(rowIndex, COL_CATEGORY), "Uncategorized")

    trans.Amount = CCur(dataArray(rowIndex, COL_AMOUNT))
    If Err.Number <> 0 Then trans.Amount = 0: Err.Clear ' Default if conversion fails

    trans.Account = Utilities.SafeString(dataArray(rowIndex, COL_ACCOUNT))
    trans.Owner = Utilities.SafeString(dataArray(rowIndex, COL_OWNER))
    
    trans.IsShared = CBool(dataArray(rowIndex, COL_IS_SHARED))
    If Err.Number <> 0 Then trans.IsShared = True: Err.Clear ' Default if conversion fails

    trans.Notes = Utilities.SafeString(dataArray(rowIndex, COL_NOTES))
    trans.SourceFile = Utilities.SafeString(dataArray(rowIndex, COL_SOURCE_FILE))

    On Error GoTo 0 ' Restore default error handling
End Sub


' Generates a unique ID for a transaction
Private Function GenerateUniqueID() As String
    ' Create a GUID-like ID with timestamp to ensure uniqueness
    Dim timeStamp As String
    timeStamp = Format(Now, "yyyymmddhhnnss")
    
    Dim randomPart As String
    randomPart = Format(Rnd(), "0000000000") ' More random digits
    
    GenerateUniqueID = "TX-" & timeStamp & "-" & randomPart
End Function

' Checks for duplicate transactions and resolves conflicts
' This implementation assumes newTrans is the definitive version if IDs match.
' Duplicate checks based on content are simple; could be enhanced.
Private Function ResolveDuplicate(ByVal newTrans As Transaction) As Transaction
    On Error GoTo ErrorHandler
    
    ' If transaction has an ID, check if it already exists
    If Len(newTrans.ID) > 0 Then
        On Error Resume Next ' Check if key exists
        Dim existingTrans As ITransaction ' Use Interface
        Set existingTrans = m_Transactions(newTrans.ID)
        On Error GoTo ErrorHandler ' Restore error handling
        
        ' If found, this is an update. Return the new transaction.
        If Not existingTrans Is Nothing Then
            If Not m_Logger Is Nothing Then m_Logger.LogInfo "TransactionRepository.ResolveDuplicate", "Updating existing transaction ID: " & newTrans.ID
            Set ResolveDuplicate = newTrans
            Exit Function
        End If
    End If
    
    ' --- Simple Content-Based Duplicate Check (Optional - can be expensive) ---
    ' This checks Date, Merchant, Amount. Could be expanded or removed for performance.
'    Dim potentialDuplicate As Boolean
'    potentialDuplicate = False
'    Dim dateStr As String
'    dateStr = Format(newTrans.TransactionDate, "yyyy-mm-dd")
'    Dim dupeKey As String
'    dupeKey = dateStr & "|" & newTrans.Merchant & "|" & newTrans.Amount
'
'    Dim trans As ITransaction ' Use Interface
'    For Each trans In m_Transactions
'        Dim transDateStr As String
'        transDateStr = Format(trans.TransactionDate, "yyyy-mm-dd")
'        Dim transKey As String
'        transKey = transDateStr & "|" & trans.Merchant & "|" & trans.Amount
'
'        If transKey = dupeKey Then
'            potentialDuplicate = True
'            If Not m_Logger Is Nothing Then
'                m_Logger.LogWarning "TransactionRepository.ResolveDuplicate", _
'                    "Potential content duplicate found for transaction on " & dateStr & _
'                    " at " & newTrans.Merchant & " for " & Format(newTrans.Amount, "$#,##0.00") & ". Keeping existing ID: " & trans.ID
'            End If
'            ' Resolution: Keep the existing transaction, discard the new one by returning the existing one
'            Set ResolveDuplicate = trans ' Return the existing transaction
'            Exit Function
'        End If
'    Next trans
    ' --- End Optional Content-Based Check ---

    ' If no duplicate found (by ID or content), return the new transaction
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
    ' Save any pending changes on termination
    If m_IsDirty Then
        On Error Resume Next ' Attempt to save, but don't stop termination if it fails
        Me.SaveChanges ' Use public wrapper which calls interface method
        On Error GoTo 0
    End If
    
    ' Clean up object references to prevent memory leaks
    Set m_Transactions = Nothing
    Set m_DataSheet = Nothing
    Set m_DataTable = Nothing
    Set m_Logger = Nothing
End Sub