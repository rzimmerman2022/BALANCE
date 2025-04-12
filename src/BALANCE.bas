Attribute VB_Name = "BALANCE"
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' Main module - Entry points and global functions

' ===== Public Type Definitions =====
' Note: Consider moving these Types to Module_Types.bas for better organization

' Log levels enum (Moved from ErrorLogger.cls comments, assumed needed)
Public Enum LogLevel
    LogLevel_Error = 1
    LogLevel_Warning = 2
    LogLevel_Info = 3
    LogLevel_Debug = 4
End Enum

' Event types enum (Moved from EventManager.cls comments, assumed needed)
Public Enum EventType
    TransactionsChanged = 1
    BalanceUpdated = 2
    CategoryDataChanged = 3
    ImportCompleted = 4
    SettingsChanged = 5
    ExportCompleted = 6
    FilterApplied = 7
    ViewChanged = 8
    DashboardRefreshed = 9
    ' Add more event types as needed
End Enum

' For TransactionAnalyzer results (Consider moving to Module_Types.bas)
Public Type CategorySummary
    Category As String
    TotalAmount As Currency
    Percentage As Double
End Type

Public Type MonthSummary
    MonthKey As String ' YYYY-MM
    TotalAmount As Currency
    IncomeAmount As Currency
    ExpenseAmount As Currency
End Type

Public Type DayOfWeekSummary
    DayName As String
    TotalAmount As Currency
    Percentage As Double
End Type

Public Type BalanceSummary
    NetBalance As Currency
    OwedAmount As Currency
    OwingUser As String
    OwedUser As String
    WhoOwes As String ' Formatted string like "User1 owes User2 $X.XX"
End Type

' For CSVImportEngine results (Consider moving to Module_Types.bas)
Public Type ImportResult
    Success As Boolean
    TransactionsAdded As Long
    DuplicatesSkipped As Long
    ErrorsEncountered As Long
    ElapsedSeconds As Double
    ErrorMessages As Collection ' Collection of error strings
    ImportedFiles As Collection ' Collection of file paths successfully imported
End Type

' ===== End Public Type Definitions =====

' Initialize the BALANCE system
Public Sub InitializeBALANCE()
    On Error GoTo ErrorHandler
    
    ' Initialize settings first
    AppSettings.Initialize
    
    ' Set up error logging (ServiceLocator will initialize if needed)
    ErrorLogger.LogInfo "BALANCE", "Initializing BALANCE system" ' Use the initialized logger
    
    ' Initialize repository (ServiceLocator will initialize if needed)
    TransactionRepository.Initialize ' Ensure it's initialized via its own method or locator
    
    ' Set up dashboard (ServiceLocator will initialize if needed)
    DashboardManager.Initialize
    ' DashboardManager.SetupDashboard ' *** REMOVED THIS LINE *** - Setup is handled within Initialize/GetOrCreateSheet
    
    ' Log initialization complete
    ErrorLogger.LogInfo "BALANCE", "BALANCE system initialized successfully"
    
    ' Optionally activate dashboard on init
    ShowDashboard

    Exit Sub
    
ErrorHandler:
    Dim errorMsg As String
    errorMsg = "ERROR initializing BALANCE: (" & Err.Number & ") " & Err.Description
    Debug.Print errorMsg
    MsgBox errorMsg, vbExclamation, "BALANCE Initialization Error"
    ' Log error if logger was initialized before the error occurred
    If Not ServiceLocator.GetErrorLogger Is Nothing Then
        On Error Resume Next ' Avoid error loop if logging fails
        ServiceLocator.GetErrorLogger.LogError "InitializeBALANCE", Err.Number, Err.Description
        On Error GoTo 0
    End If
End Sub


' Refresh data and update dashboard
Public Sub RefreshData()
    On Error GoTo ErrorHandler
    
    ' Mark dashboard as needing refresh (Property might not exist - use RefreshDashboard directly)
    ' DashboardManager.NeedsRefresh = True ' Assuming NeedsRefresh property doesn't exist based on previous code
    
    ' Update dashboard
    DashboardManager.RefreshDashboard ' Directly call the refresh method
    
    ' Update last update date in settings
    AppSettings.LastUpdateDate = Now
    
    ' Log refresh
    ErrorLogger.LogInfo "BALANCE", "Data refreshed"
    
    ' Update status
    MsgBox "Data refreshed successfully!", vbInformation, AppSettings.AppTitle
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.RefreshData", Err.Number, Err.Description
    MsgBox "Error refreshing data: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub

' Show import dialog
Public Sub ShowImportDialog()
    On Error GoTo ErrorHandler
    
    ' Get CSVImportEngine via ServiceLocator (it handles initialization)
    Dim csvEngine As ICSVImportEngine ' Use interface
    Set csvEngine = ServiceLocator.GetCSVImportEngine()
    
    ' Create a simple dialog using InputBox
    Dim choice As String
    choice = InputBox("Import options:" & vbCrLf & _
                      "1. Import single CSV file" & vbCrLf & _
                      "2. Import all CSVs from folder" & vbCrLf & _
                      vbCrLf & _
                      "Enter 1 or 2:", AppSettings.AppTitle, "1")
    
    If choice = "" Then Exit Sub ' Cancelled
    
    Select Case choice
        Case "1"
            ImportSingleCSV ' Call helper sub
        Case "2"
            ImportCSVFolder ' Call helper sub
        Case Else
            MsgBox "Invalid choice. Please try again.", vbExclamation, AppSettings.AppTitle
    End Select
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.ShowImportDialog", Err.Number, Err.Description
    MsgBox "Error showing import dialog: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub

' Import a single CSV file
Private Sub ImportSingleCSV()
    On Error GoTo ErrorHandler
    
    ' Show file dialog
    Dim filePath As String
    filePath = ShowOpenFileDialog("CSV Files (*.csv),*.csv", "Select CSV file to import")
    If filePath = "" Then Exit Sub ' Cancelled
    
    ' Get owner
    Dim owner As String
    owner = GetOwnerSelection()
    If owner = "" Then Exit Sub ' Cancelled

    ' Get CSVImportEngine via ServiceLocator
    Dim csvEngine As ICSVImportEngine ' Use interface
    Set csvEngine = ServiceLocator.GetCSVImportEngine()
       
    ' Import file using the engine's public method
    Dim resultCol As Collection
    Set resultCol = csvEngine.ImportTransactionsFromFile(filePath, owner)
    
    ' Show result (Simplified message for now)
    If Err.Number = 0 Then
         MsgBox "Import complete!" & vbCrLf & _
                "Transactions Processed: " & resultCol.Count, _
                vbInformation, AppSettings.AppTitle
         ' Refresh dashboard
         DashboardManager.RefreshDashboard ' Directly refresh
    Else
        GoTo ErrorHandler ' Jump to error handler if import function failed
    End If

    Exit Sub
    
ErrorHandler:
    Dim errorNum As Long: errorNum = Err.Number
    Dim errorDesc As String: errorDesc = Err.Description
    ErrorLogger.LogError "BALANCE.ImportSingleCSV", errorNum, errorDesc
    MsgBox "Error importing CSV: (" & errorNum & ") " & errorDesc, vbExclamation, AppSettings.AppTitle
End Sub

' Import CSVs from a folder
Private Sub ImportCSVFolder()
    On Error GoTo ErrorHandler
    
    ' Show folder dialog
    Dim folderPath As String
    folderPath = Utilities.BrowseForFolder("Select folder with CSV files")
    If folderPath = "" Then Exit Sub ' Cancelled
    
    ' Get owner
    Dim owner As String
    owner = GetOwnerSelection()
    If owner = "" Then Exit Sub ' Cancelled

    ' Get CSVImportEngine via ServiceLocator
    Dim csvEngine As ICSVImportEngine ' Use interface
    Set csvEngine = ServiceLocator.GetCSVImportEngine()
    
    ' Import folder using the engine's public method
    Dim resultCol As Collection
    Set resultCol = csvEngine.ImportTransactionsFromDirectory(folderPath, owner)

    ' Show result (Simplified message for now)
     If Err.Number = 0 Then
         MsgBox "Import complete!" & vbCrLf & _
                "Transactions Processed from folder: " & resultCol.Count, _
                vbInformation, AppSettings.AppTitle
         ' Refresh dashboard
         DashboardManager.RefreshDashboard ' Directly refresh
    Else
        GoTo ErrorHandler ' Jump to error handler if import function failed
    End If

    Exit Sub
    
ErrorHandler:
    Dim errorNum As Long: errorNum = Err.Number
    Dim errorDesc As String: errorDesc = Err.Description
    ErrorLogger.LogError "BALANCE.ImportCSVFolder", errorNum, errorDesc
    MsgBox "Error importing CSV folder: (" & errorNum & ") " & errorDesc, vbExclamation, AppSettings.AppTitle
End Sub


' Get owner selection
Private Function GetOwnerSelection() As String
    On Error Resume Next ' Use Resume Next for InputBox cancel
    
    Dim choice As String
    choice = InputBox("Select expense owner:" & vbCrLf & _
                      "1. " & AppSettings.User1Name & vbCrLf & _
                      "2. " & AppSettings.User2Name & vbCrLf & _
                      vbCrLf & _
                      "Enter 1 or 2:", AppSettings.AppTitle, "1")
    
    If choice = "" Then
        GetOwnerSelection = "" ' Cancelled
        Exit Function
    End If
    
    Select Case choice
        Case "1"
            GetOwnerSelection = AppSettings.User1Name
        Case "2"
            GetOwnerSelection = AppSettings.User2Name
        Case Else
            MsgBox "Invalid choice. Defaulting to " & AppSettings.User1Name & ".", _
                   vbExclamation, AppSettings.AppTitle
            GetOwnerSelection = AppSettings.User1Name
    End Select
    
    On Error GoTo 0 ' Restore error handling
End Function

' Show transaction list sheet
Public Sub ShowTransactionList()
    On Error GoTo ErrorHandler
    
    ' Get repository via ServiceLocator
    Dim repo As ITransactionRepository ' Use interface
    Set repo = ServiceLocator.GetTransactionRepository()
    
    ' Get or create Transactions sheet
    Dim transSheet As Worksheet
    Set transSheet = Utilities.GetOrCreateSheet("Transactions", True)
    If transSheet Is Nothing Then Err.Raise vbObjectError + 515, "ShowTransactionList", "Failed to get or create Transactions sheet."
    
    ' Clear existing content
    Utilities.ClearSheet transSheet
    
    ' Set up header
    Dim headers As Variant
    headers = Array("Date", "Merchant", "Category", "Amount", "Account", "Owner", "Shared", "Notes", "ID")
    transSheet.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    
    ' Format header
    With transSheet.Range("A1").Resize(1, UBound(headers) + 1)
        .Font.Bold = True
        .Interior.Color = AppSettings.ColorPrimary
        .Font.Color = AppSettings.ColorLightText
    End With
    
    ' Get transactions
    Dim transactions As Collection
    Set transactions = repo.GetTransactions() ' Use GetTransactions method
    
    ' Populate transactions if any exist
    If transactions.Count > 0 Then
        Dim row As Long
        row = 2
        Dim dataArray() As Variant
        ReDim dataArray(1 To transactions.Count, 1 To UBound(headers) + 1)
        
        Dim trans As ITransaction ' Use interface
        Dim i As Long: i = 1
        For Each trans In transactions
            dataArray(i, 1) = trans.TransactionDate
            dataArray(i, 2) = trans.Merchant
            dataArray(i, 3) = trans.Category
            dataArray(i, 4) = trans.Amount
            dataArray(i, 5) = trans.Account
            dataArray(i, 6) = trans.Owner
            dataArray(i, 7) = trans.IsShared
            dataArray(i, 8) = trans.Notes
            dataArray(i, 9) = trans.ID
            i = i + 1
        Next trans
        
        ' Write data array to sheet
        transSheet.Range("A2").Resize(transactions.Count, UBound(headers) + 1).Value = dataArray
        
        ' Format columns after data is populated
        transSheet.Columns("A").NumberFormat = "yyyy-mm-dd"
        transSheet.Columns("D").NumberFormat = "$#,##0.00;[Red]($#,##0.00)" ' Use standard accounting format
        transSheet.Columns("G").HorizontalAlignment = xlCenter
        
        ' Color formatting for amount (Example - might be slow for large datasets)
'        Dim cell As Range
'        For Each cell In transSheet.Range("D2:D" & transactions.Count + 1)
'            If cell.Value < 0 Then cell.Font.Color = AppSettings.ColorDanger Else cell.Font.Color = AppSettings.ColorSuccess
'        Next cell
    Else
         transSheet.Range("A2").Value = "No transactions found."
    End If
    
    ' Format table
    transSheet.Columns("A:I").AutoFit
    
    ' Add instructions for editing
    Dim nextRow As Long
    nextRow = transSheet.Cells(transSheet.Rows.Count, "A").End(xlUp).Row + 2
    transSheet.Cells(nextRow, 1).Value = "To edit transactions, modify the values directly in this sheet, then click 'Save Changes' below."
    transSheet.Range("A" & nextRow).Font.Italic = True
    
    ' Add buttons (adjust positioning)
    Dim buttonTop As Double: buttonTop = transSheet.Cells(nextRow + 1, 1).Top + 10
    Utilities.AddButton transSheet, 20, buttonTop, 150, 30, "Save Changes", "BALANCE.SaveTransactionChanges", AppSettings.ColorSuccess
    Utilities.AddButton transSheet, 190, buttonTop, 150, 30, "Delete Selected", "BALANCE.DeleteSelectedTransaction", AppSettings.ColorDanger
    Utilities.AddButton transSheet, 360, buttonTop, 150, 30, "Back to Dashboard", "BALANCE.ShowDashboard", AppSettings.ColorInfo
    
    ' Activate the sheet
    transSheet.Activate
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.ShowTransactionList", Err.Number, Err.Description
    MsgBox "Error showing transaction list: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub


' Show insights sheet
Public Sub ShowInsights()
    On Error GoTo ErrorHandler
    
    ' Get analyzer via ServiceLocator (it handles initialization)
    Dim analyzer As TransactionAnalyzer ' Use concrete type if DisplayInsights is specific to it
    Set analyzer = ServiceLocator.GetCategoryAnalyzer ' Assuming TransactionAnalyzer provides all analysis interfaces
    
    ' Generate and display insights
    analyzer.DisplayInsights
    
    ' Activate the insights sheet
    analyzer.InsightsSheet.Activate
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.ShowInsights", Err.Number, Err.Description
    MsgBox "Error showing insights: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub

' Show settings sheet
Public Sub ShowSettings()
    On Error GoTo ErrorHandler
    
    ' Ensure settings sheet exists (SaveSettings creates it if needed)
    AppSettings.SaveSettings ' Saving ensures sheet exists and has current values
    
    ' Activate the settings sheet
    ThisWorkbook.Sheets(AppSettings.SettingsSheetName).Visible = xlSheetVisible
    ThisWorkbook.Sheets(AppSettings.SettingsSheetName).Activate
    
    ' Add a back button if it doesn't exist
    On Error Resume Next ' Ignore error if shape already exists
    Dim btnShape As Shape
    Set btnShape = ThisWorkbook.Sheets(AppSettings.SettingsSheetName).Shapes("Button_Back_to_Dashboard")
    If btnShape Is Nothing Then
        Utilities.AddButton ThisWorkbook.Sheets(AppSettings.SettingsSheetName), 20, 200, 150, 30, _
                           "Back to Dashboard", "BALANCE.ShowDashboard", AppSettings.ColorInfo
    End If
    On Error GoTo ErrorHandler ' Restore error handling

    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.ShowSettings", Err.Number, Err.Description
    MsgBox "Error showing settings: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub


' Show the dashboard sheet
Public Sub ShowDashboard()
    On Error GoTo ErrorHandler
    
    ' Get dashboard manager via ServiceLocator (it handles initialization)
    Dim dbManager As IDashboardManager ' Use Interface
    Set dbManager = ServiceLocator.GetDashboardManager()
    
    ' Refresh the dashboard data
    dbManager.RefreshDashboard ' Refresh data and UI elements
    
    ' Activate dashboard sheet (Get sheet object from manager)
    Dim dbSheet As Worksheet
    Set dbSheet = DashboardManager.DashboardSheet ' Assumes DashboardSheet property exists on DashboardManager instance
    If Not dbSheet Is Nothing Then
       dbSheet.Activate
    Else
        Err.Raise vbObjectError + 516, "ShowDashboard", "Failed to get Dashboard sheet object."
    End If

    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.ShowDashboard", Err.Number, Err.Description
    MsgBox "Error showing dashboard: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub

' Save transaction changes from the transaction list sheet
Public Sub SaveTransactionChanges()
    On Error GoTo ErrorHandler
    
    ' Get the transactions sheet
    Dim transSheet As Worksheet
    On Error Resume Next
    Set transSheet = ThisWorkbook.Worksheets("Transactions")
    On Error GoTo ErrorHandler ' Restore error handling
    If transSheet Is Nothing Then
        MsgBox "Transactions sheet not found. Cannot save changes.", vbExclamation, AppSettings.AppTitle
        Exit Sub
    End If
    
    ' Get repository via ServiceLocator
    Dim repo As ITransactionRepository ' Use Interface
    Set repo = ServiceLocator.GetTransactionRepository()
        
    ' --- Efficient Update Strategy: Read sheet data into array, compare with existing repo data ---
    ' --- Simpler (less efficient) Strategy: Clear repo, re-add all from sheet ---
    ' Using the simpler strategy for now:
    
    ' Clear existing transactions in memory (but not the persistent storage yet)
    ' This requires a method in the repository or careful handling.
    ' Let's assume the repository handles updates/adds correctly without needing full clear.
    ' repo.ClearAll ' Avoid clearing all if possible, aim for updates/adds
    
    ' Read transactions from sheet
    Dim lastRow As Long
    lastRow = Utilities.GetLastRow(transSheet, 1) ' Last row in column A (Date)
    
    If lastRow <= 1 Then
        MsgBox "No transactions found on sheet to save.", vbInformation, AppSettings.AppTitle
        ' Decide if this should clear the repo or not - current repo state is kept
        ' repo.ClearAll ' Uncomment if empty sheet means delete all transactions
        ' repo.SaveChanges ' Uncomment if empty sheet means delete all transactions
        Exit Sub
    End If
    
    ' Process each row (more robust approach)
    Dim row As Long
    Dim changedCount As Long: changedCount = 0
    Dim addedCount As Long: addedCount = 0
    Dim errorCount As Long: errorCount = 0
    Dim transactionsToAdd As New Collection ' Collection to hold valid transactions from sheet

    For row = 2 To lastRow
        ' Basic check if row seems valid (e.g., has a date)
        If IsDate(transSheet.Cells(row, 1).Value) Then
             On Error Resume Next ' Handle errors creating/populating single transaction gracefully
             Dim trans As New Transaction ' Use concrete type for InitFromRow
             trans.ID = Utilities.SafeString(transSheet.Cells(row, 9).Value) ' Get existing ID if present

             trans.InitFromRow _
                 transSheet.Cells(row, 1).Value, _
                 transSheet.Cells(row, 2).Value, _
                 transSheet.Cells(row, 3).Value, _
                 transSheet.Cells(row, 4).Value, _
                 transSheet.Cells(row, 5).Value, _
                 transSheet.Cells(row, 6).Value, _
                 transSheet.Cells(row, 7).Value, _
                 transSheet.Cells(row, 8).Value, _
                 "" ' No source file for manually edited transactions

             If Err.Number = 0 Then
                transactionsToAdd.Add trans ' Add valid transaction to collection
             Else
                errorCount = errorCount + 1
                ErrorLogger.LogWarning "SaveTransactionChanges", Err.Number, "Skipping row " & row & " due to error: " & Err.Description
                Err.Clear
             End If
             On Error GoTo ErrorHandler ' Restore main error handling
        End If
    Next row

    ' --- Update Repository More Efficiently (Conceptual) ---
    ' Instead of ClearAll/AddTransactions, ideally:
    ' 1. Get existing transactions from repo into a Dictionary keyed by ID.
    ' 2. Loop through sheet data (transactionsToAdd collection).
    ' 3. If transaction ID exists in repo dict, compare values. If different, update object in repo dict. Remove from dict.
    ' 4. If transaction ID does *not* exist in repo dict, AddTransaction to repo.
    ' 5. Any transactions remaining in repo dict after loop were deleted from sheet -> RemoveTransaction from repo.
    ' This is more complex to implement here.

    ' --- Simple Clear and Re-Add Approach ---
    repo.ClearAll ' Clear existing in-memory transactions
    addedCount = repo.AddTransactions(transactionsToAdd) ' Add all valid transactions from sheet
    
    ' Save changes to persistent storage
    repo.SaveChanges
    
    ' Show result
    Dim resultMsg As String
    resultMsg = "Changes saved!" & vbCrLf & _
               "Transactions Saved/Updated: " & addedCount
    If errorCount > 0 Then resultMsg = resultMsg & vbCrLf & "Rows Skipped due to errors: " & errorCount
    MsgBox resultMsg, vbInformation, AppSettings.AppTitle

    ' Refresh the transaction list UI to reflect saved state (optional)
    ' ShowTransactionList ' Could call this to redraw the sheet from the repo

    ' Mark dashboard as needing refresh
    DashboardManager.RefreshDashboard ' Refresh dashboard as data changed

    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.SaveTransactionChanges", Err.Number, Err.Description
    MsgBox "Error saving changes: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub


' Delete selected transaction from the transaction list sheet
Public Sub DeleteSelectedTransaction()
    On Error GoTo ErrorHandler
    
    ' Get the transactions sheet
    Dim transSheet As Worksheet
    On Error Resume Next
    Set transSheet = ThisWorkbook.Worksheets("Transactions")
    On Error GoTo ErrorHandler ' Restore error handling
    If transSheet Is Nothing Then
        MsgBox "Transactions sheet not found. Cannot delete.", vbExclamation, AppSettings.AppTitle
        Exit Sub
    End If
    
    ' Ensure a single cell is selected
    If TypeName(Selection) <> "Range" Then
         MsgBox "Please select a cell within the transaction row you want to delete.", vbExclamation, AppSettings.AppTitle
         Exit Sub
    End If
    If Selection.Cells.CountLarge > 1 Then
         MsgBox "Please select only one cell within the transaction row.", vbExclamation, AppSettings.AppTitle
         Exit Sub
    End If

    ' Get selected row
    Dim selectedRow As Long
    selectedRow = Selection.Row
    
    ' Check if selection is within the data range
    If selectedRow < 2 Or selectedRow > Utilities.GetLastRow(transSheet, 9) Then ' Check against ID column
        MsgBox "Please select a cell within a valid transaction row to delete.", vbExclamation, AppSettings.AppTitle
        Exit Sub
    End If
    
    ' Get transaction ID from column I (9)
    Dim transactionID As String
    transactionID = Utilities.SafeString(transSheet.Cells(selectedRow, 9).Value) ' Column I for ID
    
    If Len(transactionID) = 0 Then
        MsgBox "Cannot delete this row. No valid transaction ID found in column I.", vbExclamation, AppSettings.AppTitle
        Exit Sub
    End If
    
    ' Confirm deletion
    Dim result As VbMsgBoxResult
    result = MsgBox("Are you sure you want to delete this transaction?" & vbCrLf & vbCrLf & _
                    "Merchant: " & transSheet.Cells(selectedRow, 2).Value & vbCrLf & _
                    "Amount: " & Format(transSheet.Cells(selectedRow, 4).Value, "$#,##0.00"), _
                    vbYesNo + vbQuestion + vbDefaultButton2, AppSettings.AppTitle & " - Confirm Delete")
    
    If result <> vbYes Then Exit Sub ' User cancelled
    
    ' Get repository via ServiceLocator
    Dim repo As ITransactionRepository ' Use Interface
    Set repo = ServiceLocator.GetTransactionRepository()
    
    ' Remove transaction from repository
    If repo.RemoveTransaction(transactionID) Then
        ' Save changes to persistent storage
        repo.SaveChanges
        
        ' Remove row from sheet visually
        Application.EnableEvents = False ' Prevent potential Worksheet_Change events
        transSheet.Rows(selectedRow).Delete
        Application.EnableEvents = True
                
        ' Show result
        MsgBox "Transaction deleted successfully!", vbInformation, AppSettings.AppTitle

        ' Refresh dashboard
        DashboardManager.RefreshDashboard

    Else
        MsgBox "Failed to delete transaction from repository. It might have already been deleted.", vbExclamation, AppSettings.AppTitle
    End If
        
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.DeleteSelectedTransaction", Err.Number, Err.Description
    MsgBox "Error deleting transaction: " & Err.Description, vbExclamation, AppSettings.AppTitle
    Application.EnableEvents = True ' Ensure events are re-enabled on error
End Sub

' Create sample data for testing
Public Sub CreateSampleData()
    On Error GoTo ErrorHandler
    
    ' Confirm creation
    Dim result As VbMsgBoxResult
    result = MsgBox("This will add 50 sample transaction records for testing." & vbCrLf & _
                    "Existing data will NOT be affected." & vbCrLf & vbCrLf & _
                    "Do you want to continue?", _
                    vbYesNo + vbQuestion, AppSettings.AppTitle & " - Create Sample Data")
    
    If result <> vbYes Then Exit Sub
    
    ' Get repository via ServiceLocator
    Dim repo As ITransactionRepository ' Use Interface
    Set repo = ServiceLocator.GetTransactionRepository()
        
    ' Create sample transactions
    Dim transactionsToAdd As New Collection
    Dim trans As Transaction ' Use Concrete type for InitFromRow
    Dim i As Long
    Dim currentDate As Date
    Dim categories As Variant
    Dim merchants As Variant
    Dim accounts As Variant
    
    ' Define sample categories, merchants, accounts
    categories = Array("Groceries", "Dining", "Utilities", "Rent", "Transportation", "Entertainment", "Shopping", "Health", "Travel", "Miscellaneous", "Income")
    merchants = Array("Walmart", "Target", "Kroger", "Amazon", "Netflix", "Gas Station", "Electric Co", "Restaurant XYZ", "Pharmacy", "Dept Store", "Paycheck", "Side Gig")
    accounts = Array("Checking", "Credit Card A", "Credit Card B", "Cash")
    
    ' Start from 3 months ago
    currentDate = DateAdd("m", -3, Date)
    
    ' Create 50 sample transactions
    For i = 1 To 50
        Set trans = New Transaction ' Create new instance each time
        
        Dim transDate As Date: transDate = DateAdd("d", Int(Rnd * 90), currentDate)
        Dim transMerchant As String: transMerchant = merchants(Int(Rnd * UBound(merchants) + LBound(merchants)))
        Dim transCategory As String: transCategory = categories(Int(Rnd * UBound(categories) + LBound(categories)))
        Dim transAmount As Currency
        If transCategory = "Income" Or Rnd < 0.1 Then ' 10% chance of income unless category is Income
            transAmount = Round((Rnd * 500) + 100, 2) ' $100 to $600
        Else
            transAmount = -Round((Rnd * 150) + 5, 2) ' $5 to $155 expense
        End If
        Dim transAccount As String: transAccount = accounts(Int(Rnd * UBound(accounts) + LBound(accounts)))
        Dim transOwner As String
        If Rnd < 0.5 Then transOwner = AppSettings.User1Name Else transOwner = AppSettings.User2Name
        Dim transIsShared As Boolean: transIsShared = (transCategory <> "Income" And Rnd < 0.8) ' 80% of expenses are shared
        Dim transNotes As String: If Rnd < 0.2 Then transNotes = "Sample note " & i Else transNotes = ""
        Dim transSourceFile As String: transSourceFile = "SampleData"

        ' Use InitFromRow
        trans.InitFromRow transDate, transMerchant, transCategory, transAmount, transAccount, transOwner, transIsShared, transNotes, transSourceFile
        
        ' Add to collection for bulk add
        transactionsToAdd.Add trans
    Next i
    
    ' Add transactions to repository
    Dim addedCount As Long
    addedCount = repo.AddTransactions(transactionsToAdd)
    
    ' Save changes
    repo.SaveChanges
        
    ' Show result
    MsgBox "Sample data created!" & vbCrLf & _
           addedCount & " sample transactions added.", _
           vbInformation, AppSettings.AppTitle

    ' Refresh dashboard
    DashboardManager.RefreshDashboard
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.CreateSampleData", Err.Number, Err.Description
    MsgBox "Error creating sample data: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub


' Show open file dialog using FileDialog object
Private Function ShowOpenFileDialog(filter As String, title As String) As String
    On Error GoTo ErrorHandler
    
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        .Title = title
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv" ' Add specific filter description
        If filter <> "*.csv" Then .Filters.Add "All Files", "*.*" ' Optionally add All Files
        .FilterIndex = 1 ' Default to CSV filter
        
        If .Show = True Then
            ShowOpenFileDialog = .SelectedItems(1)
        Else
            ShowOpenFileDialog = "" ' Return empty string if cancelled
        End If
    End With

    Set fd = Nothing ' Clean up
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.ShowOpenFileDialog", Err.Number, Err.Description
    ShowOpenFileDialog = "" ' Return empty string on error
    Set fd = Nothing
End Function