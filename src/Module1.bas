Attribute VB_Name = "BALANCE"
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' Main module - Entry points and global functions

' ===== Public Type Definitions =====

' Log levels enum (Moved from ErrorLogger.cls)
Public Enum LogLevel
    LogLevel_Error = 1
    LogLevel_Warning = 2
    LogLevel_Info = 3
    LogLevel_Debug = 4
End Enum

' Event types enum (Moved from EventManager.cls)
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

' For TransactionAnalyzer results
Public Type CategorySummary
    Category As String
    TotalAmount As Currency
    Percentage As Double
End Type

Public Type MonthSummary
    MonthKey As String ' yyyy-mm
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

' For CSVImportEngine results
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
    
    ' Set up error logging
    ErrorLogger.Initialize
    ErrorLogger.LogInfo "BALANCE", "Initializing BALANCE system"
    
    ' Initialize repository
    TransactionRepository.Initialize
    
    ' Set up dashboard
    DashboardManager.Initialize
    DashboardManager.SetupDashboard
    
    ' Log initialization complete
    ErrorLogger.LogInfo "BALANCE", "BALANCE system initialized successfully"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR initializing BALANCE: " & Err.Description
    MsgBox "Error initializing BALANCE: " & Err.Description, vbExclamation, "BALANCE Initialization Error"
End Sub

' Refresh data and update dashboard
Public Sub RefreshData()
    On Error GoTo ErrorHandler
    
    ' Mark dashboard as needing refresh
    DashboardManager.NeedsRefresh = True
    
    ' Update dashboard
    DashboardManager.RefreshDashboard
    
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
    
    ' Get CSVImportEngine
    Dim csvEngine As CSVImportEngine
    Set csvEngine = CSVImportEngine
    csvEngine.Initialize
    
    ' Create a simple dialog using InputBox
    Dim choice As String
    choice = InputBox("Import options:" & vbCrLf & _
                     "1. Import single CSV file" & vbCrLf & _
                     "2. Import all CSVs from folder" & vbCrLf & _
                     vbCrLf & _
                     "Enter 1 or 2:", AppSettings.AppTitle, "1")
    
    If choice = "" Then
        ' Cancelled
        Exit Sub
    End If
    
    Select Case choice
        Case "1"
            ' Import single file
            ImportSingleCSV
            
        Case "2"
            ' Import folder
            ImportCSVFolder
            
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
    
    If filePath = "" Then
        ' Cancelled
        Exit Sub
    End If
    
    ' Get owner
    Dim owner As String
    owner = GetOwnerSelection()
    
    If owner = "" Then
        ' Cancelled
        Exit Sub
    End If

    ' Get CSVImportEngine
    Dim csvEngine As CSVImportEngine
    Set csvEngine = CSVImportEngine ' Assumes PredeclaredId=True
    ' csvEngine.Initialize ' Initialization might be handled elsewhere or implicitly

    ' Set default owner (Example - might need adjustment based on actual flow)
    csvEngine.DefaultOwner = owner
    
    ' Import file
    Dim result As ImportResult
    result = csvEngine.ImportCSVFile(filePath, owner)
    
    ' Show result
    If result.Success Then
        MsgBox "Import successful!" & vbCrLf & _
              "Transactions added: " & result.TransactionsAdded & vbCrLf & _
              "Duplicates skipped: " & result.DuplicatesSkipped & vbCrLf & _
              "Processing time: " & Format(result.ElapsedSeconds, "0.00") & " seconds", _
              vbInformation, AppSettings.AppTitle
        
        ' Refresh dashboard
        DashboardManager.NeedsRefresh = True
        DashboardManager.RefreshDashboard
    Else
        Dim errorMsg As String
        errorMsg = "Import failed!" & vbCrLf & _
                  "Errors: " & result.ErrorsEncountered & vbCrLf & vbCrLf
        
        ' Add error messages
        Dim msg As Variant
        For Each msg In result.ErrorMessages
            errorMsg = errorMsg & "- " & msg & vbCrLf
        Next msg
        
        MsgBox errorMsg, vbExclamation, AppSettings.AppTitle
    End If
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.ImportSingleCSV", Err.Number, Err.Description
    MsgBox "Error importing CSV: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub

' Import CSVs from a folder
Private Sub ImportCSVFolder()
    On Error GoTo ErrorHandler
    
    ' Show folder dialog
    Dim folderPath As String
    folderPath = Utilities.BrowseForFolder("Select folder with CSV files")
    
    If folderPath = "" Then
        ' Cancelled
        Exit Sub
    End If
    
    ' Get owner
    Dim owner As String
    owner = GetOwnerSelection()
    
    If owner = "" Then
        ' Cancelled
        Exit Sub
    End If

    ' Get CSVImportEngine
    Dim csvEngine As CSVImportEngine
    Set csvEngine = CSVImportEngine ' Assumes PredeclaredId=True
    ' csvEngine.Initialize ' Initialization might be handled elsewhere or implicitly

    ' Set default owner (Example - might need adjustment based on actual flow)
    csvEngine.DefaultOwner = owner
    
    ' Import folder
    Dim result As ImportResult
    result = csvEngine.ImportCSVFolder(folderPath, owner)
    
    ' Show result
    If result.Success Then
        MsgBox "Import successful!" & vbCrLf & _
              "Files imported: " & result.ImportedFiles.Count & vbCrLf & _
              "Transactions added: " & result.TransactionsAdded & vbCrLf & _
              "Duplicates skipped: " & result.DuplicatesSkipped & vbCrLf & _
              "Processing time: " & Format(result.ElapsedSeconds, "0.00") & " seconds", _
              vbInformation, AppSettings.AppTitle
        
        ' Refresh dashboard
        DashboardManager.NeedsRefresh = True
        DashboardManager.RefreshDashboard
    Else
        Dim errorMsg As String
        errorMsg = "Import failed or partially succeeded!" & vbCrLf & _
                  "Files imported: " & result.ImportedFiles.Count & vbCrLf & _
                  "Transactions added: " & result.TransactionsAdded & vbCrLf & _
                  "Errors: " & result.ErrorsEncountered & vbCrLf & vbCrLf
        
        ' Add error messages (up to 5)
        Dim msg As Variant
        Dim msgCount As Long
        msgCount = 0
        For Each msg In result.ErrorMessages
            If msgCount < 5 Then
                errorMsg = errorMsg & "- " & msg & vbCrLf
                msgCount = msgCount + 1
            Else
                errorMsg = errorMsg & "- (and " & (result.ErrorMessages.Count - 5) & " more errors)" & vbCrLf
                Exit For
            End If
        Next msg
        
        MsgBox errorMsg, vbExclamation, AppSettings.AppTitle
    End If
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.ImportCSVFolder", Err.Number, Err.Description
    MsgBox "Error importing CSV folder: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub

' Get owner selection
Private Function GetOwnerSelection() As String
    On Error Resume Next
    
    Dim choice As String
    choice = InputBox("Select expense owner:" & vbCrLf & _
                     "1. " & AppSettings.User1Name & vbCrLf & _
                     "2. " & AppSettings.User2Name & vbCrLf & _
                     vbCrLf & _
                     "Enter 1 or 2:", AppSettings.AppTitle, "1")
    
    If choice = "" Then
        ' Cancelled
        GetOwnerSelection = ""
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
    
    On Error GoTo 0
End Function

' Show transaction list
Public Sub ShowTransactionList()
    On Error GoTo ErrorHandler
    
    ' This would normally use a UserForm, but for now use the Transactions worksheet
    Dim repo As TransactionRepository
    Set repo = TransactionRepository
    repo.Initialize
    
    ' Get or create Transactions sheet
    Dim transSheet As Worksheet
    Set transSheet = Utilities.GetOrCreateSheet("Transactions", True)
    
    ' Clear existing content
    Utilities.ClearSheet transSheet
    
    ' Set up header
    transSheet.Range("A1").Value = "Date"
    transSheet.Range("B1").Value = "Merchant"
    transSheet.Range("C1").Value = "Category"
    transSheet.Range("D1").Value = "Amount"
    transSheet.Range("E1").Value = "Account"
    transSheet.Range("F1").Value = "Owner"
    transSheet.Range("G1").Value = "Shared"
    transSheet.Range("H1").Value = "Notes"
    transSheet.Range("I1").Value = "ID"
    
    ' Format header
    transSheet.Range("A1:I1").Font.Bold = True
    transSheet.Range("A1:I1").Interior.Color = AppSettings.ColorPrimary
    transSheet.Range("A1:I1").Font.Color = AppSettings.ColorLightText
    
    ' Populate transactions
    Dim row As Long
    row = 2
    
    Dim trans As Transaction
    For Each trans In repo.Transactions
        transSheet.Cells(row, 1).Value = trans.TransactionDate
        transSheet.Cells(row, 1).NumberFormat = "mm/dd/yyyy"
        
        transSheet.Cells(row, 2).Value = trans.Merchant
        transSheet.Cells(row, 3).Value = trans.Category
        
        transSheet.Cells(row, 4).Value = trans.Amount
        transSheet.Cells(row, 4).NumberFormat = "$#,##0.00;($#,##0.00)"
        
        If trans.IsExpense Then
            transSheet.Cells(row, 4).Font.Color = AppSettings.ColorDanger
        Else
            transSheet.Cells(row, 4).Font.Color = AppSettings.ColorSuccess
        End If
        
        transSheet.Cells(row, 5).Value = trans.Account
        transSheet.Cells(row, 6).Value = trans.Owner
        transSheet.Cells(row, 7).Value = trans.IsShared
        transSheet.Cells(row, 8).Value = trans.Notes
        transSheet.Cells(row, 9).Value = trans.ID
        
        row = row + 1
    Next trans
    
    ' Format table
    transSheet.Range("A:I").Columns.AutoFit
    
    ' Add instructions for editing
    transSheet.Cells(row + 1, 1).Value = "To edit transactions, modify the values directly in this sheet, then click 'Save Changes' below."
    transSheet.Range("A" & (row + 1)).Font.Italic = True
    
    ' Add save button
    Utilities.AddButton transSheet, 20, (row + 3) * 15, 150, 30, "Save Changes", "BALANCE.SaveTransactionChanges", _
                      AppSettings.ColorSuccess
    
    ' Add delete button
    Utilities.AddButton transSheet, 190, (row + 3) * 15, 150, 30, "Delete Selected", "BALANCE.DeleteSelectedTransaction", _
                      AppSettings.ColorDanger
    
    ' Add back button
    Utilities.AddButton transSheet, 360, (row + 3) * 15, 150, 30, "Back to Dashboard", "BALANCE.ShowDashboard", _
                      AppSettings.ColorInfo
    
    ' Activate the sheet
    transSheet.Activate
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.ShowTransactionList", Err.Number, Err.Description
    MsgBox "Error showing transaction list: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub

' Show insights
Public Sub ShowInsights()
    On Error GoTo ErrorHandler
    
    ' Initialize the analyzer
    Dim analyzer As TransactionAnalyzer
    Set analyzer = TransactionAnalyzer
    analyzer.Initialize
    
    ' Generate and display insights
    analyzer.DisplayInsights
    
    ' Activate the insights sheet
    analyzer.InsightsSheet.Activate
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.ShowInsights", Err.Number, Err.Description
    MsgBox "Error showing insights: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub

' Show settings
Public Sub ShowSettings()
    On Error GoTo ErrorHandler
    
    ' Check if the settings sheet exists
    If Not Utilities.SheetExists(AppSettings.SettingsSheetName) Then
        AppSettings.SaveSettings
    End If
    
    ' Activate the settings sheet
    ThisWorkbook.Sheets(AppSettings.SettingsSheetName).Visible = xlSheetVisible
    ThisWorkbook.Sheets(AppSettings.SettingsSheetName).Activate
    
    ' Add a back button if it doesn't exist
    On Error Resume Next
    If ThisWorkbook.Sheets(AppSettings.SettingsSheetName).Shapes("Button_Back_to_Dashboard") Is Nothing Then
        Utilities.AddButton ThisWorkbook.Sheets(AppSettings.SettingsSheetName), 20, 300, 150, 30, _
                          "Back to Dashboard", "BALANCE.ShowDashboard", AppSettings.ColorInfo
    End If
    On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.ShowSettings", Err.Number, Err.Description
    MsgBox "Error showing settings: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub

' Show the dashboard
Public Sub ShowDashboard()
    On Error GoTo ErrorHandler
    
    ' Make sure dashboard is initialized
    DashboardManager.Initialize
    
    ' Refresh if needed
    If DashboardManager.NeedsRefresh Then
        DashboardManager.RefreshDashboard
    End If
    
    ' Activate dashboard sheet
    DashboardManager.DashboardSheet.Activate
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.ShowDashboard", Err.Number, Err.Description
    MsgBox "Error showing dashboard: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub

' Save transaction changes from the transaction list
Public Sub SaveTransactionChanges()
    On Error GoTo ErrorHandler
    
    ' Get the transactions sheet
    Dim transSheet As Worksheet
    Set transSheet = Utilities.GetOrCreateSheet("Transactions", True)
    
    ' Get repository
    Dim repo As TransactionRepository
    Set repo = TransactionRepository
    repo.Initialize
    
    ' Clear existing transactions
    repo.ClearAll
    
    ' Read transactions from sheet
    Dim lastRow As Long
    lastRow = Utilities.GetLastRow(transSheet, 1)
    
    If lastRow <= 1 Then
        ' No transactions
        MsgBox "No transactions to save.", vbInformation, AppSettings.AppTitle
        Exit Sub
    End If
    
    ' Process each row
    Dim row As Long
    Dim trans As Transaction
    Dim changedCount As Long
    
    changedCount = 0
    
    For row = 2 To lastRow
        ' Skip empty rows
        If Len(Trim(transSheet.Cells(row, 1).Value)) = 0 Then
            GoTo ContinueForLoop
        End If
        
        ' Create new transaction
        Set trans = New Transaction
        
        ' Set ID (or generate new one if empty)
        If Len(Trim(transSheet.Cells(row, 9).Value)) > 0 Then
            trans.ID = Trim(transSheet.Cells(row, 9).Value)
        End If
        
        ' Set properties
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
        
        ' Add to repository
        repo.AddTransaction trans
        
        changedCount = changedCount + 1
        
ContinueForLoop:
    Next row
    
    ' Save changes
    repo.SaveChanges
    
    ' Mark dashboard as needing refresh
    DashboardManager.NeedsRefresh = True
    
    ' Show result
    MsgBox "Changes saved!" & vbCrLf & _
          "Transactions updated: " & changedCount, _
          vbInformation, AppSettings.AppTitle
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.SaveTransactionChanges", Err.Number, Err.Description
    MsgBox "Error saving changes: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub

' Delete selected transaction
Public Sub DeleteSelectedTransaction()
    On Error GoTo ErrorHandler
    
    ' Get the transactions sheet
    Dim transSheet As Worksheet
    Set transSheet = Utilities.GetOrCreateSheet("Transactions", True)
    
    ' Get selected row
    Dim selectedRow As Long
    selectedRow = Selection.Row
    
    ' Check if selection is valid
    If selectedRow < 2 Or selectedRow > Utilities.GetLastRow(transSheet, 1) Then
        MsgBox "Please select a transaction row to delete.", vbExclamation, AppSettings.AppTitle
        Exit Sub
    End If
    
    ' Get transaction ID
    Dim transactionID As String
    transactionID = Trim(transSheet.Cells(selectedRow, 9).Value)
    
    If Len(transactionID) = 0 Then
        MsgBox "Cannot delete this row. No valid transaction ID found.", vbExclamation, AppSettings.AppTitle
        Exit Sub
    End If
    
    ' Confirm deletion
    Dim result As VbMsgBoxResult
    result = MsgBox("Are you sure you want to delete this transaction?" & vbCrLf & _
                  "Merchant: " & transSheet.Cells(selectedRow, 2).Value & vbCrLf & _
                  "Amount: " & transSheet.Cells(selectedRow, 4).Value, _
                  vbYesNo + vbQuestion, AppSettings.AppTitle)
    
    If result <> vbYes Then
        Exit Sub
    End If
    
    ' Get repository
    Dim repo As TransactionRepository
    Set repo = TransactionRepository
    repo.Initialize
    
    ' Remove transaction
    repo.RemoveTransaction transactionID
    
    ' Save changes
    repo.SaveChanges
    
    ' Remove row from sheet
    transSheet.Rows(selectedRow).Delete
    
    ' Mark dashboard as needing refresh
    DashboardManager.NeedsRefresh = True
    
    ' Show result
    MsgBox "Transaction deleted successfully!", vbInformation, AppSettings.AppTitle
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.DeleteSelectedTransaction", Err.Number, Err.Description
    MsgBox "Error deleting transaction: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub

' Create sample data for testing
Public Sub CreateSampleData()
    On Error GoTo ErrorHandler
    
    ' Confirm creation
    Dim result As VbMsgBoxResult
    result = MsgBox("This will create sample transaction data for testing." & vbCrLf & _
                  "Existing data will not be affected." & vbCrLf & vbCrLf & _
                  "Do you want to continue?", _
                  vbYesNo + vbQuestion, AppSettings.AppTitle)
    
    If result <> vbYes Then
        Exit Sub
    End If
    
    ' Get repository
    Dim repo As TransactionRepository
    Set repo = TransactionRepository
    repo.Initialize
    
    ' Create sample transactions
    Dim trans As Transaction
    Dim i As Long
    Dim currentDate As Date
    Dim categories As Variant
    Dim merchants As Variant
    Dim accounts As Variant
    
    ' Define sample categories
    categories = Array("Groceries", "Dining", "Utilities", "Rent", "Transportation", _
                      "Entertainment", "Shopping", "Health", "Travel", "Miscellaneous")
    
    ' Define sample merchants
    merchants = Array("Walmart", "Target", "Kroger", "Amazon", "Netflix", _
                     "Gas Station", "Electric Company", "Restaurant", "Pharmacy", "Department Store")
    
    ' Define sample accounts
    accounts = Array("Checking", "Credit Card", "Cash")
    
    ' Start from 3 months ago
    currentDate = DateAdd("m", -3, Date)
    
    ' Create 50 sample transactions
    For i = 1 To 50
        ' Create new transaction
        Set trans = New Transaction
        
        ' Set date (random day in the last 3 months)
        trans.TransactionDate = DateAdd("d", Int(Rnd * 90), currentDate)
        
        ' Set merchant
        trans.Merchant = merchants(Int(Rnd * UBound(merchants)))
        
        ' Set category
        trans.Category = categories(Int(Rnd * UBound(categories)))
        
        ' Set amount (negative for expense, positive for income)
        If Rnd < 0.9 Then
            ' 90% chance of expense
            trans.Amount = -Round((Rnd * 100) + 5, 2) ' $5 to $105
        Else
            ' 10% chance of income
            trans.Amount = Round((Rnd * 500) + 100, 2) ' $100 to $600
        End If
        
        ' Set account
        trans.Account = accounts(Int(Rnd * UBound(accounts)))
        
        ' Set owner (random between User1 and User2)
        If Rnd < 0.5 Then
            trans.Owner = AppSettings.User1Name
        Else
            trans.Owner = AppSettings.User2Name
        End If
        
        ' Set as shared by default
        trans.IsShared = True
        
        ' Set source file
        trans.SourceFile = "SampleData"
        
        ' Add to repository
        repo.AddTransaction trans
    Next i
    
    ' Save changes
    repo.SaveChanges
    
    ' Mark dashboard as needing refresh
    DashboardManager.NeedsRefresh = True
    DashboardManager.RefreshDashboard
    
    ' Show result
    MsgBox "Sample data created!" & vbCrLf & _
          "50 transactions added with random dates, amounts, and categories.", _
          vbInformation, AppSettings.AppTitle
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.CreateSampleData", Err.Number, Err.Description
    MsgBox "Error creating sample data: " & Err.Description, vbExclamation, AppSettings.AppTitle
End Sub

' Show open file dialog
Private Function ShowOpenFileDialog(filter As String, title As String) As String
    On Error GoTo ErrorHandler
    
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        .Title = title
        .Filters.Clear
        .Filters.Add filter
        
        If .Show = True Then
            ShowOpenFileDialog = .SelectedItems(1)
        Else
            ShowOpenFileDialog = ""
        End If
    End With
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "BALANCE.ShowOpenFileDialog", Err.Number, Err.Description
    ShowOpenFileDialog = ""
End Function
