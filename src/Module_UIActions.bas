' File: src/Module_UIActions.bas
'---------------------------------------
Attribute VB_Name = "Module_UIActions"
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' UI Actions Module
'
' Purpose: Contains public functions that can be assigned to UI elements like
' buttons, handling user interactions with a link to the event-driven architecture.

' Public function for refreshing the dashboard
Public Sub RefreshDashboardAction()
    On Error Resume Next
    
    ' Call the dashboard manager to refresh
    DashboardManager.RefreshDashboard
    
    ' If error, show a message
    If Err.Number <> 0 Then
        MsgBox "Error refreshing dashboard: " & Err.Description, vbExclamation, "Error"
    End If
End Sub

' Public function for applying date filter
Public Sub ApplyDateFilterAction()
    On Error Resume Next
    
    ' Get the dashboard sheet
    Dim dashboardSheet As Worksheet
    Set dashboardSheet = ThisWorkbook.Worksheets("Dashboard")
    
    ' Get dates from UI
    Dim startDate As Date
    Dim endDate As Date
    
    startDate = dashboardSheet.Range("H2").Value
    endDate = dashboardSheet.Range("H3").Value
    
    ' Apply the filter
    DashboardManager.ApplyDateFilter startDate, endDate
    
    ' If error, show a message
    If Err.Number <> 0 Then
        MsgBox "Error applying filter: " & Err.Description, vbExclamation, "Error"
    End If
End Sub

' Public function for importing transactions
Public Sub ImportTransactionsAction()
    On Error Resume Next
    
    ' Show file picker
    Dim filePath As Variant
    filePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select CSV File")
    
    ' If user canceled, exit
    If filePath = False Then Exit Sub
    
    ' Get owner for transactions
    Dim owner As String
    owner = InputBox("Enter the transaction owner:", "Import Transactions")
    
    ' If user canceled, exit
    If owner = "" Then Exit Sub
    
    ' Import the transactions
    CSVImportEngine.ImportTransactionsFromFile filePath, owner
    
    ' If error, show a message
    If Err.Number <> 0 Then
        MsgBox "Error importing transactions: " & Err.Description, vbExclamation, "Error"
    Else
        MsgBox "Transactions imported successfully.", vbInformation, "Import Complete"
    End If
End Sub

' Public function for exporting data
Public Sub ExportDataAction()
    On Error Resume Next
    
    ' Show file picker for save location
    Dim filePath As Variant
    filePath = Application.GetSaveAsFilename("BALANCE_Export.xlsx", "Excel Files (*.xlsx), *.xlsx", , "Save Export As")
    
    ' If user canceled, exit
    If filePath = False Then Exit Sub
    
    ' Create new workbook for export
    Dim exportWb As Workbook
    Set exportWb = Workbooks.Add
    
    ' Get transactions
    Dim transactions As Collection
    Set transactions = TransactionRepository.GetTransactions()
    
    ' Create sheet for transactions
    Dim transSheet As Worksheet
    Set transSheet = exportWb.Sheets(1)
    transSheet.Name = "Transactions"
    
    ' Add headers
    transSheet.Range("A1").Value = "Date"
    transSheet.Range("B1").Value = "Merchant"
    transSheet.Range("C1").Value = "Category"
    transSheet.Range("D1").Value = "Amount"
    transSheet.Range("E1").Value = "Account"
    transSheet.Range("F1").Value = "Owner"
    transSheet.Range("G1").Value = "Shared"
    transSheet.Range("H1").Value = "Notes"
    
    ' Format headers
    transSheet.Range("A1:H1").Font.Bold = True
    
    ' Add data
    Dim row As Long
    row = 2
    
    Dim trans As Transaction
    For Each trans In transactions
        transSheet.Cells(row, 1).Value = trans.TransactionDate
        transSheet.Cells(row, 2).Value = trans.Merchant
        transSheet.Cells(row, 3).Value = trans.Category
        transSheet.Cells(row, 4).Value = trans.Amount
        transSheet.Cells(row, 5).Value = trans.Account
        transSheet.Cells(row, 6).Value = trans.Owner
        transSheet.Cells(row, 7).Value = trans.IsShared
        transSheet.Cells(row, 8).Value = trans.Notes
        
        row = row + 1
    Next trans
    
    ' Format date column
    transSheet.Columns(1).NumberFormat = "yyyy-mm-dd"
    
    ' Format amount column
    transSheet.Columns(4).NumberFormat = "$#,##0.00;($#,##0.00)"
    
    ' Autofit columns
    transSheet.Columns("A:H").AutoFit
    
    ' Save the workbook
    exportWb.SaveAs filePath
    exportWb.Close
    
    ' Trigger export completed event
    EventManager.PublishEvent EventType.ExportCompleted
    
    ' If error, show a message
    If Err.Number <> 0 Then
        MsgBox "Error exporting data: " & Err.Description, vbExclamation, "Error"
    Else
        MsgBox "Data exported successfully.", vbInformation, "Export Complete"
    End If
End Sub