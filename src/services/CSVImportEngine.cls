VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVImportEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' CSVImportEngine Class - Handles importing transactions from CSV files

' Column mapping structure
Public Type ColumnMapping
    DateColumn As Long
    MerchantColumn As Long
    CategoryColumn As Long
    AmountColumn As Long
    AccountColumn As Long
    NotesColumn As Long
    OwnerColumn As Long
    DateFormat As String
    AmountMultiplier As Integer ' 1 for positive values, -1 for negative
End Type

' Import result structure
Public Type ImportResult
    Success As Boolean
    TransactionsAdded As Long
    DuplicatesSkipped As Long
    ErrorsEncountered As Long
    ImportedFiles As Collection ' Collection of filenames
    ElapsedSeconds As Double
    ErrorMessages As Collection ' Collection of error messages
End Type

' Known CSV formats
Private Const FORMAT_AUTO As String = "Auto-Detect"
Private Const FORMAT_EXCEL As String = "Excel Export"
Private Const FORMAT_BANK As String = "Bank Statement"
Private Const FORMAT_CREDIT_CARD As String = "Credit Card"
Private Const FORMAT_CUSTOM As String = "Custom"

' Private variables
Private m_IsInitialized As Boolean
Private m_DefaultOwner As String
Private m_LastMappingUsed As String

' ===== Properties =====

' Default owner for imported transactions
Public Property Get DefaultOwner() As String
    If Not m_IsInitialized Then Initialize
    DefaultOwner = m_DefaultOwner
End Property

Public Property Let DefaultOwner(value As String)
    m_DefaultOwner = value
End Property

' Last mapping used
Public Property Get LastMappingUsed() As String
    If Not m_IsInitialized Then Initialize
    LastMappingUsed = m_LastMappingUsed
End Property

' ===== Methods =====

' Initialize the engine
Public Sub Initialize()
    On Error Resume Next
    
    ' Set default values
    m_DefaultOwner = AppSettings.User1Name
    m_LastMappingUsed = FORMAT_AUTO
    
    m_IsInitialized = True
    
    On Error GoTo 0
End Sub

' Import transactions from a CSV file
Public Function ImportCSVFile(filePath As String, Optional owner As String = "", _
                             Optional formatType As String = FORMAT_AUTO, _
                             Optional markAsShared As Boolean = True) As ImportResult
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    Dim result As ImportResult
    result.Success = False
    result.TransactionsAdded = 0
    result.DuplicatesSkipped = 0
    result.ErrorsEncountered = 0
    Set result.ImportedFiles = New Collection
    Set result.ErrorMessages = New Collection
    
    Dim startTime As Double
    startTime = Timer
    
    ' Validate parameters
    If Len(filePath) = 0 Then
        result.ErrorMessages.Add "No file path provided"
        result.ErrorsEncountered = result.ErrorsEncountered + 1
        ImportCSVFile = result
        Exit Function
    End If
    
    ' Check if file exists
    If Not FileExists(filePath) Then
        result.ErrorMessages.Add "File not found: " & filePath
        result.ErrorsEncountered = result.ErrorsEncountered + 1
        ImportCSVFile = result
        Exit Function
    End If
    
    ' Set owner
    If Len(owner) = 0 Then
        owner = m_DefaultOwner
    End If
    
    ' Get file content
    Dim fileContent As String
    fileContent = ReadTextFile(filePath)
    
    If Len(fileContent) = 0 Then
        result.ErrorMessages.Add "Empty or unreadable file: " & filePath
        result.ErrorsEncountered = result.ErrorsEncountered + 1
        ImportCSVFile = result
        Exit Function
    End If
    
    ' Determine the format and get column mapping
    Dim mapping As ColumnMapping
    If formatType = FORMAT_AUTO Then
        ' Auto-detect format
        m_LastMappingUsed = DetectCSVFormat(fileContent)
        mapping = GetColumnMapping(m_LastMappingUsed)
    Else
        ' Use specified format
        m_LastMappingUsed = formatType
        mapping = GetColumnMapping(formatType)
    End If
    
    ' Parse CSV data
    Dim transactions As Collection
    Set transactions = ParseCSV(fileContent, mapping, owner, markAsShared, filePath, result)
    
    ' Add to repository if we have transactions
    If transactions.Count > 0 Then
        Dim repo As TransactionRepository
        Set repo = TransactionRepository
        repo.Initialize
        
        ' Add transactions and save
        repo.AddTransactions transactions
        
        ' Update result
        result.TransactionsAdded = transactions.Count
        result.ImportedFiles.Add filePath
        result.Success = True
    End If
    
    ' Calculate elapsed time
    result.ElapsedSeconds = Timer - startTime
    
    ' Return result
    ImportCSVFile = result
    
    ' Clean up
    Set transactions = Nothing
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "CSVImportEngine.ImportCSVFile", Err.Number, Err.Description
    
    ' Update result with error
    result.ErrorsEncountered = result.ErrorsEncountered + 1
    result.ErrorMessages.Add "Error importing file: " & Err.Description
    result.Success = False
    result.ElapsedSeconds = Timer - startTime
    
    ImportCSVFile = result
End Function

' Import all CSV files from a folder
Public Function ImportCSVFolder(folderPath As String, Optional owner As String = "", _
                              Optional formatType As String = FORMAT_AUTO, _
                              Optional markAsShared As Boolean = True) As ImportResult
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    Dim result As ImportResult
    result.Success = False
    result.TransactionsAdded = 0
    result.DuplicatesSkipped = 0
    result.ErrorsEncountered = 0
    Set result.ImportedFiles = New Collection
    Set result.ErrorMessages = New Collection
    
    Dim startTime As Double
    startTime = Timer
    
    ' Validate parameters
    If Len(folderPath) = 0 Then
        result.ErrorMessages.Add "No folder path provided"
        result.ErrorsEncountered = result.ErrorsEncountered + 1
        ImportCSVFolder = result
        Exit Function
    End If
    
    ' Check if folder exists
    If Not Utilities.FolderExists(folderPath) Then
        result.ErrorMessages.Add "Folder not found: " & folderPath
        result.ErrorsEncountered = result.ErrorsEncountered + 1
        ImportCSVFolder = result
        Exit Function
    End If
    
    ' Set owner
    If Len(owner) = 0 Then
        owner = m_DefaultOwner
    End If
    
    ' Get list of CSV files in the folder
    Dim csvFiles As Collection
    Set csvFiles = GetCSVFilesInFolder(folderPath)
    
    If csvFiles.Count = 0 Then
        result.ErrorMessages.Add "No CSV files found in folder: " & folderPath
        result.ErrorsEncountered = result.ErrorsEncountered + 1
        ImportCSVFolder = result
        Exit Function
    End If
    
    ' Import each file
    Dim filePath As Variant
    Dim fileResult As ImportResult
    
    For Each filePath In csvFiles
        ' Import the file
        fileResult = ImportCSVFile(filePath, owner, formatType, markAsShared)
        
        ' Update result
        result.TransactionsAdded = result.TransactionsAdded + fileResult.TransactionsAdded
        result.DuplicatesSkipped = result.DuplicatesSkipped + fileResult.DuplicatesSkipped
        result.ErrorsEncountered = result.ErrorsEncountered + fileResult.ErrorsEncountered
        
        ' Add imported file if successful
        If fileResult.Success Then
            result.ImportedFiles.Add filePath
        End If
        
        ' Add any error messages
        Dim errorMsg As Variant
        For Each errorMsg In fileResult.ErrorMessages
            result.ErrorMessages.Add errorMsg
        Next errorMsg
    Next filePath
    
    ' Set success if we imported at least one file
    result.Success = (result.ImportedFiles.Count > 0)
    
    ' Calculate elapsed time
    result.ElapsedSeconds = Timer - startTime
    
    ' Return result
    ImportCSVFolder = result
    
    ' Clean up
    Set csvFiles = Nothing
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "CSVImportEngine.ImportCSVFolder", Err.Number, Err.Description
    
    ' Update result with error
    result.ErrorsEncountered = result.ErrorsEncountered + 1
    result.ErrorMessages.Add "Error importing folder: " & Err.Description
    result.Success = False
    result.ElapsedSeconds = Timer - startTime
    
    ImportCSVFolder = result
End Function

' Get column mapping for a specific format
Private Function GetColumnMapping(formatType As String) As ColumnMapping
    On Error GoTo ErrorHandler
    
    Dim mapping As ColumnMapping
    
    Select Case formatType
        Case FORMAT_EXCEL
            ' Excel export format
            mapping.DateColumn = 1
            mapping.MerchantColumn = 2
            mapping.CategoryColumn = 3
            mapping.AmountColumn = 4
            mapping.AccountColumn = 5
            mapping.NotesColumn = 6
            mapping.OwnerColumn = 0 ' Not used
            mapping.DateFormat = "yyyy-mm-dd"
            mapping.AmountMultiplier = 1 ' Preserve sign
            
        Case FORMAT_BANK
            ' Bank statement format
            mapping.DateColumn = 1
            mapping.MerchantColumn = 3
            mapping.CategoryColumn = 0 ' Not used
            mapping.AmountColumn = 4
            mapping.AccountColumn = 2
            mapping.NotesColumn = 5
            mapping.OwnerColumn = 0 ' Not used
            mapping.DateFormat = "mm/dd/yyyy"
            mapping.AmountMultiplier = 1 ' Preserve sign
            
        Case FORMAT_CREDIT_CARD
            ' Credit card statement format
            mapping.DateColumn = 1
            mapping.MerchantColumn = 2
            mapping.CategoryColumn = 0 ' Not used
            mapping.AmountColumn = 3
            mapping.AccountColumn = 0 ' Not used
            mapping.NotesColumn = 4
            mapping.OwnerColumn = 0 ' Not used
            mapping.DateFormat = "mm/dd/yyyy"
            mapping.AmountMultiplier = -1 ' Reverse sign (positive is debit for credit cards)
            
        Case Else ' FORMAT_CUSTOM or unknown
            ' Default/custom format - will try to detect
            mapping.DateColumn = 1
            mapping.MerchantColumn = 2
            mapping.CategoryColumn = 3
            mapping.AmountColumn = 4
            mapping.AccountColumn = 5
            mapping.NotesColumn = 6
            mapping.OwnerColumn = 0 ' Not used
            mapping.DateFormat = "mm/dd/yyyy"
            mapping.AmountMultiplier = 1 ' Preserve sign
    End Select
    
    GetColumnMapping = mapping
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "CSVImportEngine.GetColumnMapping", Err.Number, Err.Description
    ' Return default mapping in case of error
    GetColumnMapping = mapping
End Function

' Detect CSV format from file content
Private Function DetectCSVFormat(fileContent As String) As String
    On Error GoTo ErrorHandler
    
    ' Get the first line (header) and rows
    Dim lines As Variant
    lines = Split(fileContent, vbCrLf)
    
    If UBound(lines) < 1 Then
        ' Not enough lines to detect format
        DetectCSVFormat = FORMAT_CUSTOM
        Exit Function
    End If
    
    Dim header As String
    header = lines(0)
    
    ' Check if header matches known formats
    
    ' Excel export format (has specific headers)
    If InStr(1, header, "Date,Merchant,Category,Amount,Account", vbTextCompare) > 0 Then
        DetectCSVFormat = FORMAT_EXCEL
        Exit Function
    End If
    
    ' Bank statement format (has transaction ID, posting date, etc.)
    If InStr(1, header, "Transaction ID,Posting Date", vbTextCompare) > 0 Or _
       InStr(1, header, "Date,Description,Amount", vbTextCompare) > 0 Then
        DetectCSVFormat = FORMAT_BANK
        Exit Function
    End If
    
    ' Credit card format (has transaction date, posting date, etc.)
    If InStr(1, header, "Transaction Date,Post Date", vbTextCompare) > 0 Or _
       InStr(1, header, "Date,Description,Debit,Credit", vbTextCompare) > 0 Then
        DetectCSVFormat = FORMAT_CREDIT_CARD
        Exit Function
    End If
    
    ' Default to custom format
    DetectCSVFormat = FORMAT_CUSTOM
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "CSVImportEngine.DetectCSVFormat", Err.Number, Err.Description
    DetectCSVFormat = FORMAT_CUSTOM
End Function

' Parse CSV data into transactions
Private Function ParseCSV(csvContent As String, mapping As ColumnMapping, _
                         owner As String, markAsShared As Boolean, _
                         sourceFile As String, ByRef result As ImportResult) As Collection
    On Error GoTo ErrorHandler
    
    Dim transactions As New Collection
    
    ' Split content into lines
    Dim lines As Variant
    lines = Split(Replace(csvContent, vbLf, vbCrLf), vbCrLf)
    
    ' Skip header row
    Dim i As Long
    For i = 1 To UBound(lines)
        ' Skip empty lines
        If Len(Trim(lines(i))) = 0 Then GoTo ContinueForLoop
        
        ' Parse CSV row
        Dim csv As Variant
        csv = ParseCSVRow(lines(i))
        
        ' Make sure we have enough columns
        Dim maxColumn As Long
        maxColumn = WorksheetFunction.Max(mapping.DateColumn, mapping.MerchantColumn, _
                                         mapping.CategoryColumn, mapping.AmountColumn, _
                                         mapping.AccountColumn, mapping.NotesColumn, _
                                         mapping.OwnerColumn)
        If UBound(csv) < maxColumn Then
            ' Not enough columns in this row - log error and skip
            result.ErrorsEncountered = result.ErrorsEncountered + 1
            result.ErrorMessages.Add "Line " & i & " has incomplete data"
            GoTo ContinueForLoop
        End If
        
        ' Create a new transaction
        Dim trans As New Transaction
        
        ' Set transaction date
        If mapping.DateColumn > 0 Then
            ' Try to parse date with the specified format
            Dim dateStr As String
            dateStr = Trim(csv(mapping.DateColumn - 1))
            
            ' Handle different date formats
            If IsDate(dateStr) Then
                trans.TransactionDate = CDate(dateStr)
            Else
                ' Log error and skip
                result.ErrorsEncountered = result.ErrorsEncountered + 1
                result.ErrorMessages.Add "Line " & i & " has invalid date: " & dateStr
                GoTo ContinueForLoop
            End If
        End If
        
        ' Set merchant
        If mapping.MerchantColumn > 0 Then
            trans.Merchant = Trim(csv(mapping.MerchantColumn - 1))
        End If
        
        ' Set category
        If mapping.CategoryColumn > 0 Then
            trans.Category = Trim(csv(mapping.CategoryColumn - 1))
        End If
        
        ' Set amount
        If mapping.AmountColumn > 0 Then
            Dim amountStr As String
            amountStr = Trim(csv(mapping.AmountColumn - 1))
            
            ' Clean amount string (remove $, commas, etc.)
            amountStr = Utilities.CleanAmountString(amountStr)
            
            ' Try to parse amount
            If IsNumeric(amountStr) Then
                trans.Amount = CDbl(amountStr) * mapping.AmountMultiplier
            Else
                ' Log error and skip
                result.ErrorsEncountered = result.ErrorsEncountered + 1
                result.ErrorMessages.Add "Line " & i & " has invalid amount: " & amountStr
                GoTo ContinueForLoop
            End If
        End If
        
        ' Set account
        If mapping.AccountColumn > 0 Then
            trans.Account = Trim(csv(mapping.AccountColumn - 1))
        End If
        
        ' Set notes
        If mapping.NotesColumn > 0 Then
            trans.Notes = Trim(csv(mapping.NotesColumn - 1))
        End If
        
        ' Set owner
        If mapping.OwnerColumn > 0 Then
            trans.Owner = Trim(csv(mapping.OwnerColumn - 1))
        Else
            trans.Owner = owner
        End If
        
        ' Set shared flag
        trans.IsShared = markAsShared
        
        ' Set source file
        trans.SourceFile = sourceFile
        
        ' Add to collection
        transactions.Add trans
        
ContinueForLoop:
    Next i
    
    Set ParseCSV = transactions
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "CSVImportEngine.ParseCSV", Err.Number, Err.Description
    
    ' Return empty collection in case of error
    Set ParseCSV = New Collection
End Function

' Parse a single CSV row into an array of values
Private Function ParseCSVRow(csvRow As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim values() As String
    Dim valueCount As Long
    Dim i As Long, j As Long
    Dim inQuotes As Boolean
    Dim currentValue As String
    
    ' Initialize result array
    ReDim values(0 To 20) ' Start with space for 20 values
    valueCount = 0
    
    ' Parse the row character by character
    inQuotes = False
    currentValue = ""
    
    For i = 1 To Len(csvRow)
        Dim c As String
        c = Mid(csvRow, i, 1)
        
        If c = """" Then
            If inQuotes And Mid(csvRow, i + 1, 1) = """" Then
                ' Double quote inside quoted value - add a single quote
                currentValue = currentValue & """"
                i = i + 1 ' Skip the next quote
            Else
                ' Toggle quote mode
                inQuotes = Not inQuotes
            End If
        ElseIf c = "," And Not inQuotes Then
            ' End of value
            If valueCount > UBound(values) Then
                ' Expand array if needed
                ReDim Preserve values(0 To UBound(values) + 10)
            End If
            
            values(valueCount) = currentValue
            valueCount = valueCount + 1
            currentValue = ""
        Else
            ' Add character to current value
            currentValue = currentValue & c
        End If
    Next i
    
    ' Add the last value
    If valueCount > UBound(values) Then
        ReDim Preserve values(0 To valueCount)
    End If
    values(valueCount) = currentValue
    valueCount = valueCount + 1
    
    ' Resize array to actual number of values
    ReDim Preserve values(0 To valueCount - 1)
    
    ParseCSVRow = values
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "CSVImportEngine.ParseCSVRow", Err.Number, Err.Description
    ' Return empty array in case of error
    Dim emptyArray(0 To 0) As String
    emptyArray(0) = ""
    ParseCSVRow = emptyArray
End Function

' Get list of CSV files in a folder
Private Function GetCSVFilesInFolder(folderPath As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim result As New Collection
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' Loop through all files in the folder
    For Each file In folder.Files
        ' Check if file is a CSV
        If LCase(fso.GetExtensionName(file.Name)) = "csv" Then
            result.Add file.Path
        End If
    Next file
    
    ' Clean up
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
    
    Set GetCSVFilesInFolder = result
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "CSVImportEngine.GetCSVFilesInFolder", Err.Number, Err.Description
    ' Return empty collection in case of error
    Set GetCSVFilesInFolder = New Collection
End Function

' Check if a file exists
Private Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
    Set fso = Nothing
    
    On Error GoTo 0
End Function

' Read text file content
Private Function ReadTextFile(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim ts As Object
    Dim content As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1) ' 1 = ForReading
    content = ts.ReadAll
    ts.Close
    
    ' Clean up
    Set ts = Nothing
    Set fso = Nothing
    
    ReadTextFile = content
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "CSVImportEngine.ReadTextFile", Err.Number, Err.Description
    ReadTextFile = ""
End Function

' Class initialize
Private Sub Class_Initialize()
    ' Nothing needed here as we're using PredeclaredId = True
    ' and explicit initialization via Initialize method
    m_IsInitialized = False
End Sub

' Class terminate
Private Sub Class_Terminate()
    ' Nothing to clean up
End Sub