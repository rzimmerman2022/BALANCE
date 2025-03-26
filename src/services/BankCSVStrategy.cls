' File: src/services/BankCSVStrategy.cls
'---------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BankCSVStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' BankCSVStrategy Class - Implements ICSVImportStrategy
'
' Purpose: Strategy for importing CSV files from a typical bank account.
' This strategy handles CSV files with a specific format used by many banks,
' including date, description, amount, and other typical banking fields.

' Interface implementation
Implements ICSVImportStrategy

' CSV format details
Private m_ExpectedHeaders As Variant
Private m_ColumnMapping As Object ' Dictionary
Private m_Logger As IErrorLogger

' Bank-specific parsing rules
Private Const DATE_FORMAT As String = "mm/dd/yyyy"
Private Const DEBIT_INDICATOR As String = "-"
Private Const FILE_SIGNATURE As String = "Bank Account Transactions"

'=========================================================================
' Initialization
'=========================================================================

Private Sub Class_Initialize()
    ' Set up expected headers for this format
    m_ExpectedHeaders = Array("Date", "Description", "Amount", "Running Bal")
    
    ' Set up column mapping (CSV column -> Transaction property)
    Set m_ColumnMapping = CreateObject("Scripting.Dictionary")
    m_ColumnMapping.Add "Date", "TransactionDate"
    m_ColumnMapping.Add "Description", "Merchant"
    m_ColumnMapping.Add "Amount", "Amount"
    
    ' Additional mapping information for special handling
    m_ColumnMapping.Add "DateFormat", DATE_FORMAT
    m_ColumnMapping.Add "DebitIndicator", DEBIT_INDICATOR
End Sub

'=========================================================================
' Interface Implementation
'=========================================================================

' Check if this strategy can process the given file
Private Function ICSVImportStrategy_CanProcessFile(ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Read the first few lines of the file
    Dim fileNum As Integer
    Dim headerLine As String, firstDataLine As String
    Dim firstFewLines As String
    Dim lineCount As Integer
    
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    ' Read up to 10 lines
    lineCount = 0
    firstFewLines = ""
    
    Do Until EOF(fileNum) Or lineCount >= 10
        Line Input #fileNum, headerLine
        firstFewLines = firstFewLines & headerLine & vbCrLf
        lineCount = lineCount + 1
    Loop
    
    Close #fileNum
    
    ' Check for bank signature or known headers
    If InStr(1, firstFewLines, FILE_SIGNATURE, vbTextCompare) > 0 Then
        ICSVImportStrategy_CanProcessFile = True
        Exit Function
    End If
    
    ' Check for expected headers
    Dim header As Variant
    Dim headerCount As Integer
    
    headerCount = 0
    For Each header In m_ExpectedHeaders
        If InStr(1, firstFewLines, header, vbTextCompare) > 0 Then
            headerCount = headerCount + 1
        End If
    Next header
    
    ' If most expected headers are found, assume this format
    ICSVImportStrategy_CanProcessFile = (headerCount >= UBound(m_ExpectedHeaders) * 0.75)
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BankCSVStrategy.CanProcessFile", Err.Number, Err.Description
    End If
    ICSVImportStrategy_CanProcessFile = False
End Function

' Get the name of this strategy
Private Property Get ICSVImportStrategy_StrategyName() As String
    ICSVImportStrategy_StrategyName = "StandardBankFormat"
End Property

' Get the description of this strategy
Private Property Get ICSVImportStrategy_StrategyDescription() As String
    ICSVImportStrategy_StrategyDescription = "Standard bank account CSV format with date, description, and amount columns"
End Property

' Import transactions from the CSV file
Private Function ICSVImportStrategy_ImportTransactions(ByVal filePath As String, ByVal owner As String, _
                                                    Optional ByVal logger As IErrorLogger = Nothing) As Collection
    On Error GoTo ErrorHandler
    
    ' Set logger if provided
    If Not logger Is Nothing Then
        Set m_Logger = logger
    End If
    
    ' Create collection for transactions
    Dim transactions As New Collection
    
    ' Log start of import
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "BankCSVStrategy.ImportTransactions", _
            "Starting import from " & filePath
    End If
    
    ' Read the CSV file
    Dim fileContent As String
    fileContent = ReadEntireFile(filePath)
    
    ' Parse CSV using array-based approach for performance
    Dim csvData As Variant
    csvData = ParseCSV(fileContent)
    
    ' Check if parsing succeeded
    If Not IsArray(csvData) Then
        If Not m_Logger Is Nothing Then
            m_Logger.LogError "BankCSVStrategy.ImportTransactions", 0, _
                "Failed to parse CSV file: " & filePath
        End If
        Set ICSVImportStrategy_ImportTransactions = transactions
        Exit Function
    End If
    
    ' Find header row and create column index mapping
    Dim headerRowIndex As Long
    Dim columnIndices As Object ' Dictionary
    Set columnIndices = CreateObject("Scripting.Dictionary")
    
    headerRowIndex = FindHeaderRow(csvData, columnIndices)
    
    If headerRowIndex < 0 Then
        If Not m_Logger Is Nothing Then
            m_Logger.LogError "BankCSVStrategy.ImportTransactions", 0, _
                "Could not find header row in CSV file: " & filePath
        End If
        Set ICSVImportStrategy_ImportTransactions = transactions
        Exit Function
    End If
    
    ' Process data rows
    Dim i As Long
    Dim rowCount As Long
    rowCount = UBound(csvData, 1)
    
    ' Process each row after the header
    For i = headerRowIndex + 1 To rowCount
        ' Create a transaction from the row
        Dim trans As Transaction
        Set trans = CreateTransactionFromRow(csvData, i, columnIndices, owner, filePath)
        
        ' Add to collection if valid
        If Not trans Is Nothing Then
            transactions.Add trans
        End If
    Next i
    
    ' Log completion
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "BankCSVStrategy.ImportTransactions", _
            "Imported " & transactions.Count & " transactions from " & filePath
    End If
    
    Set ICSVImportStrategy_ImportTransactions = transactions
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BankCSVStrategy.ImportTransactions", Err.Number, Err.Description
    End If
    Set ICSVImportStrategy_ImportTransactions = New Collection
End Function

' Get the column mapping for this strategy
Private Function ICSVImportStrategy_GetColumnMapping() As Object
    Set ICSVImportStrategy_GetColumnMapping = m_ColumnMapping
End Function

' Get expected headers for this strategy
Private Function ICSVImportStrategy_GetExpectedHeaders() As Variant
    ICSVImportStrategy_GetExpectedHeaders = m_ExpectedHeaders
End Function

'=========================================================================
' Helper Methods
'=========================================================================

' Read the entire file content
Private Function ReadEntireFile(ByVal filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim fileContent As String
    
    fileNum = FreeFile
    Open filePath For Binary As #fileNum
    fileContent = Space$(LOF(fileNum))
    Get #fileNum, , fileContent
    Close #fileNum
    
    ReadEntireFile = fileContent
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BankCSVStrategy.ReadEntireFile", Err.Number, Err.Description
    End If
    ReadEntireFile = ""
End Function

' Parse CSV content into a 2D array
Private Function ParseCSV(ByVal csvContent As String) As Variant
    On Error GoTo ErrorHandler
    
    ' Standardize line endings
    csvContent = Replace(csvContent, vbCrLf, vbLf)
    csvContent = Replace(csvContent, vbCr, vbLf)
    
    ' Split into lines
    Dim lines As Variant
    lines = Split(csvContent, vbLf)
    
    ' Count non-empty lines
    Dim lineCount As Long
    Dim i As Long
    
    lineCount = 0
    For i = LBound(lines) To UBound(lines)
        If Trim$(lines(i)) <> "" Then
            lineCount = lineCount + 1
        End If
    Next i
    
    ' If no lines, return empty array
    If lineCount = 0 Then
        ParseCSV = Array()
        Exit Function
    End If
    
    ' Create 2D array for CSV data
    Dim maxColumns As Long
    maxColumns = 1 ' Start with at least 1 column
    
    ' Find the maximum number of columns
    For i = LBound(lines) To UBound(lines)
        If Trim$(lines(i)) <> "" Then
            Dim columnCount As Long
            columnCount = CountCSVColumns(lines(i))
            If columnCount > maxColumns Then
                maxColumns = columnCount
            End If
        End If
    Next i
    
    ' Create the result array
    Dim result() As Variant
    ReDim result(1 To lineCount, 1 To maxColumns)
    
    ' Fill the array with CSV data
    Dim currentRow As Long
    currentRow = 1
    
    For i = LBound(lines) To UBound(lines)
        If Trim$(lines(i)) <> "" Then
            Dim columns As Variant
            columns = ParseCSVLine(lines(i))
            
            Dim j As Long
            For j = LBound(columns) To UBound(columns)
                If j <= maxColumns Then
                    result(currentRow, j) = columns(j)
                End If
            Next j
            
            currentRow = currentRow + 1
        End If
    Next i
    
    ParseCSV = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BankCSVStrategy.ParseCSV", Err.Number, Err.Description
    End If
    ParseCSV = Array()
End Function

' Count the number of columns in a CSV line
Private Function CountCSVColumns(ByVal line As String) As Long
    On Error GoTo ErrorHandler
    
    ' Handle quoted fields correctly
    Dim columns As Variant
    columns = ParseCSVLine(line)
    
    CountCSVColumns = UBound(columns) - LBound(columns) + 1
    
    Exit Function
    
ErrorHandler:
    CountCSVColumns = 0
End Function

' Parse a CSV line into an array of fields
Private Function ParseCSVLine(ByVal line As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim result As Variant
    Dim fields As New Collection
    Dim fieldValue As String
    Dim inQuotes As Boolean
    Dim i As Long, currentChar As String
    
    fieldValue = ""
    inQuotes = False
    
    For i = 1 To Len(line)
        currentChar = Mid$(line, i, 1)
        
        Select Case currentChar
            Case """"
                ' Handle quotes (double quote is an escaped quote in CSV)
                If i < Len(line) AndAlso Mid$(line, i + 1, 1) = """" Then
                    fieldValue = fieldValue & """"
                    i = i + 1 ' Skip the next quote
                Else
                    inQuotes = Not inQuotes
                End If
                
            Case ","
                If inQuotes Then
                    ' Comma inside quotes is part of the field
                    fieldValue = fieldValue & currentChar
                Else
                    ' End of field
                    fields.Add fieldValue
                    fieldValue = ""
                End If
                
            Case Else
                ' Any other character
                fieldValue = fieldValue & currentChar
        End Select
    Next i
    
    ' Add the last field
    fields.Add fieldValue
    
    ' Convert collection to array
    ReDim result(1 To fields.Count)
    For i = 1 To fields.Count
        result(i) = fields(i)
    Next i
    
    ParseCSVLine = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BankCSVStrategy.ParseCSVLine", Err.Number, Err.Description
    End If
    
    ' Return a single-element array with the original line on error
    ParseCSVLine = Array(line)
End Function

' Find the header row and map column indices
Private Function FindHeaderRow(ByVal csvData As Variant, ByRef columnIndices As Object) As Long
    On Error GoTo ErrorHandler
    
    Dim rowCount As Long, colCount As Long
    rowCount = UBound(csvData, 1)
    colCount = UBound(csvData, 2)
    
    ' Check each row for matching headers
    Dim i As Long, j As Long
    Dim bestRowIndex As Long
    Dim bestMatchCount As Long
    
    bestRowIndex = -1
    bestMatchCount = 0
    
    For i = 1 To rowCount
        Dim matchCount As Long
        matchCount = 0
        
        ' Clear previous indices
        columnIndices.RemoveAll
        
        ' Check each column in this row against expected headers
        For j = 1 To colCount
            Dim headerValue As String
            headerValue = Trim$(csvData(i, j))
            
            Dim expectedHeader As Variant
            For Each expectedHeader In m_ExpectedHeaders
                If StrComp(headerValue, expectedHeader, vbTextCompare) = 0 Then
                    matchCount = matchCount + 1
                    columnIndices.Add expectedHeader, j
                    Exit For
                End If
            Next expectedHeader
        Next j
        
        ' If this row has more matches than the best so far, update
        If matchCount > bestMatchCount Then
            bestMatchCount = matchCount
            bestRowIndex = i
        End If
        
        ' If we found a perfect match, stop searching
        If matchCount = UBound(m_ExpectedHeaders) - LBound(m_ExpectedHeaders) + 1 Then
            Exit For
        End If
    Next i
    
    ' If we found a good match, return the row index
    If bestMatchCount >= (UBound(m_ExpectedHeaders) - LBound(m_ExpectedHeaders) + 1) * 0.75 Then
        FindHeaderRow = bestRowIndex
    Else
        FindHeaderRow = -1
    End If
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BankCSVStrategy.FindHeaderRow", Err.Number, Err.Description
    End If
    FindHeaderRow = -1
End Function

' Create a Transaction object from a CSV row
Private Function CreateTransactionFromRow(ByVal csvData As Variant, ByVal rowIndex As Long, _
                                         ByVal columnIndices As Object, _
                                         ByVal owner As String, ByVal filePath As String) As Transaction
    On Error GoTo ErrorHandler
    
    ' Create a new transaction
    Dim trans As New Transaction
    
    ' Set transaction properties from CSV data
    
    ' Date
    If columnIndices.Exists("Date") Then
        Dim dateStr As String
        dateStr = csvData(rowIndex, columnIndices("Date"))
        
        ' Parse date according to expected format
        trans.TransactionDate = ParseDate(dateStr, m_ColumnMapping("DateFormat"))
    End If
    
    ' Description/Merchant
    If columnIndices.Exists("Description") Then
        trans.Merchant = csvData(rowIndex, columnIndices("Description"))
    End If
    
    ' Amount
    If columnIndices.Exists("Amount") Then
        Dim amountStr As String
        amountStr = csvData(rowIndex, columnIndices("Amount"))
        
        ' Parse amount, handling debits correctly
        trans.Amount = ParseAmount(amountStr, m_ColumnMapping("DebitIndicator"))
    End If
    
    ' Set the owner
    trans.Owner = owner
    
    ' Set default category - to be refined later
    trans.Category = "Uncategorized"
    
    ' Set source file
    trans.SourceFile = filePath
    
    ' Assume not shared by default
    trans.IsShared = False
    
    ' Return the transaction
    Set CreateTransactionFromRow = trans
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BankCSVStrategy.CreateTransactionFromRow", Err.Number, Err.Description
    End If
    Set CreateTransactionFromRow = Nothing
End Function

' Parse a date string according to the specified format
Private Function ParseDate(ByVal dateStr As String, ByVal dateFormat As String) As Date
    On Error GoTo ErrorHandler
    
    ' Handle different date formats
    Select Case dateFormat
        Case "mm/dd/yyyy"
            ' US format: MM/DD/YYYY
            Dim parts As Variant
            parts = Split(dateStr, "/")
            
            If UBound(parts) = 2 Then
                ParseDate = DateSerial(CInt(parts(2)), CInt(parts(0)), CInt(parts(1)))
            Else
                ParseDate = CDate(dateStr) ' Try default conversion
            End If
            
        Case "dd/mm/yyyy"
            ' European format: DD/MM/YYYY
            parts = Split(dateStr, "/")
            
            If UBound(parts) = 2 Then
                ParseDate = DateSerial(CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
            Else
                ParseDate = CDate(dateStr) ' Try default conversion
            End If
            
        Case Else
            ' Try default conversion
            ParseDate = CDate(dateStr)
    End Select
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BankCSVStrategy.ParseDate", Err.Number, _
            "Could not parse date: " & dateStr
    End If
    
    ' Return current date on error
    ParseDate = Date
End Function

' Parse an amount string, handling different formats
Private Function ParseAmount(ByVal amountStr As String, ByVal debitIndicator As String) As Double
    On Error GoTo ErrorHandler
    
    ' Remove currency symbols, commas, etc.
    amountStr = Replace(amountStr, "$", "")
    amountStr = Replace(amountStr, ",", "")
    amountStr = Replace(amountStr, " ", "")
    
    ' Check if this is a debit (expense)
    Dim isDebit As Boolean
    isDebit = (InStr(1, amountStr, debitIndicator) > 0)
    
    ' Remove any parentheses or negative signs
    amountStr = Replace(amountStr, "(", "")
    amountStr = Replace(amountStr, ")", "")
    amountStr = Replace(amountStr, "-", "")
    
    ' Convert to number
    Dim amount As Double
    amount = CDbl(amountStr)
    
    ' Apply sign based on debit/credit
    If isDebit Then
        amount = -amount
    End If
    
    ParseAmount = amount
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "BankCSVStrategy.ParseAmount", Err.Number, _
            "Could not parse amount: " & amountStr
    End If
    
    ' Return 0 on error
    ParseAmount = 0
End Function

'=========================================================================
' Public Methods
'=========================================================================

' Public method to check if this strategy can process a file
Public Function CanProcessFile(ByVal filePath As String) As Boolean
    CanProcessFile = ICSVImportStrategy_CanProcessFile(filePath)
End Function

' Public getter for strategy name
Public Property Get StrategyName() As String
    StrategyName = ICSVImportStrategy_StrategyName
End Property

' Public getter for strategy description
Public Property Get StrategyDescription() As String
    StrategyDescription = ICSVImportStrategy_StrategyDescription
End Property

' Public method to import transactions
Public Function ImportTransactions(ByVal filePath As String, ByVal owner As String, _
                                 Optional ByVal logger As IErrorLogger = Nothing) As Collection
    Set ImportTransactions = ICSVImportStrategy_ImportTransactions(filePath, owner, logger)
End Function