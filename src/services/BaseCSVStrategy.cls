' File: src/services/BaseCSVStrategy.cls
'----------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BaseCSVStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' BaseCSVStrategy Class - Abstract base class for CSV import strategies
'
' Purpose: Provides common functionality for all CSV import strategies.
' This class extracts shared code from the various CSV strategy implementations,
' promoting code reuse and consistency across strategies.

' Interface implementation
Implements ICSVImportStrategy

' Member variables (to be used by subclasses)
Public m_ExpectedHeaders As Variant
Public m_ColumnMapping As Object ' Dictionary
Public m_Logger As IErrorLogger
Public m_StrategyName As String
Public m_StrategyDescription As String

'=========================================================================
' Initialization
'=========================================================================

Private Sub Class_Initialize()
    ' Create column mapping dictionary
    Set m_ColumnMapping = CreateObject("Scripting.Dictionary")
    
    ' Set default values
    m_StrategyName = "BaseStrategy"
    m_StrategyDescription = "Base CSV import strategy"
    
    ' Call the abstract initialization method
    InitializeStrategy
End Sub

' Abstract method to be overridden by subclasses
Public Sub InitializeStrategy()
    ' To be implemented by subclasses
    ' This method should set m_ExpectedHeaders, m_ColumnMapping, etc.
End Sub

' Initialize with logger
Public Sub Initialize(Optional ByVal logger As IErrorLogger = Nothing)
    If Not logger Is Nothing Then
        Set m_Logger = logger
    End If
End Sub

'=========================================================================
' Interface Implementation
'=========================================================================

' Check if this strategy can process the given file
Private Function ICSVImportStrategy_CanProcessFile(ByVal filePath As String) As Boolean
    ' Call the abstract method that subclasses must implement
    ICSVImportStrategy_CanProcessFile = CanProcessFileImpl(filePath)
End Function

' Abstract method to be implemented by subclasses
Public Function CanProcessFileImpl(ByVal filePath As String) As Boolean
    ' To be implemented by subclasses
    Err.Raise 5, "BaseCSVStrategy", "Abstract method must be overridden"
End Function

' Get the name of this strategy
Private Property Get ICSVImportStrategy_StrategyName() As String
    ICSVImportStrategy_StrategyName = m_StrategyName
End Property

' Get the description of this strategy
Private Property Get ICSVImportStrategy_StrategyDescription() As String
    ICSVImportStrategy_StrategyDescription = m_StrategyDescription
End Property

' Import transactions from the CSV file
Private Function ICSVImportStrategy_ImportTransactions(ByVal filePath As String, ByVal owner As String, _
                                                    Optional ByVal logger As IErrorLogger = Nothing) As Collection
    ' Set logger if provided
    If Not logger Is Nothing Then
        Set m_Logger = logger
    End If
    
    ' Create collection for transactions
    Dim transactions As New Collection
    
    ' Log start of import
    LogInfo "Starting import from " & filePath
    
    ' Read the CSV file
    Dim fileContent As String
    fileContent = ReadEntireFile(filePath)
    
    ' Parse CSV using array-based approach for performance
    Dim csvData As Variant
    csvData = ParseCSV(fileContent)
    
    ' Check if parsing succeeded
    If Not IsArray(csvData) Then
        LogError 0, "Failed to parse CSV file: " & filePath
        Set ICSVImportStrategy_ImportTransactions = transactions
        Exit Function
    End If
    
    ' Find header row and create column index mapping
    Dim headerRowIndex As Long
    Dim columnIndices As Object ' Dictionary
    Set columnIndices = CreateObject("Scripting.Dictionary")
    
    headerRowIndex = FindHeaderRow(csvData, columnIndices)
    
    If headerRowIndex < 0 Then
        LogError 0, "Could not find header row in CSV file: " & filePath
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
        Dim trans As ITransaction
        Set trans = CreateTransactionFromRow(csvData, i, columnIndices, owner, filePath)
        
        ' Add to collection if valid
        If Not trans Is Nothing Then
            transactions.Add trans
        End If
    Next i
    
    ' Log completion
    LogInfo "Imported " & transactions.Count & " transactions from " & filePath
    
    Set ICSVImportStrategy_ImportTransactions = transactions
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
' Common Helper Methods
'=========================================================================

' Read the entire file content
Public Function ReadEntireFile(ByVal filePath As String) As String ' Changed Protected to Public
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
    LogError Err.Number, "Error reading file: " & Err.Description
    ReadEntireFile = ""
End Function

' Parse CSV content into a 2D array
Public Function ParseCSV(ByVal csvContent As String) As Variant ' Changed Protected to Public
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
    LogError Err.Number, "Error parsing CSV: " & Err.Description
    ParseCSV = Array()
End Function

' Count the number of columns in a CSV line
Public Function CountCSVColumns(ByVal line As String) As Long ' Changed Protected to Public
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
Public Function ParseCSVLine(ByVal line As String) As Variant ' Changed Protected to Public
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
                If i < Len(line) And Mid$(line, i + 1, 1) = """" Then
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
    LogError Err.Number, "Error parsing CSV line: " & Err.Description
    
    ' Return a single-element array with the original line on error
    ParseCSVLine = Array(line)
End Function

' Find the header row and map column indices
Public Function FindHeaderRow(ByVal csvData As Variant, ByRef columnIndices As Object) As Long ' Changed Protected to Public
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
    LogError Err.Number, "Error finding header row: " & Err.Description
    FindHeaderRow = -1
End Function

' Abstract method to create a Transaction object from a CSV row
Public Function CreateTransactionFromRow(ByVal csvData As Variant, ByVal rowIndex As Long, _
                                       ByVal columnIndices As Object, _
                                       ByVal owner As String, ByVal filePath As String) As ITransaction
    ' To be implemented by subclasses
    Err.Raise 5, "BaseCSVStrategy", "Abstract method must be overridden"
End Function

'=========================================================================
' Logging Helper Methods
'=========================================================================

' Log an error message
Public Sub LogError(ByVal errorNumber As Long, ByVal errorMessage As String) ' Changed Protected to Public
    If Not m_Logger Is Nothing Then
        m_Logger.LogError m_StrategyName, errorNumber, errorMessage
    End If
End Sub

' Log an informational message
Public Sub LogInfo(ByVal message As String) ' Changed Protected to Public
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo m_StrategyName, message
    End If
End Sub

' Log a warning message
Public Sub LogWarning(ByVal message As String) ' Changed Protected to Public
    If Not m_Logger Is Nothing Then
        m_Logger.LogWarning m_StrategyName, message
    End If
End Sub

'=========================================================================
' Public Methods
'=========================================================================

' Public method to check if this strategy can process a file
Public Function CanProcessFile(ByVal filePath As String) As Boolean
    CanProcessFile = ICSVImportStrategy_CanProcessFile(filePath)
End Function

' Public getter for strategy name
Public Property Get StrategyName() As String
    StrategyName = m_StrategyName
End Property

' Public getter for strategy description
Public Property Get StrategyDescription() As String
    StrategyDescription = m_StrategyDescription
End Property

' Public method to import transactions
Public Function ImportTransactions(ByVal filePath As String, ByVal owner As String, _
                                 Optional ByVal logger As IErrorLogger = Nothing) As Collection
    Set ImportTransactions = ICSVImportStrategy_ImportTransactions(filePath, owner, logger)
End Function

' Public getter for column mapping
Public Function GetColumnMapping() As Object
    Set GetColumnMapping = m_ColumnMapping
End Function

' Public getter for expected headers
Public Function GetExpectedHeaders() As Variant
    GetExpectedHeaders = m_ExpectedHeaders
End Function
