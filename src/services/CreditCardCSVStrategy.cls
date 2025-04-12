' File: src/services/CreditCardCSVStrategy.cls
'------------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CreditCardCSVStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' CreditCardCSVStrategy Class - Implements ICSVImportStrategy
'
' Purpose: Strategy for importing CSV files from a typical credit card statement.
' This strategy handles the specific format used by many credit card companies,
' with transaction date, posting date, description, category, and amount.

' Interface implementation
Implements ICSVImportStrategy

' CSV format details (Re-add private members)
Private m_ExpectedHeaders As Variant
Private m_ColumnMapping As Object ' Dictionary
Private m_Logger As IErrorLogger

' Credit card-specific parsing rules
Private Const DATE_FORMAT As String = "yyyy-mm-dd"
Private Const CREDIT_INDICATOR As String = "CREDIT"
Private Const FILE_SIGNATURE As String = "Credit Card Transactions"

'=========================================================================
' Initialization
'=========================================================================

' Revert to standard Class_Initialize
Private Sub Class_Initialize()
    ' Set up expected headers for this format
    m_ExpectedHeaders = Array("Transaction Date", "Posting Date", "Description", "Category", "Amount")
    
    ' Set up column mapping (CSV column -> Transaction property)
    ' Set up column mapping (CSV column -> Transaction property)
    Set m_ColumnMapping = CreateObject("Scripting.Dictionary")
    m_ColumnMapping.Add "Transaction Date", "TransactionDate"
    m_ColumnMapping.Add "Description", "Merchant"
    m_ColumnMapping.Add "Amount", "Amount"
    m_ColumnMapping.Add "Category", "Category"
    
    ' Additional mapping information for special handling
    m_ColumnMapping.Add "DateFormat", DATE_FORMAT
    m_ColumnMapping.Add "CreditIndicator", CREDIT_INDICATOR
End Sub

' Add Initialize method for logger injection
Public Sub Initialize(Optional ByVal logger As IErrorLogger = Nothing)
    If Not logger Is Nothing Then
        Set m_Logger = logger
    End If
End Sub

'=========================================================================
'=========================================================================
' Interface Implementation (Required by ICSVImportStrategy)
'=========================================================================

' Implementation for CanProcessFile that's specific to credit card format
Private Function ICSVImportStrategy_CanProcessFile(ByVal filePath As String) As Boolean
    ' Implementation similar to BankCSVStrategy but looks for credit card
    ' specific signatures in the file (e.g., expected headers)
    
    ' Basic check: See if the file contains the expected headers
    On Error Resume Next ' Use error handling carefully
    Dim baseHelper As New BaseCSVStrategy
    Dim fileContent As String
    fileContent = baseHelper.ReadEntireFile(filePath) ' Use baseHelper
    If Err.Number <> 0 Then GoTo HandleError

    Dim csvData As Variant
    csvData = baseHelper.ParseCSV(fileContent) ' Use baseHelper
    If Err.Number <> 0 Then GoTo HandleError

    Dim headerRowIndex As Long
    Dim tempColumnIndices As Object
    Set tempColumnIndices = CreateObject("Scripting.Dictionary")

    ' Pass this object's expected headers to the helper
    baseHelper.m_ExpectedHeaders = m_ExpectedHeaders
    headerRowIndex = baseHelper.FindHeaderRow(csvData, tempColumnIndices) ' Use baseHelper

    If headerRowIndex > 0 Then
        ICSVImportStrategy_CanProcessFile = True ' Assign to correct function name
    Else
        ICSVImportStrategy_CanProcessFile = False ' Assign to correct function name
    End If

    Exit Function

HandleError:
    ICSVImportStrategy_CanProcessFile = False
End Function


' Implementation for creating a transaction from a row (Helper, not directly from interface)
Private Function CreateTransactionFromRow(ByVal csvData As Variant, ByVal rowIndex As Long, _
                                        ByVal columnIndices As Object, _
                                        ByVal owner As String, ByVal filePath As String) As Transaction ' Return concrete type
    On Error GoTo ErrorHandler
    
    Dim trans As New Transaction
    Dim colIndex As Long
    Dim value As String
    
    ' Set owner and source file
    trans.Owner = owner
    trans.SourceFile = filePath
    
    ' Process mapped columns
    Dim key As Variant
    For Each key In m_ColumnMapping.Keys
        ' Skip special mapping info
        If key = "DateFormat" Or key = "CreditIndicator" Then GoTo NextKey
        
        ' Get column index for this header
        If columnIndices.Exists(key) Then
            colIndex = columnIndices(key)
            
            ' Get value from CSV data
            value = Trim$(csvData(rowIndex, colIndex))
            
            ' Set transaction property based on mapping
            Select Case m_ColumnMapping(key)
                Case "TransactionDate"
                    trans.TransactionDate = CDate(value) ' Assuming standard date format or use DATE_FORMAT
                Case "Merchant"
                    trans.Merchant = value
                Case "Amount"
                    trans.Amount = ParseCreditCardAmount(value, m_ColumnMapping("CreditIndicator"))
                Case "Category"
                    ProcessCategoryColumn trans, value
                ' Add other cases as needed
            End Select
        Else
            ' Log warning if expected column is missing
            If Not m_Logger Is Nothing Then m_Logger.LogWarning "CreditCardCSVStrategy", "Column not found in CSV: " & key ' Use m_Logger
        End If
NextKey:
    Next key
    
    ' Set default values if needed (e.g., Account)
    trans.Account = "Credit Card" ' Default account for this strategy
    
    ' Return the created transaction
    Set CreateTransactionFromRow = trans
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then m_Logger.LogError "CreditCardCSVStrategy", Err.Number, "Error creating transaction from row " & rowIndex & ": " & Err.Description ' Use m_Logger
    Set CreateTransactionFromRow = Nothing
End Function

'=========================================================================
' Credit Card Specific Helper Methods
'=========================================================================

' Credit cards often have a Category column already
Private Sub ProcessCategoryColumn(ByRef trans As Transaction, ByVal categoryValue As String) ' Use concrete type
    ' Map credit card categories to our own categories
    Select Case UCase(categoryValue)
        Case "RESTAURANTS", "DINING"
            trans.Category = "Dining"
            
        Case "GROCERIES", "SUPERMARKET"
            trans.Category = "Groceries"
            
        Case "TRAVEL", "AIRLINE", "HOTEL", "CAR RENTAL"
            trans.Category = "Travel"
            
        Case "GAS", "AUTOMOTIVE", "FUEL"
            trans.Category = "Transportation"
            
        Case "ENTERTAINMENT", "MOVIES", "MUSIC"
            trans.Category = "Entertainment"
            
        Case "MEDICAL", "HEALTHCARE", "PHARMACY"
            trans.Category = "Healthcare"
            
        Case "MERCHANDISE", "SHOPPING", "DEPARTMENT STORE"
            trans.Category = "Shopping"
            
        Case Else
            ' Keep original or set to uncategorized if empty
            If Len(Trim$(categoryValue)) > 0 Then
                trans.Category = categoryValue
            Else
                trans.Category = "Uncategorized"
            End If
    End Select
End Sub

' Credit card amounts are often positive for charges and negative for payments
Private Function ParseCreditCardAmount(ByVal amountStr As String, ByVal creditIndicator As String) As Double
    On Error GoTo ErrorHandler

    ' Remove currency symbols, commas, etc.
    amountStr = Replace(amountStr, "$", "")
    amountStr = Replace(amountStr, ",", "")
    
    ' Check if this is a credit (negative expense, i.e., a payment or refund)
    Dim isCredit As Boolean
    isCredit = (InStr(1, amountStr, creditIndicator, vbTextCompare) > 0)
    
    ' Remove any text indicators
    amountStr = Replace(amountStr, creditIndicator, "", , , vbTextCompare)
    amountStr = Trim$(amountStr)
    
    ' Convert to number
    Dim amount As Double
    amount = CDbl(amountStr)
    
    ' For credit cards, charges are positive in the statement but should be
    ' negative expenses in our system, and payments/refunds are the opposite
    If isCredit Then
        amount = amount ' Payments/refunds are positive
    Else
        amount = -amount ' Charges are negative (expenses)
    End If
    
    ParseCreditCardAmount = amount
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "CreditCardCSVStrategy.ParseCreditCardAmount", Err.Number, _
            "Could not parse amount: " & amountStr
    End If
    
    ' Return 0 on error
    ParseCreditCardAmount = 0
End Function

'=========================================================================
'=========================================================================
' Interface Implementation (Continued)
'=========================================================================

' Get the name of this strategy
Private Property Get ICSVImportStrategy_StrategyName() As String
    ICSVImportStrategy_StrategyName = "CreditCardFormat"
End Property

' Get the description of this strategy
Private Property Get ICSVImportStrategy_StrategyDescription() As String
    ICSVImportStrategy_StrategyDescription = "Credit card statement CSV format with transaction date, description, category, and amount columns"
End Property

' Get the column mapping for this strategy
Private Function ICSVImportStrategy_GetColumnMapping() As Object
    Set ICSVImportStrategy_GetColumnMapping = m_ColumnMapping
End Function

' Get expected headers for this strategy
Private Function ICSVImportStrategy_GetExpectedHeaders() As Variant
    ICSVImportStrategy_GetExpectedHeaders = m_ExpectedHeaders
End Function

' Import transactions from the CSV file (Direct Implementation)
Private Function ICSVImportStrategy_ImportTransactions(ByVal filePath As String, ByVal owner As String, _
                                                    Optional ByVal logger As IErrorLogger = Nothing) As Collection
    ' Set logger if provided
    If Not logger Is Nothing Then
        Set m_Logger = logger
    End If

    ' Create collection for transactions
    Dim transactions As New Collection
    Dim baseHelper As New BaseCSVStrategy ' Use composition for helpers

    ' Log start of import
    If Not m_Logger Is Nothing Then m_Logger.LogInfo "CreditCardCSVStrategy", "Starting import from " & filePath

    ' Read the CSV file using base helper
    Dim fileContent As String
    fileContent = baseHelper.ReadEntireFile(filePath)
    If Len(fileContent) = 0 Then GoTo ImportFailed

    ' Parse CSV using base helper
    Dim csvData As Variant
    csvData = baseHelper.ParseCSV(fileContent)
    If Not IsArray(csvData) Then GoTo ImportFailed

    ' Find header row and create column index mapping using base helper logic
    Dim headerRowIndex As Long
    Dim columnIndices As Object ' Dictionary
    Set columnIndices = CreateObject("Scripting.Dictionary")

    ' Need to pass *this* object's expected headers to the helper
    baseHelper.m_ExpectedHeaders = m_ExpectedHeaders
    headerRowIndex = baseHelper.FindHeaderRow(csvData, columnIndices)

    If headerRowIndex < 0 Then
        If Not m_Logger Is Nothing Then m_Logger.LogError "CreditCardCSVStrategy", 0, "Could not find header row in CSV file: " & filePath
        GoTo ImportFailed
    End If

    ' Process data rows
    Dim i As Long
    Dim rowCount As Long
    rowCount = UBound(csvData, 1)

    ' Process each row after the header
    For i = headerRowIndex + 1 To rowCount
        ' Create a transaction from the row using *this* class's method
        Dim trans As Transaction ' Use concrete type
        Set trans = CreateTransactionFromRow(csvData, i, columnIndices, owner, filePath)

        ' Add to collection if valid
        If Not trans Is Nothing Then
            transactions.Add trans
        End If
    Next i

    ' Log completion
    If Not m_Logger Is Nothing Then m_Logger.LogInfo "CreditCardCSVStrategy", "Imported " & transactions.Count & " transactions from " & filePath

    Set ICSVImportStrategy_ImportTransactions = transactions
    Exit Function

ImportFailed:
    ' Return empty collection on failure
    Set ICSVImportStrategy_ImportTransactions = New Collection
End Function
