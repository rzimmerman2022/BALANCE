' File: src/services/CreditCardCSVStrategy.cls
'------------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CreditCardCSVStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

' CSV format details
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

Private Sub Class_Initialize()
    ' Set up expected headers for this format
    m_ExpectedHeaders = Array("Transaction Date", "Posting Date", "Description", "Category", "Amount")
    
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

'=========================================================================
' Interface Implementation
'=========================================================================

' Similar methods to BankCSVStrategy, but with credit card specific logic
' [Implementation details would be similar to BankCSVStrategy but with
' adjustments for the credit card format]

' For brevity, I'll just include the unique aspects of this strategy
' that differ from BankCSVStrategy

' Credit cards often have a Category column already
Private Sub ProcessCategoryColumn(ByRef trans As Transaction, ByVal categoryValue As String)
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

' Implementation for CanProcessFile that's specific to credit card format
Private Function ICSVImportStrategy_CanProcessFile(ByVal filePath As String) As Boolean
    ' Implementation similar to BankCSVStrategy but looks for credit card
    ' specific signatures in the file
    
    ' For demonstration purposes, returning True
    ICSVImportStrategy_CanProcessFile = True
End Function

' Get the name of this strategy
Private Property Get ICSVImportStrategy_StrategyName() As String
    ICSVImportStrategy_StrategyName = "CreditCardFormat"
End Property

' Get the description of this strategy
Private Property Get ICSVImportStrategy_StrategyDescription() As String
    ICSVImportStrategy_StrategyDescription = "Credit card statement CSV format with transaction date, description, category, and amount columns"
End Property