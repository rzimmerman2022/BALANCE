VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' Transaction Class - Implements ITransaction interface
'
' Purpose: Represents a financial transaction with all its properties and behaviors.
' This class has been updated to implement the ITransaction interface,
' following the Interface Segregation Principle.

Implements ITransaction

' Private member variables
Private m_ID As String
Private m_Date As Date
Private m_Merchant As String
Private m_Category As String
Private m_Amount As Currency
Private m_Account As String
Private m_Owner As String
Private m_IsShared As Boolean
Private m_Notes As String
Private m_SourceFile As String

' ===== Properties =====

' Transaction ID
Public Property Get ID() As String
    ID = m_ID
End Property

Public Property Let ID(value As String)
    m_ID = value
End Property

' Transaction Date
Public Property Get TransactionDate() As Date
    TransactionDate = m_Date
End Property

Public Property Let TransactionDate(value As Date)
    m_Date = value
End Property

' Merchant
Public Property Get Merchant() As String
    Merchant = m_Merchant
End Property

Public Property Let Merchant(value As String)
    m_Merchant = value
End Property

' Category
Public Property Get Category() As String
    Category = m_Category
End Property

Public Property Let Category(value As String)
    m_Category = value
End Property

' Amount
Public Property Get Amount() As Currency
    Amount = m_Amount
End Property

Public Property Let Amount(value As Currency)
    m_Amount = value
End Property

' Account
Public Property Get Account() As String
    Account = m_Account
End Property

Public Property Let Account(value As String)
    m_Account = value
End Property

' Owner
Public Property Get Owner() As String
    Owner = m_Owner
End Property

Public Property Let Owner(value As String)
    m_Owner = value
End Property

' IsShared
Public Property Get IsShared() As Boolean
    IsShared = m_IsShared
End Property

Public Property Let IsShared(value As Boolean)
    m_IsShared = value
End Property

' Notes
Public Property Get Notes() As String
    Notes = m_Notes
End Property

Public Property Let Notes(value As String)
    m_Notes = value
End Property

' SourceFile
Public Property Get SourceFile() As String
    SourceFile = m_SourceFile
End Property

Public Property Let SourceFile(value As String)
    m_SourceFile = value
End Property

' ===== Methods =====

' Determine if transaction is an expense
Public Function IsExpense() As Boolean
    IsExpense = (m_Amount < 0)
End Function

' Determine if transaction is income
Public Function IsIncome() As Boolean
    IsIncome = (m_Amount > 0)
End Function

' Get the absolute amount (positive value)
Public Function GetAbsoluteAmount() As Currency
    GetAbsoluteAmount = Abs(m_Amount)
End Function

' Calculate split amounts based on ownership and shared status
Public Sub CalculateSplit(ByRef user1Amount As Currency, ByRef user2Amount As Currency)
    On Error GoTo ErrorHandler
    
    Dim absAmount As Currency
    absAmount = Abs(m_Amount)
    
    ' Handle expenses (negative amounts)
    If IsExpense() Then
        If m_IsShared Then
            ' Shared expense - split 50/50
            user1Amount = Round(absAmount / 2, 2)
            user2Amount = Round(absAmount / 2, 2)
        Else
            ' Individual expense - owner pays all
            If LCase(m_Owner) = LCase(AppSettings.User1Name) Then
                user1Amount = absAmount
                user2Amount = 0
            Else
                user1Amount = 0
                user2Amount = absAmount
            End If
        End If
    ' Handle income (positive amounts)
    Else
        ' Income is credited to the owner
        If LCase(m_Owner) = LCase(AppSettings.User1Name) Then
            user1Amount = -absAmount  ' Credit (negative value represents credit)
            user2Amount = 0
        Else
            user1Amount = 0
            user2Amount = -absAmount  ' Credit (negative value represents credit)
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "Transaction.CalculateSplit", Err.Number, Err.Description
End Sub

' Creates a dictionary representing the transaction
Public Function ToDictionary() As Object
    On Error GoTo ErrorHandler
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    dict.Add "ID", m_ID
    dict.Add "Date", m_Date
    dict.Add "Merchant", m_Merchant
    dict.Add "Category", m_Category
    dict.Add "Amount", m_Amount
    dict.Add "Account", m_Account
    dict.Add "Owner", m_Owner
    dict.Add "IsShared", m_IsShared
    dict.Add "Notes", m_Notes
    dict.Add "SourceFile", m_SourceFile
    
    Set ToDictionary = dict
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "Transaction.ToDictionary", Err.Number, Err.Description
    Set ToDictionary = Nothing
End Function

' Initialize a transaction from row data
Public Sub InitFromRow(dateVal As Variant, merchant As Variant, category As Variant, _
                      amount As Variant, account As Variant, owner As Variant, _
                      isShared As Variant, notes As Variant, sourceFile As Variant)
    On Error GoTo ErrorHandler
    
    ' Generate a unique ID if not already set
    If Len(m_ID) = 0 Then
        m_ID = GenerateID()
    End If
    
    ' Set properties with validation
    
    ' Date
    If IsDate(dateVal) Then
        m_Date = dateVal
    Else
        m_Date = Date ' Default to today
    End If
    
    ' Merchant
    If Not IsEmpty(merchant) Then
        m_Merchant = CStr(merchant)
    Else
        m_Merchant = "Unknown"
    End If
    
    ' Category
    If Not IsEmpty(category) Then
        m_Category = CStr(category)
    Else
        m_Category = "Uncategorized"
    End If
    
    ' Amount
    If IsNumeric(amount) Then
        m_Amount = CCur(amount)
    Else
        ' Try to clean the amount string
        Dim amtStr As String
        amtStr = CStr(amount)
        amtStr = Replace(amtStr, "$", "")
        amtStr = Replace(amtStr, ",", "")
        amtStr = Replace(amtStr, " ", "")
        
        ' Handle European format with comma as decimal
        If InStr(amtStr, ".") = 0 And InStr(amtStr, ",") > 0 Then
            amtStr = Replace(amtStr, ",", ".")
        End If
        
        ' Some CSVs use parentheses for negative numbers
        If Left(amtStr, 1) = "(" And Right(amtStr, 1) = ")" Then
            amtStr = "-" & Mid(amtStr, 2, Len(amtStr) - 2)
        End If
        
        If IsNumeric(amtStr) Then
            m_Amount = CCur(amtStr)
        Else
            m_Amount = 0 ' Default
        End If
    End If
    
    ' Account
    If Not IsEmpty(account) Then
        m_Account = CStr(account)
    Else
        m_Account = ""
    End If
    
    ' Owner
    If Not IsEmpty(owner) Then
        m_Owner = CStr(owner)
    Else
        m_Owner = AppSettings.User1Name ' Default to User 1
    End If
    
    ' IsShared
    If TypeName(isShared) = "Boolean" Then
        m_IsShared = CBool(isShared)
    ElseIf IsNumeric(isShared) Then
        m_IsShared = (CInt(isShared) <> 0)
    ElseIf TypeName(isShared) = "String" Then
        Dim sharedStr As String
        sharedStr = LCase(CStr(isShared))
        m_IsShared = (sharedStr = "true" Or sharedStr = "yes" Or sharedStr = "1")
    Else
        ' Default based on category
        If Not IsEmpty(category) Then
            Dim cat As String
            cat = LCase(CStr(category))
            
            If cat = "loan repayment" Or cat = "financial fees" Or cat = "income" Or _
               InStr(1, cat, "personal", vbTextCompare) > 0 Then
                m_IsShared = False
            Else
                m_IsShared = True
            End If
        Else
            m_IsShared = True ' Default to shared
        End If
    End If
    
    ' Notes
    If Not IsEmpty(notes) Then
        m_Notes = CStr(notes)
    Else
        m_Notes = ""
    End If
    
    ' SourceFile
    If Not IsEmpty(sourceFile) Then
        m_SourceFile = CStr(sourceFile)
    Else
        m_SourceFile = ""
    End If
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "Transaction.InitFromRow", Err.Number, Err.Description
End Sub

' Generate a unique ID for the transaction
Private Function GenerateID() As String
    On Error GoTo ErrorHandler
    
    ' Create a GUID-like ID
    Dim guidPart1 As String: guidPart1 = Hex(Int(Rnd * &HFFFF&))
    Dim guidPart2 As String: guidPart2 = Hex(Int(Rnd * &HFFFF&))
    Dim guidPart3 As String: guidPart3 = Hex(Int(Rnd * &HFFFF&))
    Dim guidPart4 As String: guidPart4 = Hex(Int(Rnd * &HFFFF&))
    Dim timeStamp As String: timeStamp = Hex(CLng(Now * 86400))
    
    ' Format parts to ensure consistent length
    While Len(guidPart1) < 4: guidPart1 = "0" & guidPart1: Wend
    While Len(guidPart2) < 4: guidPart2 = "0" & guidPart2: Wend
    While Len(guidPart3) < 4: guidPart3 = "0" & guidPart3: Wend
    While Len(guidPart4) < 4: guidPart4 = "0" & guidPart4: Wend
    
    GenerateID = guidPart1 & "-" & guidPart2 & "-" & guidPart3 & "-" & guidPart4 & "-" & timeStamp
    
    Exit Function
    
ErrorHandler:
    ErrorLogger.LogError "Transaction.GenerateID", Err.Number, Err.Description
    GenerateID = "TX-" & Format(Now(), "yyyymmddhhnnss") & "-" & Int(Rnd * 1000)
End Function

' Class initialize
Private Sub Class_Initialize()
    On Error Resume Next
    
    ' Set default values
    m_ID = ""
    m_Date = Date
    m_Merchant = ""
    m_Category = "Uncategorized"
    m_Amount = 0
    m_Account = ""
    m_Owner = ""
    m_IsShared = True
    m_Notes = ""
    m_SourceFile = ""
    
    ' Seed the random number generator for ID generation
    Randomize
    
    On Error GoTo 0
End Sub

' Class terminate
Private Sub Class_Terminate()
    ' Nothing to clean up
End Sub

'=========================================================================
' ITransaction Interface Implementation
'=========================================================================

' ID property
Private Property Get ITransaction_ID() As String
    ITransaction_ID = ID
End Property

Private Property Let ITransaction_ID(ByVal value As String)
    ID = value
End Property

' TransactionDate property
Private Property Get ITransaction_TransactionDate() As Date
    ITransaction_TransactionDate = TransactionDate
End Property

Private Property Let ITransaction_TransactionDate(ByVal value As Date)
    TransactionDate = value
End Property

' Merchant property
Private Property Get ITransaction_Merchant() As String
    ITransaction_Merchant = Merchant
End Property

Private Property Let ITransaction_Merchant(ByVal value As String)
    Merchant = value
End Property

' Category property
Private Property Get ITransaction_Category() As String
    ITransaction_Category = Category
End Property

Private Property Let ITransaction_Category(ByVal value As String)
    Category = value
End Property

' Amount property
Private Property Get ITransaction_Amount() As Currency
    ITransaction_Amount = Amount
End Property

Private Property Let ITransaction_Amount(ByVal value As Currency)
    Amount = value
End Property

' Account property
Private Property Get ITransaction_Account() As String
    ITransaction_Account = Account
End Property

Private Property Let ITransaction_Account(ByVal value As String)
    Account = value
End Property

' Owner property
Private Property Get ITransaction_Owner() As String
    ITransaction_Owner = Owner
End Property

Private Property Let ITransaction_Owner(ByVal value As String)
    Owner = value
End Property

' IsShared property
Private Property Get ITransaction_IsShared() As Boolean
    ITransaction_IsShared = IsShared
End Property

Private Property Let ITransaction_IsShared(ByVal value As Boolean)
    IsShared = value
End Property

' Notes property
Private Property Get ITransaction_Notes() As String
    ITransaction_Notes = Notes
End Property

Private Property Let ITransaction_Notes(ByVal value As String)
    Notes = value
End Property

' SourceFile property
Private Property Get ITransaction_SourceFile() As String
    ITransaction_SourceFile = SourceFile
End Property

Private Property Let ITransaction_SourceFile(ByVal value As String)
    SourceFile = value
End Property

' IsExpense method
Private Function ITransaction_IsExpense() As Boolean
    ITransaction_IsExpense = IsExpense()
End Function

' IsIncome method
Private Function ITransaction_IsIncome() As Boolean
    ITransaction_IsIncome = IsIncome()
End Function

' GetAbsoluteAmount method
Private Function ITransaction_GetAbsoluteAmount() As Currency
    ITransaction_GetAbsoluteAmount = GetAbsoluteAmount()
End Function

' CalculateSplit method
Private Sub ITransaction_CalculateSplit(ByRef user1Amount As Currency, ByRef user2Amount As Currency)
    CalculateSplit user1Amount, user2Amount
End Sub

' ToDictionary method
Private Function ITransaction_ToDictionary() As Object
    Set ITransaction_ToDictionary = ToDictionary()
End Function

' InitFromRow method
Private Sub ITransaction_InitFromRow(dateVal As Variant, merchant As Variant, category As Variant, _
                                   amount As Variant, account As Variant, owner As Variant, _
                                   isShared As Variant, notes As Variant, sourceFile As Variant)
    ' Call the public InitFromRow method of this class
    Me.InitFromRow dateVal, merchant, category, amount, account, owner, isShared, notes, sourceFile
End Sub
