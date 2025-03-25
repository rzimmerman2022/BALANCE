DashboardManagerVERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DashboardManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' DashboardManager Class - Manages the main dashboard and UI elements

' Private variables
Private m_IsInitialized As Boolean
Private m_DashboardSheet As Worksheet
Private m_NeedsRefresh As Boolean

' ===== Properties =====

' Dashboard sheet
Public Property Get DashboardSheet() As Worksheet
    If Not m_IsInitialized Then Initialize
    Set DashboardSheet = m_DashboardSheet
End Property

' Needs refresh flag
Public Property Get NeedsRefresh() As Boolean
    NeedsRefresh = m_NeedsRefresh
End Property

Public Property Let NeedsRefresh(value As Boolean)
    m_NeedsRefresh = value
End Property

' ===== Methods =====

' Initialize the dashboard
Public Sub Initialize()
    On Error GoTo ErrorHandler
    
    ' Get or create the dashboard sheet
    Set m_DashboardSheet = Utilities.GetOrCreateSheet("Dashboard", True)
    
    m_IsInitialized = True
    m_NeedsRefresh = True
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "DashboardManager.Initialize", Err.Number, Err.Description
End Sub

' Set up initial dashboard layout
Public Sub SetupDashboard()
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    ' Clear existing content
    Utilities.ClearSheet m_DashboardSheet
    
    ' Apply modern formatting
    Utilities.ApplyModernFormatting m_DashboardSheet
    
    ' Set up header
    SetupHeader
    
    ' Set up summary cards
    SetupSummaryCards
    
    ' Create charts
    Dim chartFactory As ChartFactory
    Set chartFactory = ChartFactory
    chartFactory.Initialize
    chartFactory.CreateAllDashboardCharts m_DashboardSheet
    
    ' Add control buttons
    SetupControlButtons
    
    ' Mark as refreshed
    m_NeedsRefresh = False
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "DashboardManager.SetupDashboard", Err.Number, Err.Description
End Sub

' Refresh the dashboard if needed
Public Sub RefreshDashboard()
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then Initialize
    
    ' Check if refresh is needed
    If Not m_NeedsRefresh Then Exit Sub
    
    ' Update summary cards
    UpdateSummaryCards
    
    ' Refresh charts
    Dim chartFactory As ChartFactory
    Set chartFactory = ChartFactory
    chartFactory.Initialize
    chartFactory.CreateAllDashboardCharts m_DashboardSheet
    
    ' Mark as refreshed
    m_NeedsRefresh = False
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "DashboardManager.RefreshDashboard", Err.Number, Err.Description
End Sub

' Set up the dashboard header
Private Sub SetupHeader()
    On Error GoTo ErrorHandler
    
    With m_DashboardSheet
        ' Add title
        .Range("A1").Value = AppSettings.AppTitle
        .Range("A1").Font.Size = 24
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = AppSettings.ColorPrimary
        
        ' Add subtitle
        .Range("A2").Value = AppSettings.AppSubtitle
        .Range("A2").Font.Size = 12
        .Range("A2").Font.Italic = True
        
        ' Add users info
        .Range("A4").Value = "Users: " & AppSettings.User1Name & " & " & AppSettings.User2Name
        .Range("A4").Font.Bold = True
        
        ' Add last update info
        .Range("A5").Value = "Last Updated: " & Format(AppSettings.LastUpdateDate, "mm/dd/yyyy hh:mm:ss")
        
        ' Add divider line
        With .Range("A6:J6").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = AppSettings.ColorPrimary
        End With
    End With
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "DashboardManager.SetupHeader", Err.Number, Err.Description
End Sub

' Set up summary cards
Private Sub SetupSummaryCards()
    On Error GoTo ErrorHandler
    
    ' Get balance data
    Dim balanceCalc As BalanceCalculator
    Set balanceCalc = BalanceCalculator
    
    Dim repo As TransactionRepository
    Set repo = TransactionRepository
    repo.Initialize
    
    ' Create card for transaction count
    Utilities.CreateCard m_DashboardSheet, "D1", "F3", "Transaction Count", _
                       repo.Transactions.Count, "", AppSettings.ColorPrimary
    
    ' Create card for total spending
    Dim totalSpending As Currency
    Dim trans As Transaction
    
    For Each trans In repo.Transactions
        If trans.IsExpense Then
            totalSpending = totalSpending + Abs(trans.Amount)
        End If
    Next trans
    
    Utilities.CreateCard m_DashboardSheet, "G1", "I3", "Total Spending", _
                       totalSpending, "$#,##0.00", AppSettings.ColorDanger
    
    ' Create balance summary cards
    UpdateSummaryCards
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "DashboardManager.SetupSummaryCards", Err.Number, Err.Description
End Sub

' Update balance summary cards
Private Sub UpdateSummaryCards()
    On Error GoTo ErrorHandler
    
    ' Get balance data
    Dim balanceCalc As BalanceCalculator
    Set balanceCalc = BalanceCalculator
    
    Dim repo As TransactionRepository
    Set repo = TransactionRepository
    repo.Initialize
    
    Dim balance As BalanceSummary
    balance = balanceCalc.CalculateBalance(repo.Transactions)
    
    ' Remove existing balance cards
    On Error Resume Next
    m_DashboardSheet.Shapes("Card_D4").Delete
    m_DashboardSheet.Shapes("CardHeader_D4").Delete
    m_DashboardSheet.Shapes("Card_G4").Delete
    m_DashboardSheet.Shapes("CardHeader_G4").Delete
    On Error GoTo ErrorHandler
    
    ' Create card for who owes
    Utilities.CreateCard m_DashboardSheet, "D4", "F6", "Balance Status", _
                       balance.WhoOwes, "", AppSettings.ColorInfo
    
    ' Create card for amount owed
    Utilities.CreateCard m_DashboardSheet, "G4", "I6", "Amount Owed", _
                       balance.OwedAmount, "$#,##0.00", AppSettings.ColorWarning
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "DashboardManager.UpdateSummaryCards", Err.Number, Err.Description
End Sub

' Set up control buttons
Private Sub SetupControlButtons()
    On Error GoTo ErrorHandler
    
    ' Import CSV button
    Utilities.AddButton m_DashboardSheet, 20, 50, 120, 30, "Import CSV", "BALANCE.ShowImportDialog", _
                      AppSettings.ColorSecondary
    
    ' Refresh Data button
    Utilities.AddButton m_DashboardSheet, 150, 50, 120, 30, "Refresh Data", "BALANCE.RefreshData", _
                      AppSettings.ColorInfo
    
    ' Edit Transactions button
    Utilities.AddButton m_DashboardSheet, 280, 50, 120, 30, "Edit Transactions", "BALANCE.ShowTransactionList", _
                      AppSettings.ColorPrimary
    
    ' Settings button
    Utilities.AddButton m_DashboardSheet, 410, 50, 120, 30, "Settings", "BALANCE.ShowSettings", _
                      AppSettings.ColorDark
    
    Exit Sub
    
ErrorHandler:
    ErrorLogger.LogError "DashboardManager.SetupControlButtons", Err.Number, Err.Description
End Sub

' Class initialize
Private Sub Class_Initialize()
    ' Nothing needed here as we're using PredeclaredId = True
    ' and explicit initialization via Initialize method
    m_IsInitialized = False
    m_NeedsRefresh = True
End Sub

' Class terminate
Private Sub Class_Terminate()
    ' Clean up to prevent memory leaks
    Set m_DashboardSheet = Nothing
End Sub