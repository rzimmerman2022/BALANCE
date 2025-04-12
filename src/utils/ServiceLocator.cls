VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServiceLocator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' ServiceLocator Class
'
' Purpose: Provides centralized access to service instances.
' This class implements the Service Locator pattern to manage dependencies
' and remove the need for Singleton patterns in service classes.
'
' Note: Uses PredeclaredId = True to enable global access as ServiceLocator.GetXXX()

' Private member variables for service instances
Private m_TransactionRepository As ITransactionRepository
Private m_BalanceCalculator As IBalanceCalculator ' Keep this specific reference if needed elsewhere
Private m_CategoryAnalyzer As ICategoryAnalyzer   ' Reference for Category Analysis
Private m_TimeAnalyzer As ITimeAnalyzer         ' Reference for Time Analysis
Private m_TrendAnalyzer As ITrendAnalyzer       ' Reference for Trend Analysis
Private m_CSVImportEngine As ICSVImportEngine
Private m_ErrorLogger As IErrorLogger
Private m_DashboardManager As IDashboardManager

' Helper function to get the single BalanceCalculator instance (implements multiple interfaces)
Private Function GetBalanceCalculatorInstance() As BalanceCalculator ' Return concrete type internally
    If m_BalanceCalculator Is Nothing Then
        Set m_BalanceCalculator = BalanceCalculator ' Get PredeclaredId instance
        m_BalanceCalculator.Initialize GetErrorLogger() ' Ensure initialized
    End If
    Set GetBalanceCalculatorInstance = m_BalanceCalculator
End Function


' ===== Transaction Repository =====
Public Function GetTransactionRepository() As ITransactionRepository
    If m_TransactionRepository Is Nothing Then
        Set m_TransactionRepository = TransactionRepository
        m_TransactionRepository.Initialize GetErrorLogger()
    End If
    Set GetTransactionRepository = m_TransactionRepository
End Function
Public Sub SetTransactionRepository(ByVal repository As ITransactionRepository): Set m_TransactionRepository = repository: End Sub

' ===== Balance Calculator (as IBalanceCalculator) =====
Public Function GetBalanceCalculator() As IBalanceCalculator
    Set GetBalanceCalculator = GetBalanceCalculatorInstance ' Use helper
End Function
Public Sub SetBalanceCalculator(ByVal calculator As IBalanceCalculator): Set m_BalanceCalculator = calculator: End Sub

' ===== Category Analyzer =====
' ***** CORRECTED: Return BalanceCalculator instance as ICategoryAnalyzer *****
Public Function GetCategoryAnalyzer() As ICategoryAnalyzer
    If m_CategoryAnalyzer Is Nothing Then
        Set m_CategoryAnalyzer = GetBalanceCalculatorInstance ' Get calculator instance which implements this
    End If
    Set GetCategoryAnalyzer = m_CategoryAnalyzer
End Function
Public Sub SetCategoryAnalyzer(ByVal analyzer As ICategoryAnalyzer): Set m_CategoryAnalyzer = analyzer: End Sub

' ===== Time Analyzer =====
' ***** CORRECTED: Return BalanceCalculator instance as ITimeAnalyzer *****
Public Function GetTimeAnalyzer() As ITimeAnalyzer
     If m_TimeAnalyzer Is Nothing Then
        Set m_TimeAnalyzer = GetBalanceCalculatorInstance ' Get calculator instance which implements this
    End If
    Set GetTimeAnalyzer = m_TimeAnalyzer
End Function
Public Sub SetTimeAnalyzer(ByVal analyzer As ITimeAnalyzer): Set m_TimeAnalyzer = analyzer: End Sub

' ===== Trend Analyzer =====
' ***** CORRECTED: Return BalanceCalculator instance as ITrendAnalyzer *****
Public Function GetTrendAnalyzer() As ITrendAnalyzer
     If m_TrendAnalyzer Is Nothing Then
        Set m_TrendAnalyzer = GetBalanceCalculatorInstance ' Get calculator instance which implements this
    End If
    Set GetTrendAnalyzer = m_TrendAnalyzer
End Function
Public Sub SetTrendAnalyzer(ByVal analyzer As ITrendAnalyzer): Set m_TrendAnalyzer = analyzer: End Sub

' ===== CSV Import Engine =====
Public Function GetCSVImportEngine() As ICSVImportEngine
    If m_CSVImportEngine Is Nothing Then
        Set m_CSVImportEngine = CSVImportEngine
        m_CSVImportEngine.Initialize GetTransactionRepository(), GetErrorLogger()
    End If
    Set GetCSVImportEngine = m_CSVImportEngine
End Function
Public Sub SetCSVImportEngine(ByVal engine As ICSVImportEngine): Set m_CSVImportEngine = engine: End Sub

' ===== Error Logger =====
Public Function GetErrorLogger() As IErrorLogger
    If m_ErrorLogger Is Nothing Then
        Set m_ErrorLogger = ErrorLogger
        m_ErrorLogger.Initialize
    End If
    Set GetErrorLogger = m_ErrorLogger
End Function
Public Sub SetErrorLogger(ByVal logger As IErrorLogger): Set m_ErrorLogger = logger: End Sub

' ===== Dashboard Manager =====
Public Function GetDashboardManager() As IDashboardManager
    If m_DashboardManager Is Nothing Then
        Set m_DashboardManager = DashboardManager
        m_DashboardManager.Initialize ' Dependencies now resolved internally via ServiceLocator calls
    End If
    Set GetDashboardManager = m_DashboardManager
End Function
Public Sub SetDashboardManager(ByVal manager As IDashboardManager): Set m_DashboardManager = manager: End Sub

' ===== Reset All Services =====
Public Sub ResetAll()
    Set m_TransactionRepository = Nothing
    Set m_BalanceCalculator = Nothing
    Set m_CategoryAnalyzer = Nothing
    Set m_TimeAnalyzer = Nothing
    Set m_TrendAnalyzer = Nothing
    Set m_CSVImportEngine = Nothing
    Set m_ErrorLogger = Nothing
    Set m_DashboardManager = Nothing
End Sub

' ===== Cleanup =====
Private Sub Class_Terminate()
    ResetAll ' Call ResetAll to release all references
End Sub