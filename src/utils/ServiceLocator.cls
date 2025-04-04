' File: src/utils/ServiceLocator.cls
'----------------------------------------
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
Private m_BalanceCalculator As IBalanceCalculator
Private m_CategoryAnalyzer As ICategoryAnalyzer
Private m_TimeAnalyzer As ITimeAnalyzer
Private m_TrendAnalyzer As ITrendAnalyzer
Private m_CSVImportEngine As ICSVImportEngine
Private m_ErrorLogger As IErrorLogger
Private m_DashboardManager As IDashboardManager

' ===== Transaction Repository =====

' Get or create TransactionRepository
Public Function GetTransactionRepository() As ITransactionRepository
    If m_TransactionRepository Is Nothing Then
        ' Use PredeclaredId instance
        Set m_TransactionRepository = TransactionRepository ' Removed New
        
        ' Initialize with logger
        m_TransactionRepository.Initialize GetErrorLogger()
    End If
    
    Set GetTransactionRepository = m_TransactionRepository
End Function

' Set custom TransactionRepository (for testing or custom implementations)
Public Sub SetTransactionRepository(ByVal repository As ITransactionRepository)
    Set m_TransactionRepository = repository
End Sub

' ===== Balance Calculator =====

' Get or create BalanceCalculator
Public Function GetBalanceCalculator() As IBalanceCalculator
    If m_BalanceCalculator Is Nothing Then
        ' Use PredeclaredId instance
        Set m_BalanceCalculator = BalanceCalculator ' Removed New
        
        ' Initialize with logger
        m_BalanceCalculator.Initialize GetErrorLogger()
    End If
    
    Set GetBalanceCalculator = m_BalanceCalculator
End Function

' Set custom BalanceCalculator (for testing or custom implementations)
Public Sub SetBalanceCalculator(ByVal calculator As IBalanceCalculator)
    Set m_BalanceCalculator = calculator
End Sub

' ===== Category Analyzer =====

' Get or create CategoryAnalyzer
Public Function GetCategoryAnalyzer() As ICategoryAnalyzer
    If m_CategoryAnalyzer Is Nothing Then
        ' Use PredeclaredId instance of TransactionAnalyzer
        Set m_CategoryAnalyzer = TransactionAnalyzer ' Removed New
        
        ' Initialize with logger (TransactionAnalyzer's Initialize handles logger)
        m_CategoryAnalyzer.Initialize GetErrorLogger()
    End If
    
    Set GetCategoryAnalyzer = m_CategoryAnalyzer
End Function

' Set custom CategoryAnalyzer (for testing or custom implementations)
Public Sub SetCategoryAnalyzer(ByVal analyzer As ICategoryAnalyzer)
    Set m_CategoryAnalyzer = analyzer
End Sub

' ===== Time Analyzer =====

' Get or create TimeAnalyzer
Public Function GetTimeAnalyzer() As ITimeAnalyzer
    If m_TimeAnalyzer Is Nothing Then
        ' Use PredeclaredId instance of TransactionAnalyzer
        Set m_TimeAnalyzer = TransactionAnalyzer ' Removed New
        
        ' Initialize with logger (TransactionAnalyzer's Initialize handles logger)
        m_TimeAnalyzer.Initialize GetErrorLogger()
    End If
    
    Set GetTimeAnalyzer = m_TimeAnalyzer
End Function

' Set custom TimeAnalyzer (for testing or custom implementations)
Public Sub SetTimeAnalyzer(ByVal analyzer As ITimeAnalyzer)
    Set m_TimeAnalyzer = analyzer
End Sub

' ===== Trend Analyzer =====

' Get or create TrendAnalyzer
Public Function GetTrendAnalyzer() As ITrendAnalyzer
    If m_TrendAnalyzer Is Nothing Then
        ' Use PredeclaredId instance of TransactionAnalyzer
        Set m_TrendAnalyzer = TransactionAnalyzer ' Removed New
        
        ' Initialize with logger (TransactionAnalyzer's Initialize handles logger)
        m_TrendAnalyzer.Initialize GetErrorLogger()
    End If
    
    Set GetTrendAnalyzer = m_TrendAnalyzer
End Function

' Set custom TrendAnalyzer (for testing or custom implementations)
Public Sub SetTrendAnalyzer(ByVal analyzer As ITrendAnalyzer)
    Set m_TrendAnalyzer = analyzer
End Sub

' ===== CSV Import Engine =====

' Get or create CSVImportEngine
Public Function GetCSVImportEngine() As ICSVImportEngine
    If m_CSVImportEngine Is Nothing Then
        ' Use PredeclaredId instance
        Set m_CSVImportEngine = CSVImportEngine ' Removed New
        
        ' Initialize with repository and logger
        m_CSVImportEngine.Initialize GetTransactionRepository(), GetErrorLogger()
    End If
    
    Set GetCSVImportEngine = m_CSVImportEngine
End Function

' Set custom CSVImportEngine (for testing or custom implementations)
Public Sub SetCSVImportEngine(ByVal engine As ICSVImportEngine)
    Set m_CSVImportEngine = engine
End Sub

' ===== Error Logger =====

' Get or create ErrorLogger
Public Function GetErrorLogger() As IErrorLogger
    If m_ErrorLogger Is Nothing Then
        ' Use PredeclaredId instance
        Set m_ErrorLogger = ErrorLogger ' Removed New
        
        ' Initialize
        m_ErrorLogger.Initialize
    End If
    
    Set GetErrorLogger = m_ErrorLogger
End Function

' Set custom ErrorLogger (for testing or custom implementations)
Public Sub SetErrorLogger(ByVal logger As IErrorLogger)
    Set m_ErrorLogger = logger
End Sub

' ===== Dashboard Manager =====

' Get or create DashboardManager
Public Function GetDashboardManager() As IDashboardManager
    If m_DashboardManager Is Nothing Then
        ' Use PredeclaredId instance
        Set m_DashboardManager = DashboardManager ' Removed New
        
        ' Initialize with dependencies
        m_DashboardManager.Initialize GetTransactionRepository(), _
                                      GetBalanceCalculator(), _
                                      GetErrorLogger()
    End If
    
    Set GetDashboardManager = m_DashboardManager
End Function

' Set custom DashboardManager (for testing or custom implementations)
Public Sub SetDashboardManager(ByVal manager As IDashboardManager)
    Set m_DashboardManager = manager
End Sub

' ===== Reset All Services =====

' Reset all services (useful for testing)
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
    ' Clean up object references
    Set m_TransactionRepository = Nothing
    Set m_BalanceCalculator = Nothing
    Set m_CategoryAnalyzer = Nothing
    Set m_TimeAnalyzer = Nothing
    Set m_TrendAnalyzer = Nothing
    Set m_CSVImportEngine = Nothing
    Set m_ErrorLogger = Nothing
    Set m_DashboardManager = Nothing
End Sub
