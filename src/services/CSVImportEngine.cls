' File: src/services/CSVImportEngine.cls
'----------------------------------------
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
' CSVImportEngine Class - Implements ICSVImportEngine
'
' Purpose: Central engine for importing transactions from CSV files.
' This class coordinates the import process, selecting appropriate strategies
' for different CSV formats and handling the overall import workflow.
'
' Design Decisions:
' - Uses PredeclaredId = True to enable Singleton pattern
' - Implements ICSVImportEngine for dependency injection
' - Uses the Strategy pattern to handle different CSV formats
' - Supports both single file and batch directory imports
' - Performs auto-detection of CSV formats

' Interface implementation
Implements ICSVImportEngine

' Private member variables
Private m_Strategies As Collection       ' Collection of ICSVImportStrategy objects
Private m_Repository As ITransactionRepository ' Repository for storing imported transactions
Private m_Logger As IErrorLogger         ' Logger for error handling
Private m_IsInitialized As Boolean       ' Initialization flag

'=========================================================================
' Initialization
'=========================================================================

Private Sub Class_Initialize()
    ' Create collection for strategies
    Set m_Strategies = New Collection
    m_IsInitialized = False
End Sub

' Initialize the engine with optional dependencies
Private Sub ICSVImportEngine_Initialize(Optional ByVal repository As ITransactionRepository = Nothing, _
                                       Optional ByVal logger As IErrorLogger = Nothing)
    On Error GoTo ErrorHandler
    
    ' Set repository if provided
    If Not repository Is Nothing Then
        Set m_Repository = repository
    End If
    
    ' Set logger if provided
    If Not logger Is Nothing Then
        Set m_Logger = logger
    End If
    
    ' Log initialization
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "CSVImportEngine.Initialize", "CSV Import Engine initialized"
    End If
    
    ' Register built-in strategies
    If m_Strategies.Count = 0 Then
        ' Bank account strategy
        Dim bankStrategy As New BankCSVStrategy
        Me.RegisterStrategy bankStrategy
        
        ' Credit card strategy
        Dim creditCardStrategy As New CreditCardCSVStrategy
        Me.RegisterStrategy creditCardStrategy
        
        ' Additional strategies can be registered here
    End If
    
    m_IsInitialized = True
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "CSVImportEngine.Initialize", Err.Number, Err.Description
    End If
End Sub

' Public wrapper for Initialize
Public Sub Initialize(Optional ByVal repository As ITransactionRepository = Nothing, _
                     Optional ByVal logger As IErrorLogger = Nothing)
    ICSVImportEngine_Initialize repository, logger
End Sub

'=========================================================================
' Interface Implementation - Strategy Management
'=========================================================================

' Register a strategy for a particular CSV format
Private Sub ICSVImportEngine_RegisterStrategy(ByVal strategy As ICSVImportStrategy)
    On Error GoTo ErrorHandler
    
    ' Ensure engine is initialized
    If Not m_IsInitialized Then ICSVImportEngine_Initialize
    
    ' Add the strategy to the collection
    On Error Resume Next
    m_Strategies.Remove strategy.StrategyName
    On Error GoTo ErrorHandler
    
    m_Strategies.Add strategy, strategy.StrategyName
    
    ' Log registration
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "CSVImportEngine.RegisterStrategy", _
            "Registered strategy: " & strategy.StrategyName
    End If
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "CSVImportEngine.RegisterStrategy", Err.Number, Err.Description
    End If
End Sub

' Public wrapper for RegisterStrategy
Public Sub RegisterStrategy(ByVal strategy As ICSVImportStrategy)
    ICSVImportEngine_RegisterStrategy strategy
End Sub

' Unregister a strategy by name
Private Function ICSVImportEngine_UnregisterStrategy(ByVal strategyName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Ensure engine is initialized
    If Not m_IsInitialized Then ICSVImportEngine_Initialize
    
    ' Try to remove the strategy
    On Error Resume Next
    m_Strategies.Remove strategyName
    
    If Err.Number = 0 Then
        ICSVImportEngine_UnregisterStrategy = True
        
        ' Log unregistration
        If Not m_Logger Is Nothing Then
            m_Logger.LogInfo "CSVImportEngine.UnregisterStrategy", _
                "Unregistered strategy: " & strategyName
        End If
    Else
        ICSVImportEngine_UnregisterStrategy = False
        
        ' Log strategy not found
        If Not m_Logger Is Nothing Then
            m_Logger.LogWarning "CSVImportEngine.UnregisterStrategy", _
                "Strategy not found: " & strategyName
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "CSVImportEngine.UnregisterStrategy", Err.Number, Err.Description
    End If
    ICSVImportEngine_UnregisterStrategy = False
End Function

' Public wrapper for UnregisterStrategy
Public Function UnregisterStrategy(ByVal strategyName As String) As Boolean
    UnregisterStrategy = ICSVImportEngine_UnregisterStrategy(strategyName)
End Function

' Get a list of all registered strategies
Private Function ICSVImportEngine_GetRegisteredStrategies() As Collection
    On Error GoTo ErrorHandler
    
    ' Ensure engine is initialized
    If Not m_IsInitialized Then ICSVImportEngine_Initialize
    
    ' Create a new collection to avoid returning a reference to the internal collection
    Dim result As New Collection
    
    ' Add strategy info to the collection
    Dim strategy As ICSVImportStrategy
    For Each strategy In m_Strategies
        Dim strategyInfo As Object
        Set strategyInfo = CreateObject("Scripting.Dictionary")
        
        strategyInfo.Add "Name", strategy.StrategyName
        strategyInfo.Add "Description", strategy.StrategyDescription
        
        result.Add strategyInfo
    Next strategy
    
    Set ICSVImportEngine_GetRegisteredStrategies = result
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "CSVImportEngine.GetRegisteredStrategies", Err.Number, Err.Description
    End If
    Set ICSVImportEngine_GetRegisteredStrategies = New Collection
End Function

' Public wrapper for GetRegisteredStrategies
Public Function GetRegisteredStrategies() As Collection
    Set GetRegisteredStrategies = ICSVImportEngine_GetRegisteredStrategies()
End Function

'=========================================================================
' Interface Implementation - Import Methods
'=========================================================================

' Import transactions from a specific CSV file
Private Function ICSVImportEngine_ImportTransactionsFromFile(ByVal filePath As String, ByVal owner As String, _
                                                           Optional ByVal specificStrategy As String = "") As Collection
    On Error GoTo ErrorHandler
    
    ' Ensure engine is initialized
    If Not m_IsInitialized Then ICSVImportEngine_Initialize
    
    ' Create collection for imported transactions
    Dim importedTransactions As New Collection
    
    ' Log start of import
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "CSVImportEngine.ImportTransactionsFromFile", _
            "Starting import from " & filePath
    End If
    
    ' Check if file exists
    If Dir(filePath) = "" Then
        If Not m_Logger Is Nothing Then
            m_Logger.LogError "CSVImportEngine.ImportTransactionsFromFile", 53, _
                "File not found: " & filePath
        End If
        Set ICSVImportEngine_ImportTransactionsFromFile = importedTransactions
        Exit Function
    End If
    
    ' Determine which strategy to use
    Dim strategy As ICSVImportStrategy
    
    If specificStrategy <> "" Then
        ' Use specified strategy
        On Error Resume Next
        Set strategy = m_Strategies(specificStrategy)
        On Error GoTo ErrorHandler
        
        If strategy Is Nothing Then
            If Not m_Logger Is Nothing Then
                m_Logger.LogError "CSVImportEngine.ImportTransactionsFromFile", 0, _
                    "Specified strategy not found: " & specificStrategy
            End If
            Set ICSVImportEngine_ImportTransactionsFromFile = importedTransactions
            Exit Function
        End If
    Else
        ' Auto-detect strategy
        Set strategy = ICSVImportEngine_DetectStrategy(filePath)
        
        If strategy Is Nothing Then
            If Not m_Logger Is Nothing Then
                m_Logger.LogError "CSVImportEngine.ImportTransactionsFromFile", 0, _
                    "No suitable strategy found for: " & filePath
            End If
            Set ICSVImportEngine_ImportTransactionsFromFile = importedTransactions
            Exit Function
        End If
    End If
    
    ' Log strategy selection
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "CSVImportEngine.ImportTransactionsFromFile", _
            "Using strategy " & strategy.StrategyName & " for " & filePath
    End If
    
    ' Import transactions using the selected strategy
    Dim transactions As Collection
    Set transactions = strategy.ImportTransactions(filePath, owner, m_Logger)
    
    ' Add imported transactions to the result collection
    Dim trans As Transaction
    For Each trans In transactions
        importedTransactions.Add trans
    Next trans
    
    ' Add to repository if provided
    If Not m_Repository Is Nothing Then
        m_Repository.AddTransactions transactions
        m_Repository.SaveChanges
    End If
    
    ' Log import completion
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "CSVImportEngine.ImportTransactionsFromFile", _
            "Imported " & importedTransactions.Count & " transactions from " & filePath
    End If
    
    Set ICSVImportEngine_ImportTransactionsFromFile = importedTransactions
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "CSVImportEngine.ImportTransactionsFromFile", Err.Number, Err.Description
    End If
    Set ICSVImportEngine_ImportTransactionsFromFile = New Collection
End Function

' Public wrapper for ImportTransactionsFromFile
Public Function ImportTransactionsFromFile(ByVal filePath As String, ByVal owner As String, _
                                         Optional ByVal specificStrategy As String = "") As Collection
    Set ImportTransactionsFromFile = ICSVImportEngine_ImportTransactionsFromFile(filePath, owner, specificStrategy)
End Function

' Import transactions from all CSV files in a directory
Private Function ICSVImportEngine_ImportTransactionsFromDirectory(ByVal directoryPath As String, ByVal owner As String, _
                                                                Optional ByVal filePattern As String = "*.csv") As Collection
    On Error GoTo ErrorHandler
    
    ' Ensure engine is initialized
    If Not m_IsInitialized Then ICSVImportEngine_Initialize
    
    ' Create collection for imported transactions
    Dim importedTransactions As New Collection
    
    ' Log start of import
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "CSVImportEngine.ImportTransactionsFromDirectory", _
            "Starting import from directory: " & directoryPath
    End If
    
    ' Check if directory exists
    If Dir(directoryPath, vbDirectory) = "" Then
        If Not m_Logger Is Nothing Then
            m_Logger.LogError "CSVImportEngine.ImportTransactionsFromDirectory", 76, _
                "Directory not found: " & directoryPath
        End If
        Set ICSVImportEngine_ImportTransactionsFromDirectory = importedTransactions
        Exit Function
    End If
    
    ' Ensure directory path ends with a backslash
    If Right(directoryPath, 1) <> "\" Then
        directoryPath = directoryPath & "\"
    End If
    
    ' Find all matching files
    Dim filePath As String
    filePath = Dir(directoryPath & filePattern)
    
    ' Process each file
    Do While filePath <> ""
        ' Import transactions from this file
        Dim fileTransactions As Collection
        Set fileTransactions = ICSVImportEngine_ImportTransactionsFromFile(directoryPath & filePath, owner)
        
        ' Add imported transactions to the result collection
        Dim trans As Transaction
        For Each trans In fileTransactions
            importedTransactions.Add trans
        Next trans
        
        ' Get next file
        filePath = Dir()
    Loop
    
    ' Log import completion
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "CSVImportEngine.ImportTransactionsFromDirectory", _
            "Imported " & importedTransactions.Count & " transactions from directory: " & directoryPath
    End If
    
    Set ICSVImportEngine_ImportTransactionsFromDirectory = importedTransactions
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "CSVImportEngine.ImportTransactionsFromDirectory", Err.Number, Err.Description
    End If
    Set ICSVImportEngine_ImportTransactionsFromDirectory = New Collection
End Function

' Public wrapper for ImportTransactionsFromDirectory
Public Function ImportTransactionsFromDirectory(ByVal directoryPath As String, ByVal owner As String, _
                                              Optional ByVal filePattern As String = "*.csv") As Collection
    Set ImportTransactionsFromDirectory = ICSVImportEngine_ImportTransactionsFromDirectory(directoryPath, owner, filePattern)
End Function

' Auto-detect the appropriate strategy for a CSV file
Private Function ICSVImportEngine_DetectStrategy(ByVal filePath As String) As ICSVImportStrategy
    On Error GoTo ErrorHandler
    
    ' Ensure engine is initialized
    If Not m_IsInitialized Then ICSVImportEngine_Initialize
    
    ' Log detection attempt
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "CSVImportEngine.DetectStrategy", _
            "Detecting strategy for: " & filePath
    End If
    
    ' Try each strategy to find one that can process the file
    Dim strategy As ICSVImportStrategy
    For Each strategy In m_Strategies
        If strategy.CanProcessFile(filePath) Then
            ' Found a suitable strategy
            Set ICSVImportEngine_DetectStrategy = strategy
            
            ' Log detected strategy
            If Not m_Logger Is Nothing Then
                m_Logger.LogInfo "CSVImportEngine.DetectStrategy", _
                    "Detected strategy: " & strategy.StrategyName & " for " & filePath
            End If
            
            Exit Function
        End If
    Next strategy
    
    ' No suitable strategy found
    Set ICSVImportEngine_DetectStrategy = Nothing
    
    ' Log no strategy found
    If Not m_Logger Is Nothing Then
        m_Logger.LogWarning "CSVImportEngine.DetectStrategy", _
            "No suitable strategy found for: " & filePath
    End If
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "CSVImportEngine.DetectStrategy", Err.Number, Err.Description
    End If
    Set ICSVImportEngine_DetectStrategy = Nothing
End Function

' Public wrapper for DetectStrategy
Public Function DetectStrategy(ByVal filePath As String) As ICSVImportStrategy
    Set DetectStrategy = ICSVImportEngine_DetectStrategy(filePath)
End Function

'=========================================================================
' Properties
'=========================================================================

' Get/set the repository
Public Property Get Repository() As ITransactionRepository
    Set Repository = m_Repository
End Property

Public Property Set Repository(ByVal value As ITransactionRepository)
    Set m_Repository = value
End Property

' Get/set the logger
Public Property Get Logger() As IErrorLogger
    Set Logger = m_Logger
End Property

Public Property Set Logger(ByVal value As IErrorLogger)
    Set m_Logger = value
End Property

'=========================================================================
' Cleanup
'=========================================================================

Private Sub Class_Terminate()
    ' Clean up object references
    Set m_Strategies = Nothing
    Set m_Repository = Nothing
    Set m_Logger = Nothing
End Sub