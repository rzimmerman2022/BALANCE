' File: src/utils/OperationResult.cls
'----------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OperationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' OperationResult Class
'
' Purpose: Represents the result of an operation that can succeed or fail.
' This class provides a structured way to handle errors and return data,
' improving error handling throughout the application.

' Private member variables
Private m_Success As Boolean
Private m_ErrorCode As Long
Private m_ErrorMessage As String
Private m_Data As Variant

' ===== Initialization =====

Private Sub Class_Initialize()
    ' Default to success with no data
    m_Success = True
    m_ErrorCode = 0
    m_ErrorMessage = ""
End Sub

' Initialize with success
Public Sub InitSuccess(Optional ByVal data As Variant = Null)
    m_Success = True
    m_ErrorCode = 0
    m_ErrorMessage = ""
    
    If Not IsNull(data) Then
        If IsObject(data) Then
            Set m_Data = data
        Else
            m_Data = data
        End If
    End If
End Sub

' Initialize with error
Public Sub InitError(ByVal errorCode As Long, ByVal errorMessage As String)
    m_Success = False
    m_ErrorCode = errorCode
    m_ErrorMessage = errorMessage
End Sub

' ===== Properties =====

' Success flag
Public Property Get Success() As Boolean
    Success = m_Success
End Property

' Error code
Public Property Get ErrorCode() As Long
    ErrorCode = m_ErrorCode
End Property

' Error message
Public Property Get ErrorMessage() As String
    ErrorMessage = m_ErrorMessage
End Property

' Result data
Public Property Get Data() As Variant
    If IsObject(m_Data) Then
        Set Data = m_Data
    Else
        Data = m_Data
    End If
End Property

' ===== Methods =====

' Check if the operation succeeded
Public Function Succeeded() As Boolean
    Succeeded = m_Success
End Function

' Check if the operation failed
Public Function Failed() As Boolean
    Failed = Not m_Success
End Function

' Get a string representation of the result
Public Function ToString() As String
    If m_Success Then
        ToString = "Success"
    Else
        ToString = "Error " & m_ErrorCode & ": " & m_ErrorMessage
    End If
End Function

' ===== Static Factory Methods =====

' Create a success result
Public Function CreateSuccess(Optional ByVal data As Variant = Null) As OperationResult
    Dim result As New OperationResult
    result.InitSuccess data
    Set CreateSuccess = result
End Function

' Create an error result
Public Function CreateError(ByVal errorCode As Long, ByVal errorMessage As String) As OperationResult
    Dim result As New OperationResult
    result.InitError errorCode, errorMessage
    Set CreateError = result
End Function
