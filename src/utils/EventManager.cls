' File: src/utils/EventManager.cls
'---------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' EventManager Class
'
' Purpose: Manages events and subscriptions throughout the application.
' This class implements a central event management system that allows
' components to publish and subscribe to events, enabling loose coupling
' and reactive UI updates.
'
' Design Decisions:
' - Uses PredeclaredId = True to enable Singleton pattern
' - Supports multiple subscribers per event
' - Handles typed event data for specific use cases
' - Provides comprehensive error handling
' - Allows for subscriber removal to prevent memory leaks

' Constants
Private Const MAX_SUBSCRIBERS As Integer = 100

' Type to hold subscriber information
Private Type Subscriber
    EventType As EventType
    Target As Object
    MethodName As String
    IsActive As Boolean
End Type

' Private member variables
Private m_Subscribers() As Subscriber
Private m_SubscriberCount As Integer
Private m_IsInitialized As Boolean
Private m_Logger As IErrorLogger

'=========================================================================
' Initialization
'=========================================================================

Private Sub Class_Initialize()
    ' Initialize the subscribers array
    ReDim m_Subscribers(1 To MAX_SUBSCRIBERS)
    m_SubscriberCount = 0
    m_IsInitialized = True
End Sub

' Initialize with logger
Public Sub Initialize(Optional ByVal logger As IErrorLogger = Nothing)
    On Error GoTo ErrorHandler
    
    ' Set logger if provided
    If Not logger Is Nothing Then
        Set m_Logger = logger
    End If
    
    ' Log initialization
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "EventManager.Initialize", "Event Manager initialized"
    End If
    
    m_IsInitialized = True
    
    Exit Sub
    
ErrorHandler:
    ' Can't use logger here if that's what caused the error
    If Not (m_Logger Is Nothing) And Err.Source <> "m_Logger" Then
        m_Logger.LogError "EventManager.Initialize", Err.Number, Err.Description
    End If
End Sub

'=========================================================================
' Event Subscription Methods
'=========================================================================

' Subscribe to an event
Public Function Subscribe(ByVal eventType As EventType, ByVal target As Object, ByVal methodName As String) As Integer
    On Error GoTo ErrorHandler
    
    ' Check if already at max subscribers
    If m_SubscriberCount >= MAX_SUBSCRIBERS Then
        If Not m_Logger Is Nothing Then
            m_Logger.LogWarning "EventManager.Subscribe", _
                "Maximum number of subscribers reached"
        End If
        Subscribe = -1
        Exit Function
    End If
    
    ' Check for duplicates (same event, target, and method)
    Dim i As Integer
    For i = 1 To m_SubscriberCount
        If m_Subscribers(i).EventType = eventType And _
           m_Subscribers(i).Target Is target And _
           m_Subscribers(i).MethodName = methodName Then
            
            ' If this is a re-subscription, just make it active again
            If Not m_Subscribers(i).IsActive Then
                m_Subscribers(i).IsActive = True
                Subscribe = i
                
                If Not m_Logger Is Nothing Then
                    m_Logger.LogInfo "EventManager.Subscribe", _
                        "Reactivated existing subscription for event type " & eventType
                End If
                
                Exit Function
            Else
                ' Already subscribed and active
                Subscribe = i
                Exit Function
            End If
        End If
    Next i
    
    ' Add new subscriber
    m_SubscriberCount = m_SubscriberCount + 1
    
    With m_Subscribers(m_SubscriberCount)
        .EventType = eventType
        Set .Target = target
        .MethodName = methodName
        .IsActive = True
    End With
    
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "EventManager.Subscribe", _
            "Added subscription for event type " & eventType
    End If
    
    Subscribe = m_SubscriberCount
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "EventManager.Subscribe", Err.Number, Err.Description
    End If
    Subscribe = -1
End Function

' Unsubscribe from an event
Public Function Unsubscribe(ByVal eventType As EventType, ByVal target As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim found As Boolean
    found = False
    
    ' Find all subscriptions for this target and event type
    Dim i As Integer
    For i = 1 To m_SubscriberCount
        If m_Subscribers(i).EventType = eventType And _
           m_Subscribers(i).Target Is target Then
            
            ' Mark as inactive
            m_Subscribers(i).IsActive = False
            found = True
        End If
    Next i
    
    If found And Not m_Logger Is Nothing Then
        m_Logger.LogInfo "EventManager.Unsubscribe", _
            "Removed subscription for event type " & eventType
    End If
    
    Unsubscribe = found
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "EventManager.Unsubscribe", Err.Number, Err.Description
    End If
    Unsubscribe = False
End Function

' Unsubscribe all events for a target
Public Function UnsubscribeAll(ByVal target As Object) As Integer
    On Error GoTo ErrorHandler
    
    Dim count As Integer
    count = 0
    
    ' Find all subscriptions for this target
    Dim i As Integer
    For i = 1 To m_SubscriberCount
        If m_Subscribers(i).Target Is target Then
            ' Mark as inactive
            m_Subscribers(i).IsActive = False
            count = count + 1
        End If
    Next i
    
    If count > 0 And Not m_Logger Is Nothing Then
        m_Logger.LogInfo "EventManager.UnsubscribeAll", _
            "Removed " & count & " subscriptions for target"
    End If
    
    UnsubscribeAll = count
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "EventManager.UnsubscribeAll", Err.Number, Err.Description
    End If
    UnsubscribeAll = 0
End Function

'=========================================================================
' Event Publishing Methods
'=========================================================================

' Publish an event with no data
Public Sub PublishEvent(ByVal eventType As EventType)
    On Error Resume Next
    
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "EventManager.PublishEvent", _
            "Publishing event type " & eventType
    End If
    
    ' Notify all subscribers
    Dim i As Integer
    For i = 1 To m_SubscriberCount
        If m_Subscribers(i).EventType = eventType And m_Subscribers(i).IsActive Then
            ' Call the target's method with no parameters
            CallByName m_Subscribers(i).Target, m_Subscribers(i).MethodName, VbMethod
            
            ' Log any errors
            If Err.Number <> 0 And Not m_Logger Is Nothing Then
                m_Logger.LogError "EventManager.PublishEvent", Err.Number, _
                    "Error calling handler for event " & eventType & ": " & Err.Description
                Err.Clear
            End If
        End If
    Next i
End Sub

' Publish an event with data parameter
Public Sub PublishEventWithData(ByVal eventType As EventType, ByVal data As Variant)
    On Error Resume Next
    
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "EventManager.PublishEventWithData", _
            "Publishing event type " & eventType & " with data"
    End If
    
    ' Notify all subscribers
    Dim i As Integer
    For i = 1 To m_SubscriberCount
        If m_Subscribers(i).EventType = eventType And m_Subscribers(i).IsActive Then
            ' Call the target's method with the data parameter
            CallByName m_Subscribers(i).Target, m_Subscribers(i).MethodName, VbMethod, data
            
            ' Log any errors
            If Err.Number <> 0 And Not m_Logger Is Nothing Then
                m_Logger.LogError "EventManager.PublishEventWithData", Err.Number, _
                    "Error calling handler for event " & eventType & ": " & Err.Description
                Err.Clear
            End If
        End If
    Next i
End Sub

'=========================================================================
' Helper Methods
'=========================================================================

' Get the number of active subscribers for an event
Public Function GetSubscriberCount(ByVal eventType As EventType) As Integer
    On Error GoTo ErrorHandler
    
    Dim count As Integer
    count = 0
    
    ' Count active subscribers for this event type
    Dim i As Integer
    For i = 1 To m_SubscriberCount
        If m_Subscribers(i).EventType = eventType And m_Subscribers(i).IsActive Then
            count = count + 1
        End If
    Next i
    
    GetSubscriberCount = count
    
    Exit Function
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "EventManager.GetSubscriberCount", Err.Number, Err.Description
    End If
    GetSubscriberCount = 0
End Function

' Clean up inactive subscribers
Public Sub Cleanup()
    On Error GoTo ErrorHandler
    
    ' Create a new array with only active subscribers
    Dim newSubscribers() As Subscriber
    ReDim newSubscribers(1 To MAX_SUBSCRIBERS)
    
    Dim newCount As Integer
    newCount = 0
    
    ' Copy active subscribers
    Dim i As Integer
    For i = 1 To m_SubscriberCount
        If m_Subscribers(i).IsActive Then
            newCount = newCount + 1
            newSubscribers(newCount) = m_Subscribers(i)
        End If
    Next i
    
    ' Replace the old array
    m_Subscribers = newSubscribers
    m_SubscriberCount = newCount
    
    If Not m_Logger Is Nothing Then
        m_Logger.LogInfo "EventManager.Cleanup", _
            "Cleaned up event subscriptions, now have " & newCount & " active subscribers"
    End If
    
    Exit Sub
    
ErrorHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.LogError "EventManager.Cleanup", Err.Number, Err.Description
    End If
End Sub

'=========================================================================
' Cleanup
'=========================================================================

Private Sub Class_Terminate()
    ' Clean up object references
    Dim i As Integer
    For i = 1 To m_SubscriberCount
        Set m_Subscribers(i).Target = Nothing
    Next i
    
    Set m_Logger = Nothing
End Sub
