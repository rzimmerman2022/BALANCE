VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' Application Settings Class - Singleton pattern for application settings

' Constants
Private Const DEFAULT_CSV_PATH As String = "C:\Finance\CSVs"
Private Const SETTINGS_SHEET_NAME As String = "Settings"
Private Const APP_TITLE As String = "BALANCE"
Private Const APP_SUBTITLE As String = "Bilateral Accounting Ledger for Analyzing Networked Couple Expenses"

' Color constants
Private Const COLOR_PRIMARY_RGB As String = "41,128,185"    ' Blue
Private Const COLOR_SECONDARY_RGB As String = "39,174,96"   ' Green
Private Const COLOR_ACCENT_RGB As String = "243,156,18"     ' Orange
Private Const COLOR_NEUTRAL_RGB As String = "236,240,241"   ' Light Gray
Private Const COLOR_DARK_RGB As String = "52,73,94"         ' Dark Blue-Gray
Private Const COLOR_SUCCESS_RGB As String = "46,204,113"    ' Green
Private Const COLOR_WARNING_RGB As String = "241,196,15"    ' Yellow
Private Const COLOR_DANGER_RGB As String = "231,76,60"      ' Red
Private Const COLOR_INFO_RGB As String = "52,152,219"       ' Light Blue
Private Const COLOR_TEXT_RGB As String = "44,62,80"         ' Dark Text
Private Const COLOR_LIGHT_TEXT_RGB As String = "236,240,241" ' Light Text

' Private variables
Private m_User1Name As String
Private m_User2Name As String
Private m_CSVFolderPath As String
Private m_DefaultCurrency As String
Private m_LastUpdateDate As Date
Private m_IsInitialized As Boolean

' Color properties
Private m_ColorPrimary As Long
Private m_ColorSecondary As Long
Private m_ColorAccent As Long
Private m_ColorNeutral As Long
Private m_ColorDark As Long
Private m_ColorSuccess As Long
Private m_ColorWarning As Long
Private m_ColorDanger As Long
Private m_ColorInfo As Long
Private m_ColorText As Long
Private m_ColorLightText As Long

' ===== Properties =====

' Application title
Public Property Get AppTitle() As String
    AppTitle = APP_TITLE
End Property

' Application subtitle
Public Property Get AppSubtitle() As String
    AppSubtitle = APP_SUBTITLE
End Property

' User 1 Name
Public Property Get User1Name() As String
    If Not m_IsInitialized Then Initialize
    User1Name = m_User1Name
End Property

Public Property Let User1Name(value As String)
    m_User1Name = value
    SaveSettings
End Property

' User 2 Name
Public Property Get User2Name() As String
    If Not m_IsInitialized Then Initialize
    User2Name = m_User2Name
End Property

Public Property Let User2Name(value As String)
    m_User2Name = value
    SaveSettings
End Property

' CSV Folder Path
Public Property Get CSVFolderPath() As String
    If Not m_IsInitialized Then Initialize
    CSVFolderPath = m_CSVFolderPath
End Property

Public Property Let CSVFolderPath(value As String)
    m_CSVFolderPath = value
    SaveSettings
End Property

' Default Currency
Public Property Get DefaultCurrency() As String
    If Not m_IsInitialized Then Initialize
    DefaultCurrency = m_DefaultCurrency
End Property

Public Property Let DefaultCurrency(value As String)
    m_DefaultCurrency = value
    SaveSettings
End Property

' Last Update Date
Public Property Get LastUpdateDate() As Date
    If Not m_IsInitialized Then Initialize
    LastUpdateDate = m_LastUpdateDate
End Property

Public Property Let LastUpdateDate(value As Date)
    m_LastUpdateDate = value
    SaveSettings
End Property

' Settings Sheet Name
Public Property Get SettingsSheetName() As String
    SettingsSheetName = SETTINGS_SHEET_NAME
End Property

' ===== Color Properties =====
Public Property Get ColorPrimary() As Long
    If Not m_IsInitialized Then Initialize
    ColorPrimary = m_ColorPrimary
End Property

Public Property Get ColorSecondary() As Long
    If Not m_IsInitialized Then Initialize
    ColorSecondary = m_ColorSecondary
End Property

Public Property Get ColorAccent() As Long
    If Not m_IsInitialized Then Initialize
    ColorAccent = m_ColorAccent
End Property

Public Property Get ColorNeutral() As Long
    If Not m_IsInitialized Then Initialize
    ColorNeutral = m_ColorNeutral
End Property

Public Property Get ColorDark() As Long
    If Not m_IsInitialized Then Initialize
    ColorDark = m_ColorDark
End Property

Public Property Get ColorSuccess() As Long
    If Not m_IsInitialized Then Initialize
    ColorSuccess = m_ColorSuccess
End Property

Public Property Get ColorWarning() As Long
    If Not m_IsInitialized Then Initialize
    ColorWarning = m_ColorWarning
End Property

Public Property Get ColorDanger() As Long
    If Not m_IsInitialized Then Initialize
    ColorDanger = m_ColorDanger
End Property

Public Property Get ColorInfo() As Long
    If Not m_IsInitialized Then Initialize
    ColorInfo = m_ColorInfo
End Property

Public Property Get ColorText() As Long
    If Not m_IsInitialized Then Initialize
    ColorText = m_ColorText
End Property

Public Property Get ColorLightText() As Long
    If Not m_IsInitialized Then Initialize
    ColorLightText = m_ColorLightText
End Property

' ===== Methods =====
' Initialize settings
Public Sub Initialize()
    On Error Resume Next
    
    ' Initialize colors first
    InitializeColors
    
    ' Load settings
    LoadSettings
    
    m_IsInitialized = True
    
    On Error GoTo 0
End Sub

' Initialize color values
Private Sub InitializeColors()
    Dim rgbArr As Variant
    
    ' Primary color
    rgbArr = Split(COLOR_PRIMARY_RGB, ",")
    m_ColorPrimary = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2)))
    
    ' Secondary color
    rgbArr = Split(COLOR_SECONDARY_RGB, ",")
    m_ColorSecondary = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2)))
    
    ' Accent color
    rgbArr = Split(COLOR_ACCENT_RGB, ",")
    m_ColorAccent = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2)))
    
    ' Neutral color
    rgbArr = Split(COLOR_NEUTRAL_RGB, ",")
    m_ColorNeutral = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2)))
    
    ' Dark color
    rgbArr = Split(COLOR_DARK_RGB, ",")
    m_ColorDark = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2)))
    
    ' Success color
    rgbArr = Split(COLOR_SUCCESS_RGB, ",")
    m_ColorSuccess = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2)))
    
    ' Warning color
    rgbArr = Split(COLOR_WARNING_RGB, ",")
    m_ColorWarning = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2)))
    
    ' Danger color
    rgbArr = Split(COLOR_DANGER_RGB, ",")
    m_ColorDanger = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2)))
    
    ' Info color
    rgbArr = Split(COLOR_INFO_RGB, ",")
    m_ColorInfo = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2)))
    
    ' Text color
    rgbArr = Split(COLOR_TEXT_RGB, ",")
    m_ColorText = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2)))
    
    ' Light text color
    rgbArr = Split(COLOR_LIGHT_TEXT_RGB, ",")
    m_ColorLightText = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2)))
End Sub

' Load settings from sheet
Public Sub LoadSettings()
    On Error Resume Next
    
    Dim ws As Worksheet
    
    ' Set default values
    m_User1Name = "User 1"
    m_User2Name = "User 2"
    m_CSVFolderPath = DEFAULT_CSV_PATH
    m_DefaultCurrency = "$"
    m_LastUpdateDate = Now
    
    ' Try to load from settings sheet if it exists
    If SheetExists(SETTINGS_SHEET_NAME) Then
        Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
        
        ' Read settings with error handling
        On Error Resume Next
        
        ' User names
        If Not IsEmpty(ws.Range("B3").value) Then
            m_User1Name = ws.Range("B3").value
        End If
        
        If Not IsEmpty(ws.Range("B4").value) Then
            m_User2Name = ws.Range("B4").value
        End If
        
        ' CSV folder path
        If Not IsEmpty(ws.Range("B2").value) Then
            m_CSVFolderPath = ws.Range("B2").value
        End If
        
        ' Default currency (if exists)
        If Not IsEmpty(ws.Range("B5").value) Then
            m_DefaultCurrency = ws.Range("B5").value
        End If
        
        ' Last update date (if exists)
        If Not IsEmpty(ws.Range("B6").value) Then
            If IsDate(ws.Range("B6").value) Then
                m_LastUpdateDate = ws.Range("B6").value
            End If
        End If
        
        Set ws = Nothing
        On Error GoTo 0
    End If
    
    On Error GoTo 0
End Sub

' Save settings to sheet
Public Sub SaveSettings()
    On Error Resume Next
    
    Dim ws As Worksheet
    
    ' Create settings sheet if it doesn't exist
    If Not SheetExists(SETTINGS_SHEET_NAME) Then
        CreateSettingsSheet
    End If
    
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    
    ' Save all settings
    ws.Range("B2").value = m_CSVFolderPath
    ws.Range("B3").value = m_User1Name
    ws.Range("B4").value = m_User2Name
    ws.Range("B5").value = m_DefaultCurrency
    ws.Range("B6").value = Now
    
    ' Format date
    ws.Range("B6").NumberFormat = "mm/dd/yyyy hh:mm:ss"
    
    Set ws = Nothing
    
    On Error GoTo 0
End Sub

' Create settings sheet with default values
Private Sub CreateSettingsSheet()
    On Error Resume Next
    
    Dim ws As Worksheet
    
    ' Add the sheet
    Set ws = ThisWorkbook.Worksheets.Add
    
    If Not ws Is Nothing Then
        ' Rename the sheet
        ws.Name = SETTINGS_SHEET_NAME
        
        ' Setup header
        ws.Range("A1").value = "Setting"
        ws.Range("B1").value = "Value"
        ws.Range("A1:B1").Font.Bold = True
        ws.Range("A1:B1").Interior.Color = m_ColorPrimary
        ws.Range("A1:B1").Font.Color = m_ColorLightText
        
        ' Settings rows
        ws.Range("A2").value = "CSV Folder Path"
        ws.Range("B2").value = m_CSVFolderPath
        
        ws.Range("A3").value = "User 1 Name"
        ws.Range("B3").value = m_User1Name
        
        ws.Range("A4").value = "User 2 Name"
        ws.Range("B4").value = m_User2Name
        
        ws.Range("A5").value = "Default Currency"
        ws.Range("B5").value = m_DefaultCurrency
        
        ws.Range("A6").value = "Last Update"
        ws.Range("B6").value = Now
        ws.Range("B6").NumberFormat = "mm/dd/yyyy hh:mm:ss"
        
        ' Add user interface section
        ws.Range("A8").value = "User Interface Settings"
        ws.Range("A8").Font.Bold = True
        ws.Range("A8").Interior.Color = m_ColorNeutral
        
        ws.Range("A9").value = "Auto-refresh on Open"
        ws.Range("B9").value = False
        
        ' Format
        ws.Range("A:A").Font.Bold = True
        ws.Columns("A:B").AutoFit
        
        ' Add borders
        ws.Range("A1:B9").Borders.LineStyle = xlContinuous
        ws.Range("A1:B9").Borders.Weight = xlThin
        ws.Range("A1:B9").Borders.Color = RGB(191, 191, 191)
    End If
    
    Set ws = Nothing
    
    On Error GoTo 0
End Sub

' Check if sheet exists
Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    SheetExists = Not ws Is Nothing
    Set ws = Nothing
End Function

' Class initialize event
Private Sub Class_Initialize()
    ' Nothing needed here as we're using PredeclaredId = True
    ' and explicit initialization via Initialize method
    m_IsInitialized = False
End Sub

' Class terminate event
Private Sub Class_Terminate()
    ' Clean up to prevent memory leaks
    ' Nothing needed for this class
End Sub