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
Private Const DEFAULT_CSV_PATH As String = "C:\Finance\CSVs" ' Default Path Example
Private Const SETTINGS_SHEET_NAME As String = "Settings"
Private Const APP_TITLE As String = "BALANCE"
Private Const APP_SUBTITLE As String = "Bilateral Accounting Ledger for Analyzing Networked Couple Expenses"

' Color constants
Private Const COLOR_PRIMARY_RGB As String = "41,128,185"   ' Blue
Private Const COLOR_SECONDARY_RGB As String = "39,174,96"  ' Green
Private Const COLOR_ACCENT_RGB As String = "243,156,18"    ' Orange
Private Const COLOR_NEUTRAL_RGB As String = "236,240,241"  ' Light Gray
Private Const COLOR_DARK_RGB As String = "52,73,94"        ' Dark Blue-Gray
Private Const COLOR_SUCCESS_RGB As String = "46,204,113"   ' Green
Private Const COLOR_WARNING_RGB As String = "241,196,15"   ' Yellow
Private Const COLOR_DANGER_RGB As String = "231,76,60"    ' Red
Private Const COLOR_INFO_RGB As String = "52,152,219"     ' Light Blue
Private Const COLOR_TEXT_RGB As String = "44,62,80"        ' Dark Text
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
    SaveSettings ' Call SaveSettings to persist change
End Property

' User 2 Name
Public Property Get User2Name() As String
    If Not m_IsInitialized Then Initialize
    User2Name = m_User2Name
End Property

Public Property Let User2Name(value As String)
    m_User2Name = value
    SaveSettings ' Call SaveSettings to persist change
End Property

' CSV Folder Path
Public Property Get CSVFolderPath() As String
    If Not m_IsInitialized Then Initialize
    CSVFolderPath = m_CSVFolderPath
End Property

Public Property Let CSVFolderPath(value As String)
    m_CSVFolderPath = value
    SaveSettings ' Call SaveSettings to persist change
End Property

' Default Currency
Public Property Get DefaultCurrency() As String
    If Not m_IsInitialized Then Initialize
    DefaultCurrency = m_DefaultCurrency
End Property

Public Property Let DefaultCurrency(value As String)
    m_DefaultCurrency = value
    SaveSettings ' Call SaveSettings to persist change
End Property

' Last Update Date
Public Property Get LastUpdateDate() As Date
    If Not m_IsInitialized Then Initialize
    LastUpdateDate = m_LastUpdateDate
End Property

Public Property Let LastUpdateDate(value As Date)
    m_LastUpdateDate = value
    SaveSettings ' Call SaveSettings to persist change
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
    On Error Resume Next ' Use Resume Next carefully
    
    ' Initialize colors first
    InitializeColors
    
    ' Load settings
    LoadSettings
    
    m_IsInitialized = True
    
    On Error GoTo 0 ' Restore default error handling
End Sub

' Initialize color values
Private Sub InitializeColors()
    On Error Resume Next ' Handle potential errors in Split or RGB gracefully
    Dim rgbArr As Variant
    
    ' Primary color
    rgbArr = Split(COLOR_PRIMARY_RGB, ",")
    If UBound(rgbArr) = 2 Then m_ColorPrimary = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2))) Else m_ColorPrimary = RGB(0, 0, 255) ' Default Blue
    
    ' Secondary color
    rgbArr = Split(COLOR_SECONDARY_RGB, ",")
    If UBound(rgbArr) = 2 Then m_ColorSecondary = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2))) Else m_ColorSecondary = RGB(0, 255, 0) ' Default Green
    
    ' Accent color
    rgbArr = Split(COLOR_ACCENT_RGB, ",")
    If UBound(rgbArr) = 2 Then m_ColorAccent = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2))) Else m_ColorAccent = RGB(255, 165, 0) ' Default Orange
    
    ' Neutral color
    rgbArr = Split(COLOR_NEUTRAL_RGB, ",")
    If UBound(rgbArr) = 2 Then m_ColorNeutral = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2))) Else m_ColorNeutral = RGB(211, 211, 211) ' Default Light Gray
    
    ' Dark color
    rgbArr = Split(COLOR_DARK_RGB, ",")
    If UBound(rgbArr) = 2 Then m_ColorDark = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2))) Else m_ColorDark = RGB(0, 0, 128) ' Default Navy
    
    ' Success color
    rgbArr = Split(COLOR_SUCCESS_RGB, ",")
    If UBound(rgbArr) = 2 Then m_ColorSuccess = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2))) Else m_ColorSuccess = RGB(0, 128, 0) ' Default Green
    
    ' Warning color
    rgbArr = Split(COLOR_WARNING_RGB, ",")
    If UBound(rgbArr) = 2 Then m_ColorWarning = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2))) Else m_ColorWarning = RGB(255, 255, 0) ' Default Yellow
    
    ' Danger color
    rgbArr = Split(COLOR_DANGER_RGB, ",")
    If UBound(rgbArr) = 2 Then m_ColorDanger = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2))) Else m_ColorDanger = RGB(255, 0, 0) ' Default Red
    
    ' Info color
    rgbArr = Split(COLOR_INFO_RGB, ",")
    If UBound(rgbArr) = 2 Then m_ColorInfo = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2))) Else m_ColorInfo = RGB(0, 191, 255) ' Default Deep Sky Blue
    
    ' Text color
    rgbArr = Split(COLOR_TEXT_RGB, ",")
    If UBound(rgbArr) = 2 Then m_ColorText = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2))) Else m_ColorText = RGB(0, 0, 0) ' Default Black
    
    ' Light text color
    rgbArr = Split(COLOR_LIGHT_TEXT_RGB, ",")
    If UBound(rgbArr) = 2 Then m_ColorLightText = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2))) Else m_ColorLightText = RGB(255, 255, 255) ' Default White
    
    On Error GoTo 0 ' Restore default error handling
End Sub


' Load settings from sheet
Public Sub LoadSettings()
    On Error Resume Next ' Use Resume Next to handle cases where sheet or settings don't exist
    
    Dim ws As Worksheet
    
    ' Set default values first
    m_User1Name = "User 1"
    m_User2Name = "User 2"
    m_CSVFolderPath = DEFAULT_CSV_PATH
    m_DefaultCurrency = "$"
    m_LastUpdateDate = Now
    
    ' Try to load from settings sheet if it exists
    ' ***** CORRECTED LINE BELOW *****
    If Utilities.SheetExists(SETTINGS_SHEET_NAME) Then
        Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
        
        ' Read settings with error handling (Resume Next is active)
        m_User1Name = CStr(ws.Range("B3").Value)
        m_User2Name = CStr(ws.Range("B4").Value)
        m_CSVFolderPath = CStr(ws.Range("B2").Value)
        m_DefaultCurrency = CStr(ws.Range("B5").Value)
        m_LastUpdateDate = CDate(ws.Range("B6").Value)

        ' Reset to defaults if loaded values are empty or invalid
        If Len(Trim$(m_User1Name)) = 0 Or Err.Number <> 0 Then m_User1Name = "User 1": Err.Clear
        If Len(Trim$(m_User2Name)) = 0 Or Err.Number <> 0 Then m_User2Name = "User 2": Err.Clear
        If Len(Trim$(m_CSVFolderPath)) = 0 Or Err.Number <> 0 Then m_CSVFolderPath = DEFAULT_CSV_PATH: Err.Clear
        If Len(Trim$(m_DefaultCurrency)) = 0 Or Err.Number <> 0 Then m_DefaultCurrency = "$": Err.Clear
        If Err.Number <> 0 Then m_LastUpdateDate = Now: Err.Clear ' Check specifically for date error
                
        Set ws = Nothing
    Else
        ' If settings sheet doesn't exist, ensure defaults are set (already done above)
        ' Optionally, create the sheet now with defaults
        ' CreateSettingsSheet ' Uncomment if you want to create sheet on first load
    End If
    
    On Error GoTo 0 ' Restore default error handling
End Sub


' Save settings to sheet
Public Sub SaveSettings()
    On Error Resume Next ' Use Resume Next carefully
    
    Dim ws As Worksheet
    
    ' Create settings sheet if it doesn't exist
    ' ***** CORRECTED LINE BELOW *****
    If Not Utilities.SheetExists(SETTINGS_SHEET_NAME) Then
        CreateSettingsSheet
        ' Check if CreateSettingsSheet failed
        If Err.Number <> 0 Then
            Debug.Print "Error creating settings sheet in SaveSettings: " & Err.Description
            Exit Sub
        End If
    End If
    
    ' Get the sheet (it should exist now)
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    If ws Is Nothing Then
         Debug.Print "Failed to get settings sheet in SaveSettings."
         Exit Sub
    End If

    ' Save all settings (disable events briefly for performance)
    Application.EnableEvents = False
    ws.Range("B2").Value = m_CSVFolderPath
    ws.Range("B3").Value = m_User1Name
    ws.Range("B4").Value = m_User2Name
    ws.Range("B5").Value = m_DefaultCurrency
    ws.Range("B6").Value = Now ' Always save current time on save
    
    ' Format date
    ws.Range("B6").NumberFormat = "mm/dd/yyyy hh:mm:ss"
    Application.EnableEvents = True
    
    Set ws = Nothing
    
    ' Publish settings changed event
    EventManager.PublishEvent EventType.SettingsChanged
    
    On Error GoTo 0 ' Restore default error handling
End Sub


' Create settings sheet with default values
Private Sub CreateSettingsSheet()
    On Error GoTo ErrorHandler ' Use GoTo for specific error handling here
    
    Dim ws As Worksheet
    
    ' Add the sheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)) ' Add as last sheet
    
    If ws Is Nothing Then Err.Raise vbObjectError + 513, "AppSettings.CreateSettingsSheet", "Failed to add worksheet."

    ' Rename the sheet
    On Error Resume Next ' Handle potential error if name already exists (shouldn't happen if SheetExists was checked)
    ws.Name = SETTINGS_SHEET_NAME
    If Err.Number <> 0 Then
        Application.DisplayAlerts = False
        ws.Delete ' Clean up partially created sheet if rename fails
        Application.DisplayAlerts = True
        Err.Raise vbObjectError + 514, "AppSettings.CreateSettingsSheet", "Failed to rename worksheet to '" & SETTINGS_SHEET_NAME & "'. It might already exist or be invalid."
    End If
    On Error GoTo ErrorHandler ' Restore error handling

    ' Setup header
    With ws.Range("A1:B1")
        .Value = Array("Setting", "Value")
        .Font.Bold = True
        .Interior.Color = m_ColorPrimary ' Assumes colors are initialized
        .Font.Color = m_ColorLightText
    End With
    
    ' Settings rows
    ws.Range("A2").Value = "CSV Folder Path"
    ws.Range("B2").Value = m_CSVFolderPath
    
    ws.Range("A3").Value = "User 1 Name"
    ws.Range("B3").Value = m_User1Name
    
    ws.Range("A4").Value = "User 2 Name"
    ws.Range("B4").Value = m_User2Name
    
    ws.Range("A5").Value = "Default Currency"
    ws.Range("B5").Value = m_DefaultCurrency
    
    ws.Range("A6").Value = "Last Update"
    ws.Range("B6").Value = Now
    ws.Range("B6").NumberFormat = "mm/dd/yyyy hh:mm:ss"
    
    ' Add user interface section (Example - adapt as needed)
    ws.Range("A8").Value = "User Interface Settings"
    ws.Range("A8").Font.Bold = True
    ws.Range("A8").Interior.Color = m_ColorNeutral
    
    ws.Range("A9").Value = "Auto-refresh on Open"
    ws.Range("B9").Value = False ' Default value
    
    ' Format
    ws.Range("A:A").Font.Bold = True
    ws.Columns("A:B").AutoFit
    
    ' Add borders
    With ws.Range("A1:B9").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(191, 191, 191)
    End With

    ' Hide the sheet
    ws.Visible = xlSheetVeryHidden
    
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "Error in AppSettings.CreateSettingsSheet: " & Err.Number & " - " & Err.Description
    ' Attempt to clean up sheet if partially created
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        On Error Resume Next ' Ignore error if delete fails
        ws.Delete
        Application.DisplayAlerts = True
    End If
    Set ws = Nothing
    ' Optionally re-raise the error if needed by caller
    ' Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

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