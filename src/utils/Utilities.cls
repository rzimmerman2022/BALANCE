VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Utilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' BALANCE - Bilateral Accounting Ledger for Analyzing Networked Couple Expenses
' Utilities Class - Static utility functions used throughout the application

' ===== Sheet Management =====

' Check if sheet exists
Public Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    SheetExists = Not ws Is Nothing
    Set ws = Nothing
End Function

' Get or create a sheet
Public Function GetOrCreateSheet(sheetName As String, Optional visible As Boolean = True) As Worksheet
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    ' Check if sheet exists
    If SheetExists(sheetName) Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
    Else
        ' Create new sheet
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets.Add
        
        If Not ws Is Nothing Then
            ' Rename the sheet
            On Error Resume Next
            ws.Name = sheetName
            On Error GoTo ErrorHandler
        End If
        
        On Error GoTo ErrorHandler
    End If
    
    ' Set visibility
    If Not ws Is Nothing Then
        On Error Resume Next
        If visible Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetVeryHidden
        End If
        On Error GoTo ErrorHandler
    End If
    
    Set GetOrCreateSheet = ws
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in Utilities.GetOrCreateSheet: " & Err.Description
    ' Don't raise an error - just return Nothing if we can't create the sheet
    Set GetOrCreateSheet = Nothing
End Function

' Get the last used row in a column
Public Function GetLastRow(ws As Worksheet, col As Long) As Long
    On Error Resume Next
    
    If ws Is Nothing Then
        GetLastRow = 0
        Exit Function
    End If
    
    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    
    If Err.Number <> 0 Then
        GetLastRow = 0
    End If
    
    On Error GoTo 0
End Function

' Get the last used column in a row
Public Function GetLastColumn(ws As Worksheet, row As Long) As Long
    On Error Resume Next
    
    If ws Is Nothing Then
        GetLastColumn = 0
        Exit Function
    End If
    
    GetLastColumn = ws.Cells(row, ws.Columns.Count).End(xlToLeft).Column
    
    If Err.Number <> 0 Then
        GetLastColumn = 0
    End If
    
    On Error GoTo 0
End Function

' Clear a sheet's contents while preserving the sheet
Public Sub ClearSheet(ws As Worksheet)
    On Error Resume Next
    
    If Not ws Is Nothing Then
        ws.Cells.Clear
        
        ' Also clear any shapes
        Dim shp As Shape
        For Each shp In ws.Shapes
            shp.Delete
        Next shp
    End If
    
    On Error GoTo 0
End Sub

' ===== File System Utilities =====

' Check if a folder exists
Public Function FolderExists(folderPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(folderPath)
    Set fso = Nothing
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in Utilities.FolderExists: " & Err.Description
    FolderExists = False
End Function

' Create folder if it doesn't exist
Public Function EnsureFolderExists(folderPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    ' Check if successful
    EnsureFolderExists = fso.FolderExists(folderPath)
    
    Set fso = Nothing
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in Utilities.EnsureFolderExists: " & Err.Description
    EnsureFolderExists = False
End Function

' Browse for folder dialog
Public Function BrowseForFolder(Optional prompt As String = "Select a Folder") As String
    On Error GoTo ErrorHandler
    
    Dim shellApp As Object
    Dim folder As Object
    
    Set shellApp = CreateObject("Shell.Application")
    Set folder = shellApp.BrowseForFolder(0, prompt, 0)
    
    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.Path
    Else
        BrowseForFolder = ""
    End If
    
    Set folder = Nothing
    Set shellApp = Nothing
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in Utilities.BrowseForFolder: " & Err.Description
    BrowseForFolder = ""
End Function

' Extract date from filename
Public Function ExtractDateFromFilename(fileName As String) As String
    On Error Resume Next
    
    Dim parts As Variant
    
    ' Expected format: Owner_Source_YYYYMMDD.csv
    parts = Split(Replace(fileName, ".csv", ""), "_")
    
    ' Check if we have enough parts
    If UBound(parts) >= 2 Then
        Dim datePart As String
        datePart = parts(UBound(parts))
        
        ' Check if it's a date in YYYYMMDD format
        If Len(datePart) = 8 And IsNumeric(datePart) Then
            ExtractDateFromFilename = datePart
            Exit Function
        End If
    End If
    
    ' Couldn't find date
    ExtractDateFromFilename = ""
    
    On Error GoTo 0
End Function

' ===== String Utilities =====

' Safely convert variant to string with default value
Public Function SafeString(value As Variant, Optional defaultValue As String = "") As String
    On Error Resume Next
    
    If IsNull(value) Or IsEmpty(value) Then
        SafeString = defaultValue
    Else
        SafeString = CStr(value)
    End If
    
    On Error GoTo 0
End Function

' Clean amount string
Public Function CleanAmountString(amountStr As String) As String
    On Error Resume Next
    
    Dim result As String
    result = amountStr
    
    ' Remove currency symbols and commas
    result = Replace(result, "$", "")
    result = Replace(result, "€", "")
    result = Replace(result, "£", "")
    result = Replace(result, "¥", "")
    result = Replace(result, ",", "")
    result = Replace(result, " ", "")
    
    ' Handle European format with comma as decimal
    If InStr(result, ".") = 0 And InStr(result, ",") > 0 Then
        result = Replace(result, ",", ".")
    End If
    
    ' Handle parentheses for negative numbers
    If Left(result, 1) = "(" And Right(result, 1) = ")" Then
        result = "-" & Mid(result, 2, Len(result) - 2)
    End If
    
    CleanAmountString = result
    
    On Error GoTo 0
End Function

' ===== UI Utilities =====

' Apply modern formatting to a worksheet
Public Sub ApplyModernFormatting(ws As Worksheet)
    On Error Resume Next
    
    If ws Is Nothing Then Exit Sub
    
    ' Set modern font for entire sheet
    ws.Cells.Font.Name = "Segoe UI"
    
    ' Remove gridlines for cleaner look
    ws.DisplayGridlines = False
    
    ' Adjust row heights for better spacing
    ws.Rows("1:1").RowHeight = 24
    ws.StandardHeight = 20
    
    ' Autofit columns for best display
    ws.Columns.AutoFit
    
    On Error GoTo 0
End Sub

' Create a modern card UI element
Public Function CreateCard(ws As Worksheet, topLeftCell As String, bottomRightCell As String, _
                           title As String, value As Variant, Optional valueFormat As String = "", _
                           Optional colorTheme As Long = -1) As Shape
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then Exit Function
    
    If colorTheme = -1 Then colorTheme = AppSettings.ColorPrimary
    
    ' Create background shape
    Dim cardLeft As Double, cardTop As Double
    Dim cardWidth As Double, cardHeight As Double
    Dim cardShape As Shape
    Dim headerShape As Shape
    Dim valueCell As Range
    
    cardLeft = ws.Range(topLeftCell).Left
    cardTop = ws.Range(topLeftCell).Top
    cardWidth = ws.Range(topLeftCell & ":" & bottomRightCell).Width
    cardHeight = ws.Range(topLeftCell & ":" & bottomRightCell).Height
    
    Set cardShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, cardLeft, cardTop, cardWidth, cardHeight)
    
    If Not cardShape Is Nothing Then
        With cardShape
            .Fill.ForeColor.RGB = RGB(255, 255, 255) ' White background
            .Line.ForeColor.RGB = colorTheme
            .Line.Weight = 1.5
            
            ' Add shadow
            .Shadow.Type = msoShadow21
            .Shadow.Blur = 5
            .Shadow.OffsetX = 3
            .Shadow.OffsetY = 3
            .Shadow.Transparency = 0.7
            
            .Name = "Card_" & Replace(topLeftCell, "$", "")
        End With
    End If
    
    ' Add card header with theme color
    Set headerShape = ws.Shapes.AddShape(msoShapeRectangle, cardLeft, cardTop, cardWidth, 25)
    
    If Not headerShape Is Nothing Then
        With headerShape
            .Fill.ForeColor.RGB = colorTheme
            .Line.Visible = msoFalse
            .Name = "CardHeader_" & Replace(topLeftCell, "$", "")
        End With
    End If
    
    ' Add title text
    ws.Range(topLeftCell).Value = title
    ws.Range(topLeftCell).Font.Size = 10
    ws.Range(topLeftCell).Font.Bold = True
    ws.Range(topLeftCell).Font.Color = RGB(255, 255, 255) ' White text
    ws.Range(topLeftCell).Font.Name = "Segoe UI"
    ws.Range(topLeftCell).HorizontalAlignment = xlCenter
    
    ' Add value in large font
    Set valueCell = ws.Range(topLeftCell).Offset(2, 0)
    
    valueCell.Value = value
    valueCell.Font.Size = 20
    valueCell.Font.Bold = True
    valueCell.Font.Color = colorTheme
    valueCell.Font.Name = "Segoe UI"
    valueCell.HorizontalAlignment = xlCenter
    
    If Len(valueFormat) > 0 Then
        valueCell.NumberFormat = valueFormat
    End If
    
    Set CreateCard = cardShape
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in Utilities.CreateCard: " & Err.Description
    Set CreateCard = Nothing
End Function

' Add a formatted button to a worksheet
Public Function AddButton(ws As Worksheet, left As Double, top As Double, _
                         width As Double, height As Double, caption As String, _
                         macroName As String, Optional colorTheme As Long = -1) As Shape
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then Exit Function
    
    If colorTheme = -1 Then colorTheme = AppSettings.ColorPrimary
    
    Dim btn As Shape
    
    ' Create button shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, height)
    
    If Not btn Is Nothing Then
        With btn
            .Fill.ForeColor.RGB = colorTheme
            .Line.Visible = msoFalse
            .Name = "Button_" & Replace(caption, " ", "_")
            
            ' Add caption
            .TextFrame.Characters.Text = caption
            .TextFrame.Characters.Font.Name = "Segoe UI"
            .TextFrame.Characters.Font.Size = 12
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .TextFrame.HorizontalAlignment = msoTextHAlignCenter
            .TextFrame.VerticalAlignment = msoTextVAlignCenter
            
            ' Assign macro
            .OnAction = macroName
        End With
    End If
    
    Set AddButton = btn
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in Utilities.AddButton: " & Err.Description
    Set AddButton = Nothing
End Function

' ===== Date Utilities =====

' Get first day of month
Public Function FirstDayOfMonth(dateValue As Date) As Date
    FirstDayOfMonth = DateSerial(Year(dateValue), Month(dateValue), 1)
End Function

' Get last day of month
Public Function LastDayOfMonth(dateValue As Date) As Date
    LastDayOfMonth = DateSerial(Year(dateValue), Month(dateValue) + 1, 0)
End Function

' Format month for display
Public Function FormatMonth(dateValue As Date) As String
    FormatMonth = MonthName(Month(dateValue)) & " " & Year(dateValue)
End Function

' ===== Error Handling =====

' Log error to debug window (Renamed to avoid conflict with IErrorLogger)
Public Sub DebugLogError(functionName As String, errorNumber As Long, errorDescription As String, Optional lineNumber As Long = 0)
    Dim errorMsg As String

    If lineNumber > 0 Then
        errorMsg = "DEBUG ERROR in " & functionName & " (line " & lineNumber & "): " & errorNumber & " - " & errorDescription
    Else
        errorMsg = "DEBUG ERROR in " & functionName & ": " & errorNumber & " - " & errorDescription
    End If
    
    Debug.Print errorMsg
End Sub

' Display error message box
Public Sub ShowError(message As String, Optional title As String = "Error")
    MsgBox message, vbExclamation, IIf(Len(title) > 0, title, AppSettings.AppTitle & " - Error")
End Sub

' ===== Debugging Utilities =====

' Print object properties to Immediate window
Public Sub DebugPrintObject(obj As Object, Optional title As String = "Object Properties")
    On Error Resume Next
    
    Dim prop As Variant
    
    Debug.Print "=== " & title & " ==="
    
    For Each prop In obj
        Debug.Print prop & " = " & obj(prop)
    Next prop
    
    Debug.Print "===================="
    
    On Error GoTo 0
End Sub

' Class initialize
Private Sub Class_Initialize()
    ' Nothing needed here as we're using PredeclaredId = True
End Sub

' Class terminate
Private Sub Class_Terminate()
    ' Nothing to clean up
End Sub
