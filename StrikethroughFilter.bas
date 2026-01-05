
Option Explicit

Private Const HELPER_MARKER As String = "___STRIKE_FILTER_HELPER___"
Private Const FLAG_KEEP As String = "KEEP"
Private Const FLAG_HIDE As String = "HIDE"

Public Sub ToggleStrikethroughFilter()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim helperCol As Long
    Dim targetCol As Long
    Dim headerRow As Long
    Dim lastRow As Long
    Dim dataRng As Range, c As Range
    Dim arr() As Variant
    Dim idx As Long
    Dim userInput As Variant
    Dim origCalc As XlCalculation

    origCalc = Application.Calculation
    On Error GoTo CleanExit

    '==================================================
    ' 1. TOGGLE OFF — DETECT VIA ROW 1 (BULLETPROOF)
    '==================================================
    For helperCol = 1 To ws.Columns.Count
        If ws.Cells(1, helperCol).Value = HELPER_MARKER Then
            Application.ScreenUpdating = False

            If ws.AutoFilterMode Then
                If ws.FilterMode Then ws.ShowAllData
                ws.AutoFilterMode = False
            End If

            ws.Columns(helperCol).Delete
            Application.StatusBar = False
            GoTo CleanExit
        End If
    Next helperCol

    '==================================================
    ' 2. TOGGLE ON
    '==================================================
    targetCol = ActiveCell.Column

    userInput = Application.InputBox( _
        "Enter HEADER ROW number (this row will stay visible):", _
        "Strikethrough Filter", Type:=2)

    If userInput = False Then GoTo CleanExit
    If Not IsNumeric(userInput) Then GoTo CleanExit

    headerRow = CLng(userInput)
    If headerRow < 1 Or headerRow > ws.Rows.Count Then GoTo CleanExit

    lastRow = ws.Cells(ws.Rows.Count, targetCol).End(xlUp).Row
    If lastRow <= headerRow Then GoTo CleanExit

    ' Find safe helper column to the right
    helperCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set dataRng = ws.Range(ws.Cells(headerRow + 1, targetCol), _
                           ws.Cells(lastRow, targetCol))

    ReDim arr(1 To dataRng.Rows.Count + 1, 1 To 1)
    arr(1, 1) = "FILTER"   ' header row label (NOT marker)

    For Each c In dataRng.Cells
        idx = c.Row - headerRow + 1
        With c.Font
            If IsNull(.Strikethrough) Or .Strikethrough Then
                arr(idx, 1) = FLAG_KEEP
            Else
                arr(idx, 1) = FLAG_HIDE
            End If
        End With
    Next c

    ' Write helper column
    ws.Cells(headerRow, helperCol).Resize(UBound(arr), 1).Value = arr

    ' ?? Anchor marker ALWAYS in Row 1
    ws.Cells(1, helperCol).Value = HELPER_MARKER

    ' Apply filter
    ws.Cells(headerRow, helperCol).Resize(UBound(arr), 1). _
        AutoFilter Field:=1, Criteria1:=FLAG_KEEP

    ws.Columns(helperCol).Hidden = True

    Application.StatusBar = _
        "? Strikethrough Filter ON (Column " & _
        Split(ws.Cells(1, targetCol).Address, "$")(1) & _
        ", Header Row " & headerRow & ") ?"

CleanExit:
    Application.Calculation = origCalc
    Application.ScreenUpdating = True
End Sub



