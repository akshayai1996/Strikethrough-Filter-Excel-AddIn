Attribute VB_Name = "StrikethroughFilter"
Option Explicit

Private Const HELPER_HEADER As String = "__StrikeHelper__"
Private Const MENU_CAPTION As String = "Toggle Strikethrough Filter"

'==================================================
' MAIN TOGGLE (Right-Click)
'==================================================
Public Sub ToggleStrikethroughFilter()

    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim targetRange As Range
    Dim helperColIdx As Variant
    Dim lastRow As Long, colNum As Long, i As Long
    Dim arrResults() As Variant
    Dim finalCol As Long, filterField As Long
    Dim progressStep As Long

    '---------------- TOGGLE OFF ----------------
    helperColIdx = Application.Match(HELPER_HEADER, ws.Rows(1), 0)

    If Not IsError(helperColIdx) Then
        If ws.FilterMode Then ws.ShowAllData
        ws.Columns(helperColIdx).Hidden = False
        ws.Columns(helperColIdx).Delete
        Exit Sub
    End If

    '---------------- TOGGLE ON ----------------
    On Error Resume Next
    Set targetRange = Application.InputBox( _
        "Select any cell in the column to filter by Strikethrough:", _
        "Strikethrough Filter", Type:=8)
    On Error GoTo 0

    If targetRange Is Nothing Then Exit Sub

    ToggleEvents False

    colNum = targetRange.Column
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
    If lastRow < 2 Then GoTo SafeExit

    ReDim arrResults(1 To lastRow, 1 To 1)
    arrResults(1, 1) = HELPER_HEADER

    progressStep = Application.Max(1000, lastRow \ 20) ' ~5%

    ' Core detection (True / False / Null)
    For i = 2 To lastRow
        With ws.Cells(i, colNum).Font
            arrResults(i, 1) = (IsNull(.Strikethrough) Or .Strikethrough)
        End With

        If i Mod progressStep = 0 Or i = lastRow Then
            Application.StatusBar = _
                "Strikethrough Filter: Processing " & _
                Format(i / lastRow, "0%")
        End If
    Next i

    finalCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    ws.Cells(1, finalCol).Resize(lastRow, 1).Value = arrResults
    ws.Columns(finalCol).Hidden = True

    filterField = finalCol - ws.UsedRange.Column + 1
    ws.UsedRange.AutoFilter Field:=filterField, Criteria1:="TRUE"

SafeExit:
    Application.StatusBar = False
    ToggleEvents True

End Sub

'==================================================
' RIGHT-CLICK MENU
'==================================================
Public Sub AddRightClickMenu()
    Dim btn As Object
    On Error Resume Next
    Application.CommandBars("Cell").Controls(MENU_CAPTION).Delete
    On Error GoTo 0

    Set btn = Application.CommandBars("Cell").Controls.Add(msoControlButton)
    With btn
        .Caption = MENU_CAPTION
        .OnAction = "ToggleStrikethroughFilter"
        .BeginGroup = True
        .FaceId = 517
    End With
End Sub

Public Sub RemoveRightClickMenu()
    On Error Resume Next
    Application.CommandBars("Cell").Controls(MENU_CAPTION).Delete
    On Error GoTo 0
End Sub

'==================================================
' AUTO-REHIDE HELPER
'==================================================
Public Sub RehideHelperIfVisible(ByVal ws As Worksheet)
    Dim idx As Variant
    idx = Application.Match(HELPER_HEADER, ws.Rows(1), 0)

    If Not IsError(idx) Then
        If ws.Columns(idx).Hidden = False Then
            ws.Columns(idx).Hidden = True
        End If
    End If
End Sub

'==================================================
' CLEANUP ON EXIT
'==================================================
Public Sub CleanupAllStrikeHelpers()
    Dim ws As Worksheet, idx As Variant

    For Each ws In Application.Worksheets
        idx = Application.Match(HELPER_HEADER, ws.Rows(1), 0)
        If Not IsError(idx) Then
            If ws.FilterMode Then ws.ShowAllData
            ws.Columns(idx).Hidden = False
            ws.Columns(idx).Delete
        End If
    Next ws
End Sub

'==================================================
' PERFORMANCE SWITCH
'==================================================
Private Sub ToggleEvents(ByVal State As Boolean)
    With Application
        .ScreenUpdating = State
        .EnableEvents = State
        .Calculation = IIf(State, xlCalculationAutomatic, xlCalculationManual)
    End With
End Sub
