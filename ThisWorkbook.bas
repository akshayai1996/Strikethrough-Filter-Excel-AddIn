Option Explicit

Private WithEvents App As Application

Private Sub Workbook_Open()
    Set App = Application
    AddRightClickMenu
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    CleanupAllStrikeHelpers
    RemoveRightClickMenu
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    RehideHelperIfVisible Sh
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    RehideHelperIfVisible Sh
End Sub
