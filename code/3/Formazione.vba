Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim Password
    Dim rngFormazione As Range
    Dim rngTotalone As Range
    Dim rngCorsi As Range
    Dim LastRow As Integer
    Dim Row As Integer
    Password = "Ej20082018*Excel"
    
    LastRow = 4
    While Cells(LastRow, 1) <> ""
        LastRow = LastRow + 1
    Wend

    If Target.Row >= Row And Target.Row <= LastRow And Target.Column <= 15 Then
        ActiveSheet.Unprotect Password:=Password
    Else
        ActiveSheet.Protect Password:=Password
    End If
End Sub


