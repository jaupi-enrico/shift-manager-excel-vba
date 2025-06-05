Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim Password
    Dim rngFormazione As Range
    Dim rngTotalone As Range
    Dim rngCorsi As Range
    Dim LastRow As Integer
    Password = "Ej20082018*Excel"
    
    LastRow = 4
    While Cells(LastRow, 1) <> ""
        LastRow = LastRow + 1
    Wend
    LastRow = LastRow - 1
    
    While Cells(LastRow, 54) <> ""
        LastRow = LastRow + 1
    Wend
    LastRow = LastRow - 1
    
    Set rngFormazione = Range("C4", Cells(LastRow, 16))
    Set rngTotalone = Range("T4", Cells(LastRow, 55))
    Set rngCorsi = Range("BG4", Cells(LastRow, 72))
    
    If Not Intersect(Target, rngFormazione) Is Nothing Or Not Intersect(Target, rngTotalone) Is Nothing _
    Or Not Intersect(Target, rngCorsi) Is Nothing Then
        ActiveSheet.Unprotect Password:=Password
    ElseIf Target.Row >= 55 And Target.Row <= 62 Then
        ActiveSheet.Unprotect Password:=Password
    Else
        ActiveSheet.Protect Password:=Password
    End If
    
    
End Sub


