Option Explicit

Private Sub Worksheet_Activate()
    Dim Password As String
    Dim LastRow As Integer

    Password = "Ej20082018*Excel"
    Me.Unprotect Password:=Password

    LastRow = 4
    While Me.Cells(LastRow, 1).Value <> ""
        LastRow = LastRow + 1
    Wend
    LastRow = LastRow - 1

    If LastRow < 4 Then
        Me.Protect Password:=Password
        Exit Sub
    End If

    Me.Cells(LastRow + 1, 28).FormulaLocal = "=SOMMA(AB4:AB" & LastRow & ")"
    Me.Cells(LastRow + 2, 28).FormulaLocal = "=SOMMA(AA4:AA" & LastRow & ")"
    
    Me.Protect Password:=Password
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim Password
    Dim rngPermessi As Range
    Dim LastRow As Integer
    Password = "Ej20082018*Excel"
    
    LastRow = 4
    While Cells(LastRow, 1).Value <> ""
        LastRow = LastRow + 1
    Wend
    LastRow = LastRow - 1
    
    Set rngPermessi = Range("AI4", Cells(LastRow, 49))

    If Not Intersect(Target, rngPermessi) Is Nothing Or Not Intersect(Target, Range("C52")) Is Nothing _
    Or Cells(Target.Row, 1).Interior.color = RGB(241, 170, 131) Then
        Me.Unprotect Password:=Password
    Else
        Me.Protect Password:=Password
    End If

End Sub