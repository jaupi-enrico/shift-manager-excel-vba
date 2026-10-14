Option Explicit

Private Sub Worksheet_Activate()
    Dim Password As String
    Dim LastRow As Integer

    Password = SHEET_PASSWORD
    Me.Unprotect Password:=Password
    On Error GoTo ErrHandler

    LastRow = 4
    While Me.Cells(LastRow, 1).Value <> ""
        LastRow = LastRow + 1
    Wend
    LastRow = LastRow - 1

    If LastRow < 4 Then
        Me.Protect Password:=Password
        Exit Sub
    End If
    
    Me.Cells(LastRow + 1, 28).Formula = "=SUM(AB" & LastRow - 3 & ":AB" & LastRow & ")"
    Me.Cells(LastRow + 2, 28).Formula = "=SUM(AB4:AB" & LastRow - 4 & ")"
    Me.Cells(LastRow + 3, 28).Formula = "=SUM(AA4:AA" & LastRow - 4 & ")"
    
    Me.Protect Password:=Password
    Exit Sub

ErrHandler:
    ' Il foglio e' stato sbloccato: riproteggilo anche in caso di errore
    Dim ErrMsg As String
    ErrMsg = Err.Description
    On Error Resume Next
    Me.Protect Password:=Password
    MsgBox "Errore nell'aggiornamento dei totali: " & ErrMsg, vbExclamation
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim Password
    Dim rngPermessi As Range
    Dim LastRow As Integer
    Password = SHEET_PASSWORD
    
    LastRow = 4
    While Cells(LastRow, 1).Value <> ""
        LastRow = LastRow + 1
    Wend
    LastRow = LastRow - 1
    
    Set rngPermessi = Range("AI4", Cells(LastRow, 49))
    
    If Not Intersect(Target, rngPermessi) Is Nothing Or Not Intersect(Target, Range("C52")) Is Nothing _
    Or Cells(Target.Row, 1).Interior.color = RGB(241, 170, 131) Then
        ActiveSheet.Unprotect Password:=Password
    Else
        ActiveSheet.Protect Password:=Password
    End If
End Sub


