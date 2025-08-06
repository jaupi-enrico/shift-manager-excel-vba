Option Explicit

Private Sub Worksheet_Activate()
    Dim Password As String

    Password = "Ej20082018*Excel"
    Me.Protect Password:=Password
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim rngDati As Range
    Dim Password As String

    Password = "Ej20082018*Excel"

    Set rngDati = Range("C21:J26")

    If Not Intersect(Target, rngDati) Is Nothing And Not Intersect(Target, Range("B2")) Is Nothing Then
        ActiveSheet.Unprotect Password:=Password
    Else
        ActiveSheet.Protect Password:=Password
    End If
End Sub
