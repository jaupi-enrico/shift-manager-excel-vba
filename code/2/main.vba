Option Explicit

Private Sub Worksheet_Activate()
    Dim Password As String

    Password = "Ej20082018*Excel"
    Me.Protect Password:=Password
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim Password As String

    Password = "Ej20082018*Excel"
    Me.Protect Password:=Password
End Sub