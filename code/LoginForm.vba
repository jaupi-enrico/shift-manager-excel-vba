Option Explicit

Private Sub CancelButton_Click()
    ' Definisci la password per proteggere il foglio
    Dim Password
    Password = "Ej20082018*Excel"
    
    ' Torna al foglio DASHBOARD e chiudi il form
    Worksheets("DASHBOARD").Activate
    Me.Hide
    Worksheets("Tabelle").Unprotect Password:=Password
    Worksheets("Tabelle").Range("H3").Value = 0
    Worksheets("Tabelle").Protect Password:=Password
End Sub

Private Sub EnterButton_Click()
    ' Definisci la password per proteggere il foglio
    Dim Password
    Password = "Ej20082018*Excel"
    
    Const CorrectPassword As String = "1996" ' Password corretta
    If Trim(PasswordBox.Value) = CorrectPassword Then
        Worksheets("Tabelle").Unprotect Password:=Password
        Worksheets("Tabelle").Range("H3").Value = 1
        Worksheets("Tabelle").Protect Password:=Password
        Me.Hide ' Chiudi il UserForm
    Else
        MsgBox "Password errata :)", vbExclamation, "Errore di accesso"
        PasswordBox.Value = "" ' Pulisci la TextBox
        PasswordBox.SetFocus ' Torna alla TextBox
        Worksheets("Tabelle").Unprotect Password:=Password
        Worksheets("Tabelle").Range("H3").Value = 0
        Worksheets("Tabelle").Protect Password:=Password
    End If
End Sub

Private Sub UserForm_Activate()
    ' Definisci la password per proteggere il foglio
    Dim Password
    Password = "Ej20082018*Excel"
    
    Worksheets("Tabelle").Unprotect Password:=Password
    ' Reimposta lo stato di login quando il form viene mostrato
    Worksheets("Tabelle").Range("H3").Value = 0
    Worksheets("Tabelle").Protect Password:=Password
    PasswordBox.SetFocus ' Imposta il focus sulla TextBox
End Sub
