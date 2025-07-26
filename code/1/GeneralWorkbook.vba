Option Explicit

Dim sheet As Worksheet ' Dichiarazione variabile sheet
Dim Answer As Long ' Dichiarazione variabile Answer
Dim Closing As Boolean ' Dichiarazione variabile Closing
Dim Changed As Boolean
Dim ChangedAfterSave As Boolean


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Changed = False
    Answer = MsgBox("Vuoi salvare?", vbYesNoCancel + vbQuestion + vbDefaultButton1)
    Dim Password As String
    Password = "Ej20082018*Excel"
    
    If Answer = vbYes Then
        Application.ScreenUpdating = False
        Worksheets(1).Activate
        For Each sheet In Worksheets
            sheet.Protect Password:=Password
            If sheet.name = "LUN" Or _
                sheet.name = "MAR" Or _
                sheet.name = "MER" Or _
                sheet.name = "GIO" Or _
                sheet.name = "VEN" Or _
                sheet.name = "SAB" Or _
                sheet.name = "DOM" Or _
                sheet.name = "MANAGER" Then
                    sheet.Activate
                    ActiveWindow.FreezePanes = False
            End If
        Next sheet
        Application.ScreenUpdating = True
        ThisWorkbook.Save
    ElseIf Answer = vbNo Then
        If ChangedAfterSave = False Then
            Application.ScreenUpdating = False
            For Each sheet In Worksheets
                sheet.Protect Password:=Password
                If sheet.name = "LUN" Or _
                    sheet.name = "MAR" Or _
                    sheet.name = "MER" Or _
                    sheet.name = "GIO" Or _
                    sheet.name = "VEN" Or _
                    sheet.name = "SAB" Or _
                    sheet.name = "DOM" Or _
                    sheet.name = "MANAGER" Then
                        sheet.Activate
                        ActiveWindow.FreezePanes = False
                End If
            Next sheet
            Application.ScreenUpdating = True
        ThisWorkbook.Save
        End If
        Closing = True
        Cancel = False ' Continua con la chiusura
        Worksheets(1).Activate
    Else
        Cancel = True ' Annulla la chiusura
        Closing = False ' Assicurarsi che Closing sia impostato su False
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim Password As String
    Password = "Ej20082018*Excel"
    If Closing = True And Changed = False Then
        Worksheets(1).Activate
        For Each sheet In Worksheets
            sheet.Protect Password:=Password
            If sheet.name = "LUN" Or _
                sheet.name = "MAR" Or _
                sheet.name = "MER" Or _
                sheet.name = "GIO" Or _
                sheet.name = "VEN" Or _
                sheet.name = "SAB" Or _
                sheet.name = "DOM" Or _
                sheet.name = "MANAGER" Then
                    sheet.Activate
                    ActiveWindow.FreezePanes = False
            End If
        Next sheet
    End If
    Closing = False ' Reimposta Closing su False per evitare comportamenti imprevisti
    ChangedAfterSave = False
End Sub

Private Sub Workbook_Open()
    Dim Password As String
    Password = "Ej20082018*Excel"
    For Each sheet In Worksheets
        sheet.Unprotect Password:=Password
        If sheet.name = "LUN" Or _
            sheet.name = "MAR" Or _
            sheet.name = "MER" Or _
            sheet.name = "GIO" Or _
            sheet.name = "VEN" Or _
            sheet.name = "SAB" Or _
            sheet.name = "DOM" Then
                sheet.Activate
                Range("F16").Activate
                ActiveWindow.FreezePanes = True
        ElseIf sheet.name = "MANAGER" Then
            sheet.Activate
            Range("F2").Activate
            ActiveWindow.FreezePanes = True
        End If
    Next sheet
    Worksheets(1).Activate
    Call AggiornaCodiceVBA
    Call ApplyChanges
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Dim ProtectedSheets As Variant
    Dim Password As String
    Password = "Ej20082018*Excel"
    ProtectedSheets = Array("Fasce_Tot", "Percentuali", "Tabelle", "Dipendenti", "Dipendenti-M") ' Elenco dei fogli protetti
    
    If Not IsError(Application.Match(Sh.name, ProtectedSheets, 0)) Then
        LoginForm.Show ' Mostra il form di login
        ' Se il login non ha successo, torna al foglio "DASHBOARD"
        If Worksheets("Tabelle").Range("H3").Value = 0 Then
            MsgBox "Accesso negato. Verrai reindirizzato alla Dashboard.", vbExclamation, "Autenticazione richiesta"
            Worksheets("DASHBOARD").Activate
        End If
    Else
        Sh.Unprotect Password:=Password
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Changed = True
    ChangedAfterSave = True
End Sub

Private Sub Workbook_BeforePrint(Cancel As Boolean)
    Call Hide_Lines
End Sub