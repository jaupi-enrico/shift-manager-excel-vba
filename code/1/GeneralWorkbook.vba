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
    Dim sheet As Worksheet
    Dim names_range As Range
    Dim cell_y As Range
    Dim cell_x As Range
    Dim line As Range
    Dim clear As Boolean
    Dim clear_up As Boolean
    Dim clear_down As Boolean
    Dim clear_up_up As Boolean
    Dim clear_down_down As Boolean

    On Error GoTo ErrorHandler ' Gestione degli errori
    
    If ActiveSheet.name = "TOT" Then
        Exit Sub
    End If
    
    ' Loop attraverso i fogli
    For Each sheet In Worksheets
        ' Controlla se il nome del foglio   tra i giorni della settimana
        If sheet.name = "LUN" Or _
           sheet.name = "MAR" Or _
           sheet.name = "MER" Or _
           sheet.name = "GIO" Or _
           sheet.name = "VEN" Or _
           sheet.name = "SAB" Or _
           sheet.name = "DOM" Then
            ' Prova a definire l'intervallo specifico sul foglio corrente
            On Error Resume Next
            Set names_range = sheet.Range("A17:A164")
            On Error GoTo ErrorHandler

            If Not names_range Is Nothing Then
                For Each cell_y In names_range
                    If Not IsEmpty(cell_y.Interior.color) And cell_y.Interior.color = RGB(217, 217, 217) Then
                        clear = True
                        clear_up = True
                        clear_down = True
                        clear_up_up = True
                        clear_down_down = True

                        ' Definisci la riga da controllare sul foglio corrente
                        Set line = sheet.Range(sheet.Cells(cell_y.Row, 6), sheet.Cells(cell_y.Row, 70))

                        ' Controlla il contenuto della riga
                        For Each cell_x In line
                            If cell_x.Value <> "" Then
                                clear = False
                                Exit For
                            End If
                        Next cell_x

                        ' Controlla la riga sopra
                        If cell_y.Row > 1 Then ' Evita di andare sopra la prima riga
                            For Each cell_x In line
                                If sheet.Cells(cell_x.Row - 1, cell_x.Column).Value <> "" Then
                                    clear_up = False
                                    Exit For
                                End If
                            Next cell_x
                        Else
                            clear_up = False
                        End If
                        
                        If cell_y.Row > 2 And clear_up Then ' Evita di andare sopra la prima riga
                            For Each cell_x In line
                                If sheet.Cells(cell_x.Row - 2, cell_x.Column).Value <> "" Then
                                    clear_up_up = False
                                    Exit For
                                End If
                            Next cell_x
                        Else
                            clear_up_up = False
                        End If

                        ' Controlla la riga sotto
                        If cell_y.Row < sheet.Rows.Count Then ' Evita di andare oltre l'ultima riga
                            For Each cell_x In line
                                If sheet.Cells(cell_x.Row + 1, cell_x.Column).Value <> "" Then
                                    clear_down = False
                                    Exit For
                                End If
                            Next cell_x
                        Else
                            clear_down = False
                        End If
                        
                        If cell_y.Row < sheet.Rows.Count - 1 And clear_down Then ' Evita di andare oltre l'ultima riga
                            For Each cell_x In line
                                If sheet.Cells(cell_x.Row + 1, cell_x.Column).Value <> "" Then
                                    clear_down_down = False
                                    Exit For
                                End If
                            Next cell_x
                        Else
                            clear_down_down = False
                        End If

                        ' Nascondi le righe
                        
                        If clear Then
                            cell_y.EntireRow.Hidden = True
                        End If
                        
                        If clear And clear_up And clear_up_up Then
                            sheet.Rows(cell_y.Row - 1).Hidden = True
                        ElseIf clear And clear_up And Not clear_down Then
                            sheet.Rows(cell_y.Row - 1).Hidden = True
                        End If
                        
                        If clear And clear_down And clear_down_down Then
                            sheet.Rows(cell_y.Row + 1).Hidden = True
                        ElseIf clear And clear_down And Not clear_up Then
                            sheet.Rows(cell_y.Row + 1).Hidden = True
                        End If
                        
                    End If
                Next cell_y
            Else
                Debug.Print "Intervallo non trovato sul foglio: " & sheet.name
            End If
        ElseIf sheet.name = "MANAGER" Then
            ' Gestisci il foglio "MANAGER"
 ' Prova a definire l'intervallo specifico sul foglio corrente
            On Error Resume Next
            Set names_range = sheet.Range("A3:A147")
            On Error GoTo ErrorHandler

            If Not names_range Is Nothing Then
                For Each cell_y In names_range
                    If Not IsEmpty(cell_y.Interior.color) And cell_y.Interior.color = RGB(217, 217, 217) Then
                        clear = True
                        clear_up = True
                        clear_down = True
                        clear_up_up = True
                        clear_down_down = True

                        ' Definisci la riga da controllare sul foglio corrente
                        Set line = sheet.Range(sheet.Cells(cell_y.Row, 8), sheet.Cells(cell_y.Row, 72))

                        ' Controlla il contenuto della riga
                        For Each cell_x In line
                            If cell_x.Value <> "" Then
                                clear = False
                                Exit For
                            End If
                        Next cell_x

                        ' Controlla la riga sopra
                        If cell_y.Row > 1 Then ' Evita di andare sopra la prima riga
                            For Each cell_x In line
                                If sheet.Cells(cell_x.Row - 1, cell_x.Column).Value <> "" Then
                                    clear_up = False
                                    Exit For
                                End If
                            Next cell_x
                        Else
                            clear_up = False
                        End If
                        
                        If cell_y.Row > 2 And clear_up Then ' Evita di andare sopra la prima riga
                            For Each cell_x In line
                                If sheet.Cells(cell_x.Row - 2, cell_x.Column).Value <> "" Then
                                    clear_up_up = False
                                    Exit For
                                End If
                            Next cell_x
                        Else
                            clear_up_up = False
                        End If

                        ' Controlla la riga sotto
                        If cell_y.Row < sheet.Rows.Count Then ' Evita di andare oltre l'ultima riga
                            For Each cell_x In line
                                If sheet.Cells(cell_x.Row + 1, cell_x.Column).Value <> "" Then
                                    clear_down = False
                                    Exit For
                                End If
                            Next cell_x
                        Else
                            clear_down = False
                        End If
                        
                        If cell_y.Row < sheet.Rows.Count - 1 And clear_down Then ' Evita di andare oltre l'ultima riga
                            For Each cell_x In line
                                If sheet.Cells(cell_x.Row + 1, cell_x.Column).Value <> "" Then
                                    clear_down_down = False
                                    Exit For
                                End If
                            Next cell_x
                        Else
                            clear_down_down = False
                        End If

                        ' Nascondi le righe
                        
                        If clear Then
                            cell_y.EntireRow.Hidden = True
                        End If
                        
                        If clear And clear_up And clear_up_up Then
                            sheet.Rows(cell_y.Row - 1).Hidden = True
                        ElseIf clear And clear_up And Not clear_down Then
                            sheet.Rows(cell_y.Row - 1).Hidden = True
                        End If
                        
                        If clear And clear_down And clear_down_down Then
                            sheet.Rows(cell_y.Row + 1).Hidden = True
                        ElseIf clear And clear_down And Not clear_up Then
                            sheet.Rows(cell_y.Row + 1).Hidden = True
                        End If
                        
                    End If
                Next cell_y
            Else
                Debug.Print "Intervallo non trovato sul foglio: " & sheet.name
            End If
        End If
    Next sheet

    Exit Sub

ErrorHandler:
    MsgBox "Errore: " & Err.Description & " (Codice " & Err.Number & ") nel foglio " & sheet.name, vbCritical
    Exit Sub
End Sub