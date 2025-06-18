Option Explicit
Dim sheet As Worksheet ' Dichiarazione variabile sheet

Sub ShowSheets()
    Dim Password As String
    Dim sheet As Worksheet
    Dim currentSheet As Worksheet

    Application.ScreenUpdating = False
    Password = "Ej20082018*Excel"
    Set currentSheet = ActiveSheet

    For Each sheet In Worksheets
        If sheet.name <> "DASHBOARD" Then
            sheet.Unprotect Password:=Password
        End If
    Next sheet

    currentSheet.Activate
    Application.ScreenUpdating = True
End Sub


Sub HideSheets()
    Dim Password As String
    Dim sheet As Worksheet
    Dim currentSheet As Worksheet

    Application.ScreenUpdating = False
    Password = "Ej20082018*Excel"
    Set currentSheet = ActiveSheet

    For Each sheet In Worksheets
        If sheet.name <> "DASHBOARD" Then
            sheet.Protect Password:=Password
        End If
    Next sheet

    currentSheet.Activate
    Application.ScreenUpdating = True
End Sub

Sub Delete_images()
    Dim sheet As Worksheet
    Dim shp As Shape

    For Each sheet In Worksheets
        ' Escludi i fogli "DASHBOARD" e "TOT"
        If sheet.name <> "DASHBOARD" And sheet.name <> "TOT" And _
           sheet.name <> "Dipendenti" And sheet.name <> "FORMAZIONE" Then
        
            ' Cicla attraverso ogni forma presente nel foglio
            For Each shp In sheet.Shapes
                ' Elimina l'oggetto
                shp.Delete
            Next shp
        End If
    Next sheet
    
    MsgBox "Ottimizzazione finita"
End Sub

Sub Delete_names()
    Call ShowSheets
    Dim Answer As Long
    Answer = MsgBox("Sei sicuro di togliere i nomi?", vbYesNo + vbDefaultButton2)
        
    If Answer = vbYes Then
        ' Disabilita temporaneamente gli eventi
        Application.EnableEvents = False
        
        Dim name As Range
        Dim cell As Range
        Dim rngNames As Range
        Dim sheet As Worksheet
        Dim content As Variant
        Dim columnI As Integer, columnF As Integer, ColumnName As Integer
        Dim ColumnNameAproximation As Single, ColumnNameTemp As Integer
        Dim OriginalRow As Integer, i As Integer
        Dim incrementing As Boolean
        Dim CheckName As Boolean
        Dim AnswerQuestion As Long
        Dim Target As Range
        Dim Impresa As Boolean
        
        Answer = MsgBox("Anche l'impresa?", vbYesNo + vbDefaultButton2)
        If Answer = vbNo Then
            Impresa = False
        ElseIf Answer = vbYes Then
            Impresa = True
        End If
                
        ' Loop through each sheet in the workbook
        For Each sheet In Worksheets
            ' Check if the sheet name matches one of the specified days
            If sheet.name = "LUN" Or _
               sheet.name = "MAR" Or _
               sheet.name = "MER" Or _
               sheet.name = "GIO" Or _
               sheet.name = "VEN" Or _
               sheet.name = "SAB" Or _
               sheet.name = "DOM" Then
                
                If Impresa = False Then
                    Set rngNames = sheet.Range("A17:A153")
                ElseIf Impresa = True Then
                    Set rngNames = sheet.Range("A17:A164")
                End If
                
                
                ' Loop through each cell in the range and clear its contents
                For Each name In rngNames
                    If name.Value <> "" Then
                    Set Target = name
                    
                    ' Salva il contenuto della cella target (anche se   vuota dopo la cancellazione)
                    content = Target.Value
                    OriginalRow = Target.Row ' Assegna OriginalRow al numero di riga corrente
                    
                    name.Value = ""
                    
                    columnF = 0
                    columnI = 0
                    
                    ' Trova colonne "I" e "F"
                    For Each cell In sheet.Range(sheet.Cells(OriginalRow, 6), sheet.Cells(OriginalRow, 70))
                        If cell.Value = "I" Then
                            columnI = cell.Column
                        ElseIf Trim(cell.Value) = "F" Then
                            columnF = cell.Column
                        End If
                        If columnI > 0 And columnF > 0 Then Exit For
                    Next cell
                    
                    ' Aggiorna lo stile delle celle in base alla condizione
                    If sheet.Cells(OriginalRow, 1).Value = "" And columnI > 0 And columnF > 0 Then
                        sheet.Cells(OriginalRow, 1).Interior.color = RGB(255, 255, 0)
                    Else
                        sheet.Cells(OriginalRow, 1).Interior.color = RGB(217, 217, 217)
                    End If
                    
                    ' Cancella contenuti se necessario
                    For Each cell In sheet.Range(sheet.Cells(OriginalRow - 1, columnI), sheet.Cells(OriginalRow - 1, columnF))
                        If cell.Value <> "" And cell.Value = content Then
                            cell.ClearContents
                        End If
                    Next cell
                    End If
                Next name
            End If
        Next sheet
    
Cleanup:
        
    columnF = 0
    columnI = 0
    
    Call HideSheets

    ' Riabilita gli eventi
     Application.EnableEvents = True
End Sub

Sub Show_Lines()

    Call ShowSheets
    
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
                            cell_y.EntireRow.Hidden = False
                        End If
                        
                        If clear And clear_up And clear_up_up Then
                            sheet.Rows(cell_y.Row - 1).Hidden = False
                        End If
                        
                        If clear And clear_down And clear_down_down Then
                            sheet.Rows(cell_y.Row + 1).Hidden = False
                        End If
                        
                    End If
                Next cell_y
            Else
                Debug.Print "Intervallo non trovato sul foglio: " & sheet.name
            End If

        End If
    Next sheet
    
    Call HideSheets
    
Exit Sub

ErrorHandler:
    MsgBox "Errore: " & Err.Description & " (Codice " & Err.Number & ") nel foglio " & sheet.name, vbCritical
    Exit Sub
End Sub

Sub Hide_Lines()
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
    
    Call ShowSheets
    
    On Error GoTo ErrorHandler ' Gestione degli errori
    
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
                        End If
                        
                        If clear And clear_down And clear_down_down Then
                            sheet.Rows(cell_y.Row + 1).Hidden = True
                        End If
                        
                    End If
                Next cell_y
            Else
                Debug.Print "Intervallo non trovato sul foglio: " & sheet.name
            End If

        End If
    Next sheet
    
    Call HideSheets
    
Exit Sub

ErrorHandler:
    MsgBox "Errore: " & Err.Description & " (Codice " & Err.Number & ") nel foglio " & sheet.name, vbCritical
    Exit Sub
End Sub

Sub Add_Manager()
    Call ShowSheets
    Dim manager_row As Integer

    manager_row = 4

    While Cells(manager_row, 58) <> ""
        manager_row = manager_row + 1
    Wend
    manager_row = manager_row - 1


    ActiveSheet.Rows(manager_row + 1).Insert Shift:=xlDown
    DoEvents
    ActiveSheet.Rows(manager_row).Copy
    ActiveSheet.Rows(manager_row + 1).PasteSpecial xlPasteAll

    Call ShowSheets
    Range(Cells(manager_row + 1, 58), Cells(manager_row + 1, 72)).Value = ""
    Call HideSheets
End Sub