Option Explicit
Dim sheet As Worksheet ' Dichiarazione variabile sheet

' Password di protezione dei fogli: unico punto in cui e' definita
Public Const SHEET_PASSWORD As String = "Ej20082018*Excel"
' PIN richiesto dal LoginForm per i fogli riservati
Public Const LOGIN_PIN As String = "1996"

' Ripristina lo stato dell'applicazione dopo operazioni lunghe o errori
Sub RestoreAppState()
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
End Sub

' Assegna una formula "dinamica" (Formula2) e, se la versione di Excel
' non la supporta (prima del 365), ripiega su Formula
Sub Set_Formula_Compat(ByVal rng As Range, ByVal formulaEN As String)
    On Error Resume Next
    rng.Formula2 = formulaEN
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        rng.Formula = formulaEN
    End If
    On Error GoTo 0
End Sub

' Converte una formula italiana usata nelle formattazioni condizionali
' (E/O, separatore ";", decimali con ",") nell'equivalente inglese
Function Formula_To_English(ByVal formulaIT As String) As String
    Dim f As String
    f = Replace(formulaIT, "E(", "AND(")
    f = Replace(f, "O(", "OR(")
    f = Replace(f, ";", ",")
    f = Replace(f, "0,46", "0.46")
    Formula_To_English = f
End Function

' Aggiunge una formattazione condizionale basata su formula.
' Prova prima la sintassi italiana (quella usata finora), poi quella inglese
' (Excel con impostazioni regionali diverse)
Function Add_Cond_Format(ByVal rng As Range, ByVal formulaIT As String) As FormatCondition
    Dim fc As FormatCondition
    On Error Resume Next
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formulaIT)
    If fc Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=Formula_To_English(formulaIT))
    End If
    On Error GoTo 0
    Set Add_Cond_Format = fc
End Function

Sub ShowSheets()
    Dim Password As String
    Dim sheet As Worksheet
    Dim currentSheet As Worksheet

    Application.ScreenUpdating = False
    Password = SHEET_PASSWORD
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
    Password = SHEET_PASSWORD
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

    Call ShowSheets

    For Each sheet In Worksheets
        ' Escludi i fogli "DASHBOARD" e "TOT"
        If sheet.name <> "DASHBOARD" And sheet.name <> "TOT" And _
           sheet.name <> "Dipendenti" And sheet.name <> "FORMAZIONE" And _
           sheet.name <> "Dipendenti-M" And sheet.name <> "FORMAZIONE-M" And _
           sheet.name <> "TOT-M" Then

            ' Cicla attraverso ogni forma presente nel foglio
            For Each shp In sheet.Shapes
                ' Elimina l'oggetto
                shp.Delete
            Next shp
        End If
    Next sheet

    Call HideSheets

    MsgBox "Ottimizzazione finita"
End Sub

Sub Delete_names()
    Call ShowSheets
    On Error GoTo ErrHandler
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
                    
                    If columnI > 0 And columnF > 0 Then
                        ' Cancella contenuti se necessario
                        For Each cell In sheet.Range(sheet.Cells(OriginalRow - 1, columnI), sheet.Cells(OriginalRow - 1, columnF))
                            If cell.Value <> "" And cell.Value = content Then
                                cell.ClearContents
                            End If
                        Next cell
                    End If
                    End If
                Next name
            ElseIf sheet.name = "MANAGER" Then

                Set rngNames = sheet.Range("A3:A147")

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
                    For Each cell In sheet.Range(sheet.Cells(OriginalRow, 8), sheet.Cells(OriginalRow, 72))
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
                    
                    ' Cancella contenuti se necessario (solo se la riga ha sia "I" che "F")
                    If columnI > 0 And columnF > 0 Then
                        For Each cell In sheet.Range(sheet.Cells(OriginalRow - 1, columnI), sheet.Cells(OriginalRow - 1, columnF))
                            If cell.Value <> "" And cell.Value = content Then
                                cell.ClearContents
                            End If
                        Next cell
                    End If
                    End If
                Next name
            End If
        Next sheet
    End If
Cleanup:
    On Error Resume Next
    columnF = 0
    columnI = 0
    
    Call HideSheets

    ' Riabilita gli eventi
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "Errore: " & Err.Description & " (Codice " & Err.Number & ")", vbCritical
    Resume Cleanup
End Sub

Sub Show_Lines()

    Call ShowSheets
    
    Dim sheet As Worksheet
    Dim names_range As Range
    Dim cell_y As Range
    Dim cell_x As Range
    Dim line As Range

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
            Set names_range = sheet.Range("A16:A165")
            On Error GoTo ErrorHandler

            If Not names_range Is Nothing Then
                For Each cell_y In names_range  
                    If cell_y.EntireRow.Hidden = True Then
                        cell_y.EntireRow.Hidden = False
                    End If
                Next cell_y
            Else
                Debug.Print "Intervallo non trovato sul foglio: " & sheet.name
            End If
        ElseIf sheet.name = "MANAGER" Then
            On Error Resume Next
            Set names_range = sheet.Range("A2:A148")
            On Error GoTo ErrorHandler

            If Not names_range Is Nothing Then
                For Each cell_y In names_range
                    If cell_y.EntireRow.Hidden = True Then
                        cell_y.EntireRow.Hidden = False
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
    On Error Resume Next
    Call HideSheets
    Call RestoreAppState
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
    Application.ScreenUpdating = False
    
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
                    If Not IsEmpty(cell_y.Interior.color) And Not cell_y.Interior.color = RGB(255, 255, 255) Then
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
                                If sheet.Cells(cell_x.Row + 2, cell_x.Column).Value <> "" Then
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
        ElseIf sheet.name = "MANAGER" Then
            ' Gestisci la logica specifica per il foglio MANAGER
            On Error Resume Next
            Set names_range = sheet.Range("A3:A147")
            On Error GoTo ErrorHandler

            If Not names_range Is Nothing Then
                For Each cell_y In names_range
                    If Not IsEmpty(cell_y.Interior.color) And Not cell_y.Interior.color = RGB(255, 255, 255) Then
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
                                If sheet.Cells(cell_x.Row + 2, cell_x.Column).Value <> "" Then
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
    On Error Resume Next
    Call HideSheets
    Call RestoreAppState
    Exit Sub
End Sub