Option Explicit

' Definizione delle variabili globali
Dim content As Variant
Dim orario As String
Dim columnI As Integer
Dim columnF As Integer
Dim Column As Integer
Dim rngC As Range, rngD As Range, rngName As Range, rngLines As Range, rngDiff As Range, rngResp As Range
Dim Times As Range
Dim Time As Range
Dim ColumnFound As Integer
Dim OriginalRow As Integer
Dim rangeName As String
Dim Answer As Long
Dim ColumnName As Integer
Dim ColumnNameAproximation As Single
Dim ColumnNameTemp As Integer
Dim cell As Range
Dim i As Integer
Dim incrementing As Boolean
Dim CheckName As Boolean
Dim AnswerQuestion As Long
Dim AllColumnsOccupied As Boolean

Function Find_Name_Column()
    ' Se la colonna del nome � gi� occupata o � un orario di pausa
    ' cerca la colonna successiva disponibile
    If Cells(OriginalRow - 1, ColumnName).Value <> "" Or Cells(OriginalRow, ColumnName).Value = "" Then
        ColumnNameTemp = ColumnName
        incrementing = True
        AllColumnsOccupied = False
        i = 0
        Do While Cells(OriginalRow - 1, ColumnName).Value <> "" Or Cells(OriginalRow, ColumnName).Value = ""
            ' Se la colonna supera la colonna "F", inizia a decrementare
            If ColumnNameTemp + i >= columnF Then
                incrementing = False
            End If
    
            If incrementing Then
                i = i + 1
            Else
                i = i - 1
            End If
    
            ColumnName = ColumnNameTemp + i
            
            ' Se la colonna supera la colonna "I", permetti la ricerca anche se � un orario di pausa
            If ColumnName = columnI And Cells(OriginalRow - 1, ColumnName).Value <> "" Then
                i = 0
                incrementing = True
                Do While Cells(OriginalRow - 1, ColumnName).Value <> ""
                    If ColumnNameTemp + i >= columnF Then
                        incrementing = False
                    End If
            
                    If incrementing Then
                        i = i + 1
                    Else
                        i = i - 1
                    End If

                    ColumnName = ColumnNameTemp + i

                    If ColumnName = columnI - 1 Then
                        MsgBox "Tutte le colonne sono occupate, libera spazio per visulizzare il nome!", vbCritical + vbOKOnly + vbDefaultButton1
                        AllColumnsOccupied = True
                        Exit Do
                    End If
                Loop
                Exit Do
            End If
        Loop
    End If
End Function

Function CancelName(ByVal Target As Range)
    ' Cancella il nome se gi� presente in tutte le colonne
    For Each cell In Range(Cells(OriginalRow - 1, 8), Cells(OriginalRow - 1, 72))
        If (cell.Value <> "" And cell.Value = Cells(OriginalRow, 1).Value) Or (Target.Column = 1 And cell.Value = content) Then
            cell.ClearContents
            Exit For
        End If
    Next cell
End Function

Function Calculate_Name_Column(ByVal Target As Range)
    ' Calcola la colonna per il nome
    ColumnNameAproximation = (columnF - columnI) / 2 + columnI
    ColumnName = Round(ColumnNameAproximation)
    
    Call CancelName(Target)

    ' Sistema la colonna del nome se � gi� occupata
    ' oppure � un orario di pausa
    Call Find_Name_Column

    ' Mette il nome al turno e formatta la cella
    If AllColumnsOccupied = False And ColumnName > 0 And columnI > 0 And columnF > 0 Then
        Cells(OriginalRow - 1, ColumnName).Value = Cells(OriginalRow, 1).Value
        With Cells(OriginalRow - 1, ColumnName)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.name = "Verdana"
            .Font.Bold = True
            .Font.Size = 26
        End With
    End If
End Function

Function Paint_Name_Cell()
    ' Se la cella del nome � vuota e le colonne "I" e "F" sono valide
    ' allora evidenziala in giallo
    ' Altrimenti, ripristina il colore originale
    Call Find_Column_I_F
    
    If Cells(OriginalRow, 1).Value = "" And (columnI > 0 Or columnF > 0) Then
        Cells(OriginalRow, 1).Interior.color = RGB(255, 255, 0)
    Else
        Cells(OriginalRow, 1).Interior.color = RGB(217, 217, 217)
    End If
End Function

Function Find_Column_I_F()
    ' Trova la posizione delle colonne "I" e "F" nella OriginalRow
    For Each cell In Range(Cells(OriginalRow, 8), Cells(OriginalRow, 72))
        If cell.Value = "I" Then
            columnI = cell.Column
            Exit For
        End If
    Next cell

    For Each cell In Range(Cells(OriginalRow, 8), Cells(OriginalRow, 72))
        If cell.Value = "F" Then
            columnF = cell.Column
            Exit For
        End If
    Next cell
End Function

Function TextChange(ByVal Target As Range)
    ' Controlla se il valore della cella � "P" e se la cella sopra � vuota
    If (Target.Value <> "") And (Target.Value <> Cells(Target.Row + 1, 1).Value) Then
        ' Formatta la cella per un commento
        With Cells(Target.Row, Target.Column)
            .Font.name = "Calibri"
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 28
            .Interior.color = RGB(255, 255, 0)
        End With
    End If
    ' Se il valore della cella � vuoto
    If Target.Value = "" Then
        ' Ripristina il formato della cella
        If (Target.Column <= 32 And Target.Column >= 24) Or (Target.Column <= 51 And Target.Column >= 43) Then
            ' Colonne di rush
            Target.Interior.color = RGB(217, 217, 217)
        Else
            ' Colonne di lavoro normale
            Target.Interior.color = xlNone
        End If
    End If
End Function

' Questo evento viene attivato quando si cambia la cella selezionata
' e protegge o de-protegge il foglio in base alla selezione
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Definisci la password per proteggere il foglio
    Dim Password
    Password = "Ej20082018*Excel"
    ' Definisci gli intervalli di celle
    ' che devono essere de-protetti
    Set rngC = Range("C2:C148")
    Set rngD = Range("D2:D148")
    Set rngName = Range("A3:A148")
    Set rngLines = Range("H2:BT148")
    Set rngDiff = Range("B2:B148")
    Set rngResp = Range("F2:G148")
    On Error Resume Next
    ' Verifica se � stata selezionata una sola cella
    If Target.Cells.Count = 1 Then
        ' Salva il contenuto della cella selezionata
        content = Target.Formula
    End If
    On Error Resume Next
    

    ' BUG!!
    If ActiveWindow.SelectedSheets.Count = 1 Then
        ' Controlla se la cella selezionata � all'interno di uno degli intervalli specificati
        ' Se � cos�, de-protegge il foglio
        If Not Intersect(Target, rngDiff) Is Nothing Then
            ActiveSheet.Protect Password:=Password
            ActiveSheet.Cells(Target.Row, Target.Column - 1).Activate
        ElseIf Not Intersect(Target, rngC) Is Nothing Or Not Intersect(Target, rngD) Is Nothing _
        Or Not Intersect(Target, rngName) Is Nothing Or Not Intersect(Target, rngLines) Is Nothing _
        Or Not Intersect(Target, rngResp) Is Nothing Or Not Intersect(Target, Range("AQ149")) Is Nothing Then
            ActiveSheet.Unprotect Password:=Password
        Else
            ' Altrimenti, protegge il foglio
            ActiveSheet.Protect Password:=Password
        End If
    End If
    
End Sub

' Questo evento viene attivato quando si cambia il valore di una cella
' e gestisce le modifiche in base alla cella selezionata
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo ErrorHandler ' Gestione degli errori

    ' Inizializza le variabili
    columnF = 0
    columnI = 0
    
    ' Definisci gli intervalli di celle
    ' che devono essere monitorati per le modifiche
    Set rngC = Range("C2:C148")
    Set rngD = Range("D2:D148")
    Set rngName = Range("A3:A148")
    Set rngLines = Range("H2:BT148")
    Set rngDiff = Range("B2:B148")
    Set rngResp = Range("F2:G148")

    ' Non togliere questa parte, serve per evitare errori
    ' per il men� a tendina.
    ' Tempo speso = 5 Ore
    If Application.Ready = False Then Exit Sub
    If Application.CommandBars("Cell").Enabled = False Then Exit Sub

    ' Disabilita temporaneamente gli eventi
    Application.EnableEvents = False
    
    ' Verifica se la cella cambiata � all'interno di uno degli intervalli specificati
    If Not Intersect(Target, rngC) Is Nothing And Target.Interior.color <> RGB(255, 255, 255) Then
        ' Se la cella � nella colonna di inizio orario e la cella all'inizio non � bianca
        ' Imposta il valore di "orario" e la riga originale
        orario = Target.Text
        OriginalRow = Target.Row
        rangeName = "I"
    ElseIf Not Intersect(Target, rngD) Is Nothing And Target.Interior.color <> RGB(255, 255, 255) Then
        ' Se la cella � nella colonna di fine orario e la cella all'inizio non � bianca
        ' Imposta il valore di "orario" e la riga originale
        orario = Target.Text
        OriginalRow = Target.Row
        rangeName = "F"
    ElseIf Not Intersect(Target, rngDiff) Is Nothing Then
        Target.Formula = content
        GoTo Cleanup
    ElseIf Not Intersect(Target, rngName) Is Nothing And Target.Interior.color <> RGB(255, 255, 255) Then
        ' Se la cella � nella colonna del nome e la cella all'inizio non � bianca
        ' Imposta la riga originale e il flag per il controllo del nome
        OriginalRow = Target.Row
        Target.Value = UCase(Target.Value)

        With Target
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.name = "Verdana"
            .Font.Bold = True
            .Font.Size = 24
        End With

        CheckName = True
        GoTo NameChange
    ElseIf Not Intersect(Target, rngLines) Is Nothing And Cells(Target.Row, 1).Interior.color <> RGB(255, 255, 255) Then
        ' Se la cella � nella riga di lavoro e la cella all'inizio non � bianca
        ' Imposta la riga originale e gestisce il blocco degli orari
        OriginalRow = Target.Row
        GoTo BlockLine
    ElseIf Not Intersect(Target, rngLines) Is Nothing And (Target.Value = "P" Or Target.Value = "p" Or content = "P") Then
        ' Se la cella � nella riga di lavoro e il valore � "P" o il suo vecchio valore � "P"
        ' Imposta la riga originale e gestisce il blocco della pausa
        If Target.Value = "p" Then
            Target.Value = "P"
        End If
        OriginalRow = Target.Row - 1
        GoTo Pause
    
    ElseIf Not Intersect(Target, rngLines) Is Nothing And Cells(Target.Row, 1).Interior.color = RGB(255, 255, 255) Then
        ' Se la cella � nella riga di lavoro e la cella all'inizio � bianca
        ' Imposta la riga originale e gestisce il cambio di testo
        If Cells(Target.Row - 1, Target.Column).Value = "N" Or Cells(Target.Row - 1, Target.Column).Value = "I" Or Cells(Target.Row - 1, Target.Column).Value = "F" Then
            OriginalRow = Target.Row - 1
        ElseIf Cells(Target.Row + 1, Target.Column).Value = "N" Or Cells(Target.Row + 1, Target.Column).Value = "I" Or Cells(Target.Row + 1, Target.Column).Value = "F" Then
            OriginalRow = Target.Row + 1
        ElseIf Cells(Target.Row + 2, Target.Column).Value = "P" Then
            OriginalRow = Target.Row + 1
        End If

        If OriginalRow = 0 Then
            Call TextChange(Target)
            GoTo Cleanup
        End If

        CheckName = False

        GoTo TextChange
    ElseIf Target.Column < 6 And Target.Interior.color = RGB(255, 255, 255) Then
        ' Se la cella selezionata non appartiene a nessuno degli intervalli specificati
        ' e la cella all'inizio non e' bianca, cancella il contenuto della cella
        Target.Value = ""
        GoTo Cleanup
    Else
        If Not Intersect(Target, Range("AQ149")) Is Nothing Or Not Intersect(Target, rngResp) Is Nothing Then
            With Target
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.name = "Arial"
                .Font.Bold = True
                .Font.Size = 36
            End With
        End If
        GoTo Cleanup
    End If

    ' Controlla se la formula salvata � valida prima di assegnarla
    If Not IsError(Application.Evaluate(content)) Then
        Target.Formula = content
    Else
        MsgBox "Errore: La formula salvata non e' valida."
    End If

    ' Definisci l'intervallo in cui vuoi cercare (ad esempio, la riga 1)
    Set Times = Range("H1:BT1")

    ColumnFound = 0 ' Inizializza la colonna trovata a 0

    ' Cicla attraverso ogni cella nella riga specificata
    For Each Time In Times
        ' Se l'orario � alle 6, chiedi se di mattina
        If orario = "6:00" Then
            Answer = MsgBox("Le 6:00 di mattina?", vbYesNo + vbQuestion + vbDefaultButton1)
            If Answer = vbYes Then
                ColumnFound = 9
            Else
                ColumnFound = 71
            End If
            Exit For
        ElseIf orario = "6:30" Then
            Answer = MsgBox("Le 6:30 di mattina?", vbYesNo + vbQuestion + vbDefaultButton1)
            If Answer = vbYes Then
                ColumnFound = 11
            Else
                ColumnFound = 72
            End If
            Exit For
        End If
        ' Controlla se la cella contiene il valore uguale a "orario"
        If Time.Value = orario Then
            ColumnFound = Time.Column
            Exit For ' Esce dal ciclo una volta trovata la colonna
        End If
    Next Time

    ' Se � stata trovata la colonna, aggiorna la cella corrispondente
    If ColumnFound > 0 Then
        Cells(OriginalRow, ColumnFound).Value = rangeName
        ' Aggiorna la variabile Column corrispondente
        If rangeName = "I" Then
            columnI = ColumnFound
        Else
            columnF = ColumnFound
        End If
    ElseIf orario = "" Then
        ' Passa oltre se l'orario � vuoto
        ' (Nessuna azione viene eseguita in questo caso)
    Else
        ' Errore se non l'orario non viene trovato
        MsgBox "Errore: Orario non trovato nella riga di ricerca."
    End If

    ' Trova la posizione delle colonne "I" e "F" nella OriginalRow
    If rangeName = "I" Then
        columnI = ColumnFound
        For Each cell In Range(Cells(OriginalRow, 8), Cells(OriginalRow, 72))
            If cell.Value = "F" Then
                columnF = cell.Column
                Exit For
            End If
        Next cell
    ElseIf rangeName = "F" Then
        columnF = ColumnFound
        For Each cell In Range(Cells(OriginalRow, 8), Cells(OriginalRow, 72))
            If cell.Value = "I" Then
                columnI = cell.Column
                Exit For
            End If
        Next cell
    Else
        Call Find_Column_I_F
    End If

    ' Elimina i valori dalla colonna "F" alla colonna "BT" per la OriginalRow
    Range(Cells(OriginalRow, Columns("H").Column), Cells(OriginalRow, Columns("BT").Column)).ClearContents

    ' Ripristina "I" e "F" nelle loro posizioni originali
    If columnI > 0 Then Cells(OriginalRow, columnI).Value = "I"
    If columnF > 0 Then Cells(OriginalRow, columnF).Value = "F"

    ' Cambia le celle comprese tra "I" e "F" in "N"
    If columnI > 0 And columnF > 0 And columnI < columnF Then
        For Column = columnI + 1 To columnF - 1
            If Cells(OriginalRow + 1, Column) <> "P" Then
                Cells(OriginalRow, Column).Value = "N"
            End If
        Next Column
    End If
    
    ' Errore se la collona di I � maggiore della colonna di F
    If columnI > 0 And columnF > 0 And columnI > columnF Then
        MsgBox "Errore: L'ora di inizio e' dopo quella di fine.", vbRetryCancel + vbCritical
        ' Cancella il nome se gi� presente
        Call CancelName(Target)
        Cells(OriginalRow, columnF).Value = ""
        columnF = 0
        Call Paint_Name_Cell
        GoTo Cleanup
    End If
    
    If Cells(OriginalRow, 1).Value <> "" And columnI > 0 And columnF > 0 Then
        Call Calculate_Name_Column(Target)
    End If

    Call Paint_Name_Cell

    If Not (columnI > 0 And columnF > 0) Then
        Call CancelName(Target)
    End If
    
    ' Fine della gestione del blocco
    GoTo Cleanup
  
BlockLine:
    ' Rendi maiuscole le lettere "i", "n" e "f"
    If (Target.Value = "i") Then
        Target.Value = "I"
    ElseIf (Target.Value = "n") Then
        Target.Value = "N"
    ElseIf (Target.Value = "f") Then
        Target.Value = "F"
    End If
    
    ' Se il valore della cella non � "I", "N" o "F", cancella il contenuto
    If (Target.Value <> "I") And (Target.Value <> "N") And (Target.Value <> "F") Then
        Target.Value = ""
    End If

    Call Paint_Name_Cell

    ' Fine della gestione del blocco
    GoTo Cleanup
    
TextChange:
    Call TextChange(Target)

    GoTo NameChange
  
Pause:
    
    ' Se il valore della cella � "P" e la cella sopra � "N"
    ' cancella il contenuto della cella sopra e formatta la cella corrente
    If Target.Value = "P" Then
        If Cells(Target.Row - 1, Target.Column).Value = "N" Then
            Cells(Target.Row - 1, Target.Column).ClearContents
            With Cells(Target.Row, Target.Column)
                .Font.name = "Calibri"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Font.Size = 36
            End With
            If (Target.Column <= 32 And Target.Column >= 24) Or (Target.Column <= 51 And Target.Column >= 43) Then
                ' Colonne di rush
                Target.Interior.color = RGB(217, 217, 217)
            Else
                ' Colonne di lavoro normale
                Target.Interior.color = xlNone
        End If
        Else
            If Cells(Target.Row - 1, 1).Interior.color = RGB(255, 255, 255) Then
                OriginalRow = Target.Row + 1
            End If
            GoTo TextChange
        End If

    ' Se il valore della cella era "P" e la cella sopra � vuota
    ' e la colonna corrente � compresa tra "I" e "F"
    ' imposta il valore a "N" alla cella sopra
    ElseIf content = "P" Then
        If Cells(Target.Row - 1, 1).Interior.color = RGB(255, 255, 255) Then
            OriginalRow = Target.Row + 1
            GoTo TextChange
        End If
        If Cells(Target.Row, 1).Interior.color = RGB(255, 255, 255) Then
            If Cells(Target.Row + 1, Target.Column).Value = "N" Then
                OriginalRow = Target.Row - 1
                Call Find_Column_I_F
                If Not (columnI > 0 And columnF > 0) Then
                    OriginalRow = Target.Row + 1
                    GoTo TextChange
                End If
            Else
                OriginalRow = Target.Row - 1
            End If
            Call Find_Column_I_F
            If Target.Column <= columnI Or Target.Column >= columnF Then
                GoTo TextChange
            End If
        End If
        Call Find_Column_I_F
        If Cells(Target.Row - 1, Target.Column).Value = "" And _
        columnI < Target.Column And Target.Column < columnF Then
            Cells(Target.Row - 1, Target.Column).Value = "N"
        End If
        If Target.Value <> "" Then
            Call TextChange(Target)
        End If
    Else
        ' In caso di errore, mostra un messaggio
        MsgBox "Errore nel blocco pausa."
    End If

    ' Se il nome e' presente allora gestisci il cambio di posizione del nome
    If Cells(OriginalRow, 1).Value <> "" Then
        GoTo NameChange
    End If

NameChange:

    ' Se il flag CheckName e' attivo ed e' presente un nome
    ' controlla se il nome e' gia' presente in un'altra cella
    If CheckName And Cells(OriginalRow, 1).Value <> "" Then
        Dim startRow As Long, endRow As Long
        Dim blockIndex As Long
        
        ' Calcola in quale blocco da 21 righe si trova OriginalRow (partendo da riga 2)
        blockIndex = Int((OriginalRow - 2) / 21)
        startRow = 2 + blockIndex * 21
        endRow = startRow + 20 ' 21 righe
        
        For Each cell In Range(Cells(startRow, 1), Cells(endRow, 1))
            If cell.Value = Cells(OriginalRow, 1).Value And cell.Address <> Cells(OriginalRow, 1).Address Then
                AnswerQuestion = MsgBox("Nome già inserito", vbCritical + vbRetryCancel + vbDefaultButton1)
                If AnswerQuestion = vbRetry Then
                    Cells(OriginalRow, 1).ClearContents
                ElseIf AnswerQuestion = vbCancel Then
                    Cells(OriginalRow, 1).ClearContents
                End If
                Exit For
            End If
        Next cell
    End If

    
    Call Find_Column_I_F
    
    If columnI > 0 And columnF > 0 Then
        Call Paint_Name_Cell
        Call Calculate_Name_Column(Target)
        If CheckName Then
            GoTo Cleanup
        End If
    End If
    
    Dim CheckedRow As Integer
    
    CheckedRow = OriginalRow

    ' Controlla se e' stato modificato il nome di un altra riga
    ' in caso di modifica, aggiorna la riga originale e ripeti il processo
    If CheckedRow < 148 And Cells(CheckedRow + 2, 1).Interior.color <> RGB(255, 255, 255) Then
        OriginalRow = CheckedRow + 2
        Find_Column_I_F
        If columnF > 0 And columnI > 0 Then
            ' Gestione del cambio di posizione del nome
            Call Calculate_Name_Column(Target)
        End If
    End If

    If CheckedRow > 2 And Cells(CheckedRow - 2, 1).Interior.color <> RGB(255, 255, 255) Then
        OriginalRow = CheckedRow - 2
        Find_Column_I_F
        If columnF > 0 And columnI > 0 Then
            ' Gestione del cambio di posizione del nome
            Call Calculate_Name_Column(Target)
        End If
    End If

    ' Fine della gestione del blocco
    GoTo Cleanup

ErrorHandler:
    ' Gestione degli errori
    MsgBox "Errore: " & Err.Description, vbCritical + vbOKOnly + vbDefaultButton1
    GoTo Cleanup

Cleanup:
    ' Disattiva temporaneamente la gestione errori per evitare loop
    On Error Resume Next
    
    ' Ripristina le variabili globali
    columnF = 0
    columnI = 0
    
    ' Riabilita gli eventi, anche in caso di errore
    Application.EnableEvents = True
    
    ' Ripristina la gestione degli errori
    On Error GoTo 0
    
    ' Forza la fine e reset delle variabili
    Exit Sub
    
End Sub