Option Explicit

' Definizione delle variabili globali
Dim content As Variant
Dim orario As String
Dim columnI As Integer
Dim columnF As Integer
Dim Column As Integer
Dim rngStart As Range, rngEnd As Range, rngName As Range, rngLines As Range, rngRoles As Range
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
    ' Se la colonna del nome è già occupata o è un orario di pausa
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
            
            ' Se la colonna supera la colonna "I", permetti la ricerca anche se è un orario di pausa
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
    ' Cancella il nome se già presente in tutte le colonne
    For Each cell In Range(Cells(OriginalRow - 1, 9), Cells(OriginalRow - 1, 73))
        If (cell.Value <> "" And cell.Value = Cells(OriginalRow, 3).Value) Or (Target.Column = 3 And cell.Value = content) Then
            cell.ClearContents
        End If
    Next cell
End Function

Function Calculate_Name_Column(ByVal Target As Range)
    ' Calcola la colonna per il nome
    ColumnNameAproximation = (columnF - columnI) / 2 + columnI
    ColumnName = Round(ColumnNameAproximation)
    
    Call CancelName(Target)

    ' Sistema la colonna del nome se è già occupata
    ' oppure è un orario di pausa
    Call Find_Name_Column

    ' Mette il nome al turno e formatta la cella
    If AllColumnsOccupied = False And ColumnName > 0 And columnI > 0 And columnF > 0 Then
        Cells(OriginalRow - 1, ColumnName).Value = Cells(OriginalRow, 3).Value
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
    ' Se la cella del nome è vuota e le colonne "I" e "F" sono valide
    ' allora evidenziala in giallo
    ' Altrimenti, ripristina il colore originale
    If Cells(OriginalRow, 3).Value = "" And columnI > 0 And columnF > 0 Then
        Cells(OriginalRow, 3).Interior.color = RGB(255, 255, 0)
    Else
        Cells(OriginalRow, 3).Interior.color = RGB(217, 217, 217)
    End If
End Function

Function Find_Column_I_F()
    columnI = 0
    columnF = 0
    ' Trova la posizione delle colonne "I" e "F" nella OriginalRow
    For Each cell In Range(Cells(OriginalRow, 9), Cells(OriginalRow, 73))
        If cell.Value = "I" Then
            columnI = cell.Column
            Exit For
        End If
    Next cell
    
    For Each cell In Range(Cells(OriginalRow, 9), Cells(OriginalRow, 73))
        If cell.Value = "F" Then
            columnF = cell.Column
            Exit For
        End If
    Next cell
End Function

Function TextChange(ByVal Target As Range)
    ' Controlla se il valore della cella è "P" e se la cella sopra è vuota
    If (Target.Value <> "") And (Target.Value <> Cells(Target.Row + 1, 3).Value) Then
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
    ' Se il valore della cella è vuoto
    If Target.Value = "" Then
        ' Ripristina il formato della cella
        If (Target.Column <= 33 And Target.Column >= 25) Or (Target.Column <= 52 And Target.Column >= 44) Then
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
    Set rngStart = Range("E2:E120")
    Set rngEnd = Range("F2:F120")
    Set rngName = Range("C3:C120")
    Set rngLines = Range("I2:BU120")
    Set rngRoles = Range("G2:H120")
    On Error Resume Next
    ' Verifica se è stata selezionata una sola cella
    If Target.Cells.Count = 1 Then
        ' Salva il contenuto della cella selezionata
        content = Target.Formula
    End If
    On Error Resume Next

        If ActiveWindow.SelectedSheets.Count = 1 Then
        ' Controlla se la cella selezionata è all'interno di uno degli intervalli specificati
        ' Se è così, de-protegge il foglio
        If Not Intersect(Target, rngStart) Is Nothing Or Not Intersect(Target, rngEnd) Is Nothing _
        Or Not Intersect(Target, rngName) Is Nothing Or Not Intersect(Target, rngLines) Is Nothing _
        Or Not Intersect(Target, rngRoles) Is Nothing Or Not Intersect(Target, Range("A1:C2")) Is Nothing _
        Or Not Intersect(Target, Range("CL3:CL18")) Is Nothing Then
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
    Set rngStart = Range("E2:E120")
    Set rngEnd = Range("F2:F120")
    Set rngName = Range("C3:C120")
    Set rngLines = Range("I2:BU120")
    Set rngRoles = Range("G2:H120")

        ' Non togliere questa parte, serve per evitare errori
    ' per il menù a tendina.
    ' Tempo speso = 5 Ore
    If Application.Ready = False Then Exit Sub
    If Application.CommandBars("Cell").Enabled = False Then Exit Sub
    
    ' Disabilita temporaneamente gli eventi
    Application.EnableEvents = False

    ' Verifica se la cella cambiata è all'interno di uno degli intervalli specificati
    If Not Intersect(Target, rngStart) Is Nothing And Target.Interior.color <> RGB(255, 255, 255) Then
        ' Se la cella è nella colonna di inizio orario e la cella all'inizio non è bianca
        ' Imposta il valore di "orario" e la riga originale
        orario = Target.Text
        OriginalRow = Target.Row
        rangeName = "I"
    ElseIf Not Intersect(Target, rngEnd) Is Nothing And Target.Interior.color <> RGB(255, 255, 255) Then
        ' Se la cella è nella colonna di fine orario e la cella all'inizio non è bianca
        ' Imposta il valore di "orario" e la riga originale
        orario = Target.Text
        OriginalRow = Target.Row
        rangeName = "F"
    ElseIf Not Intersect(Target, rngName) Is Nothing And Target.Interior.color <> RGB(255, 255, 255) Then
        ' Se la cella è nella colonna del nome e la cella all'inizio non è bianca
        ' Imposta la riga originale e il flag per il controllo del nome
        OriginalRow = Target.Row
        Target.Value = UCase(Target.Value)
        CheckName = True
        GoTo NameChange
    ElseIf Not Intersect(Target, rngLines) Is Nothing And Cells(Target.Row, 3).Interior.color <> RGB(255, 255, 255) Then
        ' Se la cella è nella riga di lavoro e la cella all'inizio non è bianca
        ' Imposta la riga originale e gestisce il blocco degli orari
        OriginalRow = Target.Row
        GoTo BlockLine
    ElseIf Not Intersect(Target, rngLines) Is Nothing And (Target.Value = "P" Or Target.Value = "p" Or content = "P") Then
        ' Se la cella è nella riga di lavoro e il valore è "P" o il suo vecchio valore è "P"
        ' Imposta la riga originale e gestisce il blocco della pausa
        If Target.Value = "p" Then
            Target.Value = "P"
        End If
        OriginalRow = Target.Row - 1
        GoTo Pause
    
    ElseIf Not Intersect(Target, rngLines) Is Nothing And Cells(Target.Row, 3).Interior.color = RGB(255, 255, 255) Then
        ' Se la cella è nella riga di lavoro e la cella all'inizio è bianca
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
    Else
        If Not Intersect(Target, rngRoles) Is Nothing Or Not Intersect(Target, Range("A1:C2")) Then
            With Target
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.name = "Arial"
                .Font.Bold = True
            End With
        End If
        GoTo Cleanup
    End If

    ' Controlla se la formula salvata è valida prima di assegnarla
    If Not IsError(Application.Evaluate(content)) Then
        Target.Formula = content
    Else
        MsgBox "Errore: La formula salvata non è valida."
    End If

    ' Definisci l'intervallo in cui vuoi cercare (ad esempio, la riga 1)
    Set Times = Range("I1:BU1")

    ColumnFound = 0 ' Inizializza la colonna trovata a 0

    ' Cicla attraverso ogni cella nella riga specificata
    For Each Time In Times
        If orario = "5:30" Then
            Answer = MsgBox("Le 5:30 di mattina?", vbYesNo + vbQuestion + vbDefaultButton1)
            If Answer = vbYes Then
                ColumnFound = 9
            Else
                ColumnFound = 71
            End If
            Exit For
        ElseIf orario = "6:00" Then
            Answer = MsgBox("Le 6:00 di mattina?", vbYesNo + vbQuestion + vbDefaultButton1)
            If Answer = vbYes Then
                ColumnFound = 10
            Else
                ColumnFound = 72
            End If
            Exit For
        ElseIf orario = "6:30" Then
            Answer = MsgBox("Le 6:30 di mattina?", vbYesNo + vbQuestion + vbDefaultButton1)
            If Answer = vbYes Then
                ColumnFound = 11
            Else
                ColumnFound = 73
            End If
            Exit For
        ElseIf orario = "7:00" Then
            Answer = MsgBox("Le 7:00 di mattina?", vbYesNo + vbQuestion + vbDefaultButton1)
            If Answer = vbYes Then
                ColumnFound = 14
            Else
                ColumnFound = 74
            End If
            Exit For
        End If
        ' Controlla se la cella contiene il valore uguale a "orario"
        If Time.Value = orario Then
            ColumnFound = Time.Column
            Exit For ' Esce dal ciclo una volta trovata la colonna
        End If
    Next Time

    ' Se è stata trovata la colonna, aggiorna la cella corrispondente
    If ColumnFound > 0 Then
        Cells(OriginalRow, ColumnFound).Value = rangeName
        ' Aggiorna la variabile Column corrispondente
        If rangeName = "I" Then
            columnI = ColumnFound
        Else
            columnF = ColumnFound
        End If
    ElseIf orario = "" Then
        ' Passa oltre se l'orario è vuoto
        ' (Nessuna azione viene eseguita in questo caso)
    Else
        ' Errore se non l'orario non viene trovato
        MsgBox "Errore: Orario non trovato nella riga di ricerca."
    End If

    ' Trova la posizione delle colonne "I" e "F" nella OriginalRow
    If rangeName = "I" Then
        columnI = ColumnFound
        For Each cell In Range(Cells(OriginalRow, 9), Cells(OriginalRow, 73))
            If cell.Value = "F" Then
                columnF = cell.Column
                Exit For
            End If
        Next cell
    ElseIf rangeName = "F" Then
        columnF = ColumnFound
        For Each cell In Range(Cells(OriginalRow, 9), Cells(OriginalRow, 73))
            If cell.Value = "I" Then
                columnI = cell.Column
                Exit For
            End If
        Next cell
    Else
        Call Find_Column_I_F
    End If
    
    ' Elimina i valori dalla colonna "K" alla colonna "BU" per la OriginalRow
    Range(Cells(OriginalRow, Columns("I").Column), Cells(OriginalRow, Columns("BU").Column)).ClearContents

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
    
    ' Errore se la collona di I è maggiore della colonna di F
    If columnI > 0 And columnF > 0 And columnI > columnF Then
        MsgBox "Errore: L'ora di inizio è dopo quella di fine.", vbRetryCancel + vbCritical
        ' Cancella il nome se già presente
        Call CancelName(Target)
        Cells(OriginalRow, columnF).Value = ""
        columnF = 0
        Call Paint_Name_Cell
        GoTo Cleanup
    End If
    
    If Cells(OriginalRow, 3).Value <> "" And columnI > 0 And columnF > 0 Then
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
    
    ' Se il valore della cella non è "I", "N" o "F", cancella il contenuto
    If (Target.Value <> "I") And (Target.Value <> "N") And (Target.Value <> "F") Then
        Target.Value = ""
    End If
    
    ' Fine della gestione del blocco
    GoTo Cleanup
        
TextChange:
    Call TextChange(Target)

    GoTo NameChange

Pause:
    
    ' Se il valore della cella è "P" e la cella sopra è "N"
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
            If (Target.Column <= 33 And Target.Column >= 25) Or (Target.Column <= 52 And Target.Column >= 44) Then
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

    ' Se il valore della cella era "P" e la cella sopra è vuota
    ' e la colonna corrente è compresa tra "I" e "F"
    ' imposta il valore a "N" alla cella sopra
    ElseIf content = "P" Then
        If Cells(Target.Row - 1, 3).Interior.color = RGB(255, 255, 255) Then
            OriginalRow = Target.Row + 1
            GoTo TextChange
        End If
        If Cells(Target.Row, 3).Interior.color = RGB(255, 255, 255) Then
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
    
    ' Se il nome è presente allora gestisci il cambio di posizione del nome
    If Cells(OriginalRow, 3).Value <> "" Then
        GoTo NameChange
    End If

NameChange:
    Call Find_Column_I_F
    
    If columnI > 0 And columnF > 0 Then
        Call Paint_Name_Cell
    
        Call Calculate_Name_Column(Target)
    End If
    
        ' Controlla se il è stato modificato il nome di un altra riga
    ' in caso di modifica, aggiorna la riga originale e ripeti il processo
    If OriginalRow < 119 And Cells(OriginalRow + 2, 3).Interior.color <> RGB(255, 255, 255) Then
        OriginalRow = OriginalRow + 2
        Call Find_Column_I_F
        If columnF > 0 And columnI > 0 Then
            ' Gestione del cambio di posizione del nome
            Call Calculate_Name_Column(Target)
        End If
    End If

    If OriginalRow > 3 And Cells(OriginalRow - 2, 3).Interior.color <> RGB(255, 255, 255) Then
        OriginalRow = OriginalRow - 2
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