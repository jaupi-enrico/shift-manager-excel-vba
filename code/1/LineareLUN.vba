Option Explicit

' Definizione delle variabili globali
Dim content As Variant
Dim orario As String
Dim columnI As Integer
Dim columnF As Integer
Dim Column As Integer
Dim rngC As Range, rngD As Range, rngName As Range, rngLines As Range, rngDiff As Range
Dim rngE As Range
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

' Stato per le modifiche su piu' celle (cancella / incolla blocchi)
Dim InBulk As Boolean
Dim ColumnsMsgShown As Boolean
Dim oldName(15 To 166) As String        ' nomi (colonna A) presenti prima della modifica
Dim oldNameKnown(15 To 166) As Boolean
Dim shiftState(15 To 166) As Byte       ' 0 = non letta, 1 = riga turno, 2 = riga superiore
Dim bulkRow(15 To 166) As Boolean       ' righe turno da ricalcolare
Dim bulkCDel(15 To 166) As Boolean      ' cella C svuotata (inizio cancellato)
Dim bulkDDel(15 To 166) As Boolean      ' cella D svuotata (fine cancellata)
Dim bulkI(15 To 166) As Integer
Dim bulkF(15 To 166) As Integer

Function Find_Name_Column()
    ' Se la colonna del nome ? gi? occupata o ? un orario di pausa
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
            
            ' Se la colonna supera la colonna "I", permetti la ricerca anche se ? un orario di pausa
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
                        ' In una modifica a blocchi il messaggio compare una sola volta
                        If Not (InBulk And ColumnsMsgShown) Then
                            MsgBox "Tutte le colonne sono occupate, libera spazio per visualizzare il nome!", vbCritical + vbOKOnly + vbDefaultButton1
                        End If
                        ColumnsMsgShown = True
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
    ' Cancella il nome se gi? presente in tutte le colonne
    For Each cell In Range(Cells(OriginalRow - 1, 6), Cells(OriginalRow - 1, 70))
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

    ' Sistema la colonna del nome se ? gi? occupata
    ' oppure ? un orario di pausa
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
    ' Se la cella del nome ? vuota e le colonne "I" e "F" sono valide
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
    For Each cell In Range(Cells(OriginalRow, 6), Cells(OriginalRow, 70))
        If cell.Value = "I" Then
            columnI = cell.Column
            Exit For
        End If
    Next cell
    
    For Each cell In Range(Cells(OriginalRow, 6), Cells(OriginalRow, 70))
        If cell.Value = "F" Then
            columnF = cell.Column
            Exit For
        End If
    Next cell
End Function

Function TextChange(ByVal Target As Range)
    ' Controlla se il valore della cella ? "P" e se la cella sopra ? vuota
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
    ' Se il valore della cella ? vuoto
    If Target.Value = "" Then
        ' Ripristina il formato della cella
        If (Target.Column <= 30 And Target.Column >= 22) Or (Target.Column <= 49 And Target.Column >= 41) Then
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
    Password = SHEET_PASSWORD
    ' Definisci gli intervalli di celle
    ' che devono essere de-protetti
    Set rngC = Range("C17:C164")
    Set rngD = Range("D17:D164")
    Set rngName = Range("A17:A164")
    Set rngLines = Range("F16:BR165")
    Set rngDiff = Range("B16:B165")
    Set rngE = Range("E16:E165")
    On Error Resume Next
    ' Verifica se ? stata selezionata una sola cella
    If Target.Cells.Count = 1 Then
        ' Salva il contenuto della cella selezionata
        content = Target.Formula
    End If
    On Error Resume Next

    ' Salva i nomi (colonna A) della selezione: servono a togliere il nome mostrato
    ' nella riga sopra se la selezione viene cancellata o sovrascritta in blocco
    Dim rngOldNames As Range, oc As Range, k As Long
    For k = 15 To 166
        oldNameKnown(k) = False
    Next k
    Set rngOldNames = Intersect(Target, Range("A16:A165"))
    If Not rngOldNames Is Nothing Then
        If rngOldNames.Cells.Count <= 300 Then
            For Each oc In rngOldNames.Cells
                If Not IsError(oc.Value) Then
                    oldName(oc.Row) = CStr(oc.Value)
                    oldNameKnown(oc.Row) = True
                End If
            Next oc
        End If
    End If
    

    ' BUG!!
    If ActiveWindow.SelectedSheets.Count = 1 Then
        ' Controlla se la cella selezionata ? all'interno di uno degli intervalli specificati
        ' Se ? cos?, de-protegge il foglio
        If Not Intersect(Target, rngDiff) Is Nothing Then
            ActiveSheet.Protect Password:=Password
            ActiveSheet.Cells(Target.Row, Target.Column - 1).Activate
        ElseIf Not Intersect(Target, rngE) Is Nothing Then
            ActiveSheet.Protect Password:=Password
            ActiveSheet.Cells(Target.Row, Target.Column - 1).Activate
        ElseIf Not Intersect(Target, rngC) Is Nothing Or Not Intersect(Target, rngD) Is Nothing _
        Or Not Intersect(Target, rngName) Is Nothing Or Not Intersect(Target, rngLines) Is Nothing _
        Or Not Intersect(Target, Range("AN170")) Is Nothing Or Not Intersect(Target, Range("A1")) Is Nothing Then
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
    ' Una cella (o un'area unita) -> logica per singola cella;
    ' piu' celle (cancella / incolla un blocco) -> logica a blocchi
    If Target.Cells.Count > 1 Then
        If Target.Cells(1, 1).MergeArea.Address <> Target.Address Then
            Handle_Multi_Change Target
            Exit Sub
        End If
    End If
    Handle_Single_Change Target
End Sub

' Formula delle colonne C (inizio) e D (fine): ricavano l'orario dalla posizione di "I" / "F"
Private Function Time_Formula(ByVal col As Long, ByVal r As Long) As String
    Dim letter As String
    If col = 3 Then letter = "I" Else letter = "F"
    Time_Formula = "=IFERROR(TIMEVALUE(INDEX($F$1:$BR$1,MATCH(""" & letter & """,$F" & r & ":$BR" & r & ",0))),""OFF"")"
End Function

' Vero se la riga e' una riga turno (nome in A, orari I/N/F); falso per le righe superiori
Private Function Is_Shift_Row(ByVal r As Long) As Boolean
    If r < 17 Or r > 164 Then
        Is_Shift_Row = False
        Exit Function
    End If
    If shiftState(r) = 0 Then
        If Cells(r, 1).Interior.color <> RGB(255, 255, 255) Then
            shiftState(r) = 1
        Else
            shiftState(r) = 2
        End If
    End If
    Is_Shift_Row = (shiftState(r) = 1)
End Function

' Toglie dalla riga sopra a r il nome mostrato (se presente)
Private Sub Clear_Name_Display(ByVal r As Long, ByVal nm As String)
    Dim vals As Variant
    Dim k As Long
    If r < 2 Or nm = "" Then Exit Sub
    vals = Range(Cells(r - 1, 6), Cells(r - 1, 70)).Value
    For k = 1 To 65
        If Not IsError(vals(1, k)) Then
            If CStr(vals(1, k)) = nm Then
                Cells(r - 1, k + 5).ClearContents
                Exit For
            End If
        End If
    Next k
End Sub

' Formatta una cella delle righe superiori (pausa "P" tra I e F, altrimenti commento / reset)
Private Sub Bulk_Format_Upper_Cell(ByVal c As Range)
    Dim insideShift As Boolean
    If IsError(c.Value) Then Exit Sub
    insideShift = False
    If CStr(c.Value) = "P" And c.Row - 1 >= 17 Then
        If bulkI(c.Row - 1) > 0 And bulkF(c.Row - 1) > 0 Then
            If c.Column > bulkI(c.Row - 1) And c.Column < bulkF(c.Row - 1) Then insideShift = True
        End If
    End If
    If insideShift Then
        With c
            .Font.name = "Calibri"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 36
        End With
        If (c.Column <= 30 And c.Column >= 22) Or (c.Column <= 49 And c.Column >= 41) Then
            c.Interior.color = RGB(217, 217, 217)
        Else
            c.Interior.color = xlNone
        End If
    Else
        Call TextChange(c)
    End If
End Sub

' Ricalcola una riga turno dopo una modifica a blocchi: normalizza I/N/F,
' riempie con N tra I e F (tranne sotto una P), ricolloca il nome e colora la cella nome
Private Sub Bulk_Shift_Row(ByVal R As Long)
    Dim vals As Variant, below As Variant
    Dim k As Long, c As Long
    Dim u As String, oldText As String, newText As String
    Dim changed As Boolean
    Dim firstI As Long, firstF As Long

    OriginalRow = R
    columnI = 0
    columnF = 0

    vals = Range(Cells(R, 6), Cells(R, 70)).Value
    changed = False

    ' Come BlockLine: maiuscole, e cancella tutto cio' che non e' I / N / F
    For k = 1 To 65
        If IsError(vals(1, k)) Then
            oldText = "#"
            u = ""
        Else
            oldText = CStr(vals(1, k))
            u = UCase(oldText)
            If u <> "I" And u <> "N" And u <> "F" Then u = ""
        End If
        If oldText <> u Then
            vals(1, k) = u
            changed = True
        End If
    Next k

    ' Orario di inizio / fine cancellato nelle colonne C / D: come una cancellazione
    ' singola, svuota la riga e tiene solo l'altro estremo
    If bulkCDel(R) Or bulkDDel(R) Then
        For k = 1 To 65
            If firstI = 0 And CStr(vals(1, k)) = "I" Then firstI = k
            If firstF = 0 And CStr(vals(1, k)) = "F" Then firstF = k
        Next k
        For k = 1 To 65
            vals(1, k) = ""
        Next k
        If bulkCDel(R) And Not bulkDDel(R) And firstF > 0 Then vals(1, firstF) = "F"
        If bulkDDel(R) And Not bulkCDel(R) And firstI > 0 Then vals(1, firstI) = "I"
        changed = True
    End If

    firstI = 0
    firstF = 0
    For k = 1 To 65
        If firstI = 0 And CStr(vals(1, k)) = "I" Then firstI = k
        If firstF = 0 And CStr(vals(1, k)) = "F" Then firstF = k
    Next k
    If firstI > 0 Then columnI = firstI + 5
    If firstF > 0 Then columnF = firstF + 5

    ' Tra I e F: N, tranne dove sotto c'e' una P (pausa)
    If columnI > 0 And columnF > 0 And columnI < columnF Then
        below = Range(Cells(R + 1, 6), Cells(R + 1, 70)).Value
        For c = columnI + 1 To columnF - 1
            k = c - 5
            newText = "N"
            If Not IsError(below(1, k)) Then
                If CStr(below(1, k)) = "P" Then newText = ""
            End If
            If CStr(vals(1, k)) <> newText Then
                vals(1, k) = newText
                changed = True
            End If
        Next c
    End If

    If changed Then Range(Cells(R, 6), Cells(R, 70)).Value = vals

    ' Inizio dopo la fine: errore, come per la modifica singola
    If columnI > 0 And columnF > 0 And columnI > columnF Then
        If Not ColumnsMsgShown Then
            MsgBox "Errore: L'ora di inizio è dopo quella di fine.", vbCritical
            ColumnsMsgShown = True
        End If
        Call CancelName(Cells(R, 6))
        Cells(R, columnF).Value = ""
        columnF = 0
    ElseIf columnI > 0 And columnF > 0 Then
        If Cells(R, 1).Value <> "" Then
            AllColumnsOccupied = False
            Call Calculate_Name_Column(Cells(R, 6))
        End If
    Else
        Call CancelName(Cells(R, 6))
    End If

    ' Come Paint_Name_Cell, senza rileggere la riga
    If Cells(R, 1).Value = "" And (columnI > 0 Or columnF > 0) Then
        Cells(R, 1).Interior.color = RGB(255, 255, 0)
    Else
        Cells(R, 1).Interior.color = RGB(217, 217, 217)
    End If
End Sub

' Modifica su piu' celle: cancella o incolla un blocco (anche piu' aree con Ctrl)
Private Sub Handle_Multi_Change(ByVal Target As Range)
    Dim calcState As XlCalculation
    Dim ar As Range, c As Range, inter As Range
    Dim r As Long, k As Long
    Dim newName As String
    Dim allNames As Variant
    Dim dupMsgShown As Boolean
    Dim absMsgShown As Boolean
    Dim absenceReason As String

    calcState = Application.Calculation

    On Error GoTo MultiError

    Set rngC = Range("C17:C164")
    Set rngD = Range("D17:D164")
    Set rngName = Range("A17:A164")
    Set rngLines = Range("F16:BR165")

    If Application.Ready = False Then Exit Sub
    If Application.CommandBars("Cell").Enabled = False Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    InBulk = True
    ColumnsMsgShown = False
    dupMsgShown = False

    For r = 15 To 166
        shiftState(r) = 0
        bulkRow(r) = False
        bulkCDel(r) = False
        bulkDDel(r) = False
        bulkI(r) = 0
        bulkF(r) = 0
    Next r

    ' --- Colonna A: nomi ---
    Set inter = Intersect(Target, Range("A16:A165"))
    If Not inter Is Nothing Then
        allNames = Range("A16:A165").Value
        For Each c In inter.Cells
            r = c.Row
            If Is_Shift_Row(r) Then
                ' Toglie il nome vecchio mostrato nella riga sopra
                If oldNameKnown(r) Then Call Clear_Name_Display(r, UCase(oldName(r)))
                If IsError(c.Value) Then
                    c.ClearContents
                    newName = ""
                Else
                    newName = UCase(CStr(c.Value))
                End If
                If newName <> "" Then
                    ' Nome gia' presente in un'altra riga?
                    For k = 1 To 150
                        If k + 15 <> r Then
                            If Not IsError(allNames(k, 1)) Then
                                If UCase(CStr(allNames(k, 1))) = newName Then
                                    If Not dupMsgShown Then
                                        MsgBox "Nome già inserito: " & newName, vbCritical
                                        dupMsgShown = True
                                    End If
                                    c.ClearContents
                                    newName = ""
                                    Exit For
                                End If
                            End If
                        End If
                    Next k
                End If
                If newName <> "" Then
                    ' Blocca i dipendenti in FERIE / MALATTIA / CORSO in questo giorno
                    absenceReason = Lineare_Assenza(Me, newName)
                    If absenceReason <> "" Then
                        If Not absMsgShown Then
                            MsgBox newName & " è in " & absenceReason & " questo giorno e non può essere inserito nel lineare.", vbCritical + vbOKOnly
                            absMsgShown = True
                        End If
                        c.ClearContents
                        newName = ""
                    End If
                End If
                allNames(r - 15, 1) = newName
                If newName <> "" Then
                    If CStr(c.Value) <> newName Then c.Value = newName
                    With c
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .Font.name = "Verdana"
                        .Font.Bold = True
                        .Font.Size = 24
                    End With
                End If
                bulkRow(r) = True
            ElseIf r > 15 Then
                ' Righe superiori: la colonna A deve restare vuota
                If Not IsEmpty(c.Value) Then c.ClearContents
            End If
        Next c
    End If

    ' --- Colonne C e D: orari di inizio / fine (formule) ---
    For k = 3 To 4
        If k = 3 Then
            Set inter = Intersect(Target, rngC)
        Else
            Set inter = Intersect(Target, rngD)
        End If
        If Not inter Is Nothing Then
            For Each c In inter.Cells
                r = c.Row
                If Is_Shift_Row(r) Then
                    If Not c.HasFormula Then
                        If IsEmpty(c.Value) Then
                            ' Cella cancellata: ripristina la formula e segna l'orario come cancellato
                            If k = 3 Then bulkCDel(r) = True Else bulkDDel(r) = True
                            bulkRow(r) = True
                            c.Formula = Time_Formula(k, r)
                        Else
                            ' Orario scritto / incollato come valore: stessa logica della modifica singola
                            content = Time_Formula(k, r)
                            Handle_Single_Change c
                            Application.EnableEvents = False
                            Application.ScreenUpdating = False
                        End If
                    End If
                ElseIf r > 15 Then
                    If Not c.HasFormula Then
                        If Not IsEmpty(c.Value) Then c.ClearContents
                    End If
                End If
            Next c
        End If
    Next k

    ' --- Griglia F:BR: individua le righe turno coinvolte ---
    Set inter = Intersect(Target, rngLines)
    If Not inter Is Nothing Then
        For Each c In inter.Cells
            r = c.Row
            If Is_Shift_Row(r) Then
                bulkRow(r) = True
            Else
                ' Riga superiore: tocca il turno sotto (nome) e quello sopra (pause)
                If Is_Shift_Row(r - 1) Then bulkRow(r - 1) = True
                If Is_Shift_Row(r + 1) Then bulkRow(r + 1) = True
                If Not IsError(c.Value) Then
                    If c.Value = "p" Then c.Value = "P"
                End If
            End If
        Next c
    End If

    ' --- Ricalcola le righe turno ---
    For r = 17 To 164
        If bulkRow(r) Then
            Call Bulk_Shift_Row(r)
            bulkI(r) = columnI
            bulkF(r) = columnF
        End If
    Next r

    ' --- Formatta le celle delle righe superiori comprese nel blocco ---
    If Not inter Is Nothing Then
        For Each c In inter.Cells
            If Not Is_Shift_Row(c.Row) Then Call Bulk_Format_Upper_Cell(c)
        Next c
    End If

    ' Aggiorna i nomi memorizzati, cosi' una seconda cancellazione sullo stesso blocco funziona
    Set inter = Intersect(Target, Range("A16:A165"))
    If Not inter Is Nothing Then
        For Each c In inter.Cells
            If Not IsError(c.Value) Then
                oldName(c.Row) = CStr(c.Value)
                oldNameKnown(c.Row) = True
            End If
        Next c
    End If

    GoTo MultiCleanup

MultiError:
    MsgBox "Errore: " & Err.Description, vbCritical + vbOKOnly + vbDefaultButton1

MultiCleanup:
    On Error Resume Next
    InBulk = False
    columnF = 0
    columnI = 0
    OriginalRow = 0
    Application.Calculation = calcState
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

' Modifica di una sola cella
Private Sub Handle_Single_Change(ByVal Target As Range)

    On Error GoTo ErrorHandler ' Gestione degli errori

    ' Inizializza le variabili
    columnF = 0
    columnI = 0
    OriginalRow = 0
    
    ' Definisci gli intervalli di celle
    ' che devono essere monitorati per le modifiche
    Set rngC = Range("C17:C164")
    Set rngD = Range("D17:D164")
    Set rngName = Range("A17:A164")
    Set rngLines = Range("F16:BR165")
    Set rngDiff = Range("B16:B165")

    ' Non togliere questa parte, serve per evitare errori
    ' per il men? a tendina.
    ' Tempo speso = 5 Ore
    If Application.Ready = False Then Exit Sub
    If Application.CommandBars("Cell").Enabled = False Then Exit Sub

    ' Disabilita temporaneamente gli eventi
    Application.EnableEvents = False
    
    ' Verifica se la cella cambiata ? all'interno di uno degli intervalli specificati
    If Not Intersect(Target, rngC) Is Nothing And Target.Interior.color <> RGB(255, 255, 255) Then
        ' Se la cella ? nella colonna di inizio orario e la cella all'inizio non ? bianca
        ' Imposta il valore di "orario" e la riga originale
        orario = Target.Text
        OriginalRow = Target.Row
        rangeName = "I"
    ElseIf Not Intersect(Target, rngD) Is Nothing And Target.Interior.color <> RGB(255, 255, 255) Then
        ' Se la cella ? nella colonna di fine orario e la cella all'inizio non ? bianca
        ' Imposta il valore di "orario" e la riga originale
        orario = Target.Text
        OriginalRow = Target.Row
        rangeName = "F"
    ElseIf Not Intersect(Target, rngDiff) Is Nothing Then
        Target.Formula = content
        GoTo Cleanup
    ElseIf Not Intersect(Target, rngName) Is Nothing And Target.Interior.color <> RGB(255, 255, 255) Then
        ' Se la cella ? nella colonna del nome e la cella all'inizio non ? bianca
        ' Imposta la riga originale e il flag per il controllo del nome
        OriginalRow = Target.Row
        Target.Value = UCase(Target.Value)
        ' Blocca i dipendenti in FERIE / MALATTIA / CORSO in questo giorno
        Dim AbsenceReason As String
        AbsenceReason = Lineare_Assenza(Me, CStr(Target.Value))
        If AbsenceReason <> "" Then
            MsgBox Target.Value & " è in " & AbsenceReason & " questo giorno e non può essere inserito nel lineare.", vbCritical + vbOKOnly
            ' Ripristina il nome precedente (se c'era) cosi' il resto del foglio resta coerente
            Target.Value = content
            GoTo Cleanup
        End If


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
        ' Se la cella ? nella riga di lavoro e la cella all'inizio non ? bianca
        ' Imposta la riga originale e gestisce il blocco degli orari
        OriginalRow = Target.Row
        GoTo BlockLine
    ElseIf Not Intersect(Target, rngLines) Is Nothing And (Target.Value = "P" Or Target.Value = "p" Or content = "P") Then
        ' Se la cella ? nella riga di lavoro e il valore ? "P" o il suo vecchio valore ? "P"
        ' Imposta la riga originale e gestisce il blocco della pausa
        If Target.Value = "p" Then
            Target.Value = "P"
        End If
        OriginalRow = Target.Row - 1
        GoTo Pause
    ElseIf Not Intersect(Target, rngLines) Is Nothing And Cells(Target.Row, 1).Interior.color = RGB(255, 255, 255) Then
        ' Se la cella ? nella riga di lavoro e la cella all'inizio ? bianca
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
    ElseIf Target.Column < 6 And Target.Row > 15 And Target.Interior.color = RGB(255, 255, 255) Then
        ' Se la cella selezionata non appartiene a nessuno degli intervalli specificati
        ' e la cella all'inizio non e' bianca, cancella il contenuto della cella
        Target.Value = ""
        GoTo Cleanup
    Else
        If Not Intersect(Target, Range("AN170")) Is Nothing Or Not Intersect(Target, Range("A1")) Is Nothing Then
            With Target
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.name = "Arial"
                .Font.Bold = True
                .Font.Size = 36
            End With

            If Target.Address = Range("A1").Address Then
                    If IsDate(Target.Value) Then
                        Target.Value = CDate(Target.Value)  ' Mantiene il valore come data
                        Target.NumberFormat = "[$-0410]dddd dd mmmm yyyy" ' Mostra "lunedì 13 ottobre 2025"
                    Else
                        MsgBox "Errore: Inserire una data valida in formato 'gg/mm/aaaa'.", vbExclamation
                        Target.Value = DateSerial(2024, 10, 10) ' imposta una data di default
                        Target.NumberFormat = "[$-0410]dddd dd mmmm yyyy"
                    End If
                End If
        End If
        GoTo Cleanup
    End If

    ' Controlla se la formula salvata ? valida prima di assegnarla
    If Not IsError(Application.Evaluate(content)) Then
        Target.Formula = content
    Else
        MsgBox "Errore: La formula salvata non è valida."
    End If

    ' Definisci l'intervallo in cui vuoi cercare (ad esempio, la riga 1)
    Set Times = Range("F1:BR1")

    ColumnFound = 0 ' Inizializza la colonna trovata a 0

    ' Cicla attraverso ogni cella nella riga specificata
    For Each Time In Times
        ' Se l'orario ? alle 6, chiedi se di mattina
        If orario = "6:00" Then
            Answer = MsgBox("Le 6:00 di mattina?", vbYesNo + vbQuestion + vbDefaultButton1)
            If Answer = vbYes Then
                ColumnFound = 7
            Else
                ColumnFound = 69
            End If
            Exit For
        ElseIf orario = "6:30" Then
            Answer = MsgBox("Le 6:30 di mattina?", vbYesNo + vbQuestion + vbDefaultButton1)
            If Answer = vbYes Then
                ColumnFound = 9
            Else
                ColumnFound = 70
            End If
            Exit For
        End If
        ' Controlla se la cella contiene il valore uguale a "orario"
        If Time.Value = orario Then
            ColumnFound = Time.Column
            Exit For ' Esce dal ciclo una volta trovata la colonna
        End If
    Next Time

    ' Se ? stata trovata la colonna, aggiorna la cella corrispondente
    If ColumnFound > 0 Then
        Cells(OriginalRow, ColumnFound).Value = rangeName
        ' Aggiorna la variabile Column corrispondente
        If rangeName = "I" Then
            columnI = ColumnFound
        Else
            columnF = ColumnFound
        End If
    ElseIf orario = "" Then
        ' Passa oltre se l'orario ? vuoto
        ' (Nessuna azione viene eseguita in questo caso)
    Else
        ' Errore se non l'orario non viene trovato
        MsgBox "Errore: Orario non trovato nella riga di ricerca."
    End If

    ' Trova la posizione delle colonne "I" e "F" nella OriginalRow
    If rangeName = "I" Then
        columnI = ColumnFound
        For Each cell In Range(Cells(OriginalRow, 6), Cells(OriginalRow, 70))
            If cell.Value = "F" Then
                columnF = cell.Column
                Exit For
            End If
        Next cell
    ElseIf rangeName = "F" Then
        columnF = ColumnFound
        For Each cell In Range(Cells(OriginalRow, 6), Cells(OriginalRow, 70))
            If cell.Value = "I" Then
                columnI = cell.Column
                Exit For
            End If
        Next cell
    Else
        Call Find_Column_I_F
    End If
    
    ' Elimina i valori dalla colonna "F" alla colonna "BR" per la OriginalRow
    Range(Cells(OriginalRow, Columns("F").Column), Cells(OriginalRow, Columns("BR").Column)).ClearContents

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
    
    ' Errore se la collona di I ? maggiore della colonna di F
    If columnI > 0 And columnF > 0 And columnI > columnF Then
        MsgBox "Errore: L'ora di inizio è dopo quella di fine.", vbRetryCancel + vbCritical
        ' Cancella il nome se gi? presente
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
    
    ' Se il valore della cella non ? "I", "N" o "F", cancella il contenuto
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
    
    ' Se il valore della cella ? "P" e la cella sopra ? "N"
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
            If (Target.Column <= 30 And Target.Column >= 22) Or (Target.Column <= 49 And Target.Column >= 41) Then
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

    ' Se il valore della cella era "P" e la cella sopra ? vuota
    ' e la colonna corrente ? compresa tra "I" e "F"
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

    ' Se il nome ? presente allora gestisci il cambio di posizione del nome
    If Cells(OriginalRow, 1).Value <> "" Then
        GoTo NameChange
    End If

NameChange:

    ' Se il flag CheckName ? attivo ed ? presente un nome
    ' controlla se il nome ? gi? presente in un'altra cella
    If CheckName And Cells(OriginalRow, 1).Value <> "" Then
        For Each cell In Range(Cells(16, 1), Cells(165, 1))
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
    
    ' Controlla se il ? stato modificato il nome di un altra riga
    ' in caso di modifica, aggiorna la riga originale e ripeti il processo
    If CheckedRow < 164 And Cells(CheckedRow + 2, 1).Interior.color <> RGB(255, 255, 255) Then
        OriginalRow = CheckedRow + 2
        Find_Column_I_F
        If columnF > 0 And columnI > 0 Then
            ' Gestione del cambio di posizione del nome
            Call Calculate_Name_Column(Target)
        End If
    End If

    If CheckedRow > 17 And Cells(CheckedRow - 2, 1).Interior.color <> RGB(255, 255, 255) Then
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
    OriginalRow = 0
    
    ' Riabilita gli eventi, anche in caso di errore
    Application.EnableEvents = True
    
    ' Ripristina la gestione degli errori
    On Error GoTo 0
    
    ' Forza la fine e reset delle variabili
    Exit Sub
    
End Sub