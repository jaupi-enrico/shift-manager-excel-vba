Option Explicit
' Aggiunge un lavoratore alle tabelle TOT e FORMAZIONE
Sub Add_Worker(ByVal WorkerName As String, ByVal WorkerName_Surname As String, ByVal Contract As Integer, ByVal WorkerPos As Integer)
    Dim wsTOT As Worksheet, wsFORM As Worksheet
    Dim LastRow As Integer, i As Integer

    Set wsTOT = Worksheets("TOT")
    Set wsFORM = Worksheets("FORMAZIONE")

    LastRow = 4
    While wsTOT.Cells(LastRow, 2).Value <> "IMPRESA1"
        LastRow = LastRow + 1
    Wend

    If WorkerPos < 0 Or WorkerPos + 3 > Rows.Count Then
        MsgBox "Errore: posizione lavoratore non valida (" & WorkerPos & ")", vbCritical
        Exit Sub
    End If

    wsTOT.Rows(WorkerPos + 3).Insert Shift:=xlDown
    DoEvents
    wsFORM.Rows(WorkerPos + 3).Insert Shift:=xlDown
    DoEvents

    If wsTOT.Cells(LastRow, 30).Value <> "" Then
        wsTOT.Rows(LastRow).Copy
        wsTOT.Rows(WorkerPos + 3).PasteSpecial xlPasteAll
        wsFORM.Rows(LastRow).Copy
        wsFORM.Rows(WorkerPos + 3).PasteSpecial xlPasteAll
        DoEvents
    Else
        wsTOT.Rows(4).Copy
        wsTOT.Rows(WorkerPos + 3).PasteSpecial xlPasteAll
        wsFORM.Rows(4).Copy
        wsFORM.Rows(WorkerPos + 3).PasteSpecial xlPasteAll
        DoEvents
    End If


    wsTOT.Cells(WorkerPos + 3, 1).Value = WorkerPos
    wsFORM.Cells(WorkerPos + 3, 1).Value = WorkerPos

    wsTOT.Cells(WorkerPos + 3, 2).Value = WorkerName
    wsTOT.Cells(WorkerPos + 3, 26).Value = WorkerName
    wsTOT.Cells(WorkerPos + 3, 33).Value = WorkerName
    wsTOT.Cells(WorkerPos + 3, 3).Value = WorkerName_Surname
    wsTOT.Cells(WorkerPos + 3, 34).Value = WorkerName_Surname
    wsTOT.Cells(WorkerPos + 3, 53).Value = WorkerName_Surname
    wsFORM.Cells(WorkerPos + 3, 2).Value = WorkerName_Surname
    wsFORM.Cells(WorkerPos + 3, 19).Value = WorkerName_Surname
    wsFORM.Cells(WorkerPos + 3, 58).Value = WorkerName_Surname

    wsTOT.Cells(WorkerPos + 3, 27).Value = Contract

    wsTOT.Range(wsTOT.Cells(WorkerPos + 3, 35), wsTOT.Cells(WorkerPos + 3, 50)).Value = ""
    wsFORM.Range(wsFORM.Cells(WorkerPos + 3, 3), wsFORM.Cells(WorkerPos + 3, 16)).Value = ""
    wsFORM.Range(wsFORM.Cells(WorkerPos + 3, 20), wsFORM.Cells(WorkerPos + 3, 55)).Value = ""

    For i = WorkerPos To LastRow - 3
        wsTOT.Cells(i + 3, 1).Value = i
        wsFORM.Cells(i + 3, 1).Value = i
    Next i

    wsFORM.Activate: wsFORM.Range("A1").Activate
    wsTOT.Activate: wsTOT.Range("A1").Activate
End Sub

' Elimina un lavoratore
Sub Delete_Worker(ByVal WorkerPos As Integer)
    Dim wsTOT As Worksheet, wsFORM As Worksheet, LastRow As Integer, i As Integer
    Set wsTOT = Worksheets("TOT")
    Set wsFORM = Worksheets("FORMAZIONE")

    wsTOT.Rows(WorkerPos + 3).Delete
    wsFORM.Rows(WorkerPos + 3).Delete
    DoEvents

    LastRow = 4
    While wsTOT.Cells(LastRow, 2).Value <> "IMPRESA1" And LastRow < Rows.Count
        LastRow = LastRow + 1
    Wend

    For i = WorkerPos To LastRow - 4
        wsTOT.Cells(i + 3, 1).Value = i
        wsFORM.Cells(i + 3, 1).Value = i
    Next i

    wsFORM.Activate: wsFORM.Range("A1").Activate
    wsTOT.Activate: wsTOT.Range("A1").Activate
End Sub

' Trova un nome nella colonna 2 del foglio TOT a partire da una riga specifica
Function NameFound(ByVal WorkerName As String, ByVal Start As String) As Integer
    Dim wsTOT As Worksheet, WorkerRow As Integer
    Set wsTOT = Worksheets("TOT")
    WorkerRow = Start + 1
    While wsTOT.Cells(WorkerRow, 2).Value <> "IMPRESA1"
        If wsTOT.Cells(WorkerRow, 2).Value = WorkerName Then
            NameFound = WorkerRow
            Exit Function
        End If
        WorkerRow = WorkerRow + 1
    Wend
    NameFound = -1
End Function

' Trasferisce i dati di orari/formazione
Function Transfer_data(ByVal OldPos As Integer, ByVal NewPos As Integer)
    Dim wsTOT As Worksheet
    Dim wsFORM As Worksheet

    Set wsTOT = Worksheets("TOT")
    Set wsFORM = Worksheets("FORMAZIONE")

    With wsTOT
        .Range(.Cells(OldPos, 35), .Cells(OldPos, 48)).Copy
        .Cells(NewPos, 35).PasteSpecial Paste:=xlPasteAll
    End With

    With wsFORM
        .Range(.Cells(OldPos, 3), .Cells(OldPos, 16)).Copy
        .Cells(NewPos, 3).PasteSpecial Paste:=xlPasteAll

        .Range(.Cells(OldPos, 20), .Cells(OldPos, 55)).Copy
        .Cells(NewPos, 20).PasteSpecial Paste:=xlPasteAll

        .Range(.Cells(OldPos, 59), .Cells(OldPos, 72)).Copy
        .Cells(NewPos, 59).PasteSpecial Paste:=xlPasteAll
    End With

    Application.CutCopyMode = False
End Function

Function Paint_Worker(ByVal Pos As Integer, ByVal Role As Integer)
    Dim wsTOT As Worksheet
    Dim wsFORM As Worksheet
    Set wsTOT = Worksheets("TOT")
    Set wsFORM = Worksheets("FORMAZIONE")

    Dim color As Long
    Select Case Role
        Case 1: color = RGB(255, 242, 204)
        Case 2: color = RGB(255, 255, 255)
        Case 3: color = RGB(221, 235, 247)
        Case 4: color = RGB(252, 228, 214)
    End Select

    With wsTOT
        .Range(.Cells(Pos, 2), .Cells(Pos, 27)).Interior.color = color
        .Range(.Cells(Pos, 29), .Cells(Pos, 30)).Interior.color = color
        .Range(.Cells(Pos, 33), .Cells(Pos, 48)).Interior.color = color
        .Range(.Cells(Pos, 53), .Cells(Pos, 67)).Interior.color = color
    End With

    With wsFORM
        .Range(.Cells(Pos, 2), .Cells(Pos, 16)).Interior.color = color
        .Range(.Cells(Pos, 19), .Cells(Pos, 55)).Interior.color = color
        .Range(.Cells(Pos, 58), .Cells(Pos, 72)).Interior.color = color
    End With
End Function

Function TrovaRigaValida(ws As Worksheet, colOffset As Integer, workerRow As Integer, LastRowMax As Integer) As Integer
    Dim i As Integer

    ' Cerca in basso
    For i = workerRow + 1 To LastRowMax
        If ws.Cells(i, colOffset).Value <> "FERIE" And _
           ws.Cells(i, colOffset).Value <> "MALATTIA" And _
           ws.Cells(i, colOffset).Value <> "CORSO" Then
            TrovaRigaValida = i
            Exit Function
        End If
    Next i

    ' Cerca in alto
    For i = workerRow - 1 To 2 Step -1
        If ws.Cells(i, colOffset).Value <> "FERIE" And _
           ws.Cells(i, colOffset).Value <> "MALATTIA" And _
           ws.Cells(i, colOffset).Value <> "CORSO" Then
            TrovaRigaValida = i
            Exit Function
        End If
    Next i

    TrovaRigaValida = -1 ' Nessuna riga valida trovata
End Function

Sub Check_Days(ByVal Worker As Integer, ByVal WorkerRole As Integer)
    Dim wsDip As Worksheet, wsTOT As Worksheet
    Dim i As Integer, colOffset As Integer
    Dim label As String, rigaValida As Integer
    Dim LastRowMax As Integer

    Set wsDip = Worksheets("Dipendenti")
    Set wsTOT = Worksheets("TOT")

    ' Trova l'ultima riga non vuota della colonna A nel foglio TOT
    LastRowMax = 4
    While wsTOT.Cells(LastRowMax, 1).Value <> "" And LastRowMax < Rows.Count
        LastRowMax = LastRowMax + 1
    Wend
    LastRowMax = LastRowMax - 1

    ' FERIE (colonne 10–16)
    For i = 10 To 16
        colOffset = 4 + (i - 10) * 2
        label = "FERIE"

        If wsDip.Cells(Worker, i).Value = "Si" And wsTOT.Cells(Worker + 1, colOffset).Value <> label Then
            wsTOT.Range(wsTOT.Cells(Worker + 1, colOffset), wsTOT.Cells(Worker + 1, colOffset + 1)).Merge
            wsTOT.Cells(Worker + 1, colOffset).Value = label

        ElseIf wsDip.Cells(Worker, i).Value = "No" And wsTOT.Cells(Worker + 1, colOffset).Value = label Then
            wsTOT.Range(wsTOT.Cells(Worker + 1, colOffset), wsTOT.Cells(Worker + 1, colOffset + 1)).UnMerge

            rigaValida = TrovaRigaValida(wsTOT, colOffset, Worker + 1, LastRowMax)
            If rigaValida <> -1 Then
                wsTOT.Range(wsTOT.Cells(rigaValida, colOffset), wsTOT.Cells(rigaValida, colOffset + 1)).Copy
                wsTOT.Range(wsTOT.Cells(Worker + 1, colOffset), wsTOT.Cells(Worker + 1, colOffset + 1)).PasteSpecial xlPasteAll
            End If
        End If
    Next i

    ' MALATTIA (colonne 18–24)
    For i = 18 To 24
        colOffset = 4 + (i - 18) * 2
        label = "MALATTIA"

        If wsDip.Cells(Worker, i).Value = "Si" And wsTOT.Cells(Worker + 1, colOffset).Value <> label Then
            wsTOT.Range(wsTOT.Cells(Worker + 1, colOffset), wsTOT.Cells(Worker + 1, colOffset + 1)).Merge
            wsTOT.Cells(Worker + 1, colOffset).Value = label

        ElseIf wsDip.Cells(Worker, i).Value = "No" And wsTOT.Cells(Worker + 1, colOffset).Value = label Then
            wsTOT.Range(wsTOT.Cells(Worker + 1, colOffset), wsTOT.Cells(Worker + 1, colOffset + 1)).UnMerge

            rigaValida = TrovaRigaValida(wsTOT, colOffset, Worker + 1, LastRowMax)
            If rigaValida <> -1 Then
                wsTOT.Range(wsTOT.Cells(rigaValida, colOffset), wsTOT.Cells(rigaValida, colOffset + 1)).Copy
                wsTOT.Range(wsTOT.Cells(Worker + 1, colOffset), wsTOT.Cells(Worker + 1, colOffset + 1)).PasteSpecial xlPasteAll
            End If
        End If
    Next i

    ' CORSO (colonne 26–32)
    For i = 26 To 32
        colOffset = 4 + (i - 26) * 2
        label = "CORSO"

        If wsDip.Cells(Worker, i).Value = "Si" And wsTOT.Cells(Worker + 1, colOffset).Value <> label Then
            wsTOT.Range(wsTOT.Cells(Worker + 1, colOffset), wsTOT.Cells(Worker + 1, colOffset + 1)).Merge
            wsTOT.Cells(Worker + 1, colOffset).Value = label

        ElseIf wsDip.Cells(Worker, i).Value = "No" And wsTOT.Cells(Worker + 1, colOffset).Value = label Then
            wsTOT.Range(wsTOT.Cells(Worker + 1, colOffset), wsTOT.Cells(Worker + 1, colOffset + 1)).UnMerge

            rigaValida = TrovaRigaValida(wsTOT, colOffset, Worker + 1, LastRowMax)
            If rigaValida <> -1 Then
                wsTOT.Range(wsTOT.Cells(rigaValida, colOffset), wsTOT.Cells(rigaValida, colOffset + 1)).Copy
                wsTOT.Range(wsTOT.Cells(Worker + 1, colOffset), wsTOT.Cells(Worker + 1, colOffset + 1)).PasteSpecial xlPasteAll
            End If
        End If
    Next i

    ' Colore intestazione
    If wsDip.Cells(Worker, 1).Value = "Si" Then
        wsTOT.Cells(Worker + 1, 1).Interior.color = RGB(241, 170, 131)
    ElseIf wsDip.Cells(Worker, 1).Value = "No" And wsTOT.Cells(Worker + 1, 1).Interior.color = RGB(241, 170, 131) Then
        rigaValida = 4
        While wsTOT.Cells(rigaValida, 1).Interior.color = RGB(241, 170, 131) And rigaValida <= LastRowMax
            rigaValida = rigaValida + 1
        Wend
        wsTOT.Range(wsTOT.Cells(rigaValida, 4), wsTOT.Cells(Worker + 2, 17)).Copy
        wsTOT.Range(wsTOT.Cells(Worker + 1, 4), wsTOT.Cells(Worker + 1, 17)).PasteSpecial xlPasteAll
        Call Paint_Worker(Worker + 1, WorkerRole)
        wsTOT.Cells(Worker + 1, 1).Interior.color = RGB(255, 255, 255)
    End If
End Sub

Function Update_Validation()
    Dim ws As Worksheet
    Dim wsTOT As Worksheet
    Dim LastRow As Long
    Dim addressList As String
    Dim rngValidazione As Range
    
    Set wsTOT = ThisWorkbook.Sheets("TOT")
    
    ' Trova l'ultima riga non vuota della colonna A nel foglio TOT
    LastRow = 1

    While wsTOT.Cells(LastRow, 2).Value <> "IMPRESA1" And LastRow < Rows.Count
        LastRow = LastRow + 1
    Wend
    LastRow = LastRow + 3
    
    ' Costruisci l'indirizzo dell'intervallo da usare nella validazione
    addressList = "=TOT!$B$4:$B$" & LastRow
    
    ' Applica la convalida nei fogli dei giorni
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "LUN" Or ws.name = "MAR" Or ws.name = "MER" Or _
           ws.name = "GIO" Or ws.name = "VEN" Or ws.name = "SAB" Or ws.name = "DOM" Then
            
            Set rngValidazione = ws.Range("A16:A165") ' Adatta se necessario
            
            On Error Resume Next
            rngValidazione.Validation.Delete
            On Error GoTo 0
            
            With rngValidazione.Validation
                .Add Type:=xlValidateList, _
                     AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, _
                     Formula1:=addressList
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
        End If
    Next ws
End Function

Sub UpdateFormatConditions()
    Call ShowSheets
    Dim wsTOT As Worksheet
    Dim rng As Range, rngTemp As Range
    Dim LastRow As Long, i As Long
    Dim fc As FormatCondition
    Dim col1 As String, col2 As String, formula As String
    Dim col3 As String, col4 As String
    
    Set wsTOT = ThisWorkbook.Sheets("TOT")
    
    ' Trova ultima riga colonna A
    LastRow = wsTOT.Cells(wsTOT.Rows.Count, 1).End(xlUp).Row
    
    ' Definisci intervallo
    Set rng = wsTOT.Range("D4:Q" & LastRow)
    
    ' Cancella tutte le formattazioni condizionali
    rng.FormatConditions.Delete
    
    ' Ciclo sulle coppie di colonne (E:F, G:H, I:J, ...)
    For i = 0 To (rng.Columns.Count \ 2) - 2
        Set rngTemp = rng.Offset(0, i * 2 + 1).Resize(rng.Rows.Count, 2)
        
        ' Converte l’indice di colonna in lettera (es. 5 in "E")
        col1 = Split(wsTOT.Cells(1, rngTemp.Columns(1).Column).Address(True, False), "$")(0)
        col2 = Split(wsTOT.Cells(1, rngTemp.Columns(2).Column).Address(True, False), "$")(0)
        col3 = Split(wsTOT.Cells(1, rngTemp.Columns(0).Column).Address(True, False), "$")(0)

        ' Formula condizionale dinamica
        formula = "=E(" & _
                  "O($" & col1 & "4<>""OFF"";$" & col2 & "4<>""OFF"");" & _
                  "O(" & _
                     "E($" & col3 & "4>$" & col1 & "4;$" & col2 & "4-$" & col1 & "4<0,46);" & _
                     "E($" & col3 & "4<$" & col1 & "4;(1-$" & col1 & "4)+($" & col2 & "4)<0,46)" & _
                  ")" & _
                 ")"
        
        ' Applica la regola
        Set fc = rngTemp.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
        fc.Interior.color = RGB(255, 0, 0) ' Rosso
        fc.Font.Bold = True
        fc.Font.color = RGB(255, 255, 255) ' Bianco
    Next i

    For i = 0 To 4
        Set rngTemp = rng.Offset(0, i * 2 + 1).Resize(rng.Rows.Count, 4)
        
        ' Converte l’indice di colonna in lettera (es. 5 in "E")
        col1 = Split(wsTOT.Cells(1, rngTemp.Columns(0).Column).Address(True, False), "$")(0)
        col2 = Split(wsTOT.Cells(1, rngTemp.Columns(1).Column).Address(True, False), "$")(0)
        col3 = Split(wsTOT.Cells(1, rngTemp.Columns(2).Column).Address(True, False), "$")(0)
        col4 = Split(wsTOT.Cells(1, rngTemp.Columns(4).Column).Address(True, False), "$")(0)

        ' Formula condizionale dinamica
        formula = "=E($" & col2 & "4<>""OFF"";$" & col3 & "4=""OFF"";$" & col4 & "4<>""OFF"";" & _
                  "O(E($" & col1 & "4>$" & col2 & "4;$" & col4 & "4-$" & col2 & "4<0,46);" & _
                  "E($" & col1 & "4<$" & col2 & "4;(1-$" & col2 & "4)+$" & col4 & "4<0,46)))"
        ' Applica la regola
        Set fc = rngTemp.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
        fc.Interior.color = RGB(255, 0, 0) ' Rosso
        fc.Font.Bold = True
        fc.Font.color = RGB(255, 255, 255) ' Bianco
    Next i

    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""OFF""")
    fc.Font.Bold = True
    fc.Font.color = RGB(51, 51, 255) ' Blu
End Sub

' Aggiorna tutti i lavoratori dal foglio Dipendenti
Sub Update_Workers()
    Call ShowSheets

    Update.Show vbModeless
    DoEvents ' Permette il disegno della form prima di proseguire

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.StatusBar = "Aggiornamento in corso..."

    Dim wsDip As Worksheet, wsTOT As Worksheet
    Dim Worker As Integer, WorkerName As String, WorkerName_Surname As String
    Dim WorkerRole As Integer, WorkerContract As Integer, LastRow As Integer, WorkerPos As Integer

    Set wsDip = Worksheets("Dipendenti")
    Set wsTOT = Worksheets("TOT")
    Worker = 3

    While wsDip.Cells(Worker, 3).Value <> ""
        WorkerName = wsDip.Cells(Worker, 3).Value
        WorkerName_Surname = wsDip.Cells(Worker, 4).Value
        WorkerContract = wsDip.Cells(Worker, 5).Value

        Select Case wsDip.Cells(Worker, 2).Value
            Case "Gel": WorkerRole = 1
            Case "Front": WorkerRole = 2
            Case "Tutto": WorkerRole = 3
            Case "Cucina": WorkerRole = 4
        End Select

        Dim color As Long
        Select Case WorkerRole
            Case 1: color = RGB(255, 242, 204)
            Case 2: color = RGB(255, 255, 255)
            Case 3: color = RGB(221, 235, 247)
            Case 4: color = RGB(252, 228, 214)
        End Select

        If NameFound(WorkerName, 3) <> -1 Then
            WorkerPos = NameFound(WorkerName, 3) - 3
            If WorkerPos = Worker - 2 And _
            wsTOT.Cells(WorkerPos + 3, 3).Value = WorkerName_Surname And _
            wsTOT.Cells(WorkerPos + 3, 27).Value = WorkerContract And _
            wsTOT.Cells(WorkerPos + 3, 2).Interior.color = color Then
                Call Check_Days(Worker, WorkerRole)
            Else
                Call Add_Worker(WorkerName, WorkerName_Surname, WorkerContract, Worker - 2)
                WorkerPos = NameFound(WorkerName, Worker + 1) - 3
                Call Transfer_data(WorkerPos + 3, Worker + 1)
                Call Delete_Worker(WorkerPos)
                Call Check_Days(Worker, WorkerRole)
                Call Paint_Worker(Worker + 1, WorkerRole)
            End If
        Else
            Call Add_Worker(WorkerName, WorkerName_Surname, WorkerContract, Worker - 2)
            Call Paint_Worker(Worker + 1, WorkerRole)
            Call Check_Days(Worker, WorkerRole)
        End If
        Worker = Worker + 1
    Wend

    LastRow = 4
    While wsTOT.Cells(LastRow, 2).Value <> "IMPRESA1" And LastRow < Rows.Count
        LastRow = LastRow + 1
    Wend

    While Worker - 2 <> LastRow - 3
        Call Delete_Worker(Worker - 2)
        LastRow = LastRow - 1
    Wend

    Call Update_Validation

    Call UpdateFormatConditions

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

    Call HideSheets
    Update.Hide

    ' Mostra un messaggio di completamento
    MsgBox "Aggiornamento completato", vbInformation, "Aggiornamento lavoratori"
End Sub

Sub PrintRange()
    Dim ws As Worksheet
    Dim wsTemp As Worksheet
    Dim col As Integer

    Dim LastRow As Integer
    LastRow = 4
    While Cells(LastRow, 1) <> ""
        LastRow = LastRow + 1
    Wend
    LastRow = LastRow + 2

    ' Variabili per l'input dell'utente
    Dim scelta As Integer
    Dim rangeDaStampare As Range
    
    ' Imposta i fogli da stampare
    Set ws = ThisWorkbook.Sheets("TOT")

    ' Mostra una finestra di dialogo per scegliere tra due opzioni
    scelta = MsgBox("Scegli quale intervallo stampare:" & vbCrLf & _
                    "Si: Con orari" & vbCrLf & _
                    "No: Senza orari", vbYesNoCancel + vbQuestion, "Scegli intervallo")
    
    ' Controllo la scelta dell'utente
    If scelta = vbYes Then
        Set rangeDaStampare = ws.Range(Cells(1, 3), Cells(LastRow, 17))
    ElseIf scelta = vbNo Then
        Set rangeDaStampare = ws.Range(Cells(1, 53), Cells(LastRow, 67))
    Else
        MsgBox "Operazione annullata."
        Exit Sub
    End If


    ' Mostra fogli
    Call ShowSheets

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("___TEMP_STAMPA___").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Crea foglio temporaneo per stampa
    Set wsTemp = ThisWorkbook.Sheets.Add(After:=ws)
    wsTemp.name = "___TEMP_STAMPA___"

    ' === TOT-M ===

    Dim rngOrigine As Range, rngDest As Range
    Set rngOrigine = rangeDaStampare
    Set rngDest = wsTemp.Range("A1")

    ' Copia valori
    rngDest.Resize(rngOrigine.Rows.Count, rngOrigine.Columns.Count).Value = rngOrigine.Value

    ' Copia formato (colori, bordi, font, ecc.)
    rngOrigine.Copy
    rngDest.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Copia larghezze colonne
    Dim origineCol As Integer, tempCol As Integer

    For col = 1 To rngOrigine.Columns.Count
        origineCol = rngOrigine.Columns(col).Column
        tempCol = rngDest.Columns(col).Column
        wsTemp.Columns(tempCol).ColumnWidth = ws.Columns(origineCol).ColumnWidth
    Next col

    ' Copia altezze righe
    Dim r As Long
    For r = 1 To rngOrigine.Rows.Count
        wsTemp.Rows(r).RowHeight = ws.Rows(r).RowHeight
    Next r

    ' Rimuovi eventuali formattazioni condizionali ereditate
    wsTemp.Cells.FormatConditions.Delete
    DoEvents

    ' Imposta stampa su una sola pagina
    With wsTemp.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
    End With
    
    DoEvents
    ' Stampa
    On Error Resume Next
    wsTemp.PrintOut
    On Error GoTo 0

    ' Passa il focus a un foglio sicuro prima di eliminare
    ThisWorkbook.Sheets(1).Activate

    ' Elimina fogli temporanei
    Application.DisplayAlerts = False
    wsTemp.Delete
    Application.DisplayAlerts = True

    ' Ripristina
    Call HideSheets
    Call Show_Lines
    ws.Activate
End Sub