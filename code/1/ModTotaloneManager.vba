Option Explicit
' Aggiunge un lavoratore alle tabelle TOT e FORMAZIONE
Sub Add_Worker(ByVal WorkerName As String, ByVal WorkerName_Surname As String, ByVal Contract As Integer, ByVal WorkerPos As Integer)
    Dim wsTOT As Worksheet, wsFORM As Worksheet
    Dim LastRow As Integer, i As Integer

    Set wsTOT = Worksheets("TOT-M")
    Set wsFORM = Worksheets("FORMAZIONE-M")

    LastRow = 4
    While wsTOT.Cells(LastRow, 2).Value <> ""
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
    Set wsTOT = Worksheets("TOT-M")
    Set wsFORM = Worksheets("FORMAZIONE-M")

    wsTOT.Rows(WorkerPos + 3).Delete
    wsFORM.Rows(WorkerPos + 3).Delete
    DoEvents

    LastRow = 4
    While wsTOT.Cells(LastRow, 2).Value <> "" And LastRow < Rows.Count
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
    Dim wsTOT As Worksheet, workerRow As Integer
    Set wsTOT = Worksheets("TOT")
    workerRow = Start + 1
    While wsTOT.Cells(workerRow, 2).Value <> ""
        If wsTOT.Cells(workerRow, 2).Value = WorkerName Then
            NameFound = workerRow
            Exit Function
        End If
        workerRow = workerRow + 1
    Wend
    NameFound = -1
End Function

' Trasferisce i dati di orari/formazione
Function Transfer_data(ByVal OldPos As Integer, ByVal NewPos As Integer)
    Dim wsTOT As Worksheet
    Dim wsFORM As Worksheet

    Set wsTOT = Worksheets("TOT-M")
    Set wsFORM = Worksheets("FORMAZIONE-M")

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

Sub Check_Days(ByVal Worker As Integer)
    Dim wsDip As Worksheet, wsTOT As Worksheet
    Dim i As Integer, colOffset As Integer
    Dim label As String, rigaValida As Integer
    Dim LastRowMax As Integer

    Set wsDip = Worksheets("Dipendenti-M")
    Set wsTOT = Worksheets("TOT-M")

    ' Trova l'ultima riga non vuota della colonna A nel foglio TOT
    LastRowMax = 4
    While wsTOT.Cells(LastRowMax, 1).Value <> "" And LastRowMax < Rows.Count
        LastRowMax = LastRowMax + 1
    Wend
    LastRowMax = LastRowMax - 1

    ' FERIE (colonne 10–16)
    For i = 9 To 15
        colOffset = 4 + (i - 9) * 2
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
    For i = 17 To 23
        colOffset = 4 + (i - 17) * 2
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
    For i = 25 To 31
        colOffset = 4 + (i - 25) * 2
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
        wsTOT.Cells(Worker + 1, 1).Interior.color = RGB(255, 255, 255)
    End If
End Sub

Function Update_Validation()
    Dim ws As Worksheet
    Dim wsTOT As Worksheet
    Dim LastRow As Long
    Dim addressList As String
    Dim rngValidazione As Range

    Set wsTOT = ThisWorkbook.Sheets("TOT-M")

    ' Trova l'ultima riga non vuota della colonna A nel foglio TOT
    LastRow = 4

    While wsTOT.Cells(LastRow, 2).Value <> "" And LastRow < Rows.Count
        LastRow = LastRow + 1
    Wend
    LastRow = LastRow + 3
    
    ' Costruisci l'indirizzo dell'intervallo da usare nella validazione
    addressList = "='TOT-M'!$B$4:$B$" & LastRow

    ' Applica la convalida nei fogli dei giorni
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "MANAGER" Then

            Set rngValidazione = ws.Range("A2:A148") ' Adatta se necessario
            
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


' Aggiorna tutti i lavoratori dal foglio Dipendenti
Sub Update_Workers_M()
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
    Dim WorkerContract As Integer, LastRow As Integer, WorkerPos As Integer

    Set wsDip = Worksheets("Dipendenti-M")
    Set wsTOT = Worksheets("TOT-M")
    Worker = 3

    While wsDip.Cells(Worker, 2).Value <> ""
        WorkerName = wsDip.Cells(Worker, 2).Value
        WorkerName_Surname = wsDip.Cells(Worker, 3).Value
        WorkerContract = wsDip.Cells(Worker, 4).Value

        If NameFound(WorkerName, 3) <> -1 Then
            WorkerPos = NameFound(WorkerName, 3) - 3
            If WorkerPos = Worker - 2 And _
            wsTOT.Cells(WorkerPos + 3, 3).Value = WorkerName_Surname And _
            wsTOT.Cells(WorkerPos + 3, 27).Value = WorkerContract And _
            wsTOT.Cells(WorkerPos + 3, 2).Interior.color = color Then
                Call Check_Days(Worker)
            Else
                Call Add_Worker(WorkerName, WorkerName_Surname, WorkerContract, Worker - 2)
                WorkerPos = NameFound(WorkerName, Worker + 1) - 3
                Call Transfer_data(WorkerPos + 3, Worker + 1)
                Call Delete_Worker(WorkerPos)
                Call Check_Days(Worker)
            End If
        Else
            Call Add_Worker(WorkerName, WorkerName_Surname, WorkerContract, Worker - 2)
            Call Check_Days(Worker)
        End If
        Worker = Worker + 1
    Wend

    LastRow = 4
    While wsTOT.Cells(LastRow, 2).Value <> "" And LastRow < Rows.Count
        LastRow = LastRow + 1
    Wend

    While Worker - 2 <> LastRow - 3
        Call Delete_Worker(Worker - 2)
        LastRow = LastRow - 1
    Wend

    Call Update_Validation

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

Sub PrintRange_M()
    Dim LastRow As Integer
    LastRow = 4
    While Cells(LastRow, 1) <> ""
        LastRow = LastRow + 1
    Wend
    LastRow = LastRow + 2
    ' Variabili per l'input dell'utente
    Dim scelta As Integer
    Dim rangeDaStampare As Range
    
    ' Mostra una finestra di dialogo per scegliere tra due opzioni
    scelta = MsgBox("Scegli quale intervallo stampare:" & vbCrLf & _
                    "Si: Con orari" & vbCrLf & _
                    "No: Senza orari", vbYesNoCancel + vbQuestion, "Scegli intervallo")
    
    ' Controllo la scelta dell'utente
    If scelta = vbYes Then
        Set rangeDaStampare = Range(Cells(1, 3), Cells(LastRow, 17))
        rangeDaStampare.PrintOut
    ElseIf scelta = vbNo Then
        Set rangeDaStampare = Range(Cells(1, 53), Cells(LastRow, 67))
        rangeDaStampare.PrintOut
    Else
        MsgBox "Operazione annullata."
    End If
    
End Sub