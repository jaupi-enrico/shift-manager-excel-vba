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

Function Check_Days(ByVal Worker As Integer, ByVal WorkerRole As Integer)
    Dim wsDip As Worksheet, wsTOT As Worksheet
    Dim LastRow As Integer, i As Integer
    Set wsDip = Worksheets("Dipendenti")
    Set wsTOT = Worksheets("TOT")

    For i = 10 To 16
        If wsDip.Cells(Worker, i).Value = "Si" And wsTOT.Cells(Worker + 1, 4 + (i - 10) * 2).Value <> "FERIE" Then
            wsTOT.Range(wsTOT.Cells(Worker + 1, 4 + (i - 10) * 2), wsTOT.Cells(Worker + 1, 5 + (i - 10) * 2)).Merge
            wsTOT.Cells(Worker + 1, 4 + (i - 10) * 2).Value = "FERIE"
        ElseIf wsDip.Cells(Worker, i).Value = "No" And wsTOT.Cells(Worker + 1, 4 + (i - 10) * 2).Value = "FERIE" Then
            wsTOT.Range(wsTOT.Cells(Worker + 1, 4 + (i - 10) * 2), wsTOT.Cells(Worker + 1, 5 + (i - 10) * 2)).UnMerge
            wsTOT.Range(wsTOT.Cells(Worker + 2, 4 + (i - 10) * 2), wsTOT.Cells(Worker + 2, 5 + (i - 10) * 2)).Copy
            wsTOT.Range(wsTOT.Cells(Worker + 1, 4 + (i - 10) * 2), wsTOT.Cells(Worker + 1, 5 + (i - 10) * 2)).PasteSpecial Paste:=xlPasteAll
        End If
    Next i

    For i = 18 To 24
        If wsDip.Cells(Worker, i).Value = "Si" And wsTOT.Cells(Worker + 1, 4 + (i - 18) * 2).Value <> "MALATTIA" Then
            wsTOT.Range(wsTOT.Cells(Worker + 1, 4 + (i - 18) * 2), wsTOT.Cells(Worker + 1, 5 + (i - 18) * 2)).Merge
            wsTOT.Cells(Worker + 1, 4 + (i - 18) * 2).Value = "MALATTIA"
        ElseIf wsDip.Cells(Worker, i).Value = "No" And wsTOT.Cells(Worker + 1, 4 + (i - 18) * 2).Value = "MALATTIA" Then
            wsTOT.Range(wsTOT.Cells(Worker + 1, 4 + (i - 18) * 2), wsTOT.Cells(Worker + 1, 5 + (i - 18) * 2)).UnMerge
            wsTOT.Range(wsTOT.Cells(Worker + 2, 4 + (i - 18) * 2), wsTOT.Cells(Worker + 2, 5 + (i - 18) * 2)).Copy
            wsTOT.Range(wsTOT.Cells(Worker + 1, 4 + (i - 18) * 2), wsTOT.Cells(Worker + 1, 5 + (i - 18) * 2)).PasteSpecial Paste:=xlPasteAll
        End If
    Next i

    For i = 26 To 32
        If wsDip.Cells(Worker, i).Value = "Si" And wsTOT.Cells(Worker + 1, 4 + (i - 26) * 2).Value <> "CORSO" Then
            wsTOT.Range(wsTOT.Cells(Worker + 1, 4 + (i - 26) * 2), wsTOT.Cells(Worker + 1, 5 + (i - 26) * 2)).Merge
            wsTOT.Cells(Worker + 1, 4 + (i - 26) * 2).Value = "CORSO"
        ElseIf wsDip.Cells(Worker, i).Value = "No" And wsTOT.Cells(Worker + 1, 4 + (i - 26) * 2).Value = "CORSO" Then
            wsTOT.Range(wsTOT.Cells(Worker + 1, 4 + (i - 26) * 2), wsTOT.Cells(Worker + 1, 5 + (i - 26) * 2)).UnMerge
            wsTOT.Range(wsTOT.Cells(Worker + 2, 4 + (i - 26) * 2), wsTOT.Cells(Worker + 2, 5 + (i - 26) * 2)).Copy
            wsTOT.Range(wsTOT.Cells(Worker + 1, 4 + (i - 26) * 2), wsTOT.Cells(Worker + 1, 5 + (i - 26) * 2)).PasteSpecial Paste:=xlPasteAll
        End If
    Next i

    If wsDip.Cells(Worker, 1).Value = "Si" Then
        wsTOT.Cells(Worker + 1, 1).Interior.color = RGB(241, 170, 131)
    ElseIf wsDip.Cells(Worker, 1).Value = "No" And wsTOT.Cells(Worker + 1, 1).Interior.color = RGB(241, 170, 131) Then
        wsTOT.Range(wsTOT.Cells(Worker + 2, 4), wsTOT.Cells(Worker + 2, 17)).Copy
        wsTOT.Range(wsTOT.Cells(Worker + 1, 4), wsTOT.Cells(Worker + 1, 17)).PasteSpecial xlPasteAll
        Call Paint_Worker(Worker + 1, WorkerRole)
        wsTOT.Cells(Worker + 1, 1).Interior.color = RGB(255, 255, 255)
    End If
End Function

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