Option Explicit

' Ramo GitHub da cui viene scaricato il codice.
' ATTENZIONE: "main" e' il ramo distribuito a tutti gli utenti. Un ramo di prova
' va usato solo su copie di test del file e deve tornare a "main" prima del merge.
Private Const UPDATE_BRANCH As String = "merge/no-ui-alerts"
Private Const UPDATE_BASE_URL As String = "https://raw.githubusercontent.com/jaupi-enrico/shift-manager-excel-vba/"

' Aggiorna il codice VBA dei moduli dal repository GitHub.
' - scarica TUTTI i moduli prima di toccare qualcosa: se un download fallisce non cambia niente
' - sostituisce solo i moduli il cui codice e' davvero diverso
' - prima di sostituire esporta una copia di backup nella cartella temporanea
' - il modulo Updater (in esecuzione) viene sostituito per ultimo
Sub AggiornaCodiceVBA()
    Update.Show vbModeless
    DoEvents ' Permette il disegno della form prima di proseguire

    Dim vbProj As Object
    Dim vbComp As Object
    Dim Http As Object
    Dim Nuovo As Object          ' Dictionary: nome modulo -> codice scaricato
    Dim v As Variant
    Dim URL As String
    Dim Codice As String
    Dim ModuloNome As String
    Dim tipo As Integer
    Dim BackupDir As String
    Dim Fase As Integer          ' 1 = download, 2 = backup, 3 = sostituzione, 4 = sostituzione di Updater

    On Error GoTo ErrorHandler

    Set vbProj = ThisWorkbook.VBProject
    Set Nuovo = CreateObject("Scripting.Dictionary")
    Set Http = CreateObject("MSXML2.XMLHTTP")
    tipo = Worksheets("Tabelle").Range("H9").Value

    ' --- Fase 1: scarica e verifica tutti i moduli
    Fase = 1
    For Each vbComp In vbProj.VBComponents
        ModuloNome = vbComp.name
        URL = UPDATE_BASE_URL & UPDATE_BRANCH & "/code/" & tipo & "/" & ModuloNome & ".vba"

        Http.Open "GET", URL, False
        Http.Send

        If Http.Status = 200 Then
            Codice = Http.responseText

            ' Rimuove il BOM UTF-8 se presente
            If Left(Codice, 1) = ChrW(&HFEFF) Then
                Codice = Mid(Codice, 2)
            End If

            If Not Codice_Valido(Codice) Then
                Err.Raise vbObjectError + 600, "AggiornaCodiceVBA", _
                    "Il file scaricato per il modulo " & ModuloNome & " non sembra codice VBA valido."
            End If

            ' Aggiorna solo se il codice e' cambiato
            If Codice_Diverso(vbComp.codeModule, Codice) Then
                Nuovo.Add ModuloNome, Codice
            End If
        ElseIf Http.Status = 404 Then
            Debug.Print "Pagina mancante: " & ModuloNome
        Else
            Err.Raise vbObjectError + 601, "AggiornaCodiceVBA", _
                "Errore nel download del modulo " & ModuloNome & " (stato " & Http.Status & ")."
        End If
    Next vbComp

    If Nuovo.Count = 0 Then GoTo Finito ' tutto gia' aggiornato

    ' --- Fase 2: backup dei moduli che verranno sostituiti
    Fase = 2
    BackupDir = Environ$("TEMP") & "\shift-manager-backup-" & Format(Now, "yyyymmdd-hhnnss")
    MkDir BackupDir
    For Each v In Nuovo.Keys
        Set vbComp = vbProj.VBComponents(CStr(v))
        vbComp.Export BackupDir & "\" & CStr(v) & Estensione_Modulo(vbComp.Type)
    Next v

    ' --- Fase 3: sostituzione (Updater per ultimo, perche' e' in esecuzione)
    Fase = 3
    For Each v In Nuovo.Keys
        If CStr(v) <> "Updater" Then
            Sostituisci_Codice vbProj.VBComponents(CStr(v)).codeModule, CStr(Nuovo(v))
        End If
    Next v

    Unload Update
    MsgBox "Codice aggiornato con successo! (" & Nuovo.Count & " moduli)" & vbCrLf & _
           "Copia di backup: " & BackupDir, vbInformation

    If Nuovo.Exists("Updater") Then
        ' Ultima istruzione utile: dopo questa riga il codice in esecuzione viene sostituito
        Fase = 4
        Sostituisci_Codice vbProj.VBComponents("Updater").codeModule, CStr(Nuovo("Updater"))
    End If
    Exit Sub

Finito:
    Unload Update
    Exit Sub

ErrorHandler:
    ' Gestione degli errori
    Dim ErrNum As Long, ErrDesc As String
    ErrNum = Err.Number
    ErrDesc = Err.Description
    On Error Resume Next
    Unload Update
    If ErrNum = 1004 Then
        MsgBox "Abilita i permessi sugli oggetti come spiegato nella dashboard!"
    ElseIf ErrNum = -2146697211 Then
        MsgBox "Errore di connessione a internet"
    ElseIf Fase = 4 Then
        MsgBox "Tutti i moduli sono stati aggiornati ma non e' stato possibile aggiornare il modulo Updater: " & ErrNum & vbCrLf & ErrDesc & vbCrLf & _
               "Copia di backup: " & BackupDir, vbExclamation
    ElseIf Fase = 3 Then
        MsgBox "Aggiornamento interrotto durante la sostituzione dei moduli: " & ErrNum & vbCrLf & ErrDesc & vbCrLf & _
               "Alcuni moduli potrebbero essere gia' aggiornati. Copia di backup: " & BackupDir, vbCritical
    ElseIf Fase = 2 Then
        MsgBox "Impossibile creare la copia di backup, nessun modulo e' stato modificato: " & ErrNum & vbCrLf & ErrDesc, vbCritical
    ElseIf Fase = 1 Then
        MsgBox "Aggiornamento annullato, nessun modulo e' stato modificato: " & ErrNum & vbCrLf & ErrDesc, vbCritical
    Else
        MsgBox "Errore imprevisto: " & ErrNum & vbCrLf & ErrDesc, vbCritical
    End If
End Sub

' Controlla che il testo scaricato assomigli a codice VBA (non una pagina HTML di errore)
Private Function Codice_Valido(ByVal Codice As String) As Boolean
    Codice = Normalizza_Codice(Codice)
    If Len(Codice) < 10 Or Len(Codice) > 1000000 Then Exit Function
    If Left(Codice, 1) = "<" Then Exit Function
    Codice_Valido = (InStr(Codice, "Sub ") > 0 Or InStr(Codice, "Function ") > 0 Or InStr(Codice, "Option Explicit") > 0)
End Function

' True se il codice del modulo e' diverso da quello scaricato
Private Function Codice_Diverso(ByVal cm As Object, ByVal Nuovo As String) As Boolean
    Dim Attuale As String
    If cm.CountOfLines > 0 Then Attuale = cm.Lines(1, cm.CountOfLines)
    Codice_Diverso = (StrComp(Normalizza_Codice(Attuale), Normalizza_Codice(Nuovo), vbBinaryCompare) <> 0)
End Function

' Uniforma fine riga e toglie spazi/righe vuote iniziali e finali per il confronto
Private Function Normalizza_Codice(ByVal t As String) As String
    t = Replace(t, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    Do While Len(t) > 0
        If InStr(" " & vbLf & vbTab, Left(t, 1)) > 0 Then
            t = Mid(t, 2)
        Else
            Exit Do
        End If
    Loop
    Do While Len(t) > 0
        If InStr(" " & vbLf & vbTab, Right(t, 1)) > 0 Then
            t = Left(t, Len(t) - 1)
        Else
            Exit Do
        End If
    Loop
    Normalizza_Codice = t
End Function

Private Sub Sostituisci_Codice(ByVal cm As Object, ByVal Codice As String)
    With cm
        .DeleteLines 1, .CountOfLines
        .AddFromString Codice
    End With
End Sub

' Estensione del file di esportazione in base al tipo di componente
Private Function Estensione_Modulo(ByVal tipoComp As Long) As String
    Select Case tipoComp
        Case 1: Estensione_Modulo = ".bas"   ' modulo standard
        Case 3: Estensione_Modulo = ".frm"   ' UserForm
        Case Else: Estensione_Modulo = ".cls" ' classi, fogli, ThisWorkbook
    End Select
End Function
