Option Explicit

Sub AggiornaCodiceVBA()
    Update.Show vbModeless
    DoEvents ' Permette il disegno della form prima di proseguire
    
    On Error GoTo ErrorHandler

    Dim Http As Object
    Dim vbProj As Object
    Dim vbComp As Object
    Dim ModuloNome As String
    Dim URL As String
    Dim NuovoCodice As String
    Dim Updated As Boolean
    
    Updated = False
    
    ' Imposta il riferimento al progetto VBA
    Set vbProj = ThisWorkbook.VBProject

    ' Itera su tutti i moduli VBA
    For Each vbComp In vbProj.VBComponents
        ' Nome del modulo corrente
        ModuloNome = vbComp.name
        
        ' URL del file su GitHub (ora .vba)
        URL = "https://raw.githubusercontent.com/jaupi-enrico/shift-manager-excel-vba/main/code/" & ModuloNome & ".vba"
        
        ' Scarica il file
        Set Http = CreateObject("MSXML2.XMLHTTP")
        Http.Open "GET", URL, False
        Http.Send

        If Http.Status = 200 Then
            Updated = True
            
            ' Ottieni il codice dal file remoto
            NuovoCodice = Http.responseText
            
            ' Rimuove il BOM UTF-8 se presente
            If Left(NuovoCodice, 3) = ChrW(&HFEFF) Then
                NuovoCodice = Mid(NuovoCodice, 4)
            End If

            ' Sostituisci il codice del modulo
            With vbComp.CodeModule
                .DeleteLines 1, .CountOfLines
                .AddFromString NuovoCodice
            End With
        ElseIf Http.Status = 404 Then
            Debug.Print "Pagina mancante: " & ModuloNome
        Else
            MsgBox "Errore nel download del codice! Stato: " & Http.Status & " " & ModuloNome, vbCritical
        End If
    Next vbComp
    
    Unload Update
    
    If Updated Then
        MsgBox "Codice aggiornato con successo!", vbInformation
    End If
    
    GoTo Continue
    
ErrorHandler:
    ' Gestione degli errori
    MsgBox "Errore: " & Err.Description, vbCritical + vbOKOnly + vbDefaultButton1
    
Continue:

End Sub