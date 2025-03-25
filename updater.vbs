Option Explicit

Dim fso, excelFile, zipFile, estrazioneCartella, vbaProjectPath, hexContent, fileContent
Dim objShell, objFile, newZipFile

Set fso = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

' Imposta il percorso del file Excel
excelFile = "C:\Percorso\TuoFile.xlsm"
zipFile = excelFile & ".zip"
estrazioneCartella = "C:\Percorso\Estratto"

' Rinomina il file Excel in ZIP
If fso.FileExists(excelFile) Then
    fso.MoveFile excelFile, zipFile
Else
    WScript.Echo "Errore: File Excel non trovato!"
    WScript.Quit
End If

' Crea la cartella di estrazione
If Not fso.FolderExists(estrazioneCartella) Then
    fso.CreateFolder estrazioneCartella
End If

' Estrai il contenuto del file ZIP nella cartella
Set objFile = objShell.NameSpace(zipFile)
If Not objFile Is Nothing Then
    objFile.CopyHere estrazioneCartella
Else
    WScript.Echo "Errore: Impossibile estrarre il file ZIP."
    WScript.Quit
End If

' Percorso del file vbaProject.bin
vbaProjectPath = estrazioneCartella & "\xl\vbaProject.bin"

If fso.FileExists(vbaProjectPath) Then
    ' Leggi il contenuto binario e convertilo in una stringa esadecimale
    fileContent = ReadBinaryFile(vbaProjectPath)
    hexContent = ToHex(fileContent)
    
    ' Cerca e sostituisci la stringa esadecimale 'DPB='
    hexContent = ReplaceHexPassword(hexContent)
    
    ' Scrivi il nuovo contenuto nel file vbaProject.bin
    WriteBinaryFile vbaProjectPath, FromHex(hexContent)
    
    ' Ricomprimi la cartella in un nuovo ZIP e rinominalo come .xlsm
    newZipFile = zipFile
    objShell.NameSpace(newZipFile).CopyHere estrazioneCartella
    fso.MoveFile newZipFile, excelFile
    
    ' Elimina i file temporanei
    fso.DeleteFolder estrazioneCartella, True

    WScript.Echo "La password VBA è stata rimossa con successo!"
Else
    WScript.Echo "Errore: Il file vbaProject.bin non è stato trovato!"
End If

' Funzione per leggere il contenuto binario di un file
Function ReadBinaryFile(filePath)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.LoadFromFile filePath
    ReadBinaryFile = stream.Read
    stream.Close
End Function

' Funzione per scrivere il contenuto binario in un file
Sub WriteBinaryFile(filePath, data)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.Write data
    stream.SaveToFile filePath, 2 ' Overwrite
    stream.Close
End Sub

' Funzione per convertire un array binario in una stringa esadecimale
Function ToHex(bytes)
    Dim hex, i
    For i = 0 To UBound(bytes)
        hex = hex & Right("0" & Hex(bytes(i)), 2)
    Next
    ToHex = hex
End Function

' Funzione per convertire una stringa esadecimale in un array binario
Function FromHex(hex)
    Dim bytes, i
    ReDim bytes((Len(hex) \ 2) - 1)
    For i = 0 To UBound(bytes)
        bytes(i) = CByte("&H" & Mid(hex, (i * 2) + 1, 2))
    Next
    FromHex = bytes
End Function

' Funzione per rimuovere la protezione VBA (modifica esadecimale)
Function ReplaceHexPassword(hexContent)
    ReplaceHexPassword = Replace(hexContent, "4450423D", "4450423D00") ' Rimuove DPB= password
End Function