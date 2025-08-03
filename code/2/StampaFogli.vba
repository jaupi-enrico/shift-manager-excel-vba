Sub StampaFogliSovrapposti()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim wsTemp As Worksheet, ws2copy As Worksheet
    Dim shp1 As Shape, shp2 As Shape
    Dim spazioY As Double
    Dim LastRowMax As Long

    ' Mostra fogli e nasconde linee
    Call ShowSheets
    Call Hide_Lines

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("___TEMP_STAMPA___").Delete
    ThisWorkbook.Sheets("___TMP2___").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Imposta i fogli da stampare
    Set ws1 = ThisWorkbook.Sheets("MANAGER")
    Set ws2 = ThisWorkbook.Sheets("TOT-M")

    ' Crea foglio temporaneo per stampa
    Set wsTemp = ThisWorkbook.Sheets.Add(After:=ws2)
    wsTemp.name = "___TEMP_STAMPA___"

    ' === MANAGER ===
    ws1.Range("A1:BT152").CopyPicture Appearance:=xlScreen, Format:=xlPicture
    wsTemp.Range("A1").Select
    wsTemp.Paste
    DoEvents ' Previene crash
    Set shp1 = wsTemp.Shapes(wsTemp.Shapes.Count)

    ' === TOT-M ===
    ' Trova ultima riga usata nella colonna A
    LastRowMax = 4
    While ws2.Cells(LastRowMax, 1).Value <> "" And LastRowMax < Rows.Count
        LastRowMax = LastRowMax + 1
    Wend
    LastRowMax = LastRowMax - 1

    ' Crea copia semplificata di ws2 senza formattazioni condizionali
    Set ws2copy = ThisWorkbook.Sheets.Add(After:=wsTemp)
    ws2copy.name = "___TMP2___"

    Dim rngOrigine As Range, rngDest As Range
    Set rngOrigine = ws2.Range("C1:Q" & LastRowMax + 3)
    Set rngDest = ws2copy.Range("C1")

    ' Copia valori
    rngDest.Resize(rngOrigine.Rows.Count, rngOrigine.Columns.Count).Value = rngOrigine.Value

    ' Copia formato (colori, bordi, font, ecc.)
    rngOrigine.Copy
    rngDest.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Copia larghezze colonne
    Dim col As Integer
    For col = 1 To rngOrigine.Columns.Count
        ws2copy.Columns(col + 2).ColumnWidth = ws2.Columns(col + 2).ColumnWidth
    Next col

    ' Copia altezze righe
    Dim r As Long
    For r = 1 To rngOrigine.Rows.Count
        ws2copy.Rows(r).RowHeight = ws2.Rows(r).RowHeight
    Next r

    ' Rimuovi eventuali formattazioni condizionali ereditate
    ws2copy.Cells.FormatConditions.Delete
    DoEvents

    ' Copia immagine da ws2copy
    ws2copy.Range("C1:Q" & LastRowMax + 3).CopyPicture Appearance:=xlScreen, Format:=xlPicture
    wsTemp.Activate
    wsTemp.Range("A1").Select
    wsTemp.Paste
    DoEvents
    Set shp2 = wsTemp.Shapes(wsTemp.Shapes.Count)

    ' Uniforma dimensioni
    With shp1
        .LockAspectRatio = msoFalse
        .Height = .Height * 1.4
    End With
    With shp2
        .LockAspectRatio = msoTrue
        .Width = shp1.Width * 0.65
        .Top = shp1.Top + shp1.Height + 20
    End With

    ' Imposta stampa su una sola pagina
    With wsTemp.PageSetup
        .Orientation = xlLandscape
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
    Dim wsCurr As Worksheet
    Set wsCurr = ActiveSheet

    Application.ScreenUpdating = False
    wsTemp.Activate
    
    ' Attiva il foglio da stampare
    wsTemp.Select

    Application.EnableEvents = False ' disattiva macro evento
    wsTemp.PrintOut Copies:=1
    Application.EnableEvents = True  ' riattiva macro evento
    wsCurr.Activate
    Application.ScreenUpdating = True

    ' Passa il focus a un foglio sicuro prima di eliminare
    ThisWorkbook.Sheets(1).Activate

    ' Elimina fogli temporanei
    Application.DisplayAlerts = False
    wsTemp.Delete
    ws2copy.Delete
    Application.DisplayAlerts = True

    ' Ripristina
    Call HideSheets
    Call Show_Lines
    ws2.Activate
End Sub
