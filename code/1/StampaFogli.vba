Sub StampaFogliSovrapposti()
    Call ShowSheets
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim wsTemp As Worksheet
    Dim shp1 As Shape, shp2 As Shape
    Dim spazioY As Double

    ' Imposta i fogli da stampare
    Set ws1 = ThisWorkbook.Sheets("MANAGER")
    Set ws2 = ThisWorkbook.Sheets("TOT-M")

    ' Crea foglio temporaneo
    Set wsTemp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsTemp.name = "___TEMP_STAMPA___"

    ' === MANAGER ===
    ws1.Activate
    ws1.Range("A1:BT150").CopyPicture Appearance:=xlScreen, Format:=xlPicture
    wsTemp.Activate
    wsTemp.Range("A1").Select
    wsTemp.Paste
    Set shp1 = wsTemp.Shapes(wsTemp.Shapes.Count)
    
    ' Calcola posizione per seconda immagine
    spazioY = shp1.Top + shp1.Height + 20
    
    ' === TOT-M ===
    ' Trova l'ultima riga non vuota della colonna A nel foglio MANAGER
    LastRowMax = 4
    While ws2.Cells(LastRowMax, 1).Value <> "" And LastRowMax < Rows.Count
        LastRowMax = LastRowMax + 1
    Wend
    LastRowMax = LastRowMax - 1
    ws2.Activate
    ws2.Range("C1:Q" & LastRowMax + 3).CopyPicture Appearance:=xlScreen, Format:=xlPicture
    wsTemp.Activate
    wsTemp.Cells(1, 1).Select
    wsTemp.Paste
    Set shp2 = wsTemp.Shapes(wsTemp.Shapes.Count)
    
    ' Uniforma altezza di shp2 a shp1
    With shp1
        .LockAspectRatio = msoFalse
        .Height = shp1.Height * 1.2
    End With
    With shp2
        .LockAspectRatio = msoTrue
        .Width = shp1.Width
    End With
    
    ' Posiziona shp2 sotto shp1 (con 20 punti di margine)
    spazioY = shp1.Top + shp1.Height + 20
    shp2.Top = spazioY


    ' Imposta stampa su una sola pagina
    With wsTemp.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
    End With

    ' Stampa
    wsTemp.PrintOut

    ' Elimina foglio temporaneo
    Application.DisplayAlerts = False
    wsTemp.Delete
    Application.DisplayAlerts = True

    Call HideSheets
    ws2.Activate
End Sub