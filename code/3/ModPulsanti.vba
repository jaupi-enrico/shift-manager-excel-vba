Option Explicit
Dim sheet As Worksheet ' Dichiarazione variabile sheet

Sub ShowSheets()
    Dim Password As String
    Dim sheet As Worksheet
    Dim currentSheet As Worksheet

    Application.ScreenUpdating = False
    Password = "Ej20082018*Excel"
    Set currentSheet = ActiveSheet

    For Each sheet In Worksheets
        If sheet.name <> "DASHBOARD" Then
            sheet.Unprotect Password:=Password
        End If
    Next sheet

    currentSheet.Activate
    Application.ScreenUpdating = True
End Sub


Sub HideSheets()
    Dim Password As String
    Dim sheet As Worksheet
    Dim currentSheet As Worksheet

    Application.ScreenUpdating = False
    Password = "Ej20082018*Excel"
    Set currentSheet = ActiveSheet

    For Each sheet In Worksheets
        If sheet.name <> "DASHBOARD" Then
            sheet.Protect Password:=Password
        End If
    Next sheet

    currentSheet.Activate
    Application.ScreenUpdating = True
End Sub

Sub Delete_images()
    Dim sheet As Worksheet
    Dim shp As Shape

    For Each sheet In Worksheets
        ' Escludi i fogli "DASHBOARD" e "TOT"
        If sheet.name <> "DASHBOARD" And sheet.name <> "Dipendenti" And _
           sheet.name <> "FORMAZIONE" Then
            ' Cicla attraverso ogni forma presente nel foglio
            For Each shp In sheet.Shapes
                ' Elimina l'oggetto
                shp.Delete
            Next shp
        End If
    Next sheet
    
    MsgBox "Ottimizzazione finita"
End Sub

Sub Add_Manager()
    Call ShowSheets
    Dim manager_row As Integer

    manager_row = 4

    While Cells(manager_row, 1) <> ""
        manager_row = manager_row + 1
    Wend
    manager_row = manager_row - 1


    ActiveSheet.Rows(manager_row + 1).Insert Shift:=xlDown
    DoEvents
    ActiveSheet.Rows(manager_row).Copy
    ActiveSheet.Rows(manager_row + 1).PasteSpecial xlPasteAll

    Range(Cells(manager_row + 1, 1), Cells(manager_row + 1, 15)).Value = ""
    Call HideSheets
End Sub

