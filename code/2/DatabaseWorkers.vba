Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rngMalattieCorsi As Range
    Dim LastRow As Integer
    Dim Colum As Integer
    Dim r As Long
    
    If Application.Ready = False Then Exit Sub
    If Application.CommandBars("Cell").Enabled = False Then Exit Sub
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    LastRow = 3
    While Cells(LastRow, 1).Value <> ""
        LastRow = LastRow + 1
    Wend
    LastRow = LastRow - 1
    
    Set rngMalattieCorsi = Range(Cells(3, 6), Cells(LastRow, 8))
    
    If Not Intersect(Target, rngMalattieCorsi) Is Nothing Then
        If Target.Column = 6 Then
            If Target.Value = "No" Then
                Colum = 10
                While Cells(Target.Row, Columns).Value <> ""
                    Cells(Target.Row, Columns).Value = "No"
                    Colum = Colum + 1
                Wend
            End If
            If Target.Value = "Si" Then
                Colum = 10
                While Cells(Target.Row, Columns).Value <> ""
                    Cells(Target.Row, Columns).Value = "Si"
                    Colum = Colum + 1
                Wend
            End If
        ElseIf Target.Column = 7 Then
            If Target.Value = "No" Then
                Colum = 18
                While Cells(Target.Row, Columns).Value <> ""
                    Cells(Target.Row, Columns).Value = "No"
                    Colum = Colum + 1
                Wend
            End If
            If Target.Value = "Si" Then
                Colum = 18
                While Cells(Target.Row, Columns).Value <> ""
                    Cells(Target.Row, Columns).Value = "Si"
                    Colum = Colum + 1
                Wend
            End If
        ElseIf Target.Column = 8 Then
            If Target.Value = "No" Then
                Colum = 26
                While Cells(Target.Row, Columns).Value <> ""
                    Cells(Target.Row, Columns).Value = "No"
                    Colum = Colum + 1
                Wend
            End If
            If Target.Value = "Si" Then
                Colum = 26
                While Cells(Target.Row, Columns).Value <> ""
                    Cells(Target.Row, Columns).Value = "Si"
                    Colum = Colum + 1
                Wend
            End If
        End If
    End If

    ' Colonne da controllare
    Dim colCheckA As String: colCheckA = "C" ' Prima condizione
    Dim colCheckB As String: colCheckB = "B" ' Seconda condizione

    ' Colonne helper
    Dim helperCol1 As String: helperCol1 = "AH" ' per colonna A: 0=valore, 1=vuoto
    Dim helperCol2 As String: helperCol2 = "AI" ' per colonna B: priorit  custom

    ' Compila helper Y e Z
    For r = 3 To LastRow
        ' Y: 0 se colonna A ha valore, 1 se vuota
        If Trim(Cells(r, colCheckA).Value) <> "" Then
            Cells(r, helperCol1).Value = 0
        Else
            Cells(r, helperCol1).Value = 1
        End If

        ' Z: priorit  personalizzata su colonna B
        Select Case UCase(Trim(Cells(r, colCheckB).Value))
            Case "GEL":    Cells(r, helperCol2).Value = 1
            Case "FRONT":  Cells(r, helperCol2).Value = 2
            Case "TUTTO":  Cells(r, helperCol2).Value = 3
            Case "CUCINA": Cells(r, helperCol2).Value = 4
            Case "N/A":    Cells(r, helperCol2).Value = 5
            Case "":       Cells(r, helperCol2).Value = 999
            Case Else:     Cells(r, helperCol2).Value = 500
        End Select
    Next r

    ' Ordina prima per Y (celle piene prima), poi per Z (priorit  logica)
    With Me.Sort
        .SortFields.clear
        .SortFields.Add Key:=Range(helperCol1 & "3:" & helperCol1 & LastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=Range(helperCol2 & "3:" & helperCol2 & LastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending

        .SetRange Range("A2:AI" & LastRow) ' adatta H se la tua tabella   pi  larga
        .Header = xlYes
        .Apply
    End With

    ' Nasconde le colonne helper
    Columns(helperCol1 & ":" & helperCol2).EntireColumn.Hidden = True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim Password
    Dim rngDipendenti As Range
    Dim rngMalattie As Range
    Dim rngCorsi As Range
    Dim rngFerie As Range
    Dim LastRow As Integer
    Password = "Ej20082018*Excel"
    
    LastRow = 3
    While Cells(LastRow, 1).Value <> ""
        LastRow = LastRow + 1
    Wend
    LastRow = LastRow - 1
    
    Set rngDipendenti = Range(Cells(3, 1), Cells(LastRow, 8))
    Set rngFerie = Range(Cells(3, 10), Cells(LastRow, 16))
    Set rngMalattie = Range(Cells(3, 18), Cells(LastRow, 24))
    Set rngCorsi = Range(Cells(3, 26), Cells(LastRow, 32))
    
    If Not Intersect(Target, rngDipendenti) Is Nothing Or Not Intersect(Target, rngFerie) Is Nothing _
    Or Not Intersect(Target, rngMalattie) Is Nothing Or Not Intersect(Target, rngCorsi) Is Nothing Then
        ActiveSheet.Unprotect Password:=Password
    Else
        ActiveSheet.Protect Password:=Password
    End If
End Sub

