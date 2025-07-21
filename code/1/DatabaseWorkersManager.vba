Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rngMalattieCorsi As Range
    Dim LastRow As Integer
    Dim Column As Integer
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
    
    Set rngMalattieCorsi = Range(Cells(3, 5), Cells(LastRow, 7))
    
    If Not Intersect(Target, rngMalattieCorsi) Is Nothing Then
        If Target.Column = 5 Then
            If Target.Value = "No" Then
                Column = 9
                While Cells(Target.Row, Column).Value <> ""
                    Cells(Target.Row, Column).Value = "No"
                    Column = Column + 1
                Wend
            End If
            If Target.Value = "Si" Then
                Column = 9
                While Cells(Target.Row, Column).Value <> ""
                    Cells(Target.Row, Column).Value = "Si"
                    Column = Column + 1
                Wend
            End If
        ElseIf Target.Column = 6 Then
            If Target.Value = "No" Then
                Column = 17
                While Cells(Target.Row, Column).Value <> ""
                    Cells(Target.Row, Column).Value = "No"
                    Column = Column + 1
                Wend
            End If
            If Target.Value = "Si" Then
                Column = 17
                While Cells(Target.Row, Column).Value <> ""
                    Cells(Target.Row, Column).Value = "Si"
                    Column = Column + 1
                Wend
            End If
        ElseIf Target.Column = 7 Then
            If Target.Value = "No" Then
                Column = 25
                While Cells(Target.Row, Column).Value <> ""
                    Cells(Target.Row, Column).Value = "No"
                    Column = Column + 1
                Wend
            End If
            If Target.Value = "Si" Then
                Column = 25
                While Cells(Target.Row, Column).Value <> ""
                    Cells(Target.Row, Column).Value = "Si"
                    Column = Column + 1
                Wend
            End If
        End If
    End If

    ' Colonne da controllare
    Dim colCheckA As String: colCheckA = "B" ' Prima condizione
    Dim colCheckB As String: colCheckB = "C" ' Seconda condizione

    ' Colonne helper
    Dim helperCol1 As String: helperCol1 = "AG" ' per colonna A: 0=valore, 1=vuoto
    Dim helperCol2 As String: helperCol2 = "AH" ' per colonna B: priorit  custom

    ' Compila helper Y e Z
    For r = 3 To LastRow
        ' Y: 0 se colonna A ha valore, 1 se vuota
        If Trim(Cells(r, colCheckA).Value) <> "" Then
            Cells(r, helperCol1).Value = 0
        Else
            Cells(r, helperCol1).Value = 1
        End If

        ' Y: 0 se colonna A ha valore, 1 se vuota
        If Trim(Cells(r, colCheckB).Value) <> "" Then
            Cells(r, helperCol2).Value = 0
        Else
            Cells(r, helperCol2).Value = 1
        End If
    Next r

    ' Ordina prima per Y (celle piene prima), poi per Z (priorit  logica)
    With Me.Sort
        .SortFields.clear
        .SortFields.Add Key:=Range(helperCol1 & "3:" & helperCol1 & LastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=Range(helperCol2 & "3:" & helperCol2 & LastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending

        .SetRange Range("A2:AH" & LastRow) ' adatta H se la tua tabella   pi  larga
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
    
    Set rngDipendenti = Range(Cells(3, 1), Cells(LastRow, 7))
    Set rngFerie = Range(Cells(3, 9), Cells(LastRow, 15))
    Set rngMalattie = Range(Cells(3, 17), Cells(LastRow, 23))
    Set rngCorsi = Range(Cells(3, 25), Cells(LastRow, 31))
    
    If Not Intersect(Target, rngDipendenti) Is Nothing Or Not Intersect(Target, rngFerie) Is Nothing _
    Or Not Intersect(Target, rngMalattie) Is Nothing Or Not Intersect(Target, rngCorsi) Is Nothing Then
        ActiveSheet.Unprotect Password:=Password
    Else
        ActiveSheet.Protect Password:=Password
    End If
End Sub