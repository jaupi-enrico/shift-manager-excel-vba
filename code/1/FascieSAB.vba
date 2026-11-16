Option Explicit

' Foglio vendite "Sabato": incollando un'intera pagina da A1 elimina la prima riga
' e toglie gli spazi iniziali da ogni dato (vedi Pulisci_Incolla_Vendite in ModPulsanti)
Private Sub Worksheet_Change(ByVal Target As Range)
    Call Pulisci_Incolla_Vendite(Me, Target)
End Sub
