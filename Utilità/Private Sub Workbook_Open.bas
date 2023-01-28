
Option Explicit

'+++++++++++++++++++++++++++++++++++++++++
'+ All'apertura del file, esegue la      +
'+ Sub che controlla che tutti i         +
'+ riferimenti necessari siano presenti. +
'+++++++++++++++++++++++++++++++++++++++++
Private Sub Workbook_Open()

    '++++++++++++++++++++++++++
    '+ Esegue la Sub routine. +
    '++++++++++++++++++++++++++
    Call ControlloRiferimenti
    
End Sub
