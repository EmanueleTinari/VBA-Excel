
Option Explicit

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                               +
'+ All'apertura del file di Excel, esegue la Sub che controlla   +
'+ che tutti i riferimenti necessari al progetto siano presenti. +
'+                                                               +
'+ È necessaria la Subroutine ControlloRiferimenti.              +
'+                                                               +
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub Workbook_Open()

    '+++++++++++++++++++++++++
    '+ Esegue la Subroutine. +
    '+++++++++++++++++++++++++
    Call ControlloRiferimenti
    
End Sub
