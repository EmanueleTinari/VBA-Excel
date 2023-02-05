Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                              +
'+ Funzione per valutare se un dato riferimento esiste nel progetto di Excel.   +
'+ Restituisce Vero o Falso.                                                    +
'+                                                                              +
'+ Il valore di Default della Funzione EsisteRiferimento è Falso.               +
'+                                                                              +
'+ Argomenti della Funzione:                                                    +
'+                                                                              +
'+ wbk                  -   Questo file o un altro di cui si vuole controllare. +
'+ strGUID              -   La GUID del riferimento da testare.                 +
'+                                                                              +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function EsisteRiferimento(ByVal wbk As Workbook, ByVal strGUID As String) As Boolean

' Gestione errore.
On Error GoTo GesErr

' Variabile per il ciclo tra tutti i riferimenti del file.
Dim varRiferimento As Variant

    ' Inizialmente viene impostato su False il risultato della Funzione.
    EsisteRiferimento = False
    ' Cicla tutti i riferimenti presenti nel file.
    For Each varRiferimento In wbk.VBProject.References
        ' Se il GUID del riferimento in esame è uguale a quello passato alla Funzione, allora.
        If varRiferimento.GUID = strGUID Then
            ' Imposta su True il risultato della Funzione.
            EsisteRiferimento = True
            ' Svuota le variabili.
            Set varRiferimento = Nothing
            strGUID = Empty
            ' Esce dalla Funzione.
            Exit Function
        End If
    ' Prossimo riferimento in esame.
    Next varRiferimento

' Esce dalla Funzione.
Uscita: Exit Function
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Function" & vbCrLf & "'EsisteRiferimento'" & vbCrLf & vbCrLf & Err.Description
        ' Svuota le variabili.
        Set varRiferimento = Nothing
        strGUID = Empty
        Resume Uscita
' Fine della Funzione.
End Function
