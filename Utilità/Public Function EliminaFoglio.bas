
Option Explicit

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                               +
'+ Funzione per eliminare un foglio nel file aperto di Excel.                    +
'+ Elimina il foglio il cui nome è passato dalla variabile stringa strNomeFoglio +
'+ e restituisce Vero o Falso.                                                   +
'+                                                                               +
'+ Il valore di Default della Funzione EliminaFoglio è Falso.                    +
'+                                                                               +
'+ Argomenti della Funzione:                                                     +
'+                                                                               +
'+ strNomeFoglio   -   Il nome del foglio da eliminare.                          +
'+                                                                               +
'+ È necessaria la Funzione EsisteFoglio(strNomeFoglio) per testare se           +
'+ un foglio con un dato nome esiste, per evitare errori in caso non esistesse.  +
'+                                                                               +
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function EliminaFoglio(ByVal strNomeFoglio As String) As Boolean

' Gestione errore.
On Error GoTo GesErr

    ' Inizialmente viene impostato su False il risultato della Funzione.
    EliminaFoglio = False
    ' Se il risultato della Funzione EsisteFoglio è True, ovvero se esiste un foglio col nome passato dalla stringa strNomeFoglio, allora.
    If EsisteFoglio(strNomeFoglio) = True Then
        ' Disattivo gli avvisi di Excel.
        Application.DisplayAlerts = False
        ' Viene eliminato il foglio col nome passato dalla stringa strNomeFoglio.
        Application.Worksheets(strNomeFoglio).Delete
        ' Imposta su True il risultato della Funzione.
        EliminaFoglio = True
        ' Riattivo gli avvisi di Excel.
        Application.DisplayAlerts = True
        ' Svuota la variabile.
        strNomeFoglio = Empty
        ' Esce dalla Funzione.
        Exit Function
    End If

' Esce dalla Funzione.
Uscita: Exit Function
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Function" & vbCrLf & "'EliminaFoglio'" & vbCrLf & vbCrLf & Err.Description
        ' Svuota la variabile.
        strNomeFoglio = Empty
        Resume Uscita
' Fine della Funzione.
End Function
