
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                          +
'+ Funzione per valutare se un dato foglio esiste nel file aperto di Excel. +
'+ Restituisce Vero o Falso.                                                +
'+                                                                          +
'+ Uso: Prima di tentare di accedere, creare o cancellare un foglio nel     +
'+ file di Excel, si può far eseguire un controllo col nome del foglio ed   +
'+ eseguire un IF...THEN...END IF                                           +
'+                                                                          +
'+ Esempio col foglio "Pippo"                                               +
'+                                                                          +
'+ If EsisteFoglio("Pippo") = True Then                                     +
'+                                                                          +
'+      Codice da eseguire in caso la Funzione restituisca Vero             +
'+                                                                          +
'+ End If                                                                   +
'+                                                                          +
'+ Argomenti della Funzione:                                                +
'+                                                                          +
'+ strNomeFoglio   -   Il nome del foglio da testare.                       +
'+                                                                          +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function EsisteFoglio(ByVal strNomeFoglio As String) As Boolean

' Gestione errore.
On Error GoTo GesErr

' Variabile oggetto contenente il Foglio in esame.
Dim objWS As Object

    ' Inizialmente viene impostato su False il risultato della Funzione.
    EsisteFoglio = False
    ' Ciclo tra tutti i Fogli del File.
    For Each objWS In Worksheets
        ' Se trova un Foglio il cui nome è uguale alla variabile strNomeFoglio passata allora.
        If strNomeFoglio = objWS.Name Then
            ' Imposta su True il risultato della Funzione.
            EsisteFoglio = True
            ' Svuota le variabili.
            objWS = Nothing
            strNomeFoglio = Empty
            ' Esce dalla Funzione.
            Exit Function
        End If
    ' Prossimo Foglio in esame.
    Next objWS

' Esce dalla Funzione.
Uscita: Exit Function
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Function" & vbCrLf & "'EsisteFoglio'" & vbCrLf & vbCrLf & Err.Description
        ' Svuota le variabili.
        objWS = Nothing
        strNomeFoglio = Empty
        Resume Uscita
' Fine della Funzione.
End Function
