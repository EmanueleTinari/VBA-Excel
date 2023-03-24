Attribute VB_Name = "modFunEsisteFoglio"
Option Explicit
Option Private Module

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                                +
'+ Nome :                    Function EsisteFoglio(ByVal strNomeFoglio As String) As Boolean      +
'+                                                                                                +
'+ Creata da :               Emanuele Tinari                                                      +
'+                                                                                                +
'+ In data :                 10/02/2023                                                           +
'+                                                                                                +
'+ Descrizione :             Funzione per controllare se esiste un foglio col nome passato dalla  +
'+                           stringa strNomeFoglio.                                               +
'+                                                                                                +
'+ Uso :                     Nel codice, quando è necessario creare o eliminare un foglio è       +
'+                           necessario sapere a priori se questo esiste, per non incorrere       +
'+                           in errori nel codice.                                                +
'+                                                                                                +
'+ Valore restituito:        True: Il foglio esiste.                                              +
'+                           False: Il foglio non esiste.                                         +
'+                                                                                                +
'+ Esempio :                 Con un If...Then...End If è possibile utilizzare la Funzione.        +
'+                                                                                                +
'+ Valore di default :       Falso.                                                               +
'+                                                                                                +
'+ Argomento(i) :            - ByVal strNomeFoglio As String                                      +
'+                             Il nome del foglio di Excel che si vuole sapere se esiste.         +
'+                                                                                                +
'+ Riferimento(i):           Nessuno                                                              +
'+                                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function EsisteFoglio(ByVal strNomeFoglio As String) As Boolean

' Gestione errore.
On Error GoTo GesErr

' Var oggetto contenente il Foglio in esame.
Dim objWS As Object

    ' Inizialmente viene impostato su False il risultato della Funzione (il foglio NON esiste).
    EsisteFoglio = False
    ' Ciclo tra tutti i Fogli del File.
    For Each objWS In WorkSheets
        ' Se trova un Foglio il cui nome è uguale alla Var strNomeFoglio passata allora.
        If strNomeFoglio = objWS.Name Then
            ' Imposta su True il risultato della Funzione, il Foglio esiste.
            EsisteFoglio = True
            ' Esce dalla Funzione.
            GoTo Uscita
        End If
    ' Prossimo Foglio in esame.
    Next objWS

' Esce dalla Funzione, dopo aver svuotato la/e variabile/i.
Uscita: strNomeFoglio = Empty
        Set objWS = Nothing
        Exit Function
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Function" & vbCrLf & _
        "'EsisteFoglio'" & vbCrLf & vbCrLf & _
        "Errore Numero: " & Err.Number & vbCrLf & _
        "Descrizione dell'errore:" & vbCrLf & _
        Err.Description, vbCritical, "C'è stato un errore!"
        Resume Uscita
' Fine della Funzione.
End Function
