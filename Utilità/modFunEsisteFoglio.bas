Attribute VB_Name = "modFunEsisteFoglio"
Option Explicit
Option Private Module

' x testare la Function EsisteFoglio.
Sub Prova_EsisteFoglio()

Dim wbk As Workbook
Dim strNomeFoglio As String
Dim bolRisp As Boolean
    
    Set wbk = Application.Workbooks("Qui il nome del File di Excel: *.xlsx, *.xlmx, *.xlsb...")
    
    strNomeFoglio = "Qui il Nome del Foglio da testare se esiste"
    bolRisp = EsisteFoglio(wbk, strNomeFoglio)
    Debug.Print bolRisp

End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                                +
'+ Nome :                    Function EsisteFoglio _                                              +
'+                                    (ByVal wbk As Workbook, _                                   +
'+                                    ByVal strNomeFoglio As String) As Boolean                   +
'+                                                                                                +
'+ Creata da :               Emanuele Tinari                                                      +
'+                                                                                                +
'+ In data :                 10/02/2023                                                           +
'+                                                                                                +
'+ Descrizione :             Funzione per controllare se nel File di Excel in wbk esiste un       +
'+                           Foglio col nome passato dalla stringa strNomeFoglio.                 +
'+                                                                                                +
'+ Uso :                     Nel codice, quando è necessario creare o eliminare un Foglio è       +
'+                           necessario sapere a priori se questo esiste, per non incorrere       +
'+                           in errori nel codice.                                                +
'+                                                                                                +
'+ Valore restituito:        True: Il Foglio esiste.                                              +
'+                           False: Il Foglio non esiste.                                         +
'+                                                                                                +
'+ Esempio :                 ' x testare la Function EsisteFoglio.                                +
'+                           Sub Prova_EsisteFoglio()                                             +
'+                           Dim wbk As Workbook                                                  +
'+                           Dim strNomeFoglio As String                                          +
'+                           Dim bolRisp As Boolean                                               +
'+                               Set wbk = Application.Workbooks("Qui il nome del File di Excel") +
'+                               strNomeFoglio = "Qui il Nome del Foglio da testare se esiste"    +
'+                               bolRisp = EsisteFoglio(wbk, strNomeFoglio)                       +
'+                               Debug.Print bolRisp                                              +
'+                           End Sub                                                              +
'+                                                                                                +
'+ Valore di default :       False.                                                               +
'+                                                                                                +
'+ Argomento(i) :            - ByVal wbk As Workbook                                              +
'+                             Il nome del File di Excel di cui si vogliono esaminare i Fogli.    +
'+                                                                                                +
'+                           - ByVal strNomeFoglio As String                                      +
'+                             Il nome del Foglio che si vuole sapere se esiste.                  +
'+                                                                                                +
'+ Riferimento(i):           Nessuno                                                              +
'+                                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function EsisteFoglio(ByVal wbk As Workbook, ByVal strNomeFoglio As String) As Boolean

' Gestione errore.
On Error GoTo GesErr

' Var oggetto contenente il Foglio in esame.
Dim wst                       As Worksheet

    ' Inizialmente viene impostato su False il risultato della Funzione (il Foglio NON esiste).
    EsisteFoglio = False
    ' Ciclo tra tutti i Fogli del File.
    For Each wst In wbk.Worksheets
        ' Se trova un Foglio il cui nome è uguale alla Var strNomeFoglio passata allora.
        If strNomeFoglio = wst.Name Then
            ' Imposta su True il risultato della Funzione, il Foglio esiste.
            EsisteFoglio = True
            ' Esce dalla Funzione.
            GoTo Uscita
        End If
    ' Prossimo Foglio in esame.
    Next wst

' Esce dalla Funzione, dopo aver svuotato la/e variabile/i.
Uscita: strNomeFoglio = Empty
        Set wst = Nothing
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
