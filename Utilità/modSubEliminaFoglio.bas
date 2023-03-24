Attribute VB_Name = "modSubEliminaFoglio"
Option Explicit
Option Private Module

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                                +
'+ Nome :                    Sub EliminaFoglio(ByVal strNomeFoglio As String)                     +
'+                                                                                                +
'+ Creata da :               Emanuele Tinari                                                      +
'+                                                                                                +
'+ In data :                 10/02/2023                                                           +
'+                                                                                                +
'+ Descrizione :             Sub che cancella il foglio il cui nome è passato dalla stringa       +
'+                           strNomeFoglio.                                                       +
'+                                                                                                +
'+ Uso :                     Prima di procedere all'eliminazione del foglio, controlla tramite    +
'+                           la Function EsisteFoglio se questo esiste nel file attivo. Solo se   +
'+                           la risposta alla chiamata alla Funzione è Vero procede alla          +
'+                           cancellazione del foglio.                                            +
'+                                                                                                +
'+ Valore restituito:        Nessuno                                                              +
'+                                                                                                +
'+ Esempio :                 Call EliminaFoglio("Foglio 2")                                       +
'+                                                                                                +
'+ Valore di default :       Nessuno                                                              +
'+                                                                                                +
'+ Argomento(i) :            - ByVal strNomeFoglio As String                                      +
'+                             Il nome del foglio di Excel che si vuole cancellare.               +
'+                                                                                                +
'+ Riferimento(i):           - Function EsisteFoglio(ByVal strNomeFoglio As String)               +
'+                                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub EliminaFoglio(ByVal strNomeFoglio As String)

' Gestione errore.
On Error GoTo GesErr

    ' Verifica che esista il foglio.
    If EsisteFoglio(strNomeFoglio) = True Then
        ' Disattiva l'aggiornamento dello schermo.
        Application.ScreenUpdating = False
        ' Disattiva il calcolo automatico di Excel.
        Application.Calculation = xlCalculationManual
        ' Disattiva gli avvisi di Excel.
        Application.DisplayAlerts = False
        ' Cancella il foglio.
        WorkSheets(strNomeFoglio).Delete
        ' Riattiva gli avvisi di Excel.
        Application.DisplayAlerts = True
        ' Riattiva il calcolo automatico di Excel.
        Application.Calculation = xlCalculationAutomatic
        ' Riattiva l'aggiornamento dello schermo.
        Application.ScreenUpdating = True
End If

' Esce dalla Sub, dopo aver svuotato la/e variabile/i.
Uscita: strNomeFoglio = Empty
        Exit Sub
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Sub" & vbCrLf & _
        "'EliminaFoglio'" & vbCrLf & vbCrLf & _
        "Errore Numero: " & Err.Number & vbCrLf & _
        "Descrizione dell'errore:" & vbCrLf & _
        Err.Description, vbCritical, "C'è stato un errore!"
        Resume Uscita
' Fine della Sub.
End Sub
