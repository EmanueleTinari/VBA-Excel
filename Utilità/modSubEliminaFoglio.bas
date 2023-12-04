Attribute VB_Name = "modSubEliminaFoglio"
Option Explicit
Option Private Module

' x testare la Function EliminaFoglio.
Sub Prova_EliminaFoglio()

Dim wbk As Workbook
    
    Set wbk = Application.Workbooks("Cartel1.xlsx")
    Call EliminaFoglio(wbk, "File_allegati_al_Progetto")

End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                                +
'+ Nome :                    Sub EliminaFoglio(ByVal wbk As Workbook, strNomeFoglio As String)    +
'+                                                                                                +
'+ Creata da :               Emanuele Tinari                                                      +
'+                                                                                                +
'+ In data :                 10/02/2023                                                           +
'+                                                                                                +
'+ Descrizione :             Sub che cancella dal File Excel in wbk, il Foglio il cui nome è      +
'+                           passato dalla stringa strNomeFoglio.                                 +
'+                                                                                                +
'+ Uso :                     Prima di procedere all'eliminazione del Foglio, controlla tramite    +
'+                           la Function EsisteFoglio se questo esiste nel File wbk e se non è    +
'+                           l'unico Foglio. Solo se entrambe le condizioni risultano vere,       +
'+                           procede alla cancellazione del Foglio.                               +
'+                                                                                                +
'+ Valore restituito:        Nessuno                                                              +
'+                                                                                                +
'+ Esempio :                 ' x testare la Function EliminaFoglio.                               +
'+                           Sub Prova_EliminaFoglio()                                            +
'+                           Dim wbk As Workbook                                                  +
'+                               Set wbk = Application.Workbooks("Qui il nome del File di Excel") +
'+                               Call EliminaFoglio(wbk, "Qui il nome del Foglio")                +
'+                           End Sub                                                              +
'+                                                                                                +
'+ Valore di default :       Nessuno                                                              +
'+                                                                                                +
'+ Argomento(i) :            - ByVal wbk As Workbook                                              +
'+                             Il nome del File di Excel di cui si vuole eliminare il Foglio.     +
'+                                                                                                +
'+                           - ByVal strNomeFoglio As String                                      +
'+                             Il nome del Foglio di Excel che si vuole cancellare.               +
'+                                                                                                +
'+ Riferimento(i):           - Function EsisteFoglio(ByVal strNomeFoglio As String)               +
'+                                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub EliminaFoglio(ByVal wbk As Workbook, strNomeFoglio As String)

' Gestione errore.
On Error GoTo GesErr

    ' Verifica che esista il Foglio.
    If EsisteFoglio(wbk, strNomeFoglio) = True And wbk.Worksheets.Count > 1 Then
        ' Disattiva l'aggiornamento dello schermo.
        Application.ScreenUpdating = False
        ' Disattiva il calcolo automatico di Excel.
        Application.Calculation = xlCalculationManual
        ' Disattiva gli avvisi di Excel.
        Application.DisplayAlerts = False
        ' Cancella il Foglio.
        Worksheets(strNomeFoglio).Delete
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
