Attribute VB_Name = "modSubSalvataggioAlVolo"
Option Explicit
Option Private Module

' x testare la Sub SalvataggioAlVolo.OK.
Public Sub Prova_SalvataggioAlVolo()

Const strNomeFile               As String = "PERSONAL.XLSB"
    
    SalvataggioAlVolo (strNomeFile)

End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                                +
'+ Nome :                    SalvataggioAlVolo(ByVal strNomeFile As String)                       +
'+                                                                                                +
'+ Creata da :               Emanuele Tinari                                                      +
'+                                                                                                +
'+ In data :                 10/02/2023                                                           +
'+                                                                                                +
'+ Descrizione :             Sub che salva la Cartella di lavoro senza chiedere l'iterazione con  +
'+                           l'Utente.                                                            +
'+                                                                                                +
'+ Uso :                     Nel codice, quando è necessario salvare il file prima di procedere.  +
'+                                                                                                +
'+ Valore restituito:        Nessuno                                                              +
'+                                                                                                +
'+ Esempio :                 SalvataggioAlVolo(strNomeFile)                                       +
'+                                                                                                +
'+                           ' x testare la Sub SalvataggioAlVolo.                                +
'+                           Sub ProvaSalvataggioAlVolo()                                         +
'+                           Const strNomeFile As String = "nFile.xlsx" (oppure nFile.xlsm ecc.)  +
'+                               SalvataggioAlVolo (strNomeFile)                                  +
'+                           End Sub                                                              +
'+                                                                                                +
'+ Valore di default :       Nessuno                                                              +
'+                                                                                                +
'+ Argomento(i) :            - ByVal strNomeFile As String                                        +
'+                             Il nome del file aperto (xlsx, xlsm, xlsb...) da salvare.          +
'+                                                                                                +
'+ Riferimento(i):           Nessuno                                                              +
'+                                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub SalvataggioAlVolo(ByVal strNomeFile As String)
        
' Gestione errore.
On Error GoTo GesErr

    ' Disattiva l'aggiornamento dello schermo.
    Application.ScreenUpdating = False
    ' Disattiva il calcolo automatico di Excel.
    Application.Calculation = xlCalculationManual
    ' Disattiva gli avvisi di Excel.
    Application.DisplayAlerts = False
    ' Salva il file il cui nome è passato dalla stringa strNomeFile.
    Application.Workbooks(strNomeFile).Save
    ' Riattiva gli avvisi di Excel.
    Application.DisplayAlerts = True
    ' Riattiva il calcolo automatico di Excel.
    Application.Calculation = xlCalculationAutomatic
    ' Riattiva l'aggiornamento dello schermo.
    Application.ScreenUpdating = True

' Esce dalla Sub, dopo aver svuotato la/e variabile/i.
Uscita: strNomeFile = Empty
        Exit Sub
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Sub" & vbCrLf & "'SalvataggioAlVolo'" & vbCrLf & vbCrLf & "Errore Numero: " & Err.Number & vbCrLf & "Descrizione dell'errore:" & vbCrLf & Err.Description, vbCritical, "C'è stato un errore!"
        Resume Uscita
' Fine della Sub.
End Sub
