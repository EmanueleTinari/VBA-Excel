Attribute VB_Name = "modSubSalvaFileDiExcel"
Option Explicit
Option Private Module

' x testare la Sub SalvaFileDiExcel.
Sub Prova_SalvaFileDiExcel()

Const strNomeFileDiExcel As String = "Cartel1.xlsx"
Dim wbk As Workbook

    '(*.xlsx, *.xlsm, *.xlsb, ...).
    Set wbk = Application.Workbooks(strNomeFileDiExcel)
    ' Con wbk e con Nome File.
    Call SalvaFileDiExcel(wbk, strNomeFileDiExcel)
    ' Senza wbk ma con nome File.
    Call SalvaFileDiExcel(, strNomeFileDiExcel)
    ' Con wbk ma senza Nome File.
    Call SalvaFileDiExcel(wbk, "")

End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                                +
'+ Nome :                    SalvaFileDiExcel _                                                   +
'+                                         (Optional wbk As Workbook, _                           +
'+                                         Optional ByVal strNomeFile As String)                  +
'+                                                                                                +
'+ Creata da :               Emanuele Tinari                                                      +
'+                                                                                                +
'+ In data :                 28/03/2023                                                           +
'+                                                                                                +
'+ Descrizione :             Sub che salva la Cartella di lavoro indicata in wbk o in strNomeFile +
'+                           senza chiedere l'iterazione con l'Utente.                            +
'+                                                                                                +
'+ Uso :                     Nel codice, quando è necessario salvare il file prima di procedere.  +
'+                                                                                                +
'+ Valore restituito:        Nessuno                                                              +
'+                                                                                                +
'+ Esempio :                 ' x testare la Sub SalvaFileDiExcel.                                 +
'+                           Sub Prova_SalvaFileDiExcel()                                         +
'+                           Const strNomeFileDiExcel As String = "Cartel1.xlsx"                  +
'+                           Dim wbk As Workbook                                                  +
'+                               '(*.xlsx, *.xlsm, *.xlsb, ...).                                  +
'+                               Set wbk = Application.Workbooks(strNomeFileDiExcel)              +
'+                               ' Con wbk e con Nome File.                                       +
'+                               Call SalvaFileDiExcel(wbk, strNomeFileDiExcel)                   +
'+                               ' Senza wbk ma con nome File.                                    +
'+                               Call SalvaFileDiExcel(, strNomeFileDiExcel)                      +
'+                               ' Con wbk ma senza Nome File.                                    +
'+                           Call SalvaFileDiExcel(wbk, "")                                       +
'+                           End Sub                                                              +
'+                                                                                                +
'+ Valore di default :       Nessuno                                                              +
'+                                                                                                +
'+ Argomento(i) :            - Optional ByVal wbk As Workbook                                     +
'+                             Facoltativo. Il nome del file aperto (xlsx, xlsm, xlsb...)         +
'+                             da salvare.                                                        +
'+                                                                                                +
'+                           - Optional ByVal strNomeFile As String                               +
'+                             Facoltativo. Il nome del file aperto (xlsx, xlsm, xlsb...)         +
'+                             da salvare.                                                        +
'+                                                                                                +
'+ Riferimento(i):           Nessuno                                                              +
'+                                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub SalvaFileDiExcel(Optional wbk As Workbook, Optional ByVal strNomeFile As String)

' Gestione errore.
On Error GoTo GesErr
    
    ' Disattiva gli avvisi di Excel.
    Application.DisplayAlerts = False
    ' Disattiva il calcolo automatico di Excel.
    Application.Calculation = xlCalculationManual
    ' Disattiva l'aggiornamento dello schermo.
    Application.ScreenUpdating = False
    
    If wbk Is Nothing And strNomeFile = "" Then
        GoTo Ripristina
    ElseIf Not wbk Is Nothing Or strNomeFile = "" Then
        ' Se il File passato tramite il wbk non è stato salvato, allora.
        If wbk.Saved = False Then
            ' Salva il File.
            wbk.Save
            GoTo Ripristina
        End If
    ElseIf wbk Is Nothing And strNomeFile <> "" Then
        ' Se il File passato tramite la strNomeFile non è stato salvato, allora.
        If Application.Workbooks("" & strNomeFile & "").Saved = False Then
            ' Salva il File (*.xlsx, *.xlsm, *.xlsb).
            Application.Workbooks("" & strNomeFile & "").Save
            GoTo Ripristina
        End If
    End If

Ripristina: ' Ripristina l'aggiornamento dello schermo.
            Application.ScreenUpdating = True
            ' Ripristina il calcolo automatico di Excel.
            Application.Calculation = xlCalculationAutomatic
            ' Ripristina gli avvisi di Excel.
            Application.DisplayAlerts = True
        
' Esce dalla Sub, dopo aver svuotato la/e variabile/i.
Uscita: Set wbk = Nothing
        strNomeFile = Empty
        Exit Sub
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Sub" & vbCrLf & "'SalvaFileDiExcel'" & vbCrLf & vbCrLf & "Errore Numero: " & Err.Number & vbCrLf & "Descrizione dell'errore:" & vbCrLf & Err.Description, vbCritical, "C'è stato un errore!"
        Resume Uscita
' Fine della Sub.
End Sub
