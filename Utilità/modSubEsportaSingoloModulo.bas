Attribute VB_Name = "modSubEsportaSingoloModulo"
Option Explicit
Option Private Module

' x testare la Sub EsportaSingoloModulo.
Sub Prova_EsportaSingoloModulo()

' Gestione errore.
On Error GoTo GesErr

Dim wbk                         As Workbook
Dim strSerieModuli              As String
Dim nomeModuli()                As String
Dim intI                        As Integer
    
    ' Commentare i moduli che non si vuole esportare.

    strSerieModuli = strSerieModuli & "modFunEsisteFoglio,"
    strSerieModuli = strSerieModuli & "modFunEsisteModulo,"
    strSerieModuli = strSerieModuli & "modFunMesgBox,"
    strSerieModuli = strSerieModuli & "modFunScriviFileTemp,"
    strSerieModuli = strSerieModuli & "modFunSelezionaCartella"

    strSerieModuli = strSerieModuli & "modSubEliminaFoglio,"
    strSerieModuli = strSerieModuli & "modSubEsportaSingoloModulo,"
    strSerieModuli = strSerieModuli & "modSubSalvataggioAlVolo,"
    strSerieModuli = strSerieModuli & "modSubStampaInImmediata,"
    strSerieModuli = strSerieModuli & "modSubStampaRiferEsistenti,"
    strSerieModuli = strSerieModuli & "modSubVerificaCartella,"

    nomeModuli = Split(Trim(strSerieModuli), ",")
    
    Set wbk = Workbooks("PERSONAL.XLSB")
    
    For intI = LBound(nomeModuli()) To UBound(nomeModuli())
        Call EsportaSingoloModulo(nomeModuli(intI), wbk, strGitUt)
    Next intI

' Esce dalla Sub, dopo aver svuotato la/e variabile/i.
Uscita: ReDim nomeModuli(0)
        strSerieModuli = Empty
        intI = Empty
        Set wbk = Nothing
        Exit Sub
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Sub" & vbCrLf & _
        "'Prova_EsportaSingoloModulo'" & vbCrLf & vbCrLf & _
        "Errore Numero: " & Err.Number & vbCrLf & _
        "Descrizione dell'errore:" & vbCrLf & _
        Err.Description, vbCritical, "C'è stato un errore!"
        Resume Uscita
' Fine della Sub.
End Sub



'+ Descrizione :             Routine che estrae e salva il Modulo il cui nome è passato           +
'+                           dall'Argomento strNomeModulo nel File di Excel passato dall'Argomento    +
'+                           opzionale wbk. Se quest'ultimo non è fornito, viene usato il valore  +
'+                           ThisWorkbook. E' opzionale anche passare l'Argomento strPercorso per +
'+                           indicare dove salvare il file estratto. Se questo non viene fornito, +
'+                           si apre una finestra di dialogo per indicarlo.                       +
'+                                                                                                +

'+                                                                                                +
'+ Esempio :                 ' x testare la Sub EsportaSingoloModulo.                             +
'+                           Sub Prova_EsportaSingoloModulo()                                     +
'+                           Dim wbk As Workbook                                                  +
'+                           Dim strNomeModulo As String                                          +
'+                           Dim strPercorsoDiSalvataggio As String                               +
'+                               Set wbk = Workbooks("xxx")                                       +
'+                               strNomeModulo = "xxx"                                            +
'+                               strPercorsoDiSalvataggio = "x:\xx\xxx"                           +
'+                               ' Con tutti e tre gli Argomenti.                                 +
'+                               Call EsportaSingoloModulo _                                      +
'+                                    (strNomeModulo, wbk, strPercorsoDiSalvataggio)              +
'+                               ' Con solo l'Argomento strNomeModulo                             +
'+                               Call EsportaSingoloModulo(strNomeModulo)                         +
'+                               Set wbk = Nothing                                                +
'+                           End Sub                                                              +
'+                                                                                                +
'+ Valore di default :       - Se l'Argomento wbk non viene passato, viene assunto ThisWorkbook.  +
'+                                                                                                +
'+ Argomento(i) :            - ByVal strNomeModulo As String                                      +
'+                             Il nome del Modulo che si vuole esportare.                         +
'+                                                                                                +
'+                           - Optional ByRef wbk As Workbook                                     +
'+                             Facoltativo. Questo file o un altro da cui si vuole esportare      +
'+                             il Modulo in strNomeModulo.                                        +
'+                                                                                                +
'+                           - Optional ByVal strPercorso As String                               +
'+                             Facoltativo. Il percorso dove si vuole salvare il Modulo estratto. +
'+                             Se non viene indicato nella Var, si aprirà una finestra di dialogo +
'+                             per la scelta della Cartella.                                      +
'+                                                                                                +
'+ Riferimento(i):           - Riferimento a Microsoft Office 16.0 Object Library                 +
'+                             Lib.: in "C:\Program Files (x86)\Common Files\Microsoft Shared\ _  +
'+                                       OFFICE16\MSO.dll"                                        +
'+                             GUID: "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}"                     +
'+                                                                                                +
'+                           - Microsoft Visual Basic for Applications Extensibility              +
'+                             Lib.: in "C:\Program Files (x86)\Common Files\ _                   +
'+                                       Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"                   +
'+                             GUID: "{0002E157-0000-0000-C000-000000000046}"                     +
'+                                                                                                +
'+                           - Function SelezionaCartella _                                       +
'+                                               (Optional ByVal strPercorso As String) As String +
'+                                                                                                +
'+                           - Function EsisteModulo (ByVal strNomeModulo As String, _            +
'+                                                    Optional ByVal wbk As Workbook) As Boolean  +
'+                                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub EsportaSingoloModulo(ByVal strNomeModulo As String, _
                        Optional ByVal wbk As Workbook, _
                        Optional ByVal strPercorso As String)

' Gestione errore.
On Error GoTo GesErr

Dim bolEsporta                  As Boolean
' Var oggetto contenente il Componente.
Dim objVBComp                   As VBIDE.vbComponent

    ' Se l'Argomento wbk passato alla Sub è vuoto, allora.
    If wbk Is Nothing Or Null Then
        ' Imposta nella Var wbk il ThisWorkbook.
        Set wbk = ThisWorkbook
    ' Altrimenti.
    Else
        ' Imposto nella Var wbk il WorkBook passato come Argomento alla Sub.
        Set wbk = wbk
    End If
    
    ' Controllo che il File di Excel non abbia la visualizzazione protetta del Progetto VBA.
    If wbk.VBProject.Protection = 1 Then
        ' Se la visualizzazione è protetta avvisa ed esce.
        MsgBox "Il codice VBA del File di Excel " & _
        wbk.Name & _
        " è protetto, pertanto non è possibile esportarne il Modulo " & strNomeModulo & "."
        GoTo Uscita
    End If
    ' Se l'Argomento strPercorso è vuoto, si apre la finestra di selezione Cartella.
    If strPercorso = "" Then
        strPercorso = SelezionaCartella
    End If
    ' Se l'ultimo carattere della Var strPercorso è diverso da "\", allora.
    If Right(strPercorso, 1) <> "\" Then
        ' Aggiunge un "\" alla fine della Var strPercorso.
        strPercorso = strPercorso & "\"
    End If
    ' Se il Modulo strNomeModulo è nel progetto, allora.
    If EsisteModulo(strNomeModulo, wbk) = True Then
        ' Stampa un messaggio nella Finestra Immediata.
        Debug.Print "Il Modulo " & Chr(34) & strNomeModulo & Chr(34) & _
                    " è nel File "; wbk.Name
    ' Se il Modulo strNomeModulo non c'è, scrive un avviso nella Finestra Immediata.
    ElseIf EsisteModulo(strNomeModulo, wbk) = False Then
        Debug.Print "Il Modulo " & Chr(34) & strNomeModulo & Chr(34) & _
                    " non esiste nel File "; wbk.Name
        ' Esce dalla routine.
        GoTo Uscita
    End If
        Set objVBComp = wbk.VBProject.VBComponents(strNomeModulo)
            ' Pone la Var su Vero.
            bolEsporta = True
            ' Concatena l'estensione del file per l'esportazione.
            Select Case objVBComp.Type
                Case vbext_ct_ClassModule
                    strNomeModulo = strNomeModulo & ".cls"
                Case vbext_ct_MSForm
                    strNomeModulo = strNomeModulo & ".frm"
                Case vbext_ct_StdModule
                    strNomeModulo = strNomeModulo & ".bas"
                Case vbext_ct_Document
                    'Questo è il Workbook o un WorkSheet, non tentare di esportarlo.
                    bolEsporta = False
            End Select
            ' Se la Var bolEsporta è Vero, allora.
            If bolEsporta = True Then
                ' Esporta il componente in un file di testo.
                objVBComp.Export strPercorso & strNomeModulo
            End If

' Esce dalla Sub, dopo aver svuotato la/e variabile/i.
Uscita: Set wbk = Nothing
        strNomeModulo = Empty
        strPercorso = Empty
        Set objVBComp = Nothing
        Exit Sub
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Sub" & vbCrLf & _
        "'EsportaSingoloModulo'" & vbCrLf & vbCrLf & _
        "Errore Numero: " & Err.Number & vbCrLf & _
        "Descrizione dell'errore:" & vbCrLf & _
        Err.Description, vbCritical, "C'è stato un errore!"
        Resume Uscita
' Fine della Sub.
End Sub
