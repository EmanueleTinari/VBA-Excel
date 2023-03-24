Attribute VB_Name = "modSubStampaRiferEsistenti"
Option Explicit
Option Private Module

' x testare la Sub StampaRiferEsistenti.
Public Sub Prova_StampaRiferEsistenti()

' Variabile oggetto per il WorkBook.
Dim wbk As Workbook

    ' Esempio senza Argomenti (usa ThisWorkbook come wbk e invia il resoconto alla Finestra Immediata).
    StampaRiferEsistenti
    ' Esempio con argomenti settati entrambi, stampa in un file txt nella cartella Temp.
    Set wbk = Workbooks("PERSONAL.XLSB")
    Call StampaRiferEsistenti(wbk, False)

End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                                +
'+ Nome :                    Sub StampaRiferEsistenti(Optional ByRef wbk As Workbook, _           +
'+                                                    Optional ByVal bolTxt As Boolean = False)   +
'+                                                                                                +
'+ Creata da :               Emanuele Tinari                                                      +
'+                                                                                                +
'+ In data :                 11/02/2023                                                           +
'+                                                                                                +
'+ Descrizione :             Routine che compie un ciclo tra tutti i Riferimenti del file passato +
'+                           dall'Argomento wbk e crea una stringa che si può scegliere di        +
'+                           inviare ad un file di testo oppure alla finestra Immediata.          +
'+                                                                                                +
'+ Uso :                     Eseguire l'esempio, impostando manualmente il nome del file di       +
'+                           Excel oppure utilizzando la finestra di dialogo scegli file.         +
'+                           L'Argomento bolTxt può essere impostato manualmente o tramite        +
'+                           variabile.                                                           +
'+                                                                                                +
'+ Valore restituito:        Nessuno                                                              +
'+                                                                                                +
'+ Esempio :                 ' x testare la Sub StampaRiferEsistenti.                             +
'+                           Sub ProvaStampaRiferEsistenti()                                      +
'+                           Dim wbk As Workbook                                                  +
'+                               Set wbk = Workbooks("FILE.XLSM") oppure ("FILE.XLSB")...         +
'+                               ' Invia al file di testo.                                        +
'+                               Call StampaRiferEsistenti(wbk, True)                             +
'+                               ' Invia alla finestra Immediata.                                 +
'+                               Call StampaRiferEsistenti(wbk, False)                            +
'+                           End Sub                                                              +
'+                                                                                                +
'+ Valore di default :       - Se l'Argomento wbk non viene passato, viene assunto ThisWorkbook.  +
'+                                                                                                +
'+                           - Se l'Argomento bolTxt non viene passato, si assume False.          +
'+                                                                                                +
'+ Argomento(i) :            - ByVal wbk As Workbook                                              +
'+                             Questo file o un altro di cui si vogliono ottenere i Riferimenti.  +
'+                                                                                                +
'+                           - ByVal bolTxt As Boolean                                            +
'+                             True. Crea un file di testo contenente la lista dei Riferimenti.   +
'+                             False. Invia la lista alla finestra Immediata.                     +
'+                                                                                                +
'+ Riferimento(i):           - Function ScriviFileTemp(strTesto As String, _                      +
'+                                                    Optional strPercorso As String, _           +
'+                                                    Optional strNomeFile As String, _           +
'+                                                    Optional strEstensione As String = "txt") _ +
'+                                                    As String                                   +
'+                                                                                                +
'+                           - Function StampaInImmediata _                                       +
'+                                      (ByVal strDaStampareInImmediata As String)                +
'+                                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub StampaRiferEsistenti(Optional ByVal wbk As Workbook, Optional ByVal bolTxt As Boolean = False)

' Gestione errore.
On Error GoTo GesErr

' Serve per il ciclo tra tutti i Riferimenti.
Dim intI As Integer
' La Var conterrà la stringa che si andrà a creare.
Dim strTesto As String
    
    ' Se la Var wbk passata come argomento della Function è Nothing, allora.
    If wbk Is Nothing Then
        ' Imposta nella Var Oggetto il ThisWorkbook.
        Set wbk = Application.ThisWorkbook
    End If
    ' Consideriamo solo di questo file di Excel tutti i suoi Riferimenti attivi.
    With wbk.VBProject.References
        ' Ciclo tra i Riferimenti.
        For intI = 1 To .Count
            ' Inserisce la descrizione.
            strTesto = strTesto & "Descrizione: " & .item(intI).Description & vbCrLf
            ' Inserisce il relativo nome.
            strTesto = strTesto & "Nome: " & .item(intI).Name & vbCrLf
            ' Inserisce la GUID.
            strTesto = strTesto & "GUID: " & .item(intI).Guid & vbCrLf
            ' Inserisce la maggiore versione.
            strTesto = strTesto & "M: " & .item(intI).Major & vbCrLf
            ' Inserisce la minore versione.
            strTesto = strTesto & "m: " & .item(intI).Minor & vbCrLf
            ' Inserisce il percorso del file riferito.
            strTesto = strTesto & "Path: " & .item(intI).FullPath & vbCrLf
            ' Lascia una linea in bianco.
            strTesto = strTesto & vbCrLf
        ' Prossimo Riferimento.
        Next intI
            ' Se la conta dei Riferimenti è 0 allora.
            If .Count = 0 Then
                ' Il successivo messaggio viene inserito nella Var stringa.
                strTesto = strTesto & "Nel file di Excel:" & Chr(13) & wbk.Name & Chr(13) & "non ci sono Riferimenti attivi."
            ' Se la conta dei Riferimenti è 1 allora.
            ElseIf .Count = 1 Then
                ' Il successivo messaggio viene inserito nella Var stringa.
                strTesto = strTesto & "Nel file di Excel:" & Chr(13) & wbk.Name & Chr(13) & "c'è " & .Count & " riferimento attivo."
            ' Se la conta dei Riferimenti è maggiore di 1 allora.
            ElseIf .Count > 1 Then
                ' Il successivo messaggio viene inserito nella Var stringa.
                strTesto = strTesto & "Nel file di Excel:" & Chr(13) & wbk.Name & Chr(13) & "ci sono " & .Count & " Riferimenti attivi."
            End If
    End With
    ' Se è stato inviato alla Funzione Vero, allora.
    If bolTxt = True Then
        ' Crea un file di testo contenente il valore della stringa strTesto.
        ScriviFileTemp (strTesto)
    ' Se invece è stato inviato alla Funzione Falso, allora.
    ElseIf bolTxt = False Then
        ' Invia alla finestra Immediata la stringa strTesto.
        StampaInImmediata (strTesto)
    End If

' Esce dalla Sub, dopo aver svuotato la/e variabile/i.
Uscita: Set wbk = Nothing
        bolTxt = False
        intI = Empty
        strTesto = Empty
        Exit Sub
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Sub" & vbCrLf & _
        "'StampaRiferEsistenti'" & vbCrLf & vbCrLf & _
        "Errore Numero: " & Err.Number & vbCrLf & _
        "Descrizione dell'errore:" & vbCrLf & _
        Err.Description, vbCritical, "C'è stato un errore!"
        Resume Uscita
' Fine della Sub.
End Sub
