
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                +
'+ Funzione per valutare se un dato URL esiste.                                   +
'+ Valuta anche se il server del sito, di cui l'URL, è online.                    +
'+ Restituisce Vero se trova l'URL inviato tramite la stringa strTestURL,         +
'+ oppure Falso se non trova l'URL inviato.                                       +
'+                                                                                +
'+ Il valore di Default della Funzione EsisteURL è Falso.                         +
'+                                                                                +
'+ Uso: Prima di tentare il collegamento ad una fagina internet, si può far       +
'+ eseguire alla Funzione un controllo per vedere se l'indirizzo è raggiungibile  +
'+ oppure no eseguendo un IF...THEN...END IF                                      +
'+                                                                                +
'+ Esempio con Google all'indirizzo internet https://www.google.com/              +
'+                                                                                +
'+ If EsisteURL("https://www.google.com/") = True Then                            +
'+                                                                                +
'+      Codice da eseguire in caso la Funzione restituisca Vero                   +
'+                                                                                +
'+ End If                                                                         +
'+                                                                                +
'+ Argomenti della Funzione:                                                      +
'+                                                                                +
'+ strTestURL   -   l'URL (indirizzo Internet) della pagina che vogliamo testare. +
'+                                                                                +
'+ Richiede il riferimento a Microsoft WinHTTP Services, Version 5.1              +
'+ Sono necessarie le Funzioni EsisteRiferimento(wbk, strGUID) per testare se il  +
'+ riferimento, la cui GUID è passata tramite la Variabile stringa, esiste        +
'+ nel progetto, per evitare errori in caso non esistesse e                       +
'+ AggiungiRiferimento (wbk, strGUID) per aggiungere quest'ultimo in automatico,  +
'+ se non è presente il riferimento nel progetto.                                 +
'+ È possibile aggiungere manualmente il riferimento richiesto ed eliminare la    +
'+ parte di funzione che si occupa del controllo ed aggiunta Riferimenti.         +
'+                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function EsisteURL(ByVal strTestURL As String) As Boolean

' Gestione errore.
On Error GoTo GesErr

' Oggetto contenente la richiesta WinHttpRequest.
Dim Request As Object
' Conterrà la prima risposta alla richiesta WinHttpRequest.
Dim rc1 As Variant
' Conterrà la seconda risposta alla richiesta WinHttpRequest.
Dim rc2 As Variant
' Variabile del Riferimento a Microsoft WinHTTP Services.
Dim strRif As String

    ' Stringa che rappresenta la GUID del Riferimento.
    strRif = "{662901FC-6951-4854-9EB2-D9A2570F2B2E}"
        If EsisteRiferimento(ThisWorkbook, strRif) = False Then
            AggiungiRiferimento ThisWorkbook, strRif
        End If
    
    ' Inizialmente viene impostato su False il risultato della Funzione.
    EsisteURL = False
    Set Request = CreateObject("WinHttp.WinHttpRequest.5.1")
    With Request
      .Open "GET", strTestURL, False
      .send
      rc1 = .statusText
      rc2 = .Status
    End With
    ' Se la prima risposta è "OK" e la seconda risposta ha valore 200 allora.
    If rc1 = "OK" And rc2 = 200 Then
        ' Imposta su True il risultato della Funzione.
        EsisteURL = True
        ' Svuota le variabili.
        Set Request = Nothing
        Set rc1 = Nothing
        Set rc2 = Nothing
        strTestURL = Empty
        ' Esce dalla Funzione.
        Exit Function
    ' Se la prima risposta è differente da "OK" oppure se la seconda risposta ha valore diverso da 200 allora.
    ElseIf rc1 <> "OK" Or rc2 <> 200 Then
        ' Lascia invariato il valore della Funzione.
        EsisteURL = False
        ' Svuota le variabili.
        Set Request = Nothing
        Set rc1 = Nothing
        Set rc2 = Nothing
        strTestURL = Empty
        ' Esce dalla Funzione.
        Exit Function
    End If

' Esce dalla Funzione.
Uscita: Exit Function
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Function" & vbCrLf & "'EsisteURL'" & vbCrLf & vbCrLf & Err.Description
        ' Svuota le variabili.
        Set Request = Nothing
        Set rc1 = Nothing
        Set rc2 = Nothing
        strTestURL = Empty
        Resume Uscita
' Fine della Funzione.
End Function

