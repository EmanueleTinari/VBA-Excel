
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                +
'+ Funzione per valutare se un dato URL esiste.                                   +
'+ Valuta anche se il server del sito, di cui l'URL, è online.                    +
'+ Restituisce Vero o Falso.                                                      +
'+                                                                                +
'+ Uso: Prima di tentare il collegamento ad una fagina internet, si può far       +
'+ eseguire alla funzione un controllo per vedere se l'indirizzo è raggiungibile  +
'+ oppure no eseguendo un IF...THEN...END IF                                      +
'+                                                                                +
'+ Esempio con Google all'indirizzo internet https://www.google.com/              +
'+                                                                                +
'+ If EsisteURL("https://www.google.com/") = True Then                            +
'+                                                                                +
'+      Codice da eseguire in caso la funzione restituisca Vero                   +
'+                                                                                +
'+ End If                                                                         +
'+                                                                                +
'+ Argomenti della funzione:                                                      +
'+                                                                                +
'+ strTestURL   -   l'URL (indirizzo Internet) della pagina che vogliamo testare. +
'+                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function EsisteURL(ByVal strTestURL As String) As Boolean

' Gestione errore.
On Error GoTo GesErr

' Oggetto contenente la richiesta WinHttpRequest.
Dim Request As Object
' Conterrà la risposta alla richiesta WinHttpRequest.
Dim rc As Variant

    ' Inizialmente viene impostato su False il risultato della Funzione.
    EsisteURL = False
    Set Request = CreateObject("WinHttp.WinHttpRequest.5.1")
    With Request
      .Open "GET", strTestURL, False
      .send
      rc = .statusText
    End With
    ' Se la risposta è "OK" allora.
    If rc = "OK" Then
        ' Imposta su True il risultato della Funzione.
        EsisteURL = True
        ' Svuota le variabili.
        Set Request = Nothing
        Set rc = Nothing
        strTestURL = Empty
        ' Esce dalla funzione.
        Exit Function
    ' Se la risposta è differente da "OK" allora.
    ElseIf rc <> "OK" Then
        ' Lascia invariato il valore della funzione.
        EsisteURL = False
        ' Svuota le variabili.
        Set Request = Nothing
        Set rc = Nothing
        strTestURL = Empty
        ' Esce dalla funzione.
        Exit Function
    End If

' Esce dalla funzione.
Uscita: Exit Function
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Function" & vbCrLf & "'EsisteURL'" & vbCrLf & vbCrLf & Err.Description
        ' Svuota le variabili.
        Set Request = Nothing
        Set rc = Nothing
        strTestURL = Empty
        Resume Uscita
' Fine della funzione.
End Function
