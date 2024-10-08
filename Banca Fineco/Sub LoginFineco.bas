Option Explicit

                                '++++++++++++++++++++++
                                '+ Costanti pubbliche +
                                '++++++++++++++++++++++

' Contiene l'URL della pagina di Login del sito di Fineco.
Public Const strURLlogin As String = "https://finecobank.com/it/online/"

                                '+++++++++++++++++++++++
                                '+ Variabili pubbliche +
                                '+++++++++++++++++++++++

    '+++++++++++++++++++++
    '+ VARIABILI OGGETTO +
    '+++++++++++++++++++++

Public eDriver As New Selenium.EdgeDriver           ' Richiamo un nuovo oggetto contenente il Driver per Edge.

    '++++++++++++++++++++
    '+ VARIABILI STRING +
    '++++++++++++++++++++

Public strUtente As String                          ' Contiene lo Username passato dalla InputBox.
Public strPass As String                            ' Contiene la Password passata dalla InputBox.

    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '+                                                                +
    '+ Procedura di Login automatico sul sito di Banca Fineco.        +
    '+ Per funzionare necessita:                                      +
    '+                                                                +
    '+ - Che sia installato e funzionante il Selenium Driver, nel mio +
    '+ caso per Edge. Seguire le guide su Google.                     +
    '+                                                                +
    '+ - Una volta installato, deve essere attivato nel file di Excel +
    '+ il riferimento alla Selenium Type Library.                     +
    '+                                                                +
    '+ È necessaria la Funzione EsisteURL(ByVal strTestURL As String) +
    '+ per testare il collegamento prima del suo uso.                 +
    '+                                                                +
    '+ Uso: nel codice, quando serve, inserire la chiamata alla Sub:  +
    '+                                                                +
    '+ Call LoginFineco                                               +
    '+                                                                +
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Sub LoginFineco()
    ' Controllo degli errori
    On Error Resume Next
    ' Se il controllo dell'indirizzo tramite la Funzione EsisteURL dà risultato positivo (l'indirizzo esiste), allora.
    If EsisteURL(strURLlogin) = True Then
        ' Le seguenti operazioni vengono effettuate sull'oggetto pubblico contenente il Selenium Driver.
        With eDriver
            ' Avvia il browser Edge.
            .start "Edge", ""
            ' Attende 3 secondi.
            Application.Wait Now + TimeSerial(0, 0, 3)
            ' Apre l'URL della pagina di Login di Fineco contenuto nella costante strURLlogin.
            .Get strURLlogin
            ' Attende 3 secondi.
            Application.Wait Now + TimeSerial(0, 0, 3)
            ' Se è presente l'elemento Bottone per accettare i cookies allora.
            If eDriver.IsElementPresent(sBy.XPath("//button[text()='ACCETTA TUTTI I COOKIES']")) Then
                ' Lo preme.
                .FindElementByXPath("//button[text()='ACCETTA TUTTI I COOKIES']").Click
            End If
            ' Attende 3 secondi.
            Application.Wait Now + TimeSerial(0, 0, 3)
            ' Se è presente l'elemento Bottone per accedere al Login allora.
            If eDriver.IsElementPresent(sBy.XPath("//a[contains(@href, '/it/online/login/')]")) Then
                ' Lo preme.
                .FindElementByXPath("//a[contains(@href, '/it/online/login/')]").Click
            End If
            ' Attende 1 secondo.
            Application.Wait Now + TimeSerial(0, 0, 1)
            ' Se è presente l'elemento Input per l'inserimento dell'Utente allora.
            If eDriver.IsElementPresent(sBy.XPath("//input[@id='user']")) Then
                ' Pone nella variabile il valore richiesto con la InputBox.
                strUtente = InputBox("Inserisci il tuo Codice Utente", "User", "xxxxxxxx")
                ' Lo preme.
                .FindElementByXPath("//input[@id='user']").Click
                ' Invia la variabile contenente l'Utente.
                .FindElementByXPath("//input[@id='user']").SendKeys strUtente
            End If
            ' Attende 1 secondo.
            Application.Wait Now + TimeSerial(0, 0, 1)
            ' Se è presente l'elemento Input per l'inserimento della Pw allora.
            If eDriver.IsElementPresent(sBy.XPath("//input[@id='password']")) Then
                ' Pone nella variabile il valore richiesto con la InputBox.
                strPass = InputBox("Inserisci la tua Password", "xxxxxxxx")
                ' Lo preme.
                .FindElementByXPath("//input[@id='password']").Click
                ' Invia la variabile contenente la Pw.
                .FindElementByXPath("//input[@id='password']").SendKeys strPass
            End If
            ' Attende 1 secondo.
            Application.Wait Now + TimeSerial(0, 0, 1)
            ' Se è presente l'elemento bottone per inviare i due codici inseriti allora.
            If eDriver.IsElementPresent(sBy.XPath("//input[@type='submit']")) Then
                ' Lo preme.
                .FindElementByXPath("//input[@type='submit']").Click
            End If
        End With
    ' Altrimenti se il controllo dell'indirizzo tramite la Funzione EsisteURL dà risultato negativo (l'indirizzo non esiste o ci sono problemi), allora.
    ElseIf EsisteURL(strURLlogin) = False Then
        ' Avvisa l'utente ed esce dalla Sub.
        If MsgBox("L'indirizzo della Home page del sito Fineco non è stato raggiunto dalla richiesta", vbCritical + vbOKOnly, "A T T E N Z I O N E !") = vbOK Then
            ' Esce dalla Sub.
            Exit Sub
        End If
    End If
    ' Svuota le variabili pubbliche.
    strUtente = ""
    strPass = ""
' Fine della Sub.
End Sub
