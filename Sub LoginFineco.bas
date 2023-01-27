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
    '+                                                                +
    '+ È necessaria la Funzione EsisteURL(ByVal strTestURL As String) +
    '+ per testare il collegamento prima del suo uso.                 +
    '+                                                                +
    '+ Uso: nel codice, quando serve, inserire la chiamata alla sub:  +
    '+                                                                +
    '+ Call LoginFineco                                               +
    '+                                                                +
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Sub LoginFineco()
    ' Controllo degli errori
    On Error Resume Next
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
                strUtente = InputBox("Inserisci il tuo Codice Utente", "User", "<---Inserisci tra queste due virgolette, eliminando tutta questa scritta, il tuo codice utente se vuoi che ti venga proposto in automatico !--->")
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
                strPass = InputBox("Inserisci la tua Password", "Password", "<---Qui la tua PW, se vuoi esca in automatico nella InputBox !--->")
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
    Else
        If MsgBox("L'indirizzo della Home page del sito Fineco non è stato raggiunto dalla richiesta", vbCritical + vbOKOnly, "A T T E N Z I O N E !") = vbOK Then
            Exit Sub
        End If
    End If
    ' Svuota le variabili pubbliche.
    strUtente = ""
    strPass = ""
End Sub
