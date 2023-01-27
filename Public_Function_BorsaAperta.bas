Option Explicit

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                 +
'+ Funzione per valutare se in una certa data la Borsa Italiana è aperta o chiusa. +
'+ Restituisce Vero o Falso.                                                       +
'+                                                                                 +
'+ Argomenti della funzione:                                                       +
'+                                                                                 +
'+ dtData       -   La data che vogliamo valutare.                                 +
'+                                                                                 +
'+ È assolutamente necessario che nel file di Excel vi sia un foglio chiamato      +
'+ "Calendario di borsa" contenente nella colonna C tutte le date del calendario   +
'+ e nella colonna D la parola "Aperto" se in quella data la borsa è stata, o      +
'+ sarà, aperta, oppure la parola "Chiuso" se è stata, o sarà, chiusa.             +
'+                                                                                 +
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function BorsaAperta(ByVal dtData As Date) As Boolean

' Gestione errore.
On Error GoTo GesErr

' Variabile per il ciclo del valore da ricercare.
Dim C As Range
    ' Inizialmente viene impostato su False il risultato della Funzione.
    BorsaAperta = False
    ' Seleziona il Foglio Calendario di borsa.
    Worksheets("Calendario di borsa").Select
    ' Cerca la data che si vuole ricercare nella Colonna C del foglio Calendario di Borsa.
    With Range("C1:" & Range("C1").End(xlDown).Address(False, False) & "")
        ' Cerca la data in esame nel foglio Calendario di borsa.
        Set C = .Find(dtData, LookIn:=xlValues)
        ' Se trova il valore.
        If Not C Is Nothing Then
            ' Seleziona la cella trovata.
            Range("" & C.Address(False, False) & "").Select
            ' Se il giorno in esame la borsa è Aperta e la data in esame è inferiore alla data odierna allora.
            If Range("" & C.Address(False, False) & "").Offset(0, 1).Text = "Aperto" Then
                ' La borsa è aperta.
                BorsaAperta = True
                ' Svuota le variabili.
                Set C = Nothing
                dtData = Empty
                ' Esce dalla funzione.
                Exit Function
            ' La borsa è chiusa per festività o fine settimana.
            ElseIf Range("" & C.Address(False, False) & "").Offset(0, 1).Text = "Chiuso" Then
                ' La borsa è chiusa.
                BorsaAperta = False
                ' Svuota le variabili.
                Set C = Nothing
                dtData = Empty
                ' Esce dalla funzione.
                Exit Function
            End If
        End If
    End With

' Esce dalla funzione.
Uscita: Exit Function
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Function" & vbCrLf & "'BorsaAperta'" & vbCrLf & vbCrLf & Err.Description
        ' Svuota le variabili.
        Set C = Nothing
        dtData = Empty
        Resume Uscita
' Fine della funzione.
End Function
