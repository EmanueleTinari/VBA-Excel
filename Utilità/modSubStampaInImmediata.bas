Attribute VB_Name = "modSubStampaInImmediata"
Option Explicit
Option Private Module

' x testare la Funzione StampaInImmediata.
Public Sub Prova_StampaInImmediata()

' La Var conterrà una stringa da stampare.
Dim strMiaStringa As String

    ' Invio alla Sub StampaInImmediata una stringa vuota per testare l'errore.
    StampaInImmediata ("")

    ' Invio alla Sub StampaInImmediata la stringa contenuta tra parentesi e virgolette.
    StampaInImmediata ("Qui la stringa da stampare")

    ' Inserisce nella Var la stringa di testo.
    strMiaStringa = "Prova stampa"
    ' Invio alla Sub StampaInImmediata la stringa contenuta nella Var strMiaStringa.
    StampaInImmediata (strMiaStringa)

End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                                +
'+ Nome :                    Sub StampaInImmediata(ByVal strDaStampareInImmediata As String)      +
'+                                                                                                +
'+ Creata da :               Emanuele Tinari                                                      +
'+                                                                                                +
'+ In data :                 10/02/2023                                                           +
'+                                                                                                +
'+ Descrizione :             Sub per stampare nella Finestra Immediata la stringa passata dalla   +
'+                           strDaStampareInImmediata.                                            +
'+                                                                                                +
'+ Uso :                     Nel codice, quando è necessario stampare il valore di una Var        +
'+                           o qualsiasi altro dato stringa, passare il valore alla Sub.          +
'+                                                                                                +
'+ Valore restituito:        Nessuno                                                              +
'+                                                                                                +
'+ Esempio :                 StampaInImmediata (strMiaVariabileStringaDaTestare)                  +
'+                                                                                                +
'+                           ' x testare la Funzione StampaInImmediata.                           +
'+                           Sub ProvaStampaInImmediata()                                         +
'+                               StampaInImmediata ("Qui la stringa da stampare")                 +
'+                           End Sub                                                              +
'+                                                                                                +
'+ Valore di default :       Nessuno                                                              +
'+                                                                                                +
'+ Argomento(i) :            - ByVal strDaStampareInImmediata As String                           +
'+                             La Var stringa da inviare alla Finestra Immediata.                 +
'+                                                                                                +
'+ Riferimento(i):           Nessuno                                                              +
'+                                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub StampaInImmediata(ByVal strDaStampareInImmediata As String)

' Gestione errore.
On Error GoTo GesErr

    ' Se la stringa passata alla Sub come Argomento è Vuota o Nulla, allora.
    If strDaStampareInImmediata = Empty Or Null Then
        ' Invia alla Finestra Immediata un avviso.
        Debug.Print "La stringa passata come Argomento alla Var strDaStampareInImmediata è Vuota o Nulla."
        GoTo Uscita
    Else
        ' Invia alla Finestra Immediata la stringa.
        Debug.Print "" & strDaStampareInImmediata
        GoTo Uscita
    End If

' Esce dalla Sub, dopo aver svuotato la/e variabile/i.
Uscita: strDaStampareInImmediata = Empty
        Exit Sub
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Sub" & vbCrLf & _
        "'StampaInImmediata'" & vbCrLf & vbCrLf & _
        "Errore Numero: " & Err.Number & vbCrLf & _
        "Descrizione dell'errore:" & vbCrLf & _
        Err.Description, vbCritical, "C'è stato un errore!"
        Resume Uscita
' Fine della Sub.
End Sub
