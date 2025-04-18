Attribute VB_Name = "modFunMexBox"
Option Explicit
Option Private Module

' x testare la Funzione MexBox.
Public Sub Prova_MexBox()

Dim intRisposta                 As Integer
Dim strMessaggio                As String
Dim intSecondi                  As Integer
Dim strTitolo                   As String
    
    strMessaggio = "Testo messaggio"
    intSecondi = 3
    strTitolo = "Titolo MessageBox Temporizzata"
    intRisposta = MexBox(strMessaggio, intSecondi, strTitolo, 0 + 48)
    ' Stampa nella Finestra Immediata la risposta data dall'Utente.
    Debug.Print "L'Utente ha premuto " & intRisposta

End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                                +
'+ Nome :                    Function MexBox _                                                    +
'+                                   (ByVal strTestoMsg As String, _                              +
'+                                    Optional ByVal intTempoInSecondi As Integer, _              +
'+                                    Optional ByVal strTitoloMexBox As String = "WScript", _     +
'+                                    Optional ByVal intButtons) As Integer                       +
'+                                                                                                +
'+ Creata da :               Emanuele Tinari                                                      +
'+                                                                                                +
'+ In data :                 10/02/2023                                                           +
'+                                                                                                +
'+ Descrizione :             Funzione che crea una MessageBox WScript con timer che si chiude     +
'+                           in un tempo indicato in secondi nella Var intTempoInSecondi.         +
'+                                                                                                +
'+ Uso :                     Nel codice, quando è necessario avere una MessageBox a chiusura      +
'+                           temporizzata.                                                        +
'+                                                                                                +
'+                           Tipi di bottoni:                                                     +
'+                           0  (Bottone OK)                                                      +
'+                           1  (Bottoni OK e Cancella)                                           +
'+                           2  (Bottoni Interrompi, Riprova e Annulla)                           +
'+                           3  (Bottoni Si, No e Cancella)                                       +
'+                           4  (Bottoni Si e No)                                                 +
'+                           5  (Bottoni Riprova e Cancella)                                      +
'+                                                                                                +
'+                           Tipi di icone:                                                       +
'+                           16 (mostra icona X bianca in cerchio rosso)                          +
'+                           32 (mostra icona ? bianco in cerchio blu)                            +
'+                           48 (mostra icona ! nero in triangolo giallo)                         +
'+                           64 (mostra icona ! bianco in cerchio blu)                            +
'+                                                                                                +
'+ Valore restituito:        -1 : L'Utente non ha premuto nulla e la MessageBox si è chiusa       +
'+                            1 : L'Utente ha premuto Ok                                          +
'+                            2 : L'Utente ha premuto Cancella                                    +
'+                            3 : L'Utente ha premuto Interrompi                                  +
'+                            4 : L'Utente ha premuto Riprova                                     +
'+                            5 : L'Utente ha premuto Annulla                                     +
'+                            6 : L'Utente ha premuto Si                                          +
'+                            7 : L'Utente ha premuto No                                          +
'+                                                                                                +
'+ Esempio :                 ' x testare la Funzione MexBox.                                      +
'+                           Sub ProvaMexBox()                                                    +
'+                           Dim intRisposta As Integer                                           +
'+                               intRisposta = MexBox("Messaggio", 3, "Titolo", 0 + 48)           +
'+                               Debug.Print "L'Utente ha premuto " & intRisposta                 +
'+                           End Sub                                                              +
'+                                                                                                +
'+ Valore di default :       - Se l'Argomento strTitoloMexBox non viene passato, si assume        +
'+                             "WScript"                                                          +
'+                                                                                                +
'+ Argomento(i) :            - ByVal strTestoMsg As String                                        +
'+                             Il testo da scrivere nella MessageBox temporizzata.                +
'+                                                                                                +
'+                           - Optional ByVal intTempoInSecondi As Integer                        +
'+                             Il tempo, espresso in secondi, per il quale la MessageBox rimarrà  +
'+                             visibile, dopo il quale si richiuderà, restituendo -1.             +
'+                                                                                                +
'+                           - Optional strTitoloMexBox As String = "WScript"                     +
'+                             Il titolo da dare alla MessageBox temporizzata.                    +
'+                                                                                                +
'+                           - Optional ByVal intButtons = vbDefaultButton1                       +
'+                             L'eventuale icona e i(l) bottoni(e) che si vorranno avere nella    +
'+                             MessageBox Temporizzata, da ognuno dei quali è possibile ottenere  +
'+                             un valore diverso. Vedi più sopra il Valore restituito.            +
'+                                                                                                +
'+ Riferimento(i):           - Riferimento a Windows Script Host Object Model                     +
'+                             Lib.: in "C:\Windows\SysWOW64\WSHom.ocx"                           +
'+                             GUID: "{F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}"                     +
'+                                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function MexBox _
                (ByVal strTestoMsg As String, _
                Optional ByVal intTempoInSecondi As Integer, _
                Optional ByVal strTitoloMexBox As String = "WScript", _
                Optional ByVal intButtons As Integer) As Integer

' Gestione errore.
On Error GoTo GesErr

' Early Biding Windows Script, richiede il Riferimento a Windows Script Host Object Model (WSHom.ocx).
Dim objFoglioHShell                 As New WshShell
    
    ' La Funzione MexBox è una MessageBox dell'oggetto WScript personalizzabile.
    MexBox = objFoglioHShell.PopUp(strTestoMsg, intTempoInSecondi, strTitoloMexBox, intButtons)

' Esce dalla Funzione, dopo aver svuotato la/e variabile/i.
Uscita: Set objFoglioHShell = Nothing
        strTestoMsg = Empty
        intTempoInSecondi = Empty
        strTitoloMexBox = Empty
        intButtons = Empty
        Exit Function
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Function" & vbCrLf & "'MexBox'" & vbCrLf & vbCrLf & "Errore Numero: " & Err.Number & vbCrLf & "Descrizione dell'errore:" & vbCrLf & Err.Description, vbCritical, "C'è stato un errore!"
        Resume Uscita
' Fine della Funzione.
End Function
