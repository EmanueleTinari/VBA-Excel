Attribute VB_Name = "modSubSvuotaVarData"
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                                +
'+ Nome :                    SvuotaVarData()                                                      +
'+                                                                                                +
'+ Creata da :               Emanuele Tinari                                                      +
'+                                                                                                +
'+ In data :                 30/10/2023                                                           +
'+                                                                                                +
'+ Descrizione :             Sub che crea e riempie due variabili di tipo Date e poi le svuota.   +
'+                           Creata per chi chiede come portare a Null o Nothing una variabile    +
'+                           di questo tipo e non ha trovato una risposta corretta a tale         +
'+                           quesito.                                                             +
'+                                                                                                +
'+ Uso :                     Per valutare il comportamento delle variabili, eseguire la Sub con   +
'+                           F8 e guardare nella finestra Variabili locali del VBE di Excel o     +
'+                           altro editor, il loro riempirsi e successivo svuotamento.            +
'+                                                                                                +
'+ Valore restituito:        Nessuno                                                              +
'+                                                                                                +
'+ Esempio :                 Nessuno                                                              +
'+                                                                                                +
'+ Valore di default :       Nessuno                                                              +
'+                                                                                                +
'+ Argomento(i) :            Nessuno                                                              +
'+                                                                                                +
'+ Riferimento(i):           Nessuno                                                              +
'+                                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Sub SvuotaVarData()

' Gestione errore.
On Error GoTo GesErr

' Prima Var di tipo Date.
Dim dtData As Date
' Seconda Var di tipo Date.
Dim dtOra As Date
    
    ' La riempie con una data.
    dtData = CDate("30/10/2023")
    ' La riempie con un orario.
    dtOra = CDate("08:20:32")

' Esce dalla Sub, dopo aver svuotato la/e variabile/i.
Uscita: dtData = Empty
        dtOra = Empty
        Exit Sub
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Sub" & vbCrLf & _
        "'SvuotaVarData'" & vbCrLf & vbCrLf & _
        "Errore Numero: " & Err.Number & vbCrLf & _
        "Descrizione dell'errore:" & vbCrLf & _
        Err.Description, vbCritical, "C'è stato un errore!"
        Resume Uscita
' Fine della Sub.
End Sub
