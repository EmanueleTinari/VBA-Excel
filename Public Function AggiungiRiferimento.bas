Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                              +
'+ Funzione che aggiunge al progetto di Excel indicato nella variabile wbk      +
'+ un riferimento alla libreria la cui GUID è nella variabile strGUID.          +
'+                                                                              +
'+ Argomenti della funzione:                                                    +
'+                                                                              +
'+ wbk                  -   Questo file o un altro a cui si vuole aggiungere.   +
'+ strGUID              -   La GUID del riferimento da aggiungere.              +
'+                                                                              +
'+ È necessaria la Funzione EsisteRiferimento(wbk, strGUID) per testare se      +
'+ un foglio con un dato nome esiste, per evitare errori in caso non esistesse. +
'+                                                                              +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function AggiungiRiferimento(ByVal wbk As Workbook, ByVal strGUID As String)

' Gestione errore.
On Error GoTo GesErr

    ' Se il controllo tramite la funzione EsisteRiferimento è FALSO (il riferiemento non esiste nel progetto), allora.
    If EsisteRiferimento(wbk, strGUID) = False Then
        ' Aggiunge il riferimento al file tramite la GUID passata. ", 0, 0" Seleziona l'ultima versione installata sul computer.
        wbk.VBProject.References.AddFromGuid strGUID, 0, 0
    End If

' Esce dalla funzione.
Uscita: Exit Function
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Sub" & vbCrLf & "'AggiungiRiferimento'" & vbCrLf & vbCrLf & Err.Description
        ' Svuota le variabili.
        Set wbk = Nothing
        strGUID = Empty
        Resume Uscita
' Fine della funzione.
End Function
