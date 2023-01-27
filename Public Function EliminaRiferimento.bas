Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                              +
'+ Funzione che elimina dal progetto di Excel indicato nella variabile wbk      +
'+ il riferimento alla libreria la cui GUID è nella variabile strGUID.          +
'+                                                                              +
'+ Argomenti della funzione:                                                    +
'+                                                                              +
'+ wbk                  -   Questo file o un altro a cui si vuole eliminare.    +
'+ strGUID              -   La GUID del riferimento da aggiungere.              +
'+                                                                              +
'+ È necessaria la Funzione EsisteRiferimento(wbk, strGUID) per testare se      +
'+ un foglio con un dato nome esiste, per evitare errori in caso non esistesse. +
'+                                                                              +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function EliminaRiferimento(wbk As Workbook, strGUID As String)

' Gestione errore.
On Error GoTo GesErr

    ' Se il riferimento è presente nel progetto, allora.
    If EsisteRiferimento(wbk, strGUID) = True Then
        ' Rimuove lo stesso dal progetto.
        wbk.VBProject.References.Remove
    End If

' Esce dalla sub.
Uscita: Exit Function
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Sub" & vbCrLf & "'EliminaRiferimento'" & vbCrLf & vbCrLf & Err.Description
        ' Svuota le variabili.
        Set wbk = Nothing
        strGUID = Empty
        Resume Uscita
' Fine della funzione.
End Function
