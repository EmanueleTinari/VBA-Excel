
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                             +
'+ Funzione per aprire una finestra di dialogo e selezionare                   +
'+ un file il cui percorso commpleto è il risultato della funzione.            +
'+ Per funzionare necessita un riferimento a Microsoft Office 11.0 Object      +
'+ Library o versione successiva.                                              +
'+                                                                             +
'+ È possibile variare il filtro dei file, il titolo della finestra e altro.   +
'+                                                                             +
'+ Argomenti della funzione:                                                   +
'+                                                                             +
'+ = =                                                                         +
'+                                                                             +
'+ È necessaria la funzione PercorsoDesktop per recuperare il percorso         +
'+ del Desktop (può essere evitata, impostando la variabile in altro modo).    +
'+                                                                             +
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function PercorsoCompletoFile() As String

' Gestione errore.
On Error GoTo GesErr

' Richiede un riferimento a Microsoft Office 11.0 Object Library.
Dim fd As Office.FileDialog
' Contiene il percorso completo del Desktop.
Dim strPercorsoDesktop As String

    ' Recupera il percorso del desktop utilizzando la specifica funzione.
    strPercorsoDesktop = PercorsoDesktop
    ' Imposta la finestra di scelta file.
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        ' Svuota eventuali filtri presenti.
        .Filters.Clear
        ' Aggiunge i filtri seguenti, personalizzati.
        .Filters.Add "File json", "*.json", 1
        .Filters.Add "File csv", "*.csv", 2
        .Filters.Add "File txt", "*.txt", 2
        .Filters.Add "Tutti i file", "*.*", 3
        ' Imposta il titolo della finestra di dialogo.
        .Title = "Seleziona il file Json da importare"
        ' E' possibile selezionare solo un file alla volta.
        .AllowMultiSelect = False
        ' Imposta il percorso al desktop come cartella iniziale.
        .InitialFileName = "" & strPercorsoDesktop & "\"
        ' Ci si assicura che la visualizzazione dei file sia impostata su Dettagli.
        .InitialView = msoFileDialogViewDetails
        If .Show = True Then
            ' La funzione restituisce il file selezionato.
            PercorsoCompletoFile = .SelectedItems(1)
            ' Svuota le variabili.
            strPercorsoDesktop = Empty
            Set fd = Nothing
            ' Esce dalla funzione.
            Exit Function
        Else
            ' Se nessun file è stato selezionato, avvisa.
            MsgBox "Non hai selezionato nessun file." & Chr$(13) & "Esci.", vbCritical
            ' Svuota le variabili.
            strPercorsoDesktop = Empty
            Set fd = Nothing
            ' Esce dalla funzione.
            Exit Function
        End If
    End With

' Esce dalla funzione.
Uscita: Exit Function
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Function" & vbCrLf & "'PercorsoCompletoFile'" & vbCrLf & vbCrLf & Err.Description
        ' Svuota le variabili.
        strPercorsoDesktop = Empty
        Set fd = Nothing
        Resume Uscita
' Fine della funzione.
End Function
