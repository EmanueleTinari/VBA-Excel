Attribute VB_Name = "modFunSelezionaCartella"
Option Explicit
Option Private Module

' x testare la Funzione SelezionaCartella.
Sub Prova_SelezionaCartella()

' La Var conterrà il percorso scelto come Cartella iniziale.
Dim strPercorso                 As String

    ' Immetto il percorso scelto.
    strPercorso = "C:"
    ' Senza Argomento opzionale strPercorso.
    MsgBox "Hai selezionato: " & SelezionaCartella
    ' Con Argomento opzionale strPercorso.
    MsgBox "Hai selezionato: " & SelezionaCartella(strPercorso)

End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                                +
'+ Nome :                    Function SelezionaCartella _                                         +
'+                                             (Optional ByVal strPercorso As String) As String   +
'+                                                                                                +
'+ Creata da :               Emanuele Tinari                                                      +
'+                                                                                                +
'+ In data :                 14/03/2023                                                           +
'+                                                                                                +
'+ Descrizione :             Funzione che apre una finestra di dialogo per selezionare una sola   +
'+                           Cartella (non è quindi una multiselect).                             +
'+                                                                                                +
'+ Uso :                     Nel codice quando è necessario selezionare una Cartella, usare la    +
'+                           Funzione.                                                            +
'+                                                                                                +
'+ Valore restituito:        SelezionaCartella conterrà la Cartella selezionata, completa di      +
'+                           percorso.                                                            +
'+                                                                                                +
'+ Esempio :                 ' x testare la Funzione SelezionaCartella.                           +
'+                           Sub Prova_SelezionaCartella()                                        +
'+                           ' La Var conterrà il percorso scelto come Cartella iniziale.         +
'+                           Dim strPercorso As String                                            +
'+                               ' Immetto il percorso scelto.                                    +
'+                               strPercorso = "C:"                                               +
'+                               ' Senza Argomento opzionale strPercorso.                         +
'+                               MsgBox "Hai selezionato: " & SelezionaCartella                   +
'+                               ' con Argomento opzionale strPercorso.                           +
'+                               MsgBox "Hai selezionato: " & SelezionaCartella(strPercorso)      +
'+                           End Sub                                                              +
'+                                                                                                +
'+ Valore di default :       - Se l'Argomento ByVal strPercorso non viene passato, viene assunta  +
'+                             come Cartella iniziale la Cartella Documenti                       +
'+                                                                                                +
'+ Argomento(i) :            - Optional ByVal strPercorso As String                               +
'+                             Facoltativo. Il percorso da cui si vuole iniziare la selezione.    +
'+                                                                                                +
'+ Riferimento(i):           - Riferimento a Microsoft Office 16.0 Object Library                 +
'+                             Lib.: in "C:\Program Files (x86)\Common Files\Microsoft Shared\ _  +
'+                                       OFFICE16\MSO.dll"                                        +
'+                             GUID: "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}"                     +
'+                                                                                                +
'+                           - Riferimento a Windows Script Host Object Model                     +
'+                             Lib.: in "C:\Windows\SysWOW64\wshom.ocx"                           +
'+                             GUID: "{F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}"                     +
'+                                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function SelezionaCartella(Optional ByVal strPercorso As String) As String

' Gestione errore.
On Error GoTo GesErr

' Richiede il Riferimento a Microsoft Office Object Library (MSO.dll).
Dim OFd                         As Office.FileDialog
' Early Biding Windows Script, richiede il Riferimento a Windows Script Host Object Model (wshom.ocx).
Dim objFoglioHShell             As New WshShell
' La variabile conterrà il percorso alla Cartella Documenti.
Dim strPercorsoDocs             As String

    ' Imposta la finestra di scelta Cartella.
    Set OFd = Application.FileDialog(msoFileDialogFolderPicker)
    With OFd
        ' Imposta il titolo della finestra di dialogo.
        .Title = "Seleziona una Cartella"
        ' Ci si assicura che la visualizzazione dei file sia impostata su Dettagli.
        .InitialView = msoFileDialogViewDetails
        ' E' possibile selezionare solo una Cartella alla volta.
        .AllowMultiSelect = False
        ' Se la Var strPercorso passata come Argomento è vuota, allora.
        If strPercorso = "" Then
            ' Recupera il percorso alla Cartella Documenti.
            strPercorsoDocs = objFoglioHShell.SpecialFolders("MyDocuments")
            ' Imposta il percorso alla Cartella Documenti come Cartella iniziale.
            .InitialFileName = strPercorsoDocs & Application.PathSeparator
        Else
            ' Imposta il percorso fornito come Cartella iniziale.
            .InitialFileName = strPercorso & Application.PathSeparator
        End If
        If .Show = True Then
            ' La Funzione restituisce la Cartella selezionata.
            SelezionaCartella = .SelectedItems(1)
            ' Esce.
            GoTo Uscita
        Else
            ' Se nessuna Cartella è stata selezionata, avvisa.
            MsgBox "Non hai selezionato nessuna Cartella." & Chr$(13) & "Esci.", vbCritical
            ' Esce.
            GoTo Uscita
        End If
    End With

' Esce dalla Funzione, dopo aver svuotato la/e variabile/i.
Uscita: strPercorso = Empty
        Set OFd = Nothing
        Set objFoglioHShell = Nothing
        strPercorsoDocs = Empty
        Exit Function
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Function" & vbCrLf & _
        "'SelezionaCartella'" & vbCrLf & vbCrLf & _
        "Errore Numero: " & Err.Number & vbCrLf & _
        "Descrizione dell'errore:" & vbCrLf & _
        Err.Description, vbCritical, "C'è stato un errore!"
        Resume Uscita
' Fine della Funzione.
End Function
