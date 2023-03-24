Attribute VB_Name = "modFunScriviFileTemp"
Option Explicit
Option Private Module
' x testare la Funzione ScriviFileTemp.
Public Sub Prova_ScriviFileTemp()
    
Dim strFileCreato As String
    ' Esempio in cui viene indicato solo la stringa di testo da inserire.
    ' Nella Var strFileCreato è contenuto il percorso completo e il nome del file creato.
    strFileCreato = ScriviFileTemp("Ciao")

End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                                +
'+ Nome :                    Function ScriviFileTemp(strTesto As String, _                        +
'+                                                   Optional strPercorso As String, _            +
'+                                                   Optional strNomeFile As String, _            +
'+                                                   Optional strEstensione As String = "txt") _  +
'+                                                   As String                                    +
'+                                                                                                +
'+ Creata da :               Emanuele Tinari                                                      +
'+                                                                                                +
'+ In data :                 11/02/2023                                                           +
'+                                                                                                +
'+ Descrizione :             Funzione che crea in un dato percorso (o nella cartella temporanea)  +
'+                           un file con un nome passato (oppure viene creato basandosi sulla     +
'+                           data ed ora di esecuzione) con l'estensione passata (oppure viene    +
'+                           usata quella di default (txt), contenente la stringa strTesto.       +
'+                                                                                                +
'+ Uso :                     Nel codice, quando è necessario creare un file contenente una data   +
'+                           stringa, passare il valore alla Funzione.                            +
'+                                                                                                +
'+ Valore restituito:        Il percorso completo e il nome del file.                             +
'+                                                                                                +
'+ Esempio :                 ' x testare la Funzione ScriviFileTemp.                              +
'+                           Sub ProvaScriviFileTemp()                                            +
'+                           Dim strFileCreato As String                                          +
'+                               strFileCreato = ScriviFileTemp ("Ciao")                          +
'+                           End Sub                                                              +
'+                                                                                                +
'+ Valore di default :       - Se l'Argomento strEstensione non viene passato, si assume txt.     +
'+                                                                                                +
'+ Argomento(i) :            - ByVal strTesto As String                                           +
'+                             La stringa di testo da scrivere nel file di testo.                 +
'+                                                                                                +
'+                           - Optional strPercorso As String                                     +
'+                             Si può fornire un proprio percorso per la cartella di salvataggio. +
'+                                                                                                +
'+                           - Optional ByVal strNomeFile As String                               +
'+                             Si può fornire un proprio nome per il file creato.                 +
'+                                                                                                +
'+                           - Optional strEstensione As String = "txt"                           +
'+                             Si può fornire un'estensione per il file creato, altrimenti txt.   +
'+                                                                                                +
'+ Riferimento(i):           Nessuno                                                              +
'+                                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function ScriviFileTemp(ByVal strTesto As String, _
                               Optional ByVal strPercorso As String, _
                               Optional ByVal strNomeFile As String, _
                               Optional strEstensione As String = "txt") _
                               As String

' Gestione errore.
On Error GoTo GesErr

' La Var conterrà il percorso e il nome del file.
Dim strPercorsoNomeFile As String
' La Var conterrà il numero del file che stiamo andando a creare.
Dim intNumFile As Integer
    
    ' Se la Var passata alla Funzione, contenente il nome del file, è vuota, allora.
    If strNomeFile = "" Then
        ' Crea il nome del file. L'estensione se non è passata dalla Var, viene usata quella di default.
        strNomeFile = Format(Date, "ddmmmyyyy") & "_" & Format(Time, "hhmmss") & "." & strEstensione
    End If
    ' Se la Var passata alla Funzione, contenente il percorso del file, è vuota, allora.
    If strPercorso = "" Then
        ' Crea il percorso alla cartella temporanea.
        strPercorso = Environ("TMP") & Application.PathSeparator
    End If
    ' Poi concatena le due stringe per ottenere il file.
    strPercorsoNomeFile = strPercorso & strNomeFile
    
    ' Il numero del file temporareo è il prossimo numero disponibile.
    intNumFile = FreeFile()
    Open strPercorsoNomeFile For Output As intNumFile
    Print #intNumFile, strTesto;
    Close #intNumFile
    ' Apre il file creato con Notepad massimizzato.
    Shell "Notepad.exe " & strPercorsoNomeFile, vbMaximizedFocus
    ' La Funzione restituisce il percorso e il nome del file creato.
    ScriviFileTemp = strPercorsoNomeFile

' Esce dalla Funzione, dopo aver svuotato la/e variabile/i.
Uscita: strTesto = Empty
        strPercorso = Empty
        strNomeFile = Empty
        strEstensione = Empty
        strPercorsoNomeFile = Empty
        intNumFile = Empty
        Exit Function
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Function" & vbCrLf & "'ScriviFileTemp'" & vbCrLf & vbCrLf & "Errore Numero: " & Err.Number & vbCrLf & "Descrizione dell'errore:" & vbCrLf & Err.Description, vbCritical, "C'è stato un errore!"
        Resume Uscita
' Fine della Funzione.
End Function
