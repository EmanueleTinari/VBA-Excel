Attribute VB_Name = "modSubVerificaCartella"
Option Explicit
Option Private Module

' x testare la Sub VerificaCartella.OK.
Sub Prova_VerificaCartella()
    
    VerificaCartella ("C:\Cartella_001\Cartella_002\Cartella_003")

End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                                +
'+ Nome :                    Sub VerificaCartella(ByVal strPercorsoCartella As String)            +
'+                                                                                                +
'+ Creata da :               Emanuele Tinari                                                      +
'+                                                                                                +
'+ In data :                 06/03/2023                                                           +
'+                                                                                                +
'+ Descrizione :             Sub che verifica se esiste una Cartella nel File System col percorso +
'+                           indicato dalla stringa strPercorsoCartella. Se la Cartella non       +
'+                           esiste, la crea; se è una sotto-cartella e il percorso non esiste,    +
'+                           crea anche quello.                                                   +
'+                                                                                                +
'+ Uso :                     Seguire l'esempio. Si può indicare o meno lo Slash finale nel        +
'+                           percorso inserito in strPercorsoCartella.                            +
'+                                                                                                +
'+ Valore restituito:        Nessuno                                                              +
'+                                                                                                +
'+ Esempio :                 VerificaCartella ("C:\Alfa\Beta\Gamma")                              +
'+                            oppure                                                              +
'+                           VerificaCartella ("C:\Alfa\Beta\Gamma\")                             +
'+                                                                                                +
'+ Valore di default :       Nessuno                                                              +
'+                                                                                                +
'+ Argomento(i) :            ByVal strPercorsoCartella As String                                  +
'+                                                                                                +
'+ Riferimento(i):           - Microsoft Scripting Runtime Library                                +
'+                             Lib.: in "C:Windows\SysWOW64\ScrRun.dll"                           +
'+                             GUID: "{420B2830-E718-11CF-893D-00A0C9054228}"                     +
'+                                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub VerificaCartella(ByVal strPercorsoCartella As String)

' Gestione errore.
On Error GoTo GesErr

' Dichiarazione dell'oggetto FileSystem.
Dim objFSO                      As New FileSystemObject

    ' Se il percorso non inizia con una lettera seguita da ":" e "\" oppure se comincia con "\\", allora.
    If Not (strPercorsoCartella Like "?:\*" Or strPercorsoCartella Like "\\*") Then
        ' Avvisa che quello fornito non è un percorso valido.
        MsgBox (strPercorsoCartella & "non è un percorso valido.")
    End If
    ' Se non esiste il percorso fornito e la stringa strPercorsoCartella non è vuota, allora.
    If Not objFSO.FolderExists(strPercorsoCartella) And strPercorsoCartella <> "" Then
        ' Crea la cartella o la sotto-cartella nel percorso fornito, comprensivo dello "\" finale.
        VerificaCartella Left(strPercorsoCartella, InStrRev(strPercorsoCartella, "\", Len(strPercorsoCartella) - 1))
        ' Crea il percorso alla Cartella e la Cartella stessa.
        objFSO.CreateFolder strPercorsoCartella
    End If

' Esce dalla Sub, dopo aver svuotato la/e variabile/i.
Uscita: Set objFSO = Nothing
        strPercorsoCartella = Empty
        Exit Sub
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Sub" & vbCrLf & "'VerificaCartella'" & vbCrLf & vbCrLf & "Errore Numero: " & Err.Number & vbCrLf & "Descrizione dell'errore:" & vbCrLf & Err.Description, vbCritical, "C'è stato un errore!"
        Resume Uscita
' Fine della Sub.
End Sub

