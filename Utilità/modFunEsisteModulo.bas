Attribute VB_Name = "modFunEsisteModulo"
Option Explicit
Option Private Module

' x testare la Function EsisteModulo.
Sub Prova_EsisteModulo()

Dim strNomeModulo               As String
    
    strNomeModulo = "xxx"
    If EsisteModulo(strNomeModulo) = True Then
        Debug.Print "Il Modulo " & strNomeModulo & " esiste."
    Else
        Debug.Print "Il Modulo " & strNomeModulo & " non esiste."
    End If

End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                                +
'+ Nome :                    Function EsisteModulo (ByVal strNomeModulo As String, _              +
'+                                                  Optional ByVal wbk As Workbook) As Boolean    +
'+                                                                                                +
'+ Creata da :               Emanuele Tinari                                                      +
'+                                                                                                +
'+ In data :                 25/03/2023                                                           +
'+                                                                                                +
'+ Descrizione :             Funzione per controllare se esiste nel Progetto VBA un Modulo col    +
'+                           nome passato dalla stringa strNomeModulo.                            +
'+                                                                                                +
'+ Uso :                     Nel codice, quando è necessario accertarsi se un certo Modulo esiste +
'+                           nel Progetto VBA del Foglio wbk, per non incorrere in errori.        +
'+                                                                                                +
'+ Valore restituito:        True: Il Modulo esiste.                                              +
'+                           False: Il Modulo non esiste.                                         +
'+                                                                                                +
'+ Esempio :                 Con un If...Then...End If è possibile utilizzare la Funzione.        +
'+                                                                                                +
'+ Valore di default :       - Se l'Argomento wbk non viene passato, viene assunto ThisWorkbook.  +
'+                                                                                                +
'+                           - EsisteModulo = False                                               +
'+                                                                                                +
'+ Argomento(i) :            - ByVal strNomeModulo As String                                      +
'+                             Il nome del Modulo che si vuole testare se esiste.                 +
'+                                                                                                +
'+                           - Optional ByVal wbk As Workbook                                     +
'+                             Facoltativo. Questo file o un altro in cui si vuole controllare.   +
'+                                                                                                +
'+ Riferimento(i):           - Microsoft Visual Basic for Applications Extensibility              +
'+                             Lib.: in "C:\Program Files (x86)\Common Files\ _                   +
'+                                       Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"                   +
'+                             GUID: "{0002E157-0000-0000-C000-000000000046}"                     +
'+                                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function EsisteModulo(ByVal strNomeModulo As String, Optional ByVal wbk As Workbook) As Boolean

' Gestione errore.
On Error GoTo GesErr

' Richiede il Riferimento a Microsoft Visual Basic for Applications Extensibility (VBE6EXT.OLB).
Dim objVBComp                   As vbComponent

    ' Inizialmente viene impostato su Falso il risultato della Funzione (il Modulo NON esiste).
    EsisteModulo = False
    
    ' Se l'Argomento wbk passato alla Function è vuoto, allora.
    If wbk Is Nothing Or Null Then
        ' Imposta nella Var wbk il ThisWorkbook.
        Set wbk = ThisWorkbook
    ' Altrimenti.
    Else
        ' Imposto nella Var wbk il WorkBook passato come Argomento alla Function.
        Set wbk = wbk
    End If
    
    ' Ciclo tra tutti i Componenti del Progetto.
    For Each objVBComp In wbk.VBProject.VBComponents
        ' Se il nome del Componente in esame è uguale all'Argomento strNomeModulo passato, allora.
        If objVBComp.Name = strNomeModulo Then
            ' La Funzione restituisce Vero.
            EsisteModulo = True
        End If
    ' Prossimo Componente.
    Next objVBComp

' Esce dalla Funzione, dopo aver svuotato la/e variabile/i.
Uscita: Set wbk = Nothing
        strNomeModulo = Empty
        Set objVBComp = Nothing
        Exit Function
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore nella Function" & vbCrLf & _
        "'EsisteModulo'" & vbCrLf & vbCrLf & _
        "Errore Numero: " & Err.Number & vbCrLf & _
        "Descrizione dell'errore:" & vbCrLf & _
        Err.Description, vbCritical, "C'è stato un errore!"
        Resume Uscita
' Fine della Funzione.
End Function
