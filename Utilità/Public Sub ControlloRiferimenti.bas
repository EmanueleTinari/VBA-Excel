
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                      +
'+ Routine che controlla i riferimenti nel progetto VBA +
'+ eliminando quelli non necessari ed aggiungendo       +
'+ quelli mancanti.                                     +
'+                                                      +
'+ In caso si debba aggiungere un riferimento,          +
'+ è obbligatorio variare nella sezione Dichiarazioni   +
'+ il PRIMO valore nell'Array strArrayRiferimenti ed    +
'+ aggiungere alla fine il bloccho di codice relativo   +
'+ al nuovo riferimento inserito.                       +
'+                                                      +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub ControlloRiferimenti()

' Gestione errore.
On Error GoTo GesErr
        
Dim strArrayRiferimenti(14, 0) As String        ' Array dei riferimenti. Variare il numero se i riferimenti variano ATTENZIONE, parte da 0
Dim intCiclo1 As Integer                        ' Variabile per il primo ciclo.
Dim objRiferimento As Object                    ' Oggetto che punta al riferimento.
Dim wbk As Workbook                             ' Riferimento al Workbook.
Dim intCiclo2 As Integer                        ' Variabile per il secondo ciclo.
    
    ' Il primo ciclo rimuove ogni riferimento mancante.
    For intCiclo1 = ThisWorkbook.VBProject.References.Count To 1 Step -1
        Set objRiferimento = ThisWorkbook.VBProject.References.Item(intCiclo1)
        If objRiferimento.isbroken = True Then
            ThisWorkbook.VBProject.References.Remove objRiferimento
        End If
    Next intCiclo1
    
'**************************************************************************
'* Attenzione! Aumentando o diminuendo i riferimenti, aumentano o         *
'* diminuiscono i blocchi di codice che riempiono l'array                 *
'**************************************************************************
    
    ' Riferimento 0 Visual Basic For Applications, VBA
    strArrayRiferimenti(0, 0) = "{000204EF-0000-0000-C000-000000000046}"
    
    ' Riferimento 1 Microsoft Excel 16.0 Object Library, Excel
    strArrayRiferimenti(1, 0) = "{00020813-0000-0000-C000-000000000046}"
    
    ' Riferimento 2 OLE Automation, stdole
    strArrayRiferimenti(2, 0) = "{00020430-0000-0000-C000-000000000046}"

    ' Riferimento 3 Microsoft Office 16.0 Object Library, Office
    strArrayRiferimenti(3, 0) = "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}"

    ' Riferimento 4 Microsoft Forms 2.0 Object Library, MSForms
    strArrayRiferimenti(4, 0) = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}"

    ' Riferimento 5 Microsoft XML v6.0, MSXML2
    strArrayRiferimenti(5, 0) = "{F5078F18-C551-11D3-89B9-0000F81FE221}"

    ' Riferimento 6 Microsoft HTML Object Library, MSHTML
    strArrayRiferimenti(6, 0) = "{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}"

    ' Riferimento 7 Microsoft Internet Controls, SHDocVw
    strArrayRiferimenti(7, 0) = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}"

    ' Riferimento 8 Microsoft Scripting Runtime, Scripting
    strArrayRiferimenti(8, 0) = "{420B2830-E718-11CF-893D-00A0C9054228}"

    ' Riferimento 9 Microsoft Script Control 1.0, MSScriptControl
    strArrayRiferimenti(9, 0) = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}"

    ' Riferimento 10 Microsoft ActiveX Data Objects 6.1 Library, ADODB
    strArrayRiferimenti(10, 0) = "{B691E011-1797-432E-907A-4D8C69339129}"

    ' Riferimento 11 Microsoft Windows Common Controls 6.0 (SP6), MSComctlLib
    strArrayRiferimenti(11, 0) = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}"

    ' Riferimento 12 Selenium Type Library, Selenium
    strArrayRiferimenti(12, 0) = "{0277FC34-FD1B-4616-BB19-A9AABCAF2A70}"

    ' Riferimento 13 Microsoft Visual Basic for Applications Extensibility 5.3, VBIDE
    strArrayRiferimenti(13, 0) = "{0002E157-0000-0000-C000-000000000046}"

    ' Riferimento 14 Microsoft WMI Scripting V1.2 Library, WbemScripting
    strArrayRiferimenti(14, 0) = "{565783C6-CB41-11D1-8B02-00600806D9B6}"
    
    ' Imposta il riferimento a questo Workbook.
    Set wbk = ThisWorkbook
    With Application.ThisWorkbook.VBProject.References
        ' ATTENZIONE ! Variando il numero dei riferimenti, in più o in meno, variare il numero di cicli For intCiclo1... qui sotto.
        For intCiclo1 = 0 To 13     '<----- Partire da 0 a contare i riferimenti!
            For intCiclo2 = 1 To .Count
                If .Item(intCiclo2).GUID <> strArrayRiferimenti(intCiclo1, 0) Then
                    AggiungiRiferimento wbk, strArrayRiferimenti(intCiclo1, 0)
                ElseIf .Item(intCiclo2).GUID = strArrayRiferimenti(intCiclo1, 0) Then
                    Exit For
                End If
            Next intCiclo2
        Next intCiclo1
    End With
    ' Svuota le variabili.
    Set wbk = Nothing
    intCiclo1 = Empty
    intCiclo2 = Empty
    Set objRiferimento = Nothing
    Erase strArrayRiferimenti
    Exit Sub

' Esce dalla Sub.
Uscita: Exit Sub
' Questa riga di uscita viene raggiunta in caso di errore.
GesErr: MsgBox "Errore rimuovendo: " & vbCrLf & "- " & objRiferimento.Description & ". " & vbCrLf & vbCrLf & Err.Description
        ' Svuota le variabili.
        Set wbk = Nothing
        intCiclo1 = Empty
        intCiclo2 = Empty
        Set objRiferimento = Nothing
        Erase strArrayRiferimenti
        Resume Uscita
' Fine della Sub.
End Sub
