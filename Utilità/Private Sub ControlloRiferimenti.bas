
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

Private Sub ControlloRiferimenti()

' Gestione errore.
On Error GoTo GesErr

Dim strArrayRiferimenti(14, 2) As String        ' Array dei riferimenti. Variare il primo numero se i riferimenti variano.
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
'* diminuiscono i blocchi di codice che riempiono l'array tridimensionale *
'**************************************************************************
    
    ' Riferimento 0.
    strArrayRiferimenti(0, 0) = "Visual Basic For Applications"
    strArrayRiferimenti(0, 1) = "{000204EF-0000-0000-C000-000000000046}"
    strArrayRiferimenti(0, 2) = "VBA"
    ' Riferimento 1.
    strArrayRiferimenti(1, 0) = "Microsoft Excel 16.0 Object Library"
    strArrayRiferimenti(1, 1) = "{00020813-0000-0000-C000-000000000046}"
    strArrayRiferimenti(1, 2) = "Excel"
    ' Riferimento 2.
    strArrayRiferimenti(2, 0) = "OLE Automation"
    strArrayRiferimenti(2, 1) = "{00020430-0000-0000-C000-000000000046}"
    strArrayRiferimenti(2, 2) = "stdole"
    ' Riferimento 3.
    strArrayRiferimenti(3, 0) = "Microsoft Office 16.0 Object Library"
    strArrayRiferimenti(3, 1) = "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}"
    strArrayRiferimenti(3, 2) = "Office"
    ' Riferimento 4.
    strArrayRiferimenti(4, 0) = "Microsoft Forms 2.0 Object Library"
    strArrayRiferimenti(4, 1) = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}"
    strArrayRiferimenti(4, 2) = "MSForms"
    ' Riferimento 5.
    strArrayRiferimenti(5, 0) = "Microsoft XML, v6.0"
    strArrayRiferimenti(5, 1) = "{F5078F18-C551-11D3-89B9-0000F81FE221}"
    strArrayRiferimenti(5, 2) = "MSXML2"
    ' Riferimento 6.
    strArrayRiferimenti(6, 0) = "Microsoft HTML Object Library"
    strArrayRiferimenti(6, 1) = "{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}"
    strArrayRiferimenti(6, 2) = "MSHTML"
    ' Riferimento 7.
    strArrayRiferimenti(7, 0) = "Microsoft Internet Controls"
    strArrayRiferimenti(7, 1) = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}"
    strArrayRiferimenti(7, 2) = "SHDocVw"
    ' Riferimento 8.
    strArrayRiferimenti(8, 0) = "Microsoft Scripting Runtime"
    strArrayRiferimenti(8, 1) = "{420B2830-E718-11CF-893D-00A0C9054228}"
    strArrayRiferimenti(8, 2) = "Scripting"
    ' Riferimento 9.
    strArrayRiferimenti(9, 0) = "Microsoft Script Control 1.0"
    strArrayRiferimenti(9, 1) = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}"
    strArrayRiferimenti(9, 2) = "MSScriptControl"
    ' Riferimento 10.
    strArrayRiferimenti(10, 0) = "Microsoft ActiveX Data Objects 6.1 Library"
    strArrayRiferimenti(10, 1) = "{B691E011-1797-432E-907A-4D8C69339129}"
    strArrayRiferimenti(10, 2) = "ADODB"
    ' Riferimento 11.
    strArrayRiferimenti(11, 0) = "Microsoft Windows Common Controls 6.0 (SP6)"
    strArrayRiferimenti(11, 1) = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}"
    strArrayRiferimenti(11, 2) = "MSComctlLib"
    ' Riferimento 12.
    strArrayRiferimenti(12, 0) = "Selenium Type Library"
    strArrayRiferimenti(12, 1) = "{0277FC34-FD1B-4616-BB19-A9AABCAF2A70}"
    strArrayRiferimenti(12, 2) = "Selenium"
    ' Riferimento 13.
    strArrayRiferimenti(13, 0) = "Microsoft Visual Basic for Applications Extensibility 5.3"
    strArrayRiferimenti(13, 1) = "{0002E157-0000-0000-C000-000000000046}"
    strArrayRiferimenti(13, 2) = "VBIDE"
    ' Riferimento 14.
    strArrayRiferimenti(14, 0) = "Microsoft WMI Scripting V1.2 Library"
    strArrayRiferimenti(14, 1) = "{565783C6-CB41-11D1-8B02-00600806D9B6}"
    strArrayRiferimenti(14, 2) = "WbemScripting"
    
    ' Imposta il riferimento a questo Workbook.
    Set wbk = ThisWorkbook
    With Application.ThisWorkbook.VBProject.References
        For intCiclo1 = 0 To 13
            For intCiclo2 = 1 To .Count
                If .Item(intCiclo2).Description <> strArrayRiferimenti(intCiclo1, 0) Then
                    AggiungiRiferimento wbk, strArrayRiferimenti(intCiclo1, 1), strArrayRiferimenti(intCiclo1, 2)
                ElseIf .Item(intCiclo2).Description = strArrayRiferimenti(intCiclo1, 0) Then
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
