
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                      +
'+ Routine che stampa nella Finestra Immediata          +
'+ la lista dei riferimenti attivi presenti nel file.   +
'+ NON RICHIEDE il riferimento a Microsoft Visual Basic +
'+ For Application Extensibility.                       +
'+                                                      +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub DebugPrintRiferimentiEsistenti()
' Serve per il ciclo tra tutti i riferimenti.
Dim intI As Integer
    ' Consideriamo solo di questo file di Excel tutti i suoi riferimenti attivi.
    With Application.ThisWorkbook.VBProject.References
        ' Ciclo tra i riferimenti.
        For intI = 1 To .Count
            ' Stampa la descrizione.
            Debug.Print "Descrizione: ", .Item(intI).Description
            ' Stampa il relativo nome.
            Debug.Print "Nome: ", .Item(intI).Name
            ' Stampa la GUID.
            Debug.Print "GUID: ", .Item(intI).GUID
            ' Lascia una linea in bianco.
            Debug.Print
        ' Prossimo riferimento.
        Next intI
            ' Se la conta dei riferimenti è 0 allora.
            If .Count = 0 Then
                ' Il successivo messaggio viene stampato nella finestra Immediata.
                Debug.Print "Nel file di Excel:" & Chr(13) & ThisWorkbook.Name & Chr(13) & "non ci sono riferimenti attivi."
            ' Se la conta dei riferimenti è 1 allora.
            ElseIf .Count = 1 Then
                ' Il successivo messaggio viene stampato nella finestra Immediata.
                Debug.Print "Nel file di Excel:" & Chr(13) & ThisWorkbook.Name & Chr(13) & "c'è " & .Count & " riferimento attivo."
            ' Se la conta dei riferimenti è maggiore di 1 allora.
            ElseIf .Count > 1 Then
                ' Il successivo messaggio viene stampato nella finestra Immediata.
                Debug.Print "Nel file di Excel:" & Chr(13) & ThisWorkbook.Name & Chr(13) & "ci sono " & .Count & " riferimenti attivi."
            End If
    End With
' Fine della Subroutine.
End Sub

