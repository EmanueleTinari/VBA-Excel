VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Questa_cartella_di_lavoro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                                +
'+ Nome :                    Questa_cartella_di_lavoro                                            +
'+                                                                                                +
'+ Creata da :               Emanuele Tinari                                                      +
'+                                                                                                +
'+ In data :                 27/03/2023                                                           +
'+                                                                                                +
'+ Descrizione :             Questa è la cartella principale nel File di Excel "PERSONAL.XLSB".   +
'+                           La Sub Workbook_Open() viene eseguita all'apertura di Excel ed       +
'+                           esegue la chiamata alla Sub ControlloRiferimenti che si occupa di    +
'+                           controllare che tutti i Riferimenti necessari al Progetto VBA        +
'+                           siano presenti.                                                      +
'+                                                                                                +
'+ Uso :                     Importare il file e copiare il codice VBA nel proprio.               +
'+                                                                                                +
'+ Valore restituito:        Nessuno                                                              +
'+                                                                                                +
'+ Esempio :                 Nessuno                                                              +                                                                     +
'+                                                                                                +
'+ Valore di default :       Nessuno                                                              +
'+                                                                                                +
'+ Argomento(i) :            Nessuno                                                              +
'+                                                                                                +
'+ Riferimento(i):           - Sub ControlloRiferimenti                                           +
'+                                                                                                +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub Workbook_Open()
    ControlloRiferimenti
End Sub
