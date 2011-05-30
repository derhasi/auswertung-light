Attribute VB_Name = "MenuSubs"
Sub Klasse1()
Tabelle2.Activate
End Sub
Sub Klasse2()
Tabelle3.Activate
End Sub
Sub Klasse3()
Tabelle4.Activate
End Sub
Sub Klasse4()
Tabelle5.Activate
End Sub
Sub Klasse5()
Tabelle6.Activate
End Sub
Sub Mannschaft()
Tabelle8.Activate
End Sub
Sub Daten()
Tabelle7.Activate
End Sub
Sub Info()
    MsgBox "Auswertung Light für Excel - Version 0.17" & Chr(10) & "von Johannes Haseitl - www.derhasi.de - www.zugspitzpokal.de", vbInformation, "Information"
End Sub
Sub Hilfe()
    Tabelle10.Activate
End Sub

Sub Einstellungen()
    Tabelle1.Activate
End Sub
Sub ZPOutput()
    Tabelle9.Activate
End Sub

Sub ZPOutput_Erstellen_Speichern()
    Tabelle9.ZP_Output_Erstellen
    Tabelle9.ZP_Output_Speichern
End Sub

Sub Nix()
End Sub

Sub Save()
ThisWorkbook.Save
End Sub

