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
    MsgBox "Auswertung Light für Excel - Version " & versionNr() & Chr(10) & "von Johannes Haseitl - derhasi.de - www.zugspitzpokal.de", vbInformation, "Information"
End Sub
Sub Hilfe()
    Tabelle10.Activate
End Sub

Sub Shortcuts()
    MsgBox "Strg + L : Markierte Zeilen Löschen" & Chr(10) _
    & "Strg + T : Zeit importieren " & Chr(10) _
    & "Strg + 0 : Zeit in Training importieren" & Chr(10) _
    & "Strg + 1 : Zeit in Wertung 1 importieren" & Chr(10) _
    & "Strg + 2 : Zeit in Wertung 2 importieren", vbInformation, "Tastenkürzel"
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


'''New in 19d
Function versionNr() As String
  versionNr = "0.20a"
End Function

'''New in 20
Sub exportUrkunden(Klasse As String)

    Dim ES As Worksheet
    Dim WS As Worksheet
    
    Set ES = ThisWorkbook.Worksheets("Einstellungen")
    
    For Each WS In ThisWorkbook.Worksheets
        If Klasse = "" Or WS.name = Klasse Then
        
        End If
    Next WS


End Sub

