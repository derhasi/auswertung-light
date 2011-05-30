Attribute VB_Name = "FuncSubs"
Function HoleZeile(Wert As String, Spalte As Integer, Tabellenblatt As String)

Dim TB As Worksheet
Dim Zelle As Range

Set TB = ThisWorkbook.Worksheets(Tabellenblatt)

HoleDaten = Null

For Each Zelle In TB.Columns(Spalte).Rows
    If Zelle.Value = Wert Then
        HoleZeile = Zelle.Row
        Exit Function
    End If
Next Zelle

End Function

Function eTEXT(Wert, eTextformat As String)

    eTEXT = WorksheetFunction.Text(Wert, eTextformat)

End Function


Sub DeleteZeile()
Attribute DeleteZeile.VB_ProcData.VB_Invoke_Func = "l\n14"

WSName = Selection.Parent.name

If WSName = "Klasse 1" Or WSName = "Klasse 2" Or WSName = "Klasse 3" Or WSName = "Klasse 4" Or WSName = "Klasse 5" Then
    If Selection.Row > 7 Then
        Selection.EntireRow.Delete
    End If
End If

End Sub

Sub EnableEvents()

Application.EnableEvents = True

End Sub

Function Komma2Point(Wert) As String
Wert = "" & Wert
Komma2Point = Replace(Wert, ",", ".")
End Function

Sub ZPOutput_Raus()
Dim name As String
    Sheets("zp_output").Select
    Sheets("zp_output").Copy
    name = InputBox("Dateiname")
    ActiveWorkbook.SaveAs Filename:=CurDir() & "/" & name & ".csv", FileFormat:=xlCSV, _
        CreateBackup:=False
    ActiveWorkbook.Close
        
End Sub


Sub RangeReplace(Bereich As Range, FromT As String, ToT As String)
    Bereich.Replace What:=FromT, Replacement:=ToT, LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub


