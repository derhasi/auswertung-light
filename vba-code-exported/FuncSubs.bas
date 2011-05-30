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

Sub ShapeEntfernen(name As String, WS As Worksheet)
  Dim s As Shape
  For Each s In WS.Shapes
    If s.name = name Then
      s.Delete
    End If
  Next s
End Sub

'''New in 19c
''' New CVS import method
Sub readCSV(ByVal datei As String, WS As Worksheet, Optional delimiter As String = ",", Optional TextDelimiter As String = """")
   Dim strTxt As String
   Dim myarr() As String
   Dim i As Long
   Dim s As String
   Dim lngL As Long
   Open datei For Input As #1
   lngL = 1
   Do Until EOF(1)
      Line Input #1, strTxt
      myarr = Split(strTxt, delimiter)
      For i = LBound(myarr) To UBound(myarr)
        s = myarr(i)
        'remove one left text delimiters
        If (Left(s, Len(TextDelimiter)) = TextDelimiter) Then
          s = Mid(s, 1 + Len(TextDelimiter))
        End If
        If (Right(s, Len(TextDelimiter)) = TextDelimiter) Then
          s = Mid(s, 1, Len(s) - Len(TextDelimiter))
        End If
        myarr(i) = replaceUmlauts(s)
      Next i
      WS.Range(WS.Cells(lngL, 1), WS.Cells(lngL, UBound(myarr) + 1)) = myarr
      lngL = lngL + 1
   Loop
   Close #1
End Sub

'''New in 19c
Function replaceUmlauts(str As String)
  str = Replace(str, "Ã„", "Ä")
  str = Replace(str, "Ã¤", "ä")
  str = Replace(str, "Ã–", "Ö")
  str = Replace(str, "Ã¶", "ö")
  str = Replace(str, "Ãœ", "Ü")
  str = Replace(str, "Ã¼", "ü")
  str = Replace(str, "ÃŸ", "ß")
  str = Replace(str, "Â", """")
  str = Replace(str, "Ã©", "é")
  str = Replace(str, "Ã©", "è")
  str = Replace(str, "Ã¨", "á")
  str = Replace(str, "Ã ", "à")
  replaceUmlauts = str
End Function

