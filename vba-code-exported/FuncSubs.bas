Attribute VB_Name = "FuncSubs"
Function HoleZeile(Wert As String, spalte As Integer, Tabellenblatt As String)

  Dim TB As Worksheet
  Dim Zelle As Range
  
  Set TB = ThisWorkbook.Worksheets(Tabellenblatt)
  
  HoleDaten = Null
  
  For Each Zelle In TB.Columns(spalte).Rows
    If Zelle.Value = Wert Then
      HoleZeile = Zelle.Row
      Exit Function
    End If
  Next Zelle

End Function

Function HoleZeileWB(Wert As String, spalte As Integer, Tabellenblatt As String, Optional Arbeitsmappe As String = "")

  Dim TB As Worksheet
  Dim Zelle As Range
  
  If Arbeitsmappe = "" Then
    Set TB = ThisWorkbook.Worksheets(Tabellenblatt)
  Else
    Set TB = Workbooks(Arbeitsmappe).Worksheets(Tabellenblatt)
  End If
  
  HoleDaten = Null
  
  For Each Zelle In TB.Columns(spalte).Rows
    If Zelle.Value = Wert Then
      HoleZeileWB = Zelle.Row
      Exit Function
    End If
  Next Zelle

End Function

Function eTEXT(Wert, eTextformat As String)

  eTEXT = WorksheetFunction.Text(Wert, eTextformat)

End Function


Sub DeleteZeile()
Attribute DeleteZeile.VB_ProcData.VB_Invoke_Func = "l\n14"

  WSName = Selection.Parent.Name
    
  If WSName = "Klasse 1" Or WSName = "Klasse 2" Or WSName = "Klasse 3" Or WSName = "Klasse 4" Or WSName = "Klasse 5" Then
    If Selection.Row > 7 Then
      Selection.EntireRow.Delete
    End If
  End If

End Sub

Sub EnableEvents()
  Application.EnableEvents = True
End Sub
Sub DisableEvents()
  Application.EnableEvents = False
End Sub

Function Komma2Point(Wert) As String
  Wert = "" & Wert
  Komma2Point = Replace(Wert, ",", ".")
End Function

Sub ZPOutput_Raus()
  Dim Name As String
  Sheets("zp_output").Select
  Sheets("zp_output").Copy
  Name = InputBox("Dateiname")
  ActiveWorkbook.SaveAs Filename:=CurDir() & "/" & Name & ".csv", FileFormat:=xlCSV, _
    CreateBackup:=False
  ActiveWorkbook.Close
        
End Sub


Sub RangeReplace(Bereich As Range, FromT As String, ToT As String)
  Bereich.Replace What:=FromT, Replacement:=ToT, LookAt:=xlPart, SearchOrder _
    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub

Sub ShapeEntfernen(Name As String, WS As Worksheet)
  Dim S As Shape
  For Each S In WS.Shapes
    If S.Name = Name Then
      S.Delete
    End If
  Next S
End Sub

'''New in 19c
''' New CVS import method
Sub readCSV(ByVal datei As String, WS As Worksheet, Optional delimiter As String = ",", Optional TextDelimiter As String = """")
  Dim strTxt As String
  Dim myarr() As String
  Dim i As Long
  Dim S As String
  Dim lngL As Long
  Open datei For Input As #1
  lngL = 1
  Do Until EOF(1)
     Line Input #1, strTxt
     myarr = Split(strTxt, delimiter)
     For i = LBound(myarr) To UBound(myarr)
       S = myarr(i)
       'remove one left text delimiters
       If (Left(S, Len(TextDelimiter)) = TextDelimiter) Then
         S = Mid(S, 1 + Len(TextDelimiter))
       End If
       If (Right(S, Len(TextDelimiter)) = TextDelimiter) Then
         S = Mid(S, 1, Len(S) - Len(TextDelimiter))
       End If
       myarr(i) = replaceUmlauts(S)
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
