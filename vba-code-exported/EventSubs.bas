Attribute VB_Name = "EventSubs"
Public Sub WS_Klasse_Change(Target As Range)

    Dim Bereich As Range
    Dim Zelle As Range
    Dim DB As Worksheet
    Dim HZ As Integer
    
    'Holt Daten anhand der Lizenznummer und formatiert die entsprechende Zeile
    
    
    Set DB = ThisWorkbook.Worksheets("Daten")
    Set Bereich = Intersect(Target.Cells, Target.Parent.Columns(7).Cells)
    If Bereich Is Nothing Then Exit Sub
    
    Application.EnableEvents = False
    
    For Each Zelle In Bereich
        If Zelle.Row > 7 Then
            If Zelle.Value <> "" Then
                HZ = HoleZeile(Zelle.Value, 1, "Daten")
                If HZ > 0 Then
                     Zelle.Parent.Cells(Zelle.Row, 3).Value = DB.Cells(HZ, 2).Value
                    'Rookie
                    If Year(Worksheets("Einstellungen").Range("D5").Value) & "" = "" & DB.Cells(HZ, 5).Value Then
                      With Zelle.Parent.Cells(Zelle.Row, 3).Interior
                        .ColorIndex = 15
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                      End With
                    Else
                      Zelle.Parent.Cells(Zelle.Row, 3).Interior.ColorIndex = xlNone
                    End If
                    'Rest...
                    Zelle.Parent.Cells(Zelle.Row, 4).Value = DB.Cells(HZ, 3).Value & ", " & DB.Cells(HoleZeile(Zelle.Value, 1, "Daten"), 4).Value
                    Zelle.Parent.Cells(Zelle.Row, 5).Value = DB.Cells(HoleZeile(Zelle.Value, 1, "Daten"), 6).Value & " " & DB.Cells(HoleZeile(Zelle.Value, 1, "Daten"), 7).Value
                    Zelle.Parent.Cells(Zelle.Row, 6).Value = DB.Cells(HZ, 8).Value
                    Zelle.Parent.Cells(Zelle.Row, 8).FormulaR1C1 = "=IF(RC23="""",0,RC23)"
                    Zelle.Parent.Cells(Zelle.Row, 8).Font.ColorIndex = 2
                    Zelle.Parent.Cells(Zelle.Row, 12).FormulaR1C1 = "=ROUND(RC[-3]*Einstellungen!R9C4+RC[-2]*Einstellungen!R10C4+RC[-1],2)"
                    Zelle.Parent.Cells(Zelle.Row, 12).Font.ColorIndex = 2
                    Zelle.Parent.Cells(Zelle.Row, 16).FormulaR1C1 = "=ROUND(RC[-3]*Einstellungen!R9C4+RC[-2]*Einstellungen!R10C4+RC[-1],2)"
                    Zelle.Parent.Cells(Zelle.Row, 16).Font.ColorIndex = 2
                    Zelle.Parent.Cells(Zelle.Row, 20).FormulaR1C1 = "=ROUND(RC[-3]*Einstellungen!R9C4+RC[-2]*Einstellungen!R10C4+RC[-1],2)"
                    Zelle.Parent.Cells(Zelle.Row, 20).Font.ColorIndex = 2
                    Zelle.Parent.Cells(Zelle.Row, 21).FormulaR1C1 = "=RC[-5]+RC[-1]"
                    Zelle.Parent.Cells(Zelle.Row, 22).FormulaR1C1 = "=MIN(RC[-6],RC[-2])"
                    Zelle.Parent.Cells(Zelle.Row, 22).Font.ColorIndex = 2
                    Zelle.Parent.Cells(Zelle.Row, 22).EntireColumn.Hidden = True
                    Zelle.Parent.Cells(Zelle.Row, 1).FormulaR1C1 = "=IF(RC[22]<>"""",""niW"",IF(OR(ISTEXT(R[-2]C),R[-1]C[20]<RC[20],AND(R[-1]C[20]=RC[20],R[-1]C[21]<RC[21])),ROW()-7,R[-1]C))"
                    Zelle.Parent.Cells(Zelle.Row, 24).FormulaR1C1 = "=IF(ISTEXT(RC[-23]),0,ROUND((COUNTIF(R8C3:R6000C3,""<>"")-RC[-23])*10/COUNTIF(R8C3:R6000C3,""<>"")+1,2))"
                    Zelle.Parent.Cells(Zelle.Row, 25).FormulaR1C1 = "=IF(ISTEXT(RC[-24]),0,IF(RC[-24]=1,6,IF(RC[-24]>10,0.5,IF(RC[-24]<11,(12-RC[-24])/2))))"
                    With Zelle.Parent.Columns("A:Y").Rows(Zelle.Row).Cells
                        .Borders(xlDiagonalDown).LineStyle = xlNone
                        .Borders(xlDiagonalUp).LineStyle = xlNone
                        .Borders(xlEdgeLeft).LineStyle = xlNone
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).Weight = xlThin
                        .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                        .Borders(xlEdgeRight).LineStyle = xlNone
                        .Borders(xlInsideVertical).LineStyle = xlNone
                    End With
                End If
            End If
            If Zelle.Value = "" Or HZ < 1 Then
                Zelle.Parent.Cells(Zelle.Row, 3).Value = ""
                Zelle.Parent.Cells(Zelle.Row, 4).Value = ""
                Zelle.Parent.Cells(Zelle.Row, 5).Value = ""
                Zelle.Parent.Cells(Zelle.Row, 6).Value = ""
                Zelle.Parent.Cells(Zelle.Row, 12).FormulaR1C1 = ""
                Zelle.Parent.Cells(Zelle.Row, 16).FormulaR1C1 = ""
                Zelle.Parent.Cells(Zelle.Row, 20).FormulaR1C1 = ""
                Zelle.Parent.Cells(Zelle.Row, 21).FormulaR1C1 = ""
                Zelle.Parent.Cells(Zelle.Row, 22).FormulaR1C1 = ""
                Zelle.Parent.Cells(Zelle.Row, 1).FormulaR1C1 = ""
                Zelle.Parent.Cells(Zelle.Row, 24).FormulaR1C1 = ""
                Zelle.Parent.Cells(Zelle.Row, 25).FormulaR1C1 = ""
            End If
        End If
    Next Zelle
    
    Application.EnableEvents = True

End Sub

Sub WS_SelectionChange(Target As Range)

    Dim tr As Integer
    Dim tc As Integer
    Dim tr0 As Integer
    Dim tc0 As Integer
    
    tr = Selection.Row
    tc = Selection.Column
    tr0 = Selection.Row
    tc0 = Selection.Column
    
    If tr > 7 Then
        If tc = 1 Then tc = 2
        If tc > 2 And tc < 7 Then tc = 7
        If tc = 8 Then tc = 9
        If tc = 12 Then tc = 13
        If tc = 16 Then tc = 17
        If tc = 20 Then tc = 23
        If tc = 21 Then tc = 23
        If tc > 23 Then tc = 2: tr = tr + 1
    End If
    
If tr <> tr0 Or tc <> tc0 Then Target.Parent.Cells(tr, tc).Select

End Sub

''' Ergebnisliste berechnene
Sub CB1_Click(WS As Worksheet)
    CB_Ergebnisliste_Click WS
End Sub

Sub CB_Ergebnisliste_Click(WS As Worksheet)
    'Ergebnisliste berechnen
    WS.Range("A8:Y6000").Sort Key1:=WS.Range("H8"), Order1:=xlAscending, Key2:=WS.Range("U8") _
        , Order2:=xlAscending, Key3:=WS.Range("V8"), Order3:=xlAscending, Header:= _
        xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    With WS.Range("B6:B7,G6:G7,K7,I6:K7,M6:O7,Q6:S7,W6:W7,G4").Interior
      .ColorIndex = 0
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
    End With

    'Rookies markieren
    Dim Bereich As Range
    Dim Zelle As Range
    Dim DB As Worksheet
    Dim HZ As Integer
    'Holt Daten anhand der Lizenznummer und formatiert die entsprechende Zeile
    Dim actYear As String
    actYear = "" & Year(Worksheets("Einstellungen").Range("D5").Value)
    
    Set DB = ThisWorkbook.Worksheets("Daten")
    Set Bereich = Intersect(WS.UsedRange.Cells, WS.Columns(7).Cells)
    If Bereich Is Nothing Then Exit Sub
    
    Application.EnableEvents = False
    
    For Each Zelle In Bereich
        If Zelle.Row > 7 Then
            If Zelle.Value <> "" Then
                HZ = HoleZeile(Zelle.Value, 1, "Daten")
                If HZ > 0 Then
                    Zelle.Parent.Cells(Zelle.Row, 3).Value = DB.Cells(HZ, 2).Value
                    'Rookie
                    If actYear = "" & DB.Cells(HZ, 5).Value Then
                      With Zelle.Parent.Cells(Zelle.Row, 3).Interior
                        .ColorIndex = 15
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                      End With
                    Else
                      Zelle.Parent.Cells(Zelle.Row, 3).Interior.ColorIndex = xlNone
                    End If
                End If
            End If
        End If
    Next Zelle
    '''Changed 19b -> 19c
    Application.EnableEvents = True

End Sub

'''Nach Starnummern sortieren
Sub CB2_Click(WS As Worksheet)
  CB_Startliste_Click WS
End Sub
Sub CB_Startliste_Click(WS As Worksheet)
    WS.Range("A8:Y6000").Sort Key1:=WS.Range("B8"), Order1:=xlAscending, Header:= _
           xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    With WS.Range("B6:B7,G6:G7,K7,I6:K7,M6:O7,Q6:S7,W6:W7,G4").Interior
         .ColorIndex = 36
         .Pattern = xlSolid
         .PatternColorIndex = xlAutomatic
    End With
End Sub


Sub Ergebnis_berechnen()

    Select Case ActiveSheet.Name
        Case "Klasse 1" To "Klasse 5"
                Run "CB1_Click", ActiveSheet
        Case "Manschaft"
                Tabelle8.Mannschaftswertung_berechnen
        Case Else
                
    End Select

End Sub

Sub Sortieren_nach_Startnummer()

    Select Case ActiveSheet.Name
        Case "Klasse 1" To "Klasse 5"
                Run "CB2_Click", ActiveSheet
        Case Else
                
    End Select

End Sub

Sub Zeit_Importieren_T()
  Zeit_Importieren 0
End Sub

Sub Zeit_Importieren_1()
  Zeit_Importieren 1
End Sub

Sub Zeit_Importieren_2()
  Zeit_Importieren 2
End Sub

Sub Zeit_Importieren(Optional t As Integer = -1)

  If Left(ActiveSheet.Name, 7) <> "Klasse " Then Exit Sub
  If Selection.Row < 8 Then Exit Sub
  
  Dim sourceFile As String
  sourceFile = ThisWorkbook.Worksheets("Einstellungen").Range("L18").Value
  'Kein sourceFile => kein Import
  If Len(sourceFile) <= 0 Then Exit Sub
  If Len(ThisWorkbook.Worksheets("Einstellungen").Range("L19")) <= 0 Then Exit Sub
  If Len(ThisWorkbook.Worksheets("Einstellungen").Range("L20")) <= 0 Then Exit Sub
  If Len(ThisWorkbook.Worksheets("Einstellungen").Range("L21")) <= 0 Then Exit Sub
  Dim sourceSheet As String
  sourceSheet = ThisWorkbook.Worksheets("Einstellungen").Range("L19")
  Dim rangeCell As String
  rangeCell = ThisWorkbook.Worksheets("Einstellungen").Range("L20") & ThisWorkbook.Worksheets("Einstellungen").Range("L21")

  Dim SBS As Worksheet
  On Error GoTo Fehler
  Set SBS = Workbooks(sourceFile).Worksheets(sourceSheet)
  
  Dim val As Currency  ' Format muss auf zwei Dezimalzeichen gerundet werden
  Dim format As Integer
  
  format = ThisWorkbook.Worksheets("Einstellungen").Range("L22")
  
 ' val = Round(SBS.Range(rangeCell).Value, 2) -> ge�ndert, Grimm    30.05.2010
 ' val = Round(SBS.Range(rangeCell).Value * 86400, 2)  -> ge�ndert, Grimm    30.05.2010
 ' new in 20a
  val = ZeitImportHoleZeit(SBS.Range(rangeCell), format)
  
  Dim ZielR As Range
  
  Select Case t
    Case 0
      Set ZielR = Selection.Parent.Cells(Selection.Row, 11)
    Case 1
      Set ZielR = Selection.Parent.Cells(Selection.Row, 15)
    Case 2
      Set ZielR = Selection.Parent.Cells(Selection.Row, 19)
    Case Else
      If (Selection.Column <= 11) Then
        Set ZielR = Selection.Parent.Cells(Selection.Row, 11)
      ElseIf (Selection.Column <= 15) Then
        Set ZielR = Selection.Parent.Cells(Selection.Row, 15)
      Else
        Set ZielR = Selection.Parent.Cells(Selection.Row, 19)
      End If
  End Select
  
  If ZielR.Value > 0 Then
    If MsgBox(ZielR.Address & " beinhaltet bereits einen Wert: " & ZielR.Value & Chr(13) & "Soll dieser �berschrieben werden?", vbYesNo) <> vbYes Then
        MsgBox "Wert konnte nicht importiert werden!", vbCritical
        Exit Sub
    End If
  End If
  ZielR.NumberFormat = "0.00"
  ZielR.Value = val
  Exit Sub
Fehler:
    If Len(sourceFile) > 0 Then
        MsgBox "Datei " & sourceFile & "ist nicht ge�ffnet!", vbCritical, sourceFile
    End If
End Sub

' Holt eine Wert aus einer Zelle, mit einem von Zwei formaten
' und gibt es als Dezimalzahl zur�ck
' new in 20a
Function ZeitImportHoleZeit(r As Range, format As Integer) As Currency
    
    Select Case format
        ' Dezimal
        Case 0
          ZeitImportHoleZeit = Round(r.Value, 2)
        ' Zeit
        Case 1
          ZeitImportHoleZeit = Round(r.Value * 86400, 2)
    End Select

End Function
