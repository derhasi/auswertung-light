Attribute VB_Name = "Modul3"
Sub Makro2()
Attribute Makro2.VB_Description = "Makro am 30.04.2007 von Chargen aufgezeichnet"
Attribute Makro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro2 Makro
' Makro am 30.04.2007 von Chargen aufgezeichnet
'

'
    ActiveCell.FormulaR1C1 = _
        "=Einstellungen!R[2]C[3]&"" ""&Einstellungen!R[3]C[6]&"" "" &Einstellungen!R[3]C[3]& "" am ""&TEXT(Einstellungen!R[4]C[3],""TT.MM.JJJJ"")"
    Range("A2:K2").Select
    
End Sub
