Sub Macro1()
'
' Macro1 Macro
'
' Keyboard Shortcut: Ctrl+q
'
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=LEFT(RIGHT(CELL(""filename""),LEN(CELL(""filename""))-54),LEN(RIGHT(CELL(""filename""),LEN(CELL(""filename""))-54))-5)"
    Range("A1").Select
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Dim MyFileName As String
        MyFileName = "C:\Users\Jingyu\Desktop\ÊýÁªÃúÆ·\½ðÈÚ±³¾°\relationsExcel - ¸±±¾/" & [a1] & ".csv"
    ActiveWorkbook.SaveAs Filename:=MyFileName, _
        FileFormat:=xlCSV
	
End Sub