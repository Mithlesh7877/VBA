Attribute VB_Name = "Module1"
Sub test()
Attribute test.VB_Description = "test macro"
Attribute test.VB_ProcData.VB_Invoke_Func = "i\n14"
'
' test Macro
' test macro
'
' Keyboard Shortcut: Ctrl+i
'
    Sheets.Add After:=ActiveSheet
    ActiveCell.FormulaR1C1 = "Hello"
    Range("A1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    ActiveWindow.DisplayGridlines = False
End Sub


Sub test1()
Attribute test1.VB_Description = "testing macro for filling n"
Attribute test1.VB_ProcData.VB_Invoke_Func = "I\n14"
'
' test1 Macro
' testing macro for filling n
'
' Keyboard Shortcut: Ctrl+Shift+I
'
    Selection.CurrentRegion.Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "n"
End Sub
