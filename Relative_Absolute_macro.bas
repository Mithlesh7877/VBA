Attribute VB_Name = "Relative_Absolute_macro"
Sub abso()
Attribute abso.VB_ProcData.VB_Invoke_Func = " \n14"
'
' abso Macro
'

'
    Selection.End(xlToLeft).Select
    Selection.End(xlDown).Select
    Range("A7").Select
End Sub
Sub rel()
Attribute rel.VB_ProcData.VB_Invoke_Func = " \n14"
'
' rel Macro


'
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
End Sub
Sub rel_abs()
Attribute rel_abs.VB_ProcData.VB_Invoke_Func = " \n14"
'
' rel_abs Macro


'
    
    
    
    Range("E6").Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
End Sub
