Attribute VB_Name = "Inbuiltfunc"
Option Explicit

Sub counttopmsgbox()
'program to count and output top 3 number in selected range
Dim myr As Range
Dim one As Long, two As Long, three As Long
On Error GoTo leave
Set myr = Application.InputBox("Enter num", "Count", , , , , , 8)
If Application.WorksheetFunction.Count(myr) > 2 Then
one = Excel.WorksheetFunction.Max(myr)
two = Excel.WorksheetFunction.Large(myr, 2)
three = Excel.WorksheetFunction.Large(myr, 3)
MsgBox one & vbNewLine & two & vbNewLine & three
Else
MsgBox ("Please select a valid range")
End If
leave:
End Sub
