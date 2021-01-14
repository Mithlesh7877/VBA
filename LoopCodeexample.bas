Attribute VB_Name = "LoopCode"
Option Explicit
Sub ifels()
If Range("B2").Value <> "" Then
Range("C2").Value = Range("B2").Value
End If
End Sub

'With_End
Sub withchgfont()
Dim ran As Range
Set ran = Range("A4", "A" & Cells(Rows.Count, 1).End(xlUp).Row)
Debug.Print ran.Address
With ran.Font
    .Name = "Arial"
    .Size = 12
    .Bold = True
    End With
End Sub

Sub withoutchgfont()
Dim ran As Range
Set ran = Range("A4", "A" & Cells(Rows.Count, 1).End(xlUp).Row)
Debug.Print ran.Address
With ran.Font
    .Name = "Calibri"
    .Size = 11
    .Bold = False
    End With
End Sub

Sub foreachc()
Dim sh As Worksheet
For Each sh In ThisWorkbook.Worksheets
sh.Protect
Debug.Print sh.Name
Next sh
End Sub

Sub unforeachc()
Dim sh As Worksheet
For Each sh In ThisWorkbook.Worksheets
sh.Unprotect
Debug.Print sh.Name
Next sh
End Sub

