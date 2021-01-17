Attribute VB_Name = "Ranges"
Option Explicit

Sub ReferCell()
Range("A1").Value = "1st" 'Cells(1, 2) = "1"
Range("A2:C2").Value = "2nd"
Range("A3,C3").Value = "3rd"
Range("A5:C5,E5:F5") = "4th"
Range("A4", "E4").Value = "5th"
Range("A1").Offset(7, 2) = "6th"

End Sub

Sub WorksheetRef()
'Sheets(2).Select
'Sheets("Test").Select
'Sheets("Test").Range("A1").Value = "Hi"
'ThisWorkbook.Save
'Sheets.Select
'Cells.Copy
'Cells.PasteSpecial xlPasteValues
'ThisWorkbook.SaveAs Filename:=ThisWorkbook.Path & "HC" & ThisWorkbook.Name
End Sub
