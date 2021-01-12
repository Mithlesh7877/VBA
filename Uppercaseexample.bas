Attribute VB_Name = "Uppercaseexample"
Option Explicit

Sub built()
'Upper case conversion example
With Sheet3
.Range("B2").Value = VBA.Date
.Range("B3").Value = VBA.UCase(Range("A3").Value)
End With
End Sub

