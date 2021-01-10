Attribute VB_Name = "Var"
Option Explicit ' used to get popup if any var is
'not defined
Sub DefineVar()
    Dim lastrow As Long
    
    lastrow = Rows.Count
    Debug.Print lastrow
    
    Dim mymonth(1 To 12) As String
    mymonth(1) = "Jan"
    Debug.Print mymonth
    
    Const myname As String = "Mithlesh"
    Debug.Print myname
 
End Sub

Sub createworkb()
 'Dim newbook As Workbook
 'Set newbook = Workbooks.Add
 Dim newsheet As Worksheet
 'Set newsheet = Worksheets.Add
 'newsheet.Delete
 'Sheet4.delete
 
End Sub
