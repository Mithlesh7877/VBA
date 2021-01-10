Attribute VB_Name = "comment_delete_example"
Option Explicit

Sub del_comment()
'
' del_comment Macro
'

'
    Selection.SpecialCells(xlCellTypeComments).Select
    Selection.ClearComments
End Sub

