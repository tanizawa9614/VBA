Attribute VB_Name = "Module1"
Option Explicit

Sub �񕝑�()
   On Error Resume Next
   Selection.ColumnWidth = ActiveCell.ColumnWidth + 3
End Sub
Sub �񕝌�()
   On Error Resume Next
   Selection.ColumnWidth = ActiveCell.ColumnWidth - 3
End Sub
Sub �s����()
   On Error Resume Next
   Selection.RowHeight = ActiveCell.RowHeight + 7
End Sub
Sub �s����()
   On Error Resume Next
   Selection.RowHeight = ActiveCell.RowHeight - 7
End Sub
