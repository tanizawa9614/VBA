Attribute VB_Name = "Module1"
Option Explicit
Sub �񕝑�()
   Selection.ColumnWidth = Round(ActiveCell.ColumnWidth, 0) + 3
End Sub
Sub �񕝌�()
   Selection.ColumnWidth = Round(ActiveCell.ColumnWidth, 0) - 3
End Sub
Sub �s����()
   Selection.RowHeight = Round(ActiveCell.RowHeight, 0) + 5
End Sub
Sub �s����()
   Selection.RowHeight = Round(ActiveCell.RowHeight, 0) - 5
End Sub
