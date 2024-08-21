Attribute VB_Name = "S_FILEDIRCHANGE_2"
Option Explicit
Sub FILEDIRCHANGE_2()
   On Error Resume Next
   Dim i As Long
   Do While Cells(i + 2, 1) <> ""
      Name Cells(i + 2, "A") As Cells(i + 2, "M")
      i = i + 1
   Loop
End Sub
