Attribute VB_Name = "setdowhileloop"
Option Explicit

Sub test()
   Dim i As Long
   Dim A As Range
   Set A = Cells(i + 1, 1)
   Do While A <> ""
      A.Offset(, 1) = A & "dede"
      Set A = A.Offset(1, 0)
      i = i + 1
   Loop
End Sub
