Attribute VB_Name = "è„ëµÇ¶"
Option Explicit

Sub è„ëµÇ¶()
   Dim T As Double, i As Long
   On Error Resume Next
   With ActiveWindow.Selection
      T = .ShapeRange(.ShapeRange.Count).Top
      For i = 1 To .ShapeRange.Count - 1
         .ShapeRange(i).Top = T
      Next
      If .ShapeRange.Count = 1 Then
            .ShapeRange(1).Top = 0
      End If
   End With
End Sub
