Attribute VB_Name = "�㑵��"
Option Explicit

Sub �㑵��()
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
