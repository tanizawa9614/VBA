Attribute VB_Name = "ã‰º’†‰›‘µ‚¦"
Option Explicit

Sub ã‰º’†‰›‘µ‚¦()
   Dim T As Double, H As Double
   Dim M As Long, i As Long
   
   On Error Resume Next
   With ActiveWindow.Selection
      T = .ShapeRange(.ShapeRange.Count).Top
      H = .ShapeRange(.ShapeRange.Count).Height
      M = T + H * 0.5
      For i = 1 To .ShapeRange.Count - 1
         .ShapeRange(i).Top = M - .ShapeRange(i).Height * 0.5
      Next
      If .ShapeRange.Count = 1 Then
            .ShapeRange(1).Top = ActivePresentation.PageSetup.SlideHeight * 0.5 - H * 0.5
      End If
   End With
End Sub
