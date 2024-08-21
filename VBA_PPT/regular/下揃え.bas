Attribute VB_Name = "‰º‘µ‚¦"
Option Explicit

Sub ‰º‘µ‚¦()
   Dim T As Double, H As Double
   Dim M As Long, i As Long
   
   On Error Resume Next
   With ActiveWindow.Selection
      T = .ShapeRange(.ShapeRange.Count).Top
      H = .ShapeRange(.ShapeRange.Count).Height
      M = T + H
      For i = 1 To .ShapeRange.Count - 1
         .ShapeRange(i).Top = M - .ShapeRange(i).Height
      Next
      If .ShapeRange.Count = 1 Then
            .ShapeRange(1).Top = ActivePresentation.PageSetup.SlideHeight - H
      End If
   End With
End Sub
