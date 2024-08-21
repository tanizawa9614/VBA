Attribute VB_Name = "‰E‘µ‚¦"
Option Explicit

Sub ‰E‘µ‚¦()
   Dim L As Double, W As Double
   Dim M As Long, i As Long
   
   On Error Resume Next
   With ActiveWindow.Selection
      L = .ShapeRange(.ShapeRange.Count).Left
      W = .ShapeRange(.ShapeRange.Count).Width
      M = L + W
      For i = 1 To .ShapeRange.Count - 1
         .ShapeRange(i).Left = M - .ShapeRange(i).Width
      Next
      If .ShapeRange.Count = 1 Then
            .ShapeRange(1).Left = ActivePresentation.PageSetup.SlideWidth - W
      End If
   End With
End Sub

