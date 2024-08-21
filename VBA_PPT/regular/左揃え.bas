Attribute VB_Name = "ç∂ëµÇ¶"
Option Explicit

Sub ç∂ëµÇ¶()
    Dim L As Double, i As Long
    On Error Resume Next
        With ActiveWindow.Selection
        L = .ShapeRange(.ShapeRange.Count).Left
        For i = 1 To .ShapeRange.Count - 1
            .ShapeRange(i).Left = L
        Next
        If .ShapeRange.Count = 1 Then
            .ShapeRange(1).Left = 0 'ActivePresentation.PageSetup
        End If
    End With
End Sub
