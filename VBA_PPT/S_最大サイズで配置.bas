Attribute VB_Name = "S_最大サイズで配置"
Option Explicit

Sub 最大サイズで配置()
    Dim W As Double
    Dim H As Double
    Dim pW As Double
    Dim pH As Double
    With ActivePresentation.PageSetup
        W = .slideWidth
        H = .slideHeight
    End With
    
    Dim lockflg As Boolean
    Dim shp As Shape
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        lockflg = True
        With shp
            If .LockAspectRatio = msoFalse Then lockflg = False
            .LockAspectRatio = msoTrue
            pW = .Width
            pH = .Height
            If pW / pH >= W / H Then
                .Width = W
            Else
                .Height = H
            End If
            .Left = 0.5 * W - 0.5 * .Width
            .Top = 0.5 * H - 0.5 * .Height
            If lockflg = False Then .LockAspectRatio = msoFalse
        End With
    Next
End Sub

Private Function Point2Cm(s As Double)
    Point2Cm = s / 72 * 2.54
End Function

Private Function Cm2Point(s As Double)
    Cm2Point = s * 72 / 2.54
End Function

