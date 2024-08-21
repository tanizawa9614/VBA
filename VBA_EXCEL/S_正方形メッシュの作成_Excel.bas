Attribute VB_Name = "S_正方形メッシュの作成_Excel"
Option Explicit

Sub 正方形メッシュの作成_Excel()
    
    Dim nr As Long, nc As Long
    Dim i As Long, j As Long
    Dim T As Double, L As Double, cnt As Long
    Dim W As Long, Line_Weight As Double
    Dim shp As Shape
        
    nr = 20
    nc = 20
    T = 20
    L = 20
    W = 10
    Line_Weight = 0.1
    
    Set shp = ActiveSheet.Shapes.AddShape(msoShapeRectangle, T, L, 10, 10)
    With shp.Line
        .ForeColor.RGB = RGB(0, 0, 0)
        .Weight = Line_Weight
    End With
    shp.Fill.Visible = msoFalse
    
    For i = 1 To nr
        For j = 1 To nc
            If i = 1 And j = 1 Then
            Else
                With shp.Duplicate
                    .Top = T + W * (i - 1)
                    .Left = L + W * (j - 1)
                End With
            End If
        Next
    Next

    
End Sub
