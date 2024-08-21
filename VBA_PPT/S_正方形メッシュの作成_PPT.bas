Attribute VB_Name = "S_正方形メッシュの作成_PPT"
Option Explicit

Sub 正方形メッシュの作成_PPT()
    
    Dim nr As Long, nc As Long
    Dim i As Long, j As Long
    Dim T As Double, L As Double, cnt As Long
    Dim W As Long, Line_Weight As Double
    Dim shp As Shape, idx(), idx2()
        
    nr = 50 '縦のメッシュ数
    nc = 50 '横のメッシュ数
    T = 15    '上位置
    L = 20   '左位置
    W = 3 ' メッシュサイズ
    Line_Weight = 0.1
    
    ReDim idx(1 To nr * nc), idx2(1 To nc)
    
    Dim Sld As Slide, Si As Long
    Si = ActiveWindow.Selection.SlideRange.SlideIndex
    Set Sld = ActivePresentation.Slides(Si)
    
    
    With Sld.Shapes.AddShape(msoShapeRectangle, L, T, W, W)
        With .Line
            .ForeColor.RGB = RGB(0, 0, 0)
            .Weight = Line_Weight
        End With
        .Select
        .Fill.Visible = msoFalse
        
        For j = 2 To nc
            With .Duplicate
                .Top = T
                .Left = L + W * (j - 1)
            End With
        Next
        
        For i = 0 To nc - 1
            idx2(i + 1) = Sld.Shapes.Count - i
        Next
        
        For i = 2 To nr
            With Sld.Shapes.Range(idx2).Duplicate
                If nc <> 1 Then .Group
                .Top = T + W * (i - 1)
                .Left = L
                If nc <> 1 Then .Ungroup
            End With
            For j = 0 To nc - 1
                idx2(j + 1) = Sld.Shapes.Count - j
            Next
        Next
        
        For i = 0 To nc * nr - 1
            idx(i + 1) = Sld.Shapes.Count - i
        Next
        If nc = 1 And nr = 1 Then Exit Sub
        Sld.Shapes.Range(idx).Group.Select
    End With
End Sub

