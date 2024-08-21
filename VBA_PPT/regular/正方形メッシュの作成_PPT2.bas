Attribute VB_Name = "�����`���b�V���̍쐬_PPT2"
Option Explicit

Sub �����`���b�V���̍쐬_PPT()
    
    Dim nr As Long, nc As Long
    Dim i As Long, j As Long
    Dim T As Double, L As Double, cnt As Long
    Dim W As Long, Line_Weight As Double
    Dim shp As Shape, idx()
    Dim loopR As Long, loopC As Long
    Dim excessR As Long, excessC As Long
    Dim getC As Long, getR As Long, totC As Long, totR As Long
        
    nr = 10 '�c�̃��b�V����
    nc = 10 '���̃��b�V����
    T = 15    '��ʒu
    L = 20   '���ʒu
    W = 40 ' ���b�V���T�C�Y
    Line_Weight = 0.1
    
    loopR = Int(Log(nr) / Log(2))
    loopC = Int(Log(nc) / Log(2))
    excessR = nr - POWER(2, loopR)
    excessC = nc - POWER(2, loopC)
    
    Dim Sld As Slide, Si As Long
    Si = ActiveWindow.Selection.SlideRange.SlideIndex
    Set Sld = ActivePresentation.Slides(Si)
    
    With Sld.Shapes.AddShape(msoShapeRectangle, L, T, W, W)
        With .Line
            .ForeColor.RGB = RGB(0, 0, 0)
            .Weight = Line_Weight
        End With
        .Fill.Visible = msoFalse
        totC = 1
        
        For j = 1 To loopC
            getC = POWER(2, j - 1)
            Call GetIdx(Sld, getC, idx)
            With Sld.Shapes.Range(idx).Duplicate
                If getC <> 1 Then .Group
                .Top = T
                .Left = L + W * totC
                If getC <> 1 Then .Ungroup
            End With
            totC = totC + getC
        Next
        
        getC = excessC
        Call GetIdx(Sld, getC, idx)
        With Sld.Shapes.Range(idx).Duplicate
            .Group
            .Top = T
            .Left = L + W * totC
            .Ungroup
        End With
        totC = totC + getC
                      
        totR = 1
        For i = 1 To loopR
            getR = POWER(2, i - 1) * totC
            Call GetIdx(Sld, getR, idx)
            With Sld.Shapes.Range(idx).Duplicate
                .Group
                .Top = T + W * totR
                .Left = L
                .Ungroup
            End With
            totR = totR + POWER(2, i - 1)
        Next
        
        getR = excessR * totC
        Call GetIdx(Sld, getR, idx)
        With Sld.Shapes.Range(idx).Duplicate
            .Group
            .Top = T + W * totR
            .Left = L
            .Ungroup
        End With
  
        
        Call GetIdx(Sld, nc * nr, idx)
        Sld.Shapes.Range(idx).Group.Select
    End With
End Sub
Private Sub GetIdx(Sld, getC, idx)
    Dim i As Long
    ReDim idx(1 To getC)
    For i = 0 To getC - 1
        idx(i + 1) = Sld.Shapes.Count - i
    Next
End Sub

Private Function POWER(a, n) As Double
    Dim i As Long, ans As Double
    ans = 1
    For i = 1 To n
        ans = ans * a
    Next
    POWER = ans
End Function
