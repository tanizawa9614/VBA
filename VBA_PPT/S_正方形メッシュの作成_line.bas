Attribute VB_Name = "S_正方形メッシュの作成_line"
Option Explicit
Dim Sld As Slide
Dim idx()
Dim nr As Long, nc As Long
Dim T As Double, L As Double
Dim W As Long, Line_Weight As Double
      

Sub 正方形メッシュの作成_PPT_line()
    
    nr = 10 '縦のメッシュ数
    nc = nr * 16 / 9 '横のメッシュ数
    T = 0    '位置
    L = 0   '左位置
    W = 10 ' メッシュサイズ
    Line_Weight = 0.1
        
    Dim Clength As Double, Rlength As Double
    Clength = W * nc
    Rlength = W * nr
    Dim loopR As Long, loopC As Long
    loopR = Int(Log(nr) / Log(2))
    loopC = Int(Log(nc) / Log(2))
    Dim excessR As Long, excessC As Long
    excessR = nr - POWER(2, loopR)
    excessC = nc - POWER(2, loopC)
    Dim Si As Long
    Si = ActiveWindow.Selection.SlideRange.SlideIndex
    Set Sld = ActivePresentation.Slides(Si)
    
    Call ConnectorAdd(L, T, L + Clength, T)
    Call LineDuplicate(loopR, excessR, 1)
    Call ConnectorAdd(L, T, L, T + Rlength)
    Call LineDuplicate(loopC, excessC, 2)
    
    Call GetIdx(Sld, nc + nr + 2, idx)
    Sld.Shapes.Range(idx).Group.Select
End Sub
Private Sub LineDuplicate(loopcnt, excess, rc)
    
    Dim i As Long, j As Long
    Dim getC As Long, totC As Long
    Dim rflg As Long, cflg As Long
    
    If rc = 1 Then
        rflg = 1
        cflg = 0
    Else
        rflg = 0
        cflg = 1
    End If
        
    totC = 1
        
    For j = 1 To loopcnt
        getC = POWER(2, j - 1)
        Call GetIdx(Sld, getC, idx)
        With Sld.Shapes.Range(idx).Duplicate
            If getC <> 1 Then .Group
            .Top = T + W * totC * rflg
            .Left = L + W * totC * cflg
            If getC <> 1 Then .Ungroup
        End With
        totC = totC + getC
    Next
    
    getC = excess + 1
    Call GetIdx(Sld, getC, idx)
    With Sld.Shapes.Range(idx).Duplicate
        .Group
        .Top = T + W * totC * rflg
        .Left = L + W * totC * cflg
        .Ungroup
    End With
    
End Sub

Private Sub ConnectorAdd(startX, startY, endX, endY)
    With Sld.Shapes.AddLine(startX, startY, endX, endY)
        With .Line
            .ForeColor.RGB = RGB(0, 0, 0)
            .Weight = Line_Weight
        End With
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
