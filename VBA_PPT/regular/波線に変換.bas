Attribute VB_Name = "”gü‚É•ÏŠ·"
Option Explicit
Option Base 1
Dim nrate As Double, Hrate As Double

Sub ”gü‚É•ÏŠ·()
    
    nrate = 1 ' ”g‚ÌŒÂ”‚ğŒˆ‚ß‚éƒpƒ‰ƒ[ƒ^C‚P‚æ‚è‘å‚Å”g‚ÌŒÂ”‘
    Hrate = 1 ' ”g‚Ì‚‚³‚ğŒˆ‚ß‚éƒpƒ‰ƒ[ƒ^C‚P‚æ‚è‘å‚Å”g‚Ì‚‚³‘
    
    Dim i As Long, j As Long
    Dim T As Double, L As Double
    Dim Sld As Slide, Si As Long, shp As Shape
    Dim Col As Long, LW As Double, W As Double
    Dim startX As Double, startY As Double, tmp As Double
    Dim endX As Double, endY As Double
    
'    On Error GoTo L1
    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.Type = msoLine Then
            Col = shp.Line.ForeColor.RGB
            startX = shp.Left
            startY = shp.Top
            endX = startX + shp.Width
            endY = startY + shp.Height
            If shp.VerticalFlip Then
                tmp = endY
                endY = startY
                startY = tmp
            End If
            LW = shp.Line.Weight
            Call ”gü‚Ìì¬(Col, startX, startY, endX, endY, LW)
            shp.Delete
        End If
    Next
L1:
End Sub


Private Sub ”gü‚Ìì¬(Col As Long, startX As Double, startY As Double, endX As Double, endY As Double, LW As Double)

    Dim n As Long, i As Long
    Dim H As Double
    Dim x As Double, y As Double
    Dim a As Double, b As Double, L_length As Double
    Dim Si As Long, Sld As Slide
    Dim sighn As Long
    Dim vecn, vecu, normn As Double
    Dim diffx As Double, diffy As Double, dist As Double
    
    Si = ActiveWindow.Selection.SlideRange.SlideIndex
    Set Sld = ActivePresentation.Slides(Si)
    
    
    If endX = startX Then endX = endX + 0.0000000001
    
    diffx = endX - startX
    diffy = endY - startY
    dist = Sqr(diffx ^ 2 + diffy ^ 2)
    
    a = diffy / diffx
    b = startY - a * startX
    n = dist * nrate ^ 1.2 / 15 '”g‚ÌŒÂ”‚ğŒˆ’è
    H = dist / n * Hrate ^ 1.5 / 3 '”g‚Ì‚‚³
        
    vecn = Normalize(Array((endY - startY), -1 * (endX - startX)))
    
    sighn = 1
    With Sld.Shapes.BuildFreeform(msoEditingAuto, startX, startY)
        For i = 1 To 2 * n
            x = Abs(startX + (endX - startX) / (2 * n) * (i - 1))
            y = Abs(a * x + b)
            vecu = MPlus(MMultiConst(vecn, sighn * H * 0.5), Array(x, y))
            .AddNodes msoSegmentCurve, msoEditingAuto, vecu(1), vecu(2)
            sighn = sighn * -1
        Next
        .ConvertToShape
    End With
    With Sld.Shapes
        With .Range(.Count)
            .Line.ForeColor.RGB = Col
            .Line.Weight = LW
            .Select
        End With
    End With
End Sub
Private Function MPlus(M1, M2)
    Dim r1 As Long, c1 As Long, ans()
    Dim r2 As Long, c2 As Long
    Dim i As Long, j As Long
    
    r1 = UBound(M1, 1) - LBound(M1, 1) + 1
    r2 = UBound(M2, 1) - LBound(M2, 1) + 1
    If r1 <> r2 Then
        MsgBox "”z—ñ‚Í“¯‚¶ƒTƒCƒY‚Å‚È‚¯‚ê‚Î‚È‚è‚Ü‚¹‚ñ"
        Exit Function
    End If
    
    ReDim ans(1 To r1)
    
    For i = 1 To r1
        ans(i) = M1(i) + M2(i)
    Next
    MPlus = ans
End Function
Private Function MMultiConst(M, g As Double)
    Dim r As Long, ans()
    Dim i As Long
    
    r = UBound(M, 1)
    ReDim ans(1 To r)
    
    For i = 1 To r
        ans(i) = g * M(i)
    Next
    MMultiConst = ans
End Function
Private Function Normalize(vec)
    Dim r As Long, c As Long, ans
    Dim i As Long, norm As Double
    
    r = UBound(vec, 1) - LBound(vec, 1) + 1
      
    norm = 0
    ans = vec
    For i = 1 To UBound(ans)
        norm = norm + ans(i) ^ 2
    Next
    
    norm = norm ^ 0.5
    
    For i = 1 To UBound(ans)
        ans(i) = ans(i) / norm
    Next
    Normalize = ans
End Function
