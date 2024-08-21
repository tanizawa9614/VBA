Attribute VB_Name = "S_îgê¸Ç…ïœä∑"
Option Explicit

Sub îgê¸Ç…ïœä∑()
    Dim shp As Shape
    Dim n As Double
    Dim L As Double
    Dim dx As Double
    Dim dy As Double
    Dim H As Double
    Dim xy() As Double
    Dim shpW As Double
    Dim shpH As Double
    Dim Pi As Double
    Dim i As Long
    Dim rot(1 To 2, 1 To 2) As Double
    Dim theta As Double
    Dim NewXY
    Dim Si As Long, Sld As Slide
    Si = ActiveWindow.Selection.SlideRange.SlideIndex
    Set Sld = ActivePresentation.Slides(Si)
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim tmp As Double
    Dim nsub As Long
    Dim nL As Double
    
    nsub = 4 '1é¸ä˙ìñÇΩÇËÇÃì_êî
    H = 3 ' îgÇÃçÇÇ≥
    Pi = Atn(1) * 4  'Å@â~é¸ó¶
    nL = 10 '1é¸ä˙ìñÇΩÇËÇÃîgÇÃïù
       
    On Error GoTo L1
    For Each shp In ActiveWindow.Selection.ShapeRange
        
        If shp.Type = msoLine Then
            x1 = shp.Left
            x2 = x1 + shp.Width
            y1 = shp.Top
            y2 = y1 + shp.Height
            If (shp.VerticalFlip And shp.HorizontalFlip) = False And (shp.VerticalFlip Or shp.HorizontalFlip) Then
                tmp = y1
                y1 = y2
                y2 = tmp
            End If
            shpW = shp.Width
            shpH = shp.Height
            L = Sqr(shpW ^ 2 + shpH ^ 2)
            n = Int(L / nL) + 0.5 ' îgÇÃêî
            ReDim xy(1 To 2, 1 To Int(nsub * n) + 1)
            For i = 1 To UBound(xy, 2)
                xy(1, i) = L / (nsub * n) * (i - 1)
                xy(2, i) = -H * Sin(2 * n * Pi / L * xy(1, i))
            Next
            theta = -Atn((y2 - y1) / (x2 - x1))
            rot(1, 1) = Cos(theta)
            rot(1, 2) = Sin(theta)
            rot(2, 1) = -Sin(theta)
            rot(2, 2) = Cos(theta)
            NewXY = MMULT(rot, xy)
            With Sld.Shapes.BuildFreeform(msoEditingAuto, x1, y1)
                For i = 1 To UBound(xy, 2)
                    .AddNodes msoSegmentCurve, msoEditingAuto, x1 + NewXY(1, i), y1 + NewXY(2, i)
                Next
                .ConvertToShape
            End With
            With Sld.Shapes
                With .Range(.Count)
                    .Line.ForeColor.RGB = shp.Line.ForeColor.RGB
                    .Line.Weight = shp.Line.Weight
                    .Select
'                    If shp.VerticalFlip Then
'                        .Top = shp.Top + shp.Height
'                    End If
                End With
            End With
            shp.Delete
        End If
    Next
L1:
End Sub

Private Function MMULT(ByVal A1, ByVal A2)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim n As Long
    Dim M As Long
    Dim ans() As Double
    Dim tmp As Double
    
    n = UBound(A1, 1)
    M = UBound(A2, 2)
    ReDim ans(1 To n, 1 To M)
    
    For i = 1 To n
        For j = 1 To M
            tmp = 0
            For k = 1 To UBound(A1, 2)
                tmp = tmp + A1(i, k) * A2(k, j)
            Next
            ans(i, j) = tmp
        Next
    Next
    MMULT = ans
End Function
