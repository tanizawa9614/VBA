Attribute VB_Name = "F_polyxpoly"
Option Explicit

Function polyxpoly(line0, poly0) As Variant
    Dim intersections() As Variant, line, poly
    Dim i As Integer, j As Integer, k As Integer
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim x3 As Double, y3 As Double, x4 As Double, y4 As Double
    Dim denominator As Double, numerator1 As Double, numerator2 As Double
    Dim t As Double, cnt As Long
    Dim s As Double
    
    line = line0
    poly = poly0
    ReDim intersections(1 To UBound(line, 1) - 1, 1 To UBound(poly, 1))
    
    For i = 1 To UBound(line, 1) - 1
        x1 = line(i, 1)
        y1 = line(i, 2)
        x2 = line(i + 1, 1)
        y2 = line(i + 1, 2)
        
        For j = 1 To UBound(poly, 1)
            x3 = poly(j, 1)
            y3 = poly(j, 2)
            x4 = poly((j Mod UBound(poly, 1)) + 1, 1)
            y4 = poly((j Mod UBound(poly, 1)) + 1, 2)
            
            denominator = (y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1)
            
            If denominator <> 0 Then
                numerator1 = (x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)
                numerator2 = (x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)
                
                t = numerator1 / denominator
                s = numerator2 / denominator

                If t >= 0 And t <= 1 Then
                    If s >= 0 And s <= 1 Then
                        intersections(i, j) = Array(x1 + t * (x2 - x1), y1 + t * (y2 - y1))
                        cnt = cnt + 1
                    End If
                End If
            End If
        Next j
    Next i
    
    Dim ans() As Variant
    ReDim ans(1 To cnt, 1 To 2)
    cnt = 0
    
    For i = 1 To UBound(intersections, 1)
        For j = 1 To UBound(intersections, 2)
            If Not IsEmpty(intersections(i, j)) Then
                cnt = cnt + 1
                ans(cnt, 1) = intersections(i, j)(0)
                ans(cnt, 2) = intersections(i, j)(1)
            End If
        Next j
    Next i
    
    polyxpoly = ans
End Function

