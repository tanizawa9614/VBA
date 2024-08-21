Attribute VB_Name = "S_ê≥näpå`ÇÃçÏê¨2"
Option Explicit
Option Base 1
Dim n As Long, dtheta As Double

Sub oneto1000()
    For n = 3 To 20 Step 5
         Call ê≥näpå`ÇÃçÏê¨
    Next
End Sub

Sub ê≥näpå`ÇÃçÏê¨()
    
    Dim Ar1(2, 2)
    Dim ans
    Dim i As Long, x0 As Double, y0 As Double
    Dim r As Double, dist As Double, tmp
    
'    n = 9
    dtheta = 360 / n / 180 * Pi
    r = 100
    x0 = r
    y0 = r
    ReDim tmp(1 To n, 1 To 2)
    
    Dim Sld As Slide, Si As Long
    Si = ActiveWindow.Selection.SlideRange.SlideIndex
    Set Sld = ActivePresentation.Slides(Si)
    
    If n <= 2 Then
        MsgBox "n ÇÕ 2 à»è„ÇÃêîílÇë„ì¸ÇµÇƒÇ≠ÇæÇ≥Ç¢"
    End If
    
    Ar1(1, 1) = Cos(dtheta)
    Ar1(1, 2) = -Sin(dtheta)
    Ar1(2, 1) = Sin(dtheta)
    Ar1(2, 2) = Cos(dtheta)
    ans = Array(0, -r)
    
    With Sld.Shapes.BuildFreeform(msoEditingAuto, ans(1) + x0, ans(2) + y0)
        For i = 1 To n
            ans = MMulti(Ar1, ans, False)
            .AddNodes msoSegmentLine, msoEditingAuto, ans(1) + x0, ans(2) + y0
            tmp(i, 1) = ans(1)
            tmp(i, 2) = ans(2)
        Next
        .ConvertToShape.Select
    End With
    dist = Sqr((tmp(1, 1) - tmp(2, 1)) ^ 2 + (tmp(1, 2) - tmp(2, 2)) ^ 2)
'    MsgBox dist
End Sub
Function Pi() As Double
    Dim ans0 As Double, ans As Double
    Dim n As Long, iter As Double, cnt As Long
    
    iter = 0.000000000000001
    n = 1
    ans0 = 0
    Do
        ans = 16 * calcatan(0.2, n) - 4 * calcatan(1 / 239, n)
        If Abs(ans - ans0) < iter Then Exit Do
        n = n + 1
        ans0 = ans
    Loop
    Pi = ans
End Function

Function calcatan(x As Double, n As Long) As Double
    Dim i As Long, ans As Double
    ans = 0
    For i = 1 To n
        ans = ans + (-1) ^ (i - 1) / (2 * i - 1) * x ^ (2 * i - 1)
    Next
    calcatan = ans
End Function


Private Function MPlus(ByVal a, ByVal b)
    Dim ans()
    Dim i As Long, j As Long
    a = Convert2dim(a)
    b = Convert2dim(b)
    
    ReDim ans(1 To UBound(a, 1), 1 To UBound(a, 2))
    
    If UBound(a, 1) <> UBound(b, 1) Then
        GoTo ErrHdl2
    ElseIf UBound(a, 1) <> UBound(b, 1) Then
        GoTo ErrHdl2
    End If
    
    For i = 1 To UBound(a, 1)
        For j = 1 To UBound(a, 2)
            ans(i, j) = a(i, j) + b(i, j)
        Next
    Next
    MPlus = ans
ErrHdl2:
End Function

Private Function MMulti(ByVal a, ByVal b, Optional array_flg As Boolean = True)
    Dim ans
    a = Convert2dim(a)
    b = Convert2dim(b)
    If array_flg Then 'îzóÒÇÃóvëfÇ«Ç§ÇµÇÃä|ÇØéZ
        ans = array_multi(a, b)
    Else  'çsóÒÇÃä|ÇØéZ
        ans = matrix_multi(a, b)
    End If
    MMulti = ConvertVec(ans)
End Function
Private Function ConvertVec(ByVal vec)
    Dim ans
    Dim i As Long
    Dim r As Long
    Dim c As Long
    r = UBound(vec, 1)
    c = UBound(vec, 2)
    
    If r = 1 Then
        ReDim ans(1 To c)
        For i = 1 To c
            ans(i) = vec(1, i)
        Next
    ElseIf c = 1 Then
        ReDim ans(1 To r)
        For i = 1 To r
            ans(i) = vec(i, 1)
        Next
    End If
    ConvertVec = ans
End Function

Private Function matrix_multi(ByVal M1, ByVal M2)
    Dim i As Long, j As Long, k As Long
    Dim r1 As Long, c1 As Long
    Dim r2 As Long, c2 As Long
    Dim Ar1, Ar2, ans(), tmp As Double
    
    r1 = UBound(M1)
    c1 = UBound(M1, 2)
    r2 = UBound(M2)
    c2 = UBound(M2, 2)
    
    If c1 <> r2 Then Exit Function
    
    ReDim ans(1 To r1, 1 To c2)
    
    For i = 1 To r1
        For j = 1 To c2
            tmp = 0
            For k = 1 To c1
                tmp = tmp + M1(i, k) * M2(k, j)
            Next
            ans(i, j) = tmp
        Next
    Next
    matrix_multi = ans
End Function
Private Function array_multi(ByVal a, ByVal b)
    Dim const_value As Double
    Dim i As Long, j As Long, flg As Boolean
    Dim Ar1, Ar2, ans()
    
    If IsArray(a) = False Then
        const_value = a
        Ar1 = b
        Ar2 = b
        flg = True
    ElseIf IsArray(b) = False Then
        const_value = b
        Ar1 = a
        Ar2 = a
        flg = True
    Else
        Ar1 = a
        Ar2 = b
        flg = False
    End If
    
    If flg Then
        For i = 1 To UBound(Ar2, 1)
            For j = 1 To UBound(Ar2, 2)
                Ar2(i, j) = const_value
            Next
        Next
    End If
    
    ReDim ans(1 To UBound(Ar2, 1), 1 To UBound(Ar2, 2))
    For i = 1 To UBound(Ar2, 1)
        For j = 1 To UBound(Ar2, 2)
            ans(i, j) = Ar1(i, j) * Ar2(i, j)
        Next
    Next
    array_multi = ans
End Function

Private Function Convert2dim(ByVal a)
    Dim r As Long, c As Long
    Dim rl As Long, cl As Long
    Dim i As Long, j As Long
    Dim ans(), flg As Boolean
    
    If IsArray(a) = False Then
        Convert2dim = a
        Exit Function
    End If
    
    r = UBound(a) - LBound(a) + 1
    rl = LBound(a)
    On Error Resume Next
    c = UBound(a, 2) - LBound(a, 2) + 1
    cl = LBound(a, 2)
    If Err.Number > 0 Then
        c = 1
        flg = True
    End If
    On Error GoTo 0
    
    ReDim ans(1 To r, 1 To c)
    For i = 1 To r
        For j = 1 To c
            If flg Then
                ans(i, j) = a(rl + i - 1)
            Else
                ans(i, j) = a(rl + i - 1, cl + j - 1)
            End If
        Next
    Next
    
    Convert2dim = ans
End Function

