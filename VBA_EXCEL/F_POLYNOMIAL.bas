Attribute VB_Name = "F_POLYNOMIAL"
Option Explicit

Function POLYNOMIAL(X0, Aarr0)
    Dim x, Aarr, ans(), n As Long
    Dim i As Long
    Dim j As Long, jmax As Long
    x = X0
    Aarr = Aarr0
    n = UBound(x, 1)
    jmax = UBound(Aarr) - 1
    
    ReDim ans(1 To n, 1 To 1)
    Dim tmp As Double, cnt As Long
    For i = 1 To n
        tmp = 0
        cnt = 0
        For j = jmax To 0 Step -1
            cnt = cnt + 1
            tmp = tmp + Aarr(cnt, 1) * x(i, 1) ^ j
        Next
        ans(i, 1) = tmp
    Next
    
    POLYNOMIAL = ans
End Function

