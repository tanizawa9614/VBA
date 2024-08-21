Attribute VB_Name = "F_LINEST2"
Option Explicit

Function LINEST2(Y0, X0, Optional N0 As Long = 1)
    Dim y, x, n As Long
    Dim narr(), i As Long
    Dim r As Long, j As Long
    
    y = Y0
    x = X0
    n = N0
    r = UBound(y, 1)
    ReDim narr(1 To r, 1 To n)
    
    For i = 1 To r
        For j = 1 To n
            narr(i, j) = x(i, 1) ^ j
        Next
    Next
    
    Dim ans
    ans = WorksheetFunction.LinEst(y, narr)
    
    LINEST2 = WorksheetFunction.Transpose(ans)
End Function
