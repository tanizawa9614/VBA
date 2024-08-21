Attribute VB_Name = "S_calcpi"
Option Explicit

Sub calcpi()
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
End Sub

Function calcatan(x As Double, n As Long) As Double
    Dim i As Long, ans As Double
    ans = 0
    For i = 1 To n
        ans = ans + (-1) ^ (i - 1) / (2 * i - 1) * x ^ (2 * i - 1)
    Next
    calcatan = ans
End Function
