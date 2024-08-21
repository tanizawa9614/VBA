Attribute VB_Name = "S_î˜ï™åWêî"
Option Explicit

Sub î˜ï™åWêî()
    Dim h As Double
    Dim x As Double
    Dim iter As Double
    Dim i As Long
    Dim ans() As Double
    Dim f1 As Double
    Dim f2 As Double
    Dim f3 As Double
    Dim f4 As Double
    Dim hL As Long, hU As Long
    
    hL = -5
    hU = 20
    ReDim ans(1 To hU - hL + 1)
    x = -2
    iter = 0.000000000000001
    
    For i = 1 To hU - hL + 1
        h = 10 ^ (-1 * (hL + i - 1))
        f1 = fx(x - 2 * h)
        f2 = fx(x - h)
        f3 = fx(x + h)
        f4 = fx(x + 2 * h)
        ans(i) = (f1 - 8 * f2 + 8 * f3 - f4) / 12 / h
    Next
End Sub

Function fx(x As Double)
    fx = x ^ 2
End Function

