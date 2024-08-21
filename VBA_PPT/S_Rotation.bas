Attribute VB_Name = "S_Rotation"
Option Explicit

Sub Rotation()
    
    Dim Ar1(1, 1), Ar2(1, 0)
    Dim ans, theta As Double
    
    theta = 45 / 180 * Pi
    Ar1(0, 0) = Cos(theta)
    Ar1(0, 1) = -Sin(theta)
    Ar1(1, 0) = Sin(theta)
    Ar1(1, 1) = Cos(theta)
    Ar2(0, 0) = 1
    Ar2(1, 0) = 0
    ans = Mmult(Ar1, Ar2)
    
End Sub

Function Mmult(Ar1, Ar2)
    Dim newAr1, newAr2
    Dim ans()
    Dim r1 As Long, c1 As Long
    Dim r2 As Long, c2 As Long
    Dim i As Long, j As Long, k As Long
    Dim tmp As Double
    
    newAr1 = Mat1toN(Ar1)
    newAr2 = Mat1toN(Ar2)
    
    r1 = UBound(newAr1, 1)
    r2 = UBound(newAr2, 1)
    c1 = UBound(newAr1, 2)
    c2 = UBound(newAr2, 2)
    
    If c1 <> r2 Then
        GoTo ErrHdl
    End If
    ReDim ans(1 To r1, 1 To c2)
    
    For i = 1 To r1
        For j = 1 To c2
            tmp = 0
            For k = 1 To r2
                tmp = tmp + newAr1(i, k) * newAr2(k, j)
            Next
            ans(i, j) = tmp
        Next
    Next
    Mmult = ans
    Exit Function
ErrHdl:
    MsgBox "行列のサイズが適切ではありません"
End Function

Function Mat1toN(Ar)
    Dim nri As Long, nrf As Long
    Dim nci As Long, ncf As Long
    Dim nr As Long, nc As Long
    Dim i As Long, j As Long
    Dim ans()
    
    nri = LBound(Ar, 1)
    nrf = UBound(Ar, 1)
    nci = LBound(Ar, 2)
    ncf = UBound(Ar, 2)
    
    nr = nrf - nri + 1
    nc = ncf - nci + 1
    
    ReDim ans(1 To nr, 1 To nc)
    
    For i = 1 To nr
        For j = 1 To nc
            ans(i, j) = Ar(LBound(Ar, 1) + i - 1, LBound(Ar, 2) + j - 1)
        Next
    Next
    Mat1toN = ans
End Function

Function Pi()
    Pi = 3.14159265358979
End Function
