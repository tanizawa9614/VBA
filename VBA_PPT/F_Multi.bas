Attribute VB_Name = "F_Multi"
Option Explicit
Option Base 1
Sub main()
    Dim A(2, 2), B(2), ans, ans2, ans3, con As Long
    
    A(1, 1) = 1
    A(1, 2) = 2
    A(2, 1) = 0
    A(2, 2) = 1
    
    B(1) = 0
'    B(1, 2) = 1
    B(2) = 1
'    B(2, 2) = 0
    
    con = 10
    
'    ans = MMulti(A, B, True)
    ans2 = MMulti(A, B, False)
'    ans3 = MPlus(A, B)
End Sub

Private Function MPlus(ByVal A, ByVal B)
    Dim ans()
    Dim i As Long, j As Long
    A = Convert2dim(A)
    B = Convert2dim(B)
    
    ReDim ans(1 To UBound(A, 1), 1 To UBound(A, 2))
    
    If UBound(A, 1) <> UBound(B, 1) Then
        GoTo ErrHdl2
    ElseIf UBound(A, 1) <> UBound(B, 1) Then
        GoTo ErrHdl2
    End If
    
    For i = 1 To UBound(A, 1)
        For j = 1 To UBound(A, 2)
            ans(i, j) = A(i, j) + B(i, j)
        Next
    Next
    MPlus = ans
ErrHdl2:
End Function

Private Function MMulti(ByVal A, ByVal B, Optional array_flg As Boolean = True)
    Dim ans
    A = Convert2dim(A)
    B = Convert2dim(B)
    If array_flg Then 'îzóÒÇÃóvëfÇ«Ç§ÇµÇÃä|ÇØéZ
        ans = array_multi(A, B)
    Else  'çsóÒÇÃä|ÇØéZ
        ans = matrix_multi(A, B)
    End If
    MMulti = ans
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
Private Function array_multi(ByVal A, ByVal B)
    Dim const_value As Double
    Dim i As Long, j As Long, flg As Boolean
    Dim Ar1, Ar2, ans()
    
    If IsArray(A) = False Then
        const_value = A
        Ar1 = B
        Ar2 = B
        flg = True
    ElseIf IsArray(B) = False Then
        const_value = B
        Ar1 = A
        Ar2 = A
        flg = True
    Else
        Ar1 = A
        Ar2 = B
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

Private Function Convert2dim(ByVal A)
    Dim r As Long, c As Long
    Dim rl As Long, cl As Long
    Dim i As Long, j As Long
    Dim ans(), flg As Boolean
    
    If IsArray(A) = False Then
        Convert2dim = A
        Exit Function
    End If
    
    r = UBound(A) - LBound(A) + 1
    rl = LBound(A)
    On Error Resume Next
    c = UBound(A, 2) - LBound(A, 2) + 1
    cl = LBound(A, 2)
    If Err.Number > 0 Then
        c = 1
        flg = True
    End If
    On Error GoTo 0
    
    ReDim ans(1 To r, 1 To c)
    For i = 1 To r
        For j = 1 To c
            If flg Then
                ans(i, j) = A(rl + i - 1)
            Else
                ans(i, j) = A(rl + i - 1, cl + j - 1)
            End If
        Next
    Next
    
    Convert2dim = ans
End Function
