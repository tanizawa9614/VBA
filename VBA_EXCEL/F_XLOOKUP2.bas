Attribute VB_Name = "F_XLOOKUP2"
Option Explicit
Option Base 1
Function XLOOKUP2(ŒŸõ’l, ŒŸõ”ÍˆÍ, –ß‚è”ÍˆÍ)
    Dim findV, findA, retA, buf
    Dim i As Long, j As Long, k As Long, l As Long
    Dim flg As Boolean, A()
    Dim findstr
    
    Application.Volatile
    findV = MAKE2DARRAY(ŒŸõ’l)
    findA = MAKE2DARRAY(ŒŸõ”ÍˆÍ)
    retA = MAKE2DARRAY(–ß‚è”ÍˆÍ)
    
    flg = False
    
    If UBound(findA, 1) = UBound(retA, 1) Then flg = True
    If UBound(findA, 2) <> 1 Then flg = False
    If flg = False Then Exit Function
    
    ReDim A(1 To UBound(findV), 1 To UBound(retA, 2) * UBound(findV, 2))
    
    For l = 1 To UBound(findV, 2)
        For i = 1 To UBound(findV)
            findstr = findV(i, l)
            For j = 1 To UBound(findA, 1)
                If findstr = findA(j, 1) Then
                    For k = 1 To UBound(retA, 2)
                        A(i, k + UBound(retA, 2) * (l - 1)) = retA(j, k)
                    Next
                    Exit For
                End If
            Next
        Next
    Next
    XLOOKUP2 = A
End Function
Private Function MAKE2DARRAY(A0)
    Dim i As Long, j As Long
    Dim r As Long, c As Long
    Dim cnt1 As Long, cnt2 As Long
    Dim myA, A, flg As Boolean
    
    A = A0
    flg = True
    If Not IsArray(A) Then
        ReDim myA(1 To 1, 1 To 1)
        myA(1, 1) = A
        GoTo L1
    End If
    r = UBound(A)
    On Error Resume Next
    c = UBound(A, 2)
    If Err.Number > 0 Then
        flg = False
        c = 1
    End If
    On Error GoTo 0
    
    ReDim myA(1 To r, 1 To c)
    cnt1 = 1
    For i = LBound(A) To UBound(A)
        cnt2 = 1
        If flg Then
            For j = LBound(A, 2) To UBound(A, 2)
                myA(cnt1, cnt2) = A(i, j)
                cnt2 = cnt2 + 1
            Next
        Else
            myA(cnt1, cnt2) = A(i)
        End If
        cnt1 = cnt1 + 1
    Next
L1:
    MAKE2DARRAY = myA
End Function



