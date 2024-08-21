Attribute VB_Name = "F_TEXTSPLIT2"
Option Explicit
Function TEXTSPLIT2(ï∂éöóÒ, _
               ãÊêÿÇËï∂éö, _
               Optional ãÛÇÕçÌèú As Boolean = True, _
               Optional ë„ë÷ï∂éö As String = "")

    Dim A, B(), i As Long, j As Long
    Dim c As String, n As Long
    Dim l As Long, myA(), buf As Double
    Dim mymax As Long, STRARRAY, DELIMITER
    
    Application.Volatile
    STRARRAY = ï∂éöóÒ
    DELIMITER = ãÊêÿÇËï∂éö
    STRARRAY = MAKE2DARRAY(STRARRAY)
    DELIMITER = MAKE2DARRAY(DELIMITER)
    mymax = 1
    ReDim myA(1 To UBound(STRARRAY), 1 To 1)
    
    For l = 1 To UBound(STRARRAY, 1)
        c = STRARRAY(l, 1)
        For i = 1 To UBound(DELIMITER)
            c = Replace(c, DELIMITER(i, 1), vbTab)
        Next
        A = Split(c, vbTab)
        ReDim B(UBound(A))
        n = 0
        For i = 0 To UBound(A)
            On Error Resume Next
            buf = A(i)
            If Err.Number > 0 Then
                B(i) = A(i)
            Else
                B(i) = buf
            End If
            On Error GoTo 0
        Next
        A = B
        If ãÛÇÕçÌèú Then
            For i = 0 To UBound(A)
                If A(i) <> "" Then
                    B(n) = A(i)
                    n = n + 1
                End If
            Next
            ReDim Preserve B(n - 1)
            A = B
        End If
        
        If mymax < UBound(A) + 1 Then
            mymax = UBound(A) + 1
            ReDim Preserve myA(1 To UBound(STRARRAY, 1), 1 To mymax)
        End If
        
        For i = 0 To UBound(A)
           myA(l, i + 1) = A(i)
        Next
    Next
    
    For i = 1 To UBound(myA, 1)
        For j = 1 To UBound(myA, 2)
            If IsEmpty(myA(i, j)) Then myA(i, j) = ë„ë÷ï∂éö
        Next
    Next
    
    TEXTSPLIT2 = myA
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


