Attribute VB_Name = "F_FINDVALUE"
Option Explicit

Function FINDVALUE(•¶Žš—ñ As Variant)

    Dim str As Variant
    Dim slen As Long, i As Long
    Dim varr As String
    Dim ans, tmp As String, tmp2 As String
    Dim sidx As Long, fidx As Long
    
    str = •¶Žš—ñ
    
    slen = Len(str)
    sidx = 1
    fidx = 1
    For i = 1 To slen
        tmp = Mid(str, sidx, fidx - sidx + 1)
        If IsNumeric(tmp) = False Then
            tmp2 = Mid(str, sidx, fidx - sidx)
            sidx = i + 1
            fidx = sidx
            If Right(varr, 1) <> vbTab Then
                varr = varr & vbTab
            End If
            varr = varr & tmp2
        Else
            fidx = fidx + 1
            If i = slen Then
                If Right(varr, 1) <> vbTab Then
                    varr = varr & vbTab
                End If
                varr = varr & tmp
            End If
        End If
    Next
    If Left(varr, 1) = vbTab Then varr = Mid(varr, 2)
    If Right(varr, 1) = vbTab Then varr = Left(varr, Len(varr) - 1)
    
    ans = Split(varr, vbTab)
    
    Dim ans2()
    ReDim ans2(UBound(ans))
    For i = 0 To UBound(ans)
        ans2(i) = Val(ans(i))
    Next
    FINDVALUE = ans2
End Function
